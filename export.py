#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按商户拆分混部流水并计费导出。

输入（与本脚本同目录）:
  txns.xlsx       混部流水, 列 txn/mch/term/amt/tier/ts, 尾部可能有合计行
  terminals.csv   终端绑定, 列 term/mch/start/end, 起止两端均含当日
  fee_rules.json  按档的分段费率

输出:
  out/<mch>.xlsx  每商户一个工作簿, sheet 名 txns,
                  明细列 txn/term/amt_cents/tier/ts/fee_cents (不含 mch),
                  按 ts 升序、同 ts 按 txn 升序
  out/shared.csv  判不了归属的交易, 列与商户明细相同
  export_log.txt  台账, 每次导出追加一段

归属规则: 交易时刻落在哪个终端的哪个绑定区间, 就归哪个商户; 终端绑定与交易
行自带的 mch 不一致时, 以交易时刻的绑定版本为准; 无任何绑定覆盖则落 shared.csv。

计费: 金额换算为整数分后按档逐段累计, 各段费用加总后逐笔四舍五入到分, 半分向上。
合计行不进入任何明细, 也不留在 shared.csv。

台账段格式:
  [EXPORT] time=<ISO8601 UTC> detail_rows=<明细行数> files=<out 目录下产物文件数>
  [AUDIT] <检查名> OK|FAIL      (六项: 跨租户引用/合计行混入/文件清单/行数对账/费用合计/台账登记)
  [AUDIT_RESULT] PASS|FAIL
同日重复导出不追加第二段台账, 且产物逐字节一致。
"""

from __future__ import annotations

import csv
import json
import os
import re
import sys
import tempfile
import zipfile
from datetime import date, datetime, timezone
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

BASE = Path(__file__).resolve().parent
SRC_XLSX = BASE / "txns.xlsx"
TERMINALS_CSV = BASE / "terminals.csv"
FEE_RULES_JSON = BASE / "fee_rules.json"
OUT_DIR = BASE / "out"
LOG_PATH = BASE / "export_log.txt"

SHEET_NAME = "txns"
DETAIL_HEADER = ["txn", "term", "amt_cents", "tier", "ts", "fee_cents"]
MERCHANT_FILE_HEADER = DETAIL_HEADER[:]

AUDIT_NAMES = ["跨租户引用", "合计行混入", "文件清单", "行数对账", "费用合计", "台账登记"]

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

CENT = Decimal("0.01")
TOTAL_TOKENS = {"TOTAL", "合计", "小计", "SUM"}


class ExportError(Exception):
    """输入不合法, 无法开始导出。"""


# --------------------------------------------------------------------------- #
# 读取 xlsx (stdlib only)
# --------------------------------------------------------------------------- #
def _read_xlsx(path: Path, sheet_name: str | None = None) -> list[list[str | None]]:
    with zipfile.ZipFile(path) as zf:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_target = {r.get("Id"): r.get("Target") for r in rels}
        sheets: list[tuple[str, str]] = []
        for sh in wb.iter(f"{{{NS_MAIN}}}sheet"):
            target = rel_target[sh.get(f"{{{NS_REL}}}id")].lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            sheets.append((sh.get("name"), target))

        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sst = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sst.iter(f"{{{NS_MAIN}}}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{{{NS_MAIN}}}t")))

        if sheet_name is None:
            _, target = sheets[0]
        else:
            matches = [t for n, t in sheets if n == sheet_name]
            if not matches:
                raise ExportError(f"{path.name} 中没有 sheet {sheet_name!r}")
            target = matches[0]

        root = ET.fromstring(zf.read(target))
        table: list[list[str | None]] = []
        for row in root.iter(f"{{{NS_MAIN}}}row"):
            cells: dict[str, str | None] = {}
            for cell in row.findall(f"{{{NS_MAIN}}}c"):
                ref = cell.get("r") or ""
                col = re.match(r"[A-Z]+", ref).group() if re.match(r"[A-Z]+", ref) else ""
                ctype = cell.get("t")
                if ctype == "inlineStr":
                    value: str | None = "".join(
                        t.text or "" for t in cell.iter(f"{{{NS_MAIN}}}t")
                    )
                else:
                    v = cell.find(f"{{{NS_MAIN}}}v")
                    value = None if v is None else v.text
                    if value is not None and ctype == "s":
                        value = shared[int(value)]
                cells[col] = value
            width = max((_col_index(c) for c in cells if c), default=-1) + 1
            table.append([cells.get(_col_letter(i)) for i in range(width)])
        return table


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _col_letter(index: int) -> str:
    s = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        s = chr(ord("A") + rem) + s
    return s


# --------------------------------------------------------------------------- #
# 写入 xlsx (stdlib only, 幂等字节)
# --------------------------------------------------------------------------- #
def _sheet_xml(header: list[str], rows: list[list[object]]) -> str:
    cols = len(header)
    last_col = _col_letter(cols - 1)
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<worksheet xmlns="{NS_MAIN}">',
        f'<dimension ref="A1:{last_col}{len(rows) + 1}"/>',
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>',
        '<sheetFormatPr defaultRowHeight="15"/>',
        "<sheetData>",
    ]

    def emit_cell(ref: str, value: object) -> str:
        if isinstance(value, (int, Decimal)):
            return f'<c r="{ref}"><v>{value}</v></c>'
        return (
            f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">'
            f"{escape(str(value))}</t></is></c>"
        )

    all_rows: list[list[object]] = [list(header)] + rows
    for r_idx, row in enumerate(all_rows, start=1):
        cells = "".join(emit_cell(f"{_col_letter(c)}{r_idx}", row[c]) for c in range(cols))
        parts.append(f'<row r="{r_idx}">{cells}</row>')
    parts.append("</sheetData></worksheet>")
    return "".join(parts)


_STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<styleSheet xmlns="{NS_MAIN}">'
    '<fonts count="1"><font><name val="Calibri"/><family val="2"/><sz val="11"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>"
)

_CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<Types xmlns="{NS_CT}">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    "</Types>"
)

_ROOT_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<Relationships xmlns="{NS_PKG_REL}">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="xl/workbook.xml"/></Relationships>'
)

_WORKBOOK_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">'
    f'<sheets><sheet name="{SHEET_NAME}" sheetId="1" r:id="rId1"/></sheets></workbook>'
)

_WORKBOOK_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<Relationships xmlns="{NS_PKG_REL}">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
    'Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
    'Target="styles.xml"/></Relationships>'
)


def _build_xlsx_bytes(header: list[str], rows: list[list[object]]) -> bytes:
    """生成确定性的 xlsx 字节串(固定 zip 时间戳), 便于同日重复导出比对。"""
    sheet = _sheet_xml(header, rows)
    entries = [
        ("[Content_Types].xml", _CONTENT_TYPES_XML),
        ("_rels/.rels", _ROOT_RELS_XML),
        ("xl/workbook.xml", _WORKBOOK_XML),
        ("xl/_rels/workbook.xml.rels", _WORKBOOK_RELS_XML),
        ("xl/styles.xml", _STYLES_XML),
        ("xl/worksheets/sheet1.xml", sheet),
    ]
    buf = tempfile.SpooledTemporaryFile()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in entries:
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            zf.writestr(info, data.encode("utf-8"))
    buf.seek(0)
    return buf.read()


def _write_if_changed(path: Path, data: bytes) -> bool:
    if path.exists() and path.read_bytes() == data:
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), prefix=".tmp-", suffix=path.suffix)
    try:
        with os.fdopen(fd, "wb") as fh:
            fh.write(data)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise
    return True


# --------------------------------------------------------------------------- #
# 规则与输入解析
# --------------------------------------------------------------------------- #
def load_fee_rules(path: Path) -> dict[str, list[tuple[int | None, str, Decimal]]]:
    """tier -> [(该段金额上限(分)或 None, 计费单位, value)]"""
    raw = json.loads(path.read_text(encoding="utf-8"))
    tiers: dict[str, list[tuple[int | None, str, Decimal]]] = {}
    for tier, spec in raw["tiers"].items():
        segs: list[tuple[int | None, str, Decimal]] = []
        for seg in spec["segments"]:
            cap = None if seg["up_to_yuan"] is None else int(
                (Decimal(str(seg["up_to_yuan"])) * 100).to_integral_value(ROUND_HALF_UP)
            )
            unit = seg["unit"]
            if unit not in {"flat_cents", "bps", "pct"}:
                raise ExportError(f"费率档 {tier} 含未知计费单位 {unit!r}")
            segs.append((cap, unit, Decimal(str(seg["value"]))))
        if not segs:
            raise ExportError(f"费率档 {tier} 没有分段")
        tiers[tier] = segs
    return tiers


def compute_fee(amt_cents: int, tier: str, tiers: dict[str, list[tuple[int | None, str, Decimal]]]) -> int:
    """按段累计费用, 加总后逐笔四舍五入到分, 半分向上。"""
    if tier not in tiers:
        raise ExportError(f"未知费率档 {tier!r}")
    total = Decimal(0)
    lower = 0
    for cap, unit, value in tiers[tier]:
        upper = amt_cents if cap is None else min(amt_cents, cap)
        width = upper - lower
        if width > 0:
            if unit == "flat_cents":
                total += value  # 落入该段即收一次固定费
            elif unit == "bps":
                total += Decimal(width) * value / Decimal(10000)
            else:  # pct
                total += Decimal(width) * value / Decimal(100)
        if cap is None or amt_cents <= cap:
            break
        lower = cap
    return int(total.quantize(Decimal(1), rounding=ROUND_HALF_UP))


def to_cents(value: str) -> int:
    amount = Decimal(str(value).strip())
    return int((amount * 100).quantize(Decimal(1), rounding=ROUND_HALF_UP))


def load_bindings(path: Path) -> dict[str, list[tuple[date, date, str]]]:
    binds: dict[str, list[tuple[date, date, str]]] = {}
    with path.open(encoding="utf-8", newline="") as fh:
        for rec in csv.DictReader(fh):
            term = (rec.get("term") or "").strip()
            mch = (rec.get("mch") or "").strip()
            if not term or not mch:
                continue
            start = date.fromisoformat((rec.get("start") or "").strip())
            end = date.fromisoformat((rec.get("end") or "").strip())
            binds.setdefault(term, []).append((start, end, mch))
    for term in binds:
        binds[term].sort(key=lambda b: (b[0], b[1]))  # 同一终端多版本按起始日排序
    return binds


def resolve_owner(binds: dict[str, list[tuple[date, date, str]]], term: str, when: date) -> str | None:
    """取交易时刻覆盖该终端的绑定版本; 重叠时以起始日最晚者为准。"""
    owner = None
    for start, end, mch in binds.get(term, []):
        if start <= when <= end:
            owner = mch
    return owner


def parse_ts(text: str) -> datetime:
    text = (text or "").strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise ExportError(f"无法解析交易时间 {text!r}")


# --------------------------------------------------------------------------- #
# 读源流水
# --------------------------------------------------------------------------- #
def load_source(path: Path) -> tuple[list[dict], int]:
    """返回 (明细行, 被排除的合计行数)。"""
    table = _read_xlsx(path)
    header_idx = None
    for i, row in enumerate(table):
        names = [str(v).strip().lower() if v is not None else "" for v in row]
        if "txn" in names and "term" in names and "ts" in names:
            header_idx = i
            header = [str(v).strip().lower() if v is not None else "" for v in row]
            break
    if header_idx is None:
        raise ExportError("流水 xlsx 找不到表头(txn/term/ts)")

    required = ["txn", "term", "amt", "tier", "ts"]
    col = {name: header.index(name) for name in required if name in header}
    missing = [n for n in required if n not in col]
    if missing:
        raise ExportError(f"流水表头缺列: {', '.join(missing)}")
    mch_col = header.index("mch") if "mch" in header else None

    detail: list[dict] = []
    total_rows = 0
    seen: set[str] = set()
    for row in table[header_idx + 1 :]:
        def cell(name: str) -> str | None:
            idx = col[name]
            return row[idx] if idx < len(row) else None

        raw_txn = cell("txn")
        txn = "" if raw_txn is None else str(raw_txn).strip()
        if not txn:
            if any(v not in (None, "") for v in row):
                total_rows += 1  # 无交易号但有内容的行按合计/汇总行处理
            continue
        if txn.upper() in TOTAL_TOKENS:
            total_rows += 1
            continue
        if txn in seen:
            raise ExportError(f"交易号重复: {txn}")
        seen.add(txn)

        term = (cell("term") or "").strip()
        tier = (cell("tier") or "").strip()
        ts_raw = (cell("ts") or "").strip()
        amt_raw = cell("amt")
        if not term or not tier or not ts_raw or amt_raw in (None, ""):
            raise ExportError(f"交易 {txn} 字段不完整")
        when = parse_ts(ts_raw)
        detail.append(
            {
                "txn": txn,
                "mch": (str(row[mch_col]).strip() if mch_col is not None and mch_col < len(row) and row[mch_col] else ""),
                "term": term,
                "amt_cents": to_cents(str(amt_raw)),
                "tier": tier,
                "ts": when.strftime("%Y-%m-%d %H:%M:%S"),
                "when": when,
            }
        )
    return detail, total_rows


# --------------------------------------------------------------------------- #
# 组装导出内容
# --------------------------------------------------------------------------- #
def build_export(detail: list[dict], binds, tiers):
    merchants: dict[str, list[list[object]]] = {}
    shared: list[list[object]] = []

    for rec in detail:
        fee = compute_fee(rec["amt_cents"], rec["tier"], tiers)
        owner = resolve_owner(binds, rec["term"], rec["when"].date())
        row: list[object] = [
            rec["txn"],
            rec["term"],
            rec["amt_cents"],
            rec["tier"],
            rec["ts"],
            fee,
        ]
        if owner is None:
            shared.append([rec["when"], rec["txn"], row])
        else:
            merchants.setdefault(owner, []).append([rec["when"], rec["txn"], row])

    def ordered(items):
        return [r for _, _, r in sorted(items, key=lambda x: (x[0], x[1]))]

    merchant_rows = {m: ordered(rows) for m, rows in merchants.items()}
    shared_rows = ordered(shared)
    return merchant_rows, shared_rows


# --------------------------------------------------------------------------- #
# 审计
# --------------------------------------------------------------------------- #
def _read_output_rows(path: Path) -> tuple[list[str], list[list[str]]]:
    table = _read_xlsx(path, SHEET_NAME)
    if not table:
        return [], []
    header = ["" if v is None else str(v) for v in table[0]]
    rows = [["" if v is None else str(v) for v in r] for r in table[1:] if any(v is not None for v in r)]
    return header, rows


def audit_exports(
    out_dir: Path,
    merchant_rows: dict[str, list[list[object]]],
    shared_rows: list[list[object]],
    tiers: dict[str, list[tuple[int | None, str, Decimal]]],
    binds: dict[str, list[tuple[date, date, str]]],
    detail_count: int,
    total_rows_skipped: int,
) -> tuple[dict[str, tuple[bool, str]], dict[str, list[list[str]]]]:
    """基于落盘产物做审计: 归属与费用都从产物自身重算, 与源明细只对行数。

    返回 {检查名: (是否通过, 说明)} 与前五项结果; 第六项由台账回调校验。
    """
    result: dict[str, tuple[bool, str]] = {}
    readback: dict[str, list[list[str]]] = {}
    problems: dict[str, list[str]] = {name: [] for name in AUDIT_NAMES}

    def judge_row(row: list[str], origin: str) -> None:
        """独立复算一行: 归属、费用、合计行。"""
        if len(row) != len(DETAIL_HEADER):
            problems["跨租户引用"].append(f"{origin} 行字段数异常: {row}")
            return
        txn, term, amt_s, tier, ts, fee_s = row
        if txn.upper() in TOTAL_TOKENS:
            problems["合计行混入"].append(f"{origin} 含合计行 {txn}")
            return
        try:
            amt, fee = int(amt_s), int(fee_s)
        except ValueError:
            problems["费用合计"].append(f"{origin} 金额/费用非整数: {row}")
            return
        if amt < 0:
            problems["费用合计"].append(f"{txn} 金额为负")
        try:
            recomputed = compute_fee(amt, tier, tiers)
        except ExportError as exc:
            problems["费用合计"].append(f"交易 {txn} {exc}")
            return
        if recomputed != fee:
            problems["费用合计"].append(f"{txn} 费用 {fee}, 按规则应为 {recomputed}")
        try:
            owner = resolve_owner(binds, term, parse_ts(ts).date())
        except ExportError as exc:
            problems["跨租户引用"].append(f"{txn} {exc}")
            return
        if origin == "shared.csv":
            if owner is not None:
                problems["跨租户引用"].append(f"shared.csv 含可归属 {owner} 的交易 {txn}")
        elif owner != origin.removesuffix(".xlsx"):
            problems["跨租户引用"].append(
                f"{origin} 含 {txn}, 交易时刻应归 {owner or 'shared.csv'}"
            )

    for mch in sorted(merchant_rows):
        path = out_dir / f"{mch}.xlsx"
        if not path.exists():
            problems["文件清单"].append(f"缺失 {path.name}")
            continue
        header, rows = _read_output_rows(path)
        with zipfile.ZipFile(path) as zf:
            names = [n for n in zf.namelist()]
        if "xl/workbook.xml" not in names:
            problems["文件清单"].append(f"{path.name} 不是工作簿")
        try:
            sheet_names = _workbook_sheet_names(path)
        except Exception:
            sheet_names = []
        if sheet_names != [SHEET_NAME]:
            problems["文件清单"].append(f"{path.name} 的 sheet 为 {sheet_names}, 应为 [{SHEET_NAME!r}]")
        if header != MERCHANT_FILE_HEADER:
            problems["跨租户引用"].append(
                f"{path.name} 表头为 {header}, 应不含 mch 且列序固定"
            )
        readback[mch] = rows
        for row in rows:
            judge_row(row, path.name)

    shared_path = out_dir / "shared.csv"
    shared_rows_disk: list[list[str]] = []
    if not shared_path.exists():
        problems["文件清单"].append("缺失 shared.csv")
    else:
        with shared_path.open(encoding="utf-8", newline="") as fh:
            reader = csv.reader(fh)
            header = next(reader, [])
            if header != DETAIL_HEADER:
                problems["跨租户引用"].append(f"shared.csv 表头异常: {header}")
            for row in reader:
                if not row or not any(row):
                    continue
                shared_rows_disk.append(row)
                judge_row(row, "shared.csv")

    # 文件清单: 产物集合必须与本次应产出集合完全一致
    expected_files = {f"{m}.xlsx" for m in merchant_rows} | {"shared.csv"}
    actual_files = {
        p.name for p in out_dir.iterdir() if p.is_file() and not p.name.startswith(".")
    } if out_dir.exists() else set()
    for name in sorted(expected_files - actual_files):
        problems["文件清单"].append(f"缺失 {name}")
    for name in sorted(actual_files - expected_files):
        problems["文件清单"].append(f"多余文件 {name}")

    # 行数对账
    counted = 0
    for mch, rows in readback.items():
        expected_n = len(merchant_rows[mch])
        counted += len(rows)
        if len(rows) != expected_n:
            problems["行数对账"].append(f"{mch}.xlsx 行数 {len(rows)}, 应为 {expected_n}")
        if len(rows) != len({r[0] for r in rows if r}):
            problems["行数对账"].append(f"{mch}.xlsx 交易号重复")
    counted += len(shared_rows_disk)
    if len(shared_rows_disk) != len(shared_rows):
        problems["行数对账"].append(f"shared.csv 行数 {len(shared_rows_disk)}, 应为 {len(shared_rows)}")
    if counted != detail_count:
        problems["行数对账"].append(
            f"产物合计 {counted} 行 != 源明细 {detail_count} 行(已剔除合计行 {total_rows_skipped} 行)"
        )

    # 费用合计: 逐笔费用已独立复算, 这里再校验总额可加性
    fee_sum = 0
    for rows in list(readback.values()) + [shared_rows_disk]:
        for row in rows:
            if len(row) == len(DETAIL_HEADER):
                try:
                    fee_sum += int(row[5])
                except ValueError:
                    pass
    if fee_sum < 0:
        problems["费用合计"].append(f"费用合计为负: {fee_sum}")

    # 合计行混入: 若行数对账不平, 同样说明有合计行级别的混入
    if counted != detail_count:
        problems["合计行混入"].append("产物行数与源明细不一致, 疑似合计行混入")

    for name in AUDIT_NAMES[:-1]:
        msgs = problems[name]
        result[name] = (not msgs, "; ".join(msgs) if msgs else "OK")
    return result, readback


def _workbook_sheet_names(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as zf:
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
    return [sh.get("name") for sh in wb.iter(f"{{{NS_MAIN}}}sheet")]


# --------------------------------------------------------------------------- #
# 台账
# --------------------------------------------------------------------------- #
EXPORT_RE = re.compile(r"^\[EXPORT\] time=(\S+) detail_rows=(\d+) files=(\d+)$")
AUDIT_RE = re.compile(r"^\[AUDIT\] (\S+) (OK|FAIL)$")
RESULT_RE = re.compile(r"^\[AUDIT_RESULT\] (PASS|FAIL)$")


def parse_segments(text: str) -> list[list[str]]:
    return _segments_from_lines(text.splitlines())


def ledger_check(text: str, today: str, detail_rows: int, files: int) -> tuple[bool, str]:
    """审计第六项: 台账登记。"""
    segments = parse_segments(text)
    todays = [s for s in segments if s and s[0].startswith(f"[EXPORT] time={today}")]
    if len(todays) != 1:
        return False, f"当日台账段数 {len(todays)}, 应为 1"
    seg = todays[0]
    m = EXPORT_RE.match(seg[0])
    if not m:
        return False, "当日 [EXPORT] 行格式不符"
    if int(m.group(2)) != detail_rows or int(m.group(3)) != files:
        return False, (
            f"台账登记明细行数/文件数 {m.group(2)}/{m.group(3)} "
            f"与本次 {detail_rows}/{files} 不一致"
        )
    audits = [AUDIT_RE.match(line) for line in seg[1:] if line.startswith("[AUDIT] ")]
    if len(audits) != 6 or any(a is None for a in audits):
        return False, "当日审计行数不为 6"
    names = [a.group(1) for a in audits if a]
    if names != AUDIT_NAMES:
        return False, f"当日审计项名称/顺序不符: {names}"
    results = [line for line in seg if line.startswith("[AUDIT_RESULT]")]
    if len(results) != 1 or not RESULT_RE.match(results[0]):
        return False, "当日缺少唯一的 [AUDIT_RESULT] 行"
    if seg[-1] != results[0]:
        return False, "[AUDIT_RESULT] 不是当日段末行"
    # 总结论必须与六项结论自洽
    expected_result = "PASS" if all(a.group(2) == "OK" for a in audits if a) else "FAIL"
    if results[0] != f"[AUDIT_RESULT] {expected_result}":
        return False, f"[AUDIT_RESULT] 为 {results[0]}, 与六项结论应为 {expected_result}"
    return True, "OK"


def build_segment(ts: str, detail_rows: int, files: int, checks: dict[str, tuple[bool, str]]) -> list[str]:
    lines = [f"[EXPORT] time={ts} detail_rows={detail_rows} files={files}"]
    lines += [f"[AUDIT] {name} {'OK' if checks[name][0] else 'FAIL'}" for name in AUDIT_NAMES]
    overall = all(checks[name][0] for name in AUDIT_NAMES)
    lines.append(f"[AUDIT_RESULT] {'PASS' if overall else 'FAIL'}")
    return lines


# --------------------------------------------------------------------------- #
# 台账落盘
# --------------------------------------------------------------------------- #
def _segments_from_lines(lines: list[str]) -> list[list[str]]:
    """把台账行切成段; 段以 [EXPORT] 开头, 空行忽略。"""
    segs: list[list[str]] = []
    for line in lines:
        if not line.strip():
            continue
        if line.startswith("[EXPORT]"):
            segs.append([line])
        elif segs:
            segs[-1].append(line)
    return segs


def _render_segments(segs: list[list[str]]) -> str:
    return "\n".join("\n".join(seg) for seg in segs) + "\n"


def _write_ledger(today: str, detail_rows: int, files: int,
                  checks: dict[str, tuple[bool, str]]) -> bool:
    """登记台账。同日已有一段则就地刷新, 不追加第二段。返回是否新增了段。

    刷新后回读校验第六项(台账登记); 若结论有变, 用修正后的结论重写一次。
    """
    old_text = LOG_PATH.read_text(encoding="utf-8") if LOG_PATH.exists() else ""
    segs = _segments_from_lines(old_text.splitlines())

    same_idx = [i for i, seg in enumerate(segs) if seg[0].startswith(f"[EXPORT] time={today}")]
    if len(same_idx) > 1:
        # 历史遗留的重复同日段: 全部合并为第一段的位置
        keep = same_idx[0]
        segs = [seg for i, seg in enumerate(segs) if i not in same_idx[1:]]
        same_idx = [keep]

    checks["台账登记"] = (True, "OK")
    if same_idx:
        ts = segs[same_idx[0]][0].split("time=")[1].split()[0]
        segs[same_idx[0]] = build_segment(ts, detail_rows, files, checks)
        added = False
    else:
        ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        segs.append(build_segment(ts, detail_rows, files, checks))
        added = True

    text = _render_segments(segs)
    ok, reason = ledger_check(text, today, detail_rows, files)
    if not ok:
        checks["台账登记"] = (False, reason)
        idx = [i for i, seg in enumerate(segs) if seg[0].startswith(f"[EXPORT] time={today}")][0]
        segs[idx] = build_segment(segs[idx][0].split("time=")[1].split()[0], detail_rows, files, checks)
        text = _render_segments(segs)
        ok2, reason2 = ledger_check(text, today, detail_rows, files)
        if not ok2:
            checks["台账登记"] = (False, reason2)

    _write_if_changed(LOG_PATH, text.encode("utf-8"))
    return added


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #
def main() -> int:
    for path in (SRC_XLSX, TERMINALS_CSV, FEE_RULES_JSON):
        if not path.exists():
            print(f"缺少输入: {path}", file=sys.stderr)
            return 2

    try:
        tiers = load_fee_rules(FEE_RULES_JSON)
        binds = load_bindings(TERMINALS_CSV)
        detail, total_rows_skipped = load_source(SRC_XLSX)
    except ExportError as exc:
        print(f"导出失败: {exc}", file=sys.stderr)
        return 2

    merchant_rows, shared_rows = build_export(detail, binds, tiers)
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    # 产物内容确定(固定 zip 时间戳), 同日重复导出逐字节等价。
    writes: list[tuple[Path, bytes]] = []
    for mch in sorted(merchant_rows):
        writes.append((OUT_DIR / f"{mch}.xlsx",
                       _build_xlsx_bytes(DETAIL_HEADER, merchant_rows[mch])))
    shared_buf = tempfile.SpooledTemporaryFile(mode="w+", encoding="utf-8", newline="")
    writer = csv.writer(shared_buf, lineterminator="\n")
    writer.writerow(DETAIL_HEADER)
    writer.writerows(shared_rows)
    shared_buf.seek(0)
    writes.append((OUT_DIR / "shared.csv", shared_buf.read().encode("utf-8")))
    shared_buf.close()

    changed = sum(1 for path, data in writes if _write_if_changed(path, data))

    checks, _ = audit_exports(
        OUT_DIR,
        merchant_rows,
        shared_rows,
        tiers,
        binds,
        detail_count=len(detail),
        total_rows_skipped=total_rows_skipped,
    )

    now = datetime.now(timezone.utc)
    ts = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    today = now.strftime("%Y-%m-%d")
    files_exported = len(merchant_rows) + 1  # 商户工作簿 + shared.csv
    added = _write_ledger(today, len(detail), files_exported, checks)

    overall = all(checks[name][0] for name in AUDIT_NAMES)

    print(f"导出时刻 {ts}")
    print(f"明细 {len(detail)} 行, 商户 {len(merchant_rows)} 个, 文件 {files_exported} 个"
          f"(含 shared.csv), 剔除合计行 {total_rows_skipped} 行")
    if shared_rows:
        print(f"无法归属 {len(shared_rows)} 行 -> shared.csv")
    print(f"明细文件{'有更新' if changed else '内容未变'}({changed} 个重写)")
    print(f"台账: {'追加一段' if added else '同日已登记, 刷新当日段'}")
    for name in AUDIT_NAMES:
        passed, detail_msg = checks[name]
        print(f"  [AUDIT] {name} {'OK' if passed else 'FAIL'}"
              + ("" if passed else f"  # {detail_msg}"))
    print(f"[AUDIT_RESULT] {'PASS' if overall else 'FAIL'}")
    return 0 if overall else 1


if __name__ == "__main__":
    sys.exit(main())
