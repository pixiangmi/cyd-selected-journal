"""Append-only checkpoints and consistent Excel/CSV/JSON projections."""
import copy
import csv
import json
import os
import tempfile
from dataclasses import asdict, dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from .client import atomic_json

FIELD_NAMES = ("journal_name", "abbreviation", "issn", "eissn", "impact_factor", "citescore", "oa", "review_time")
SUMMARY_HEADERS = ["journal_id", "input_ids", "source_url", "fetched_at"] + [
    f"{name}{suffix}" for name in FIELD_NAMES for suffix in ("", "_status", "_raw", "_version", "_year")]
SUMMARY_HEADERS += [f"{scheme}_{part}" for scheme in ("jcr", "journal_partition", "xinrui")
                    for part in ("source_label", "version", "status", "subjects")]
PARTITION_HEADERS = ["journal_id", "journal_name", "scheme", "source_label", "version", "year", "status",
                     "level", "subject", "quartile", "basis", "source_url", "fetched_at"]
QUERY_HEADERS = ["input_id", "journal_name", "issn", "raw", "status", "journal_id", "message"]
CANDIDATE_HEADERS = ["input_id", "journal_id", "journal_name", "issn", "url", "reason"]

BASE_LABELS = {
    "journal_id": "LetPub期刊ID", "input_id": "输入ID", "input_ids": "关联输入ID",
    "journal_name": "期刊全称", "abbreviation": "期刊简称", "issn": "ISSN", "eissn": "eISSN",
    "impact_factor": "影响因子IF", "citescore": "CiteScore", "oa": "开放获取状态",
    "review_time": "审稿周期", "source_url": "来源链接", "fetched_at": "抓取时间",
    "scheme": "分区体系", "source_label": "原始体系名称", "version": "版本标签", "year": "年份",
    "status": "状态", "level": "学科层级", "subject": "学科名称", "quartile": "分区",
    "basis": "指标口径", "raw": "原始输入", "message": "查询说明", "url": "候选链接", "reason": "候选原因",
}
SUFFIX_LABELS = {"status": "状态", "raw": "页面原文", "version": "版本标签", "year": "指标年份"}
SCHEME_LABELS = {"jcr": "JCR分区", "journal_partition": "期刊分区表（中科院）", "xinrui": "新锐分区"}
PART_LABELS = {"source_label": "原始体系名称", "version": "版本标签", "status": "状态", "subjects": "学科分区"}
HEADER_LABELS = {**BASE_LABELS,
                 **{f"{name}_{suffix}": f"{BASE_LABELS[name]}_{label}"
                    for name in FIELD_NAMES for suffix, label in SUFFIX_LABELS.items()},
                 **{f"{scheme}_{part}": f"{label}_{PART_LABELS[part]}"
                    for scheme, label in SCHEME_LABELS.items() for part in PART_LABELS}}


@dataclass(frozen=True)
class ExportOptions:
    field_names: str = "original"
    drop_empty_columns: bool = False
    missing_report: str = "none"

    def __post_init__(self):
        if self.field_names not in ("original", "zh"):
            raise ValueError("字段名模式必须为 original 或 zh")
        if not isinstance(self.drop_empty_columns, bool):
            raise ValueError("删除空列选项必须为布尔值")
        if self.missing_report not in ("none", "terminal", "file", "both"):
            raise ValueError("缺失统计模式必须为 none、terminal、file 或 both")

    def to_dict(self):
        return asdict(self)

    def label(self, key):
        return HEADER_LABELS[key] if self.field_names == "zh" else key


def is_missing(value):
    """Treat only null/blank/empty containers as missing; 0 and False are data."""
    return value is None or (isinstance(value, str) and not value.strip()) or (
        isinstance(value, (list, dict, tuple)) and not value)


def export_tables(data, options):
    tables = []
    for title, headers, rows in projections(data):
        # No observations means no evidence a field is always empty. Preserve
        # the schema of an empty candidate/partition table.
        kept = [h for h in headers if not (
            options.drop_empty_columns and rows and all(is_missing(row.get(h)) for row in rows))]
        tables.append((title, kept, rows))
    return tables


def missing_statistics(data, options):
    """Count original columns, including columns removed from the export view."""
    tables = []
    for title, headers, rows in projections(data):
        columns = []
        for key in headers:
            empty = [row for row in rows if is_missing(row.get(key))]
            reasons = {}
            for row in empty:
                status = row.get(f"{key}_status")
                if key in ("quartile", "subject", "basis", "level"):
                    status = row.get("status")
                reason = status if status in ("login_required", "not_provided", "parse_error") else "unspecified"
                reasons[reason] = reasons.get(reason, 0) + 1
            columns.append({"field": key, "label": options.label(key), "row_count": len(rows),
                            "missing_count": len(empty), "present_count": len(rows) - len(empty),
                            "missing_rate": len(empty) / len(rows) if rows else None,
                            "dropped": bool(options.drop_empty_columns and rows and len(empty) == len(rows)),
                            "reasons": reasons})
        tables.append({"table": title, "row_count": len(rows), "columns": columns})
    return {"schema_version": 1, "definition": "null、空白字符串、空列表或空对象为空值；0、False、unknown 不算空值。空表缺失率为 null。",
            "field_names": options.field_names, "tables": tables}


def emit_missing_statistics(report, emit):
    emit("缺失值统计（基于删除空列前的全部字段）:")
    for table in report["tables"]:
        emit(f"{table['table']}: {table['row_count']} 行")
        if not table["row_count"]:
            emit("  无数据，缺失率不适用；保留表头")
            continue
        missing = [column for column in table["columns"] if column["missing_count"]]
        for column in missing:
            emit(f"  {column['label']}: {column['missing_count']}/{column['row_count']} "
                 f"({column['missing_rate']:.1%})" + (" [已删除空列]" if column["dropped"] else ""))
        emit(f"  无缺失字段: {len(table['columns']) - len(missing)} 列")


def write_csv(path, headers, rows):
    handle, temporary = tempfile.mkstemp(prefix=".tmp-", dir=path.parent)
    try:
        with os.fdopen(handle, "w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.writer(stream)
            writer.writerow(headers)
            writer.writerows(rows)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def save_missing_statistics(directory, report):
    atomic_json(directory / "missing_values.json", report)
    # Column labels follow the export language; original identifiers are always
    # retained alongside them so external code can locate fields reliably.
    headers = ["table", "field", "label", "row_count", "missing_count", "present_count", "missing_rate", "dropped", "reasons"]
    labels = ["工作表", "原始字段名", "输出字段名", "总行数", "缺失数", "非缺失数", "缺失率", "是否删除", "缺失原因计数"]
    rows = [[value_for_table(table["table"] if key == "table" else column[key]) for key in headers]
            for table in report["tables"] for column in table["columns"]]
    write_csv(directory / "missing_values.csv", labels if report["field_names"] == "zh" else headers, rows)


def append_event(directory, event):
    with (Path(directory) / "checkpoint.jsonl").open("a", encoding="utf-8") as stream:
        stream.write(json.dumps(event, ensure_ascii=False) + "\n")
        stream.flush()
        os.fsync(stream.fileno())


def read_events(directory):
    path = Path(directory) / "checkpoint.jsonl"
    if not path.exists():
        return {}
    data = path.read_bytes()
    events = {}
    offset = 0
    for index, line in enumerate(data.splitlines(keepends=True)):
        try:
            event = json.loads(line)
            events[event["query"]["input_id"]] = event
        except (ValueError, KeyError, TypeError):
            if index != len(data.splitlines()) - 1:
                raise ValueError("检查点中间记录损坏，无法安全恢复")
            # Only discard a torn trailing write, retaining all prior events.
            with path.open("r+b") as stream:
                stream.truncate(offset)
            break
        offset += len(line)
    else:
        if data and not data.endswith(b"\n"):
            with path.open("ab") as stream:
                stream.write(b"\n")
    return events


def result_data(inputs, events):
    queries = []
    journals = {}
    candidates = []
    for row in inputs:
        event = events.get(row["input_id"])
        if event is None:
            queries.append({**row, "status": "pending", "journal_id": None, "message": "尚未处理"})
            continue
        queries.append(event["query"])
        candidates.extend(event.get("candidates", []))
        record = event.get("journal")
        if record:
            key = record["journal_id"]
            if key not in journals:
                journals[key] = copy.deepcopy(record)
                journals[key]["input_ids"] = []
            journals[key]["input_ids"].append(row["input_id"])
    return {"schema_version": 1, "journals": list(journals.values()), "queries": queries, "candidates": candidates}


def projections(data):
    summary, partitions = [], []
    for journal in data["journals"]:
        row = {key: journal[key] for key in ("journal_id", "input_ids", "source_url", "fetched_at")}
        for name, f in journal["fields"].items():
            row[name] = f["value"]
            for part in ("status", "raw", "version", "year"):
                row[f"{name}_{part}"] = f[part]
        for partition in journal["classifications"]:
            scheme = partition["scheme"]
            for part in ("source_label", "version", "status"):
                row[f"{scheme}_{part}"] = partition[part]
            row[f"{scheme}_subjects"] = "; ".join(
                f"{e['basis'] or e['level']}: {e['subject']} {e['quartile'] or ''}" for e in partition["entries"])
            for entry in partition["entries"] or [{}]:
                partitions.append({**{k: partition[k] for k in ("scheme", "source_label", "version", "year", "status")},
                                   **entry, "journal_id": journal["journal_id"],
                                   "journal_name": journal["fields"]["journal_name"]["value"],
                                   "source_url": journal["source_url"], "fetched_at": journal["fetched_at"]})
        summary.append(row)
    return [("期刊汇总", SUMMARY_HEADERS, summary), ("分区明细", PARTITION_HEADERS, partitions),
            ("查询记录", QUERY_HEADERS, data["queries"]), ("候选清单", CANDIDATE_HEADERS, data["candidates"])]


def value_for_table(value):
    if isinstance(value, (dict, list)):
        value = json.dumps(value, ensure_ascii=False)
    if isinstance(value, str):
        # Preserve raw originals in JSON while making spreadsheet cells inert.
        value = "".join(c for c in value if ord(c) >= 32 or c in "\n\r\t")
        if value.lstrip().startswith(("=", "+", "-", "@")):
            value = "'" + value
    return value


def export_results(directory, data, options=None, emit=print):
    options = options or ExportOptions()
    directory = Path(directory)
    tables = export_tables(data, options)
    json_data = data
    if options.field_names != "original" or options.drop_empty_columns:
        json_data = {**data, "export_view": {
            "field_names": options.field_names, "drop_empty_columns": options.drop_empty_columns,
            "tables": [{"table": title,
                        "columns": [{"field": h, "label": options.label(h)} for h in headers],
                        "rows": [{options.label(h): row.get(h) for h in headers} for row in rows]}
                       for title, headers, rows in tables]}}
    atomic_json(directory / "results.json", json_data)
    workbook = Workbook()
    workbook.remove(workbook.active)
    for title, headers, rows in tables:
        sheet = workbook.create_sheet(title)
        labels = [options.label(h) for h in headers]
        sheet.append(labels)
        for row in rows:
            sheet.append([value_for_table(row.get(h)) for h in headers])
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="24557A")
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = min(50, max(14, max(len(str(c.value or "")) for c in col) + 2))
        write_csv(directory / f"{title}.csv", labels,
                  ([value_for_table(row.get(h)) for h in headers] for row in rows))
    handle, temporary = tempfile.mkstemp(prefix=".tmp-", suffix=".xlsx", dir=directory)
    os.close(handle)
    try:
        workbook.save(temporary)
        os.replace(temporary, directory / "results.xlsx")
    finally:
        workbook.close()
        if os.path.exists(temporary):
            os.unlink(temporary)
    if options.missing_report != "none":
        report = missing_statistics(data, options)
        if options.missing_report in ("file", "both"):
            save_missing_statistics(directory, report)
        if options.missing_report in ("terminal", "both"):
            emit_missing_statistics(report, emit)
    if options.missing_report not in ("file", "both"):
        # Re-exporting a resumed task must not leave an outdated report visible.
        for filename in ("missing_values.json", "missing_values.csv"):
            (directory / filename).unlink(missing_ok=True)
