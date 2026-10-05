"""Deterministic matching and interruption-safe batch execution."""
import copy

from .client import AccessStopped, FetchError
from .inputs import name_key
from .parsing import ParseError
from .storage import append_event, export_results, result_data

COMPLETED = {"matched", "not_found", "needs_confirmation", "invalid"}


def identity_matches(row, record):
    fields = record["fields"]
    if row["issn"] and row["issn"] not in (fields["issn"]["value"], fields["eissn"]["value"]):
        return False
    if row["journal_name"] and name_key(row["journal_name"]) != name_key(fields["journal_name"]["value"]):
        return False
    return True


def candidate_rows(row, candidates, reason):
    return [{**candidate, "input_id": row["input_id"], "reason": reason} for candidate in candidates]


def make_event(row, status, message, journal=None, candidates=None):
    return {"query": {**row, "status": status, "message": message,
                      "journal_id": journal["journal_id"] if journal else None},
            "journal": journal, "candidates": candidates or []}


def resolve(row, client, existing):
    if row["error"]:
        return make_event(row, "invalid", row["error"])
    for record in existing.values():
        if identity_matches(row, record):
            return matched_event(row, copy.deepcopy(record))
    search = client.search(row)
    candidates = search["candidates"]
    if not search["complete"]:
        reason = "搜索分页未完整读取，请补充 ISSN 后核对"
        return make_event(row, "needs_confirmation", reason, candidates=candidate_rows(row, candidates, reason))
    if not candidates:
        return make_event(row, "not_found", "LetPub 搜索没有返回期刊记录")
    if row["issn"]:
        # Search uses printed ISSN even when the query was an eISSN.
        matching = candidates
    else:
        matching = [c for c in candidates if name_key(c["journal_name"]) == name_key(row["journal_name"])]
    if len(matching) != 1:
        reason = "没有唯一的精确匹配，请在候选清单核对并补充 ISSN"
        return make_event(row, "needs_confirmation", reason, candidates=candidate_rows(row, candidates, reason))
    record = client.detail(matching[0]["journal_id"])
    if not identity_matches(row, record):
        reason = "详情中的名称或 ISSN 与输入不一致"
        return make_event(row, "needs_confirmation", reason, candidates=candidate_rows(row, candidates, reason))
    return matched_event(row, record)


def matched_event(row, record):
    broken = [key for key, f in record["fields"].items() if f["status"] == "parse_error"]
    broken.extend(p["scheme"] for p in record["classifications"] if p["status"] == "parse_error")
    return make_event(row, "partial" if broken else "matched",
                      "部分字段解析失败: " + ", ".join(broken) if broken else "精确匹配", record)


def run_batch(inputs, directory, client, events=None, emit=print, export_options=None):
    events = dict(events or {})
    existing = {event["journal"]["journal_id"]: event["journal"] for event in events.values()
                if event.get("journal") and event["query"]["status"] == "matched"}
    exit_code = None
    try:
        for index, row in enumerate(inputs, 1):
            previous = events.get(row["input_id"])
            if previous and previous["query"]["status"] in COMPLETED:
                emit(f"[{index}/{len(inputs)}] 跳过已完成输入 {row['input_id']}")
                continue
            try:
                event = resolve(row, client, existing)
            except AccessStopped as error:
                event = make_event(row, "blocked", str(error))
                events[row["input_id"]] = event
                append_event(directory, event)
                emit(f"[{index}/{len(inputs)}] blocked: {error}")
                break
            except (FetchError, ParseError) as error:
                event = make_event(row, "failed", str(error))
            events[row["input_id"]] = event
            append_event(directory, event)
            if event.get("journal") and event["query"]["status"] == "matched":
                existing[event["journal"]["journal_id"]] = event["journal"]
            emit(f"[{index}/{len(inputs)}] {event['query']['status']}: {row['journal_name'] or row['issn']} — {event['query']['message']}")
    except KeyboardInterrupt:
        emit("收到中断，正在保存已完成结果")
        exit_code = 130
    finally:
        data = result_data(inputs, events)
        export_results(directory, data, export_options, emit=emit)
    counts = {status: sum(q["status"] == status for q in data["queries"])
              for status in ("matched", "needs_confirmation", "not_found", "partial", "failed", "invalid", "blocked", "pending")}
    emit("完成: " + ", ".join(f"{k}={v}" for k, v in counts.items() if v) + f"; 去重期刊={len(data['journals'])}")
    if exit_code is None:
        exit_code = 0 if all(q["status"] == "matched" for q in data["queries"]) else 2
    return exit_code, data
