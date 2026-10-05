"""Parse only the result table and labelled journal detail cells."""
import re
from datetime import datetime, timezone
from urllib.parse import parse_qs, urljoin, urlsplit, urlunsplit

from bs4 import BeautifulSoup

from .inputs import normalize_name, normalize_issn

BASE = "https://www.letpub.com.cn/"
SEARCH = BASE + "index.php?page=journalapp&view=search"


class ParseError(Exception):
    pass


def detail_url(journal_id):
    return BASE + f"index.php?page=journalapp&view=detail&journalid={journal_id}"


def text(node):
    return " ".join(node.get_text(" ", strip=True).split()) if node else ""


def visible_soup(html):
    soup = BeautifulSoup(html, "html.parser")
    # Apply simple stylesheet selectors before removing styles themselves.
    for style in soup.find_all("style"):
        for selectors, rules in re.findall(r"([^{}]+)\{([^{}]*)\}", style.get_text()):
            if re.search(r"(?:display\s*:\s*none|visibility\s*:\s*hidden)", rules, re.I):
                try:
                    for node in soup.select(selectors.strip()):
                        node.decompose()
                except Exception:
                    # Unsupported CSS selectors cannot be applied by SoupSieve.
                    continue
    for node in list(soup.find_all(True)):
        if node.attrs is None:
            continue
        hidden = (node.has_attr("hidden") or node.get("aria-hidden") == "true" or
                  re.search(r"(?:display\s*:\s*none|visibility\s*:\s*hidden)", node.get("style", ""), re.I))
        if hidden or node.name in ("script", "style", "noscript"):
            node.decompose()
    return soup


def safe_link(href):
    # LetPub uses an unescaped &currentsearchpage; HTML decoding makes ¤t.
    href = href.replace("¤tsearchpage", "&currentsearchpage")
    parts = urlsplit(urljoin(BASE, href))
    if parts.hostname not in ("www.letpub.com.cn", "letpub.com.cn"):
        raise ParseError("查询链接指向非 LetPub 主机")
    return urlunsplit(("https", "www.letpub.com.cn", parts.path, parts.query, ""))


def parse_search(html, current_page=1):
    soup = visible_soup(html)
    table = next((t for t in soup.select("table.table_yjfx, table#journallisttable")
                  if "ISSN" in text(t) and "期刊名" in text(t)), None)
    if table is None:
        message = text(soup)
        if re.search(r"(?:没有|未找到|暂无|无符合|未查到|找不到).{0,15}(?:结果|期刊|记录)|查询结果为\s*0|找到\s*0\s*(?:本|条)", message):
            return {"candidates": [], "next_url": None}
        raise ParseError("无法识别期刊搜索结果表，不能判断为未收录")
    candidates = {}
    for anchor in table.find_all("a", href=True):
        href = safe_link(anchor["href"])
        query = parse_qs(urlsplit(href).query)
        identifier = query.get("journalid", [""])[0]
        if query.get("view") != ["detail"] or not identifier.isdigit() or identifier == "0":
            continue
        if "xuanxiangk_id" in query or not text(anchor) or text(anchor) == "文章":
            continue
        row = anchor.find_parent("tr")
        matches = re.findall(r"\b\d{4}-\d{3}[\dX]\b", text(row), re.I)
        candidates[identifier] = {"journal_id": identifier, "journal_name": text(anchor),
                                  "issn": matches[0].upper() if matches else None,
                                  "url": detail_url(identifier)}
    next_url = None
    for anchor in table.find_all("a", href=True):
        if text(anchor) == "下一页":
            href = safe_link(anchor["href"])
            page = parse_qs(urlsplit(href).query).get("currentsearchpage", [""])[0]
            if not page.isdigit():
                raise ParseError("无法识别搜索分页参数")
            if int(page) > current_page:
                next_url = href
            break
    return {"candidates": list(candidates.values()), "next_url": next_url}


def field(value=None, status="not_provided", raw=None, version=None, year=None):
    return {"value": value, "status": status, "raw": raw, "version": version, "year": year}


def version_info(label):
    version = re.search(r"[（(]\s*([^()（）]*\d{4}[^()（）]*)\s*[）)]", label)
    # A publication/release year is not necessarily the metric year.
    explicit = re.search(r"(?<!\d)(\d{4})\s*(?:年度|年(?:影响因子|IF|CiteScore))", label, re.I)
    return version.group(1).strip() if version else label, int(explicit.group(1)) if explicit else None


def cell_status(cell):
    raw = text(cell)
    if re.search(r"(?:登录|注册).{0,20}(?:查看|后)|(?:查看|获取).{0,15}(?:登录|注册)", raw):
        return "login_required"
    if raw.lower() in ("", "n/a", "na", "--", "-", "暂无", "无记录", "无"):
        return "not_provided"
    return "ok"


def numeric_field(label, cell):
    status = cell_status(cell)
    raw = text(cell)
    version, year = version_info(label)
    if status != "ok":
        return field(status=status, raw=raw, version=version, year=year)
    # Only the labelled cell's leading numeric value is usable.
    match = re.match(r"\s*(\d+(?:\.\d+)?)\b", raw)
    return field(float(match.group(1)) if match else None,
                 "ok" if match else "parse_error", raw, version, year)


def own_rows(table):
    return [tr for tr in table.find_all("tr") if tr.find_parent("table") is table]


def table_records(table):
    if table is None:
        return []
    headers = []
    records = []
    for tr in own_rows(table):
        cells = tr.find_all(["td", "th"], recursive=False)
        if not cells:
            continue
        if any(c.name == "th" for c in cells) or not headers:
            headers = [text(c) for c in cells]
        else:
            records.append(dict(zip(headers, cells)))
    return records


def quartile(node):
    raw = text(node)
    match = re.search(r"\bQ[1-4]\b|[1-4]\s*区", raw, re.I)
    return match.group(0).replace(" ", "").upper() if match else None


def parse_classification(label, cell, scheme):
    version, year = version_info(label)
    release = re.search(r"(\d{4})年", version)
    if scheme != "jcr" and release:
        year = int(release.group(1))
    status = cell_status(cell)
    entries = []
    if status == "ok":
        records = table_records(cell.find("table"))
        for record in records:
            if "大类学科" in record:
                big = record["大类学科"]
                raw = text(big)
                entries.append({"level": "major", "subject": re.sub(r"[1-4]\s*区", "", raw).strip(),
                                "quartile": quartile(big), "basis": None})
                small = record.get("小类学科")
                if small:
                    nested = small.find("table")
                    if nested:
                        for tr in own_rows(nested):
                            cells = tr.find_all("td", recursive=False)
                            if len(cells) >= 2:
                                entries.append({"level": "minor", "subject": text(cells[0]),
                                                "quartile": quartile(cells[1]), "basis": None})
                    elif text(small):
                        entries.append({"level": "minor", "subject": re.sub(r"[1-4]\s*区", "", text(small)).strip(),
                                        "quartile": quartile(small), "basis": None})
            else:
                subject_key = next((key for key in record if "学科" in key), None)
                quartile_key = next((key for key in record if "分区" in key and key != subject_key), None)
                if subject_key and quartile_key:
                    entries.append({"level": "subject", "subject": text(record[subject_key]),
                                    "quartile": quartile(record[quartile_key]), "basis": subject_key})
        # JCR may have multiple distinct labelled tables, e.g. JIF and JCI.
        if scheme == "jcr":
            entries = []
            for table in cell.find_all("table"):
                for record in table_records(table):
                    subject_key = next((key for key in record if "学科" in key), None)
                    qkey = next((key for key in record if "分区" in key and key != subject_key), None)
                    if subject_key and qkey:
                        basis = subject_key
                        entries.append({"level": "subject", "subject": text(record[subject_key]),
                                        "quartile": quartile(record[qkey]), "basis": basis})
        if not entries or any(not e["quartile"] for e in entries):
            status = "parse_error"
    return {"scheme": scheme, "source_label": label, "version": version, "year": year,
            "status": status, "entries": entries}


def release_key(label):
    """Compare explicit publication labels; do not assume DOM order is newest."""
    years = re.findall(r"\d{4}", label)
    month = re.search(r"\d{4}年\s*(\d{1,2})月", label)
    return (max(map(int, years)) if years else 0, int(month.group(1)) if month else 0)


def parse_detail(html, journal_id):
    soup = visible_soup(html)
    label = next((td for td in soup.find_all("td") if text(td) == "期刊名字"), None)
    if label is None:
        raise ParseError("无法识别期刊详情表")
    table = label.find_parent("table")
    pairs = []
    for tr in own_rows(table):
        cells = tr.find_all("td", recursive=False)
        if len(cells) >= 2:
            pairs.append((text(cells[0]), cells[1]))
    by_label = dict(pairs)
    name_cell = by_label["期刊名字"]
    anchor = next((a for a in name_cell.find_all("a", href=True)
                   if parse_qs(urlsplit(safe_link(a["href"])).query).get("journalid") == [str(journal_id)]), None)
    if anchor is None or not text(anchor):
        raise ParseError("详情页缺少对应期刊名称链接")
    fields = {key: field() for key in ("journal_name", "abbreviation", "issn", "eissn", "impact_factor", "citescore", "oa", "review_time")}
    fields["journal_name"] = field(text(anchor), "ok", text(anchor))
    abbr = name_cell.find("font", color=re.compile("gr[ae]y", re.I))
    fields["abbreviation"] = field(text(abbr) or None, "ok" if text(abbr) else "not_provided", text(abbr))
    for key, source in (("issn", "期刊ISSN"), ("eissn", "E-ISSN")):
        cell = by_label.get(source)
        status = cell_status(cell)
        value = None
        if status == "ok":
            try:
                value = normalize_issn(text(cell))
            except ValueError:
                status = "parse_error"
        fields[key] = field(value, status, text(cell))
    classifications = {}
    for label_text, cell in pairs:
        if re.search(r"最新IF|最新.*影响因子|\d{4}年度(?:IF|影响因子)", label_text, re.I) and "实时" not in label_text:
            fields["impact_factor"] = numeric_field(label_text, cell)
        elif label_text.startswith("CiteScore"):
            inner = cell.find("table")
            records = table_records(inner)
            number_cell = records[0].get("CiteScore") if records else None
            fields["citescore"] = numeric_field(label_text, number_cell if number_cell is not None else cell)
        elif label_text.startswith("是否OA"):
            raw = text(cell)
            state = cell_status(cell)
            values = {"yes": "open", "是": "open", "no": "closed", "否": "closed", "hybrid": "hybrid", "混合oa": "hybrid"}
            value = values.get(raw.casefold(), "unknown")
            fields["oa"] = field(value, state if state != "ok" or value != "unknown" else "parse_error", raw)
        elif label_text == "平均审稿速度":
            fields["review_time"] = field(text(cell) if cell_status(cell) == "ok" else None, cell_status(cell), text(cell))
        elif "JCR分区" in label_text:
            if "jcr" not in classifications or release_key(label_text) > release_key(classifications["jcr"]["source_label"]):
                classifications["jcr"] = parse_classification(label_text, cell, "jcr")
        elif "新锐期刊分区表" in label_text:
            if "xinrui" not in classifications or release_key(label_text) > release_key(classifications["xinrui"]["source_label"]):
                classifications["xinrui"] = parse_classification(label_text, cell, "xinrui")
        elif label_text.startswith("期刊分区表") and "预警" not in label_text:
            scheme = "journal_partition"
            if scheme not in classifications or release_key(label_text) > release_key(classifications[scheme]["source_label"]):
                classifications[scheme] = parse_classification(label_text, cell, scheme)
    for scheme in ("jcr", "journal_partition", "xinrui"):
        classifications.setdefault(scheme, {"scheme": scheme, "source_label": None, "version": None,
                                           "year": None, "status": "not_provided", "entries": []})
    return {"journal_id": str(journal_id), "fields": fields, "classifications": list(classifications.values()),
            "source_url": detail_url(journal_id), "fetched_at": datetime.now(timezone.utc).isoformat()}
