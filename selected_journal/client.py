"""Rate-limited HTTP access and mode-separated 24-hour detail cache."""
import copy
import hashlib
import http.cookiejar
import json
import os
import re
import tempfile
import time
from pathlib import Path
from urllib.parse import urlsplit

import requests

from .parsing import SEARCH, ParseError, detail_url, parse_detail, parse_search, text, visible_soup


class FetchError(Exception):
    pass


class AccessStopped(FetchError):
    """No more network access should occur during this run."""


def atomic_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(prefix=".tmp-", dir=path.parent)
    try:
        with os.fdopen(handle, "w", encoding="utf-8") as stream:
            json.dump(data, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def load_cookies(session, filename):
    jar = http.cookiejar.MozillaCookieJar(str(filename))
    try:
        jar.load(ignore_discard=True, ignore_expires=False)
    except (OSError, http.cookiejar.LoadError) as error:
        raise ValueError("无法读取 Netscape Cookie 文件，请检查路径与格式") from error
    allowed = [cookie for cookie in jar if cookie.domain.lstrip(".").lower() in
               ("letpub.com.cn", "www.letpub.com.cn") and not cookie.is_expired()]
    for cookie in allowed:
        session.cookies.set_cookie(cookie)
    # Digest separates accounts without persisting the credential itself.
    signature = "\n".join(sorted(f"{c.domain}/{c.path}/{c.name}={c.value}" for c in allowed))
    return "login-" + hashlib.sha256(signature.encode()).hexdigest()[:16] if allowed else "anonymous"


class LetPubClient:
    def __init__(self, cache_dir, cookie_file=None, refresh=False, session=None,
                 sleep=time.sleep, clock=time.monotonic):
        self.session = session or requests.Session()
        self.session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; SelectedJournal/0.1)",
                                     "Accept-Language": "zh-CN,zh;q=0.9"})
        self.mode = load_cookies(self.session, cookie_file) if cookie_file else "anonymous"
        self.cache_dir = Path(cache_dir) / self.mode
        self.refresh = refresh
        self.sleep = sleep
        self.clock = clock
        self.last_request = None
        self.ready = False
        self.stopped = False
        self.details = {}
        self.searches = {}

    def close(self):
        self.session.close()

    def fetch(self, method, url, **kwargs):
        if self.stopped:
            raise AccessStopped("本次任务已停止网络访问")
        if urlsplit(url).hostname not in ("letpub.com.cn", "www.letpub.com.cn") or not url.startswith("https://"):
            raise FetchError("仅允许 HTTPS LetPub 请求")
        for attempt in range(3):
            if self.last_request is not None:
                self.sleep(max(0, 3 - (self.clock() - self.last_request)))
            self.last_request = self.clock()
            try:
                response = self.session.request(method, url, timeout=(10, 30), **kwargs)
            except (requests.Timeout, requests.ConnectionError):
                if attempt == 2:
                    raise FetchError("网络超时或连接失败，已尝试 3 次")
                self.sleep((5, 15)[attempt])
                continue
            except requests.RequestException as error:
                raise FetchError("HTTP 请求失败") from error
            if response.status_code in (403, 429):
                self.stopped = True
                raise AccessStopped(f"HTTP {response.status_code}，已停止网络访问")
            if response.status_code in (408, 500, 502, 503, 504):
                if attempt < 2:
                    self.sleep((5, 15)[attempt])
                    continue
                raise FetchError(f"HTTP {response.status_code}，已尝试 3 次")
            if response.status_code != 200:
                raise FetchError(f"HTTP {response.status_code}")
            response.encoding = "utf-8"
            html = response.text
            soup = visible_soup(html)
            message = text(soup)
            has_data = (any(text(td) == "期刊名字" for td in soup.find_all("td")) or
                        any("ISSN" in text(t) and "期刊名" in text(t) for t in soup.select("table.table_yjfx")))
            if not has_data and re.search(r"请输入验证码|请完成验证|安全验证|访问过于频繁|异常访问|captcha", message, re.I):
                self.stopped = True
                raise AccessStopped("检测到访问验证页面，已停止网络访问")
            return html
        raise FetchError("请求失败")

    def search(self, row):
        key = (row["journal_name"].casefold(), row["issn"])
        if key in self.searches:
            return copy.deepcopy(self.searches[key])
        if not self.ready:
            self.fetch("GET", SEARCH)
            self.ready = True
        form = {"searchissn": row["issn"], "searchsort": "relevance"} if row["issn"] else {
            "searchname": row["journal_name"], "searchsort": "relevance"}
        html = self.fetch("POST", SEARCH, data=form)
        candidates = {}
        seen_pages = set()
        complete = True
        for page in range(1, 11):
            digest = hashlib.sha256(html.encode()).hexdigest()
            if digest in seen_pages:
                complete = False
                break
            seen_pages.add(digest)
            try:
                result = parse_search(html, page)
            except ParseError:
                if page == 1:
                    raise
                complete = False
                break
            for candidate in result["candidates"]:
                candidates[candidate["journal_id"]] = candidate
            if not result["next_url"]:
                break
            if page == 10:
                complete = False
                break
            try:
                html = self.fetch("GET", result["next_url"])
            except AccessStopped:
                raise
            except FetchError:
                complete = False
                break
        result = {"candidates": list(candidates.values()), "complete": complete}
        self.searches[key] = result
        return copy.deepcopy(result)

    def detail(self, identifier):
        identifier = str(identifier)
        if identifier in self.details:
            return copy.deepcopy(self.details[identifier])
        cache = self.cache_dir / f"{identifier}.json"
        record = None
        if not self.refresh and cache.exists():
            try:
                saved = json.loads(cache.read_text(encoding="utf-8"))
                age = time.time() - saved["saved_at"]
                if (0 <= age < 86400 and saved.get("schema_version") == 1 and
                        saved["record"]["journal_id"] == identifier and cacheable(saved["record"])):
                    record = saved["record"]
            except (ValueError, OSError, KeyError, TypeError):
                pass
        if record is None:
            record = parse_detail(self.fetch("GET", detail_url(identifier)), identifier)
            # Persist only structurally valid details, never HTML or Cookies.
            if cacheable(record):
                atomic_json(cache, {"schema_version": 1, "saved_at": time.time(), "record": record})
        self.details[identifier] = record
        return copy.deepcopy(record)


def cacheable(record):
    return (all(f["status"] != "parse_error" for f in record["fields"].values()) and
            all(p["status"] != "parse_error" for p in record["classifications"]))
