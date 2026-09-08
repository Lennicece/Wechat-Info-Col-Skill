# -*- coding: utf-8 -*-
"""
微信公众号近20天文章抓取 + 观点总结（极简版）

你只需要改 1 个地方：SEED_KEYWORD
运行：python wechat_article_analyzer_simple.py

输出：
- 微信公众号抓取_关键词_YYYYMMDD.xlsx（单一文件）
  - Sheet1：发布时间/发布人/发布标题/发布连接/阅读量（按阅读量降序）
  - Sheet2：符合条件文章的 takeaways（逐篇内容总结）
"""

from __future__ import annotations

import random
import re
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Optional, Tuple
from urllib.parse import quote, urlparse, parse_qs

import jieba
import pandas as pd
import requests
from bs4 import BeautifulSoup
from sklearn.cluster import KMeans
from sklearn.feature_extraction.text import TfidfVectorizer


# ========== 只改这里 ==========
SEED_KEYWORD = "漫剧工业化"
PAGES_PER_KEYWORD = 2
MAX_URLS = 50
OUTPUT_DIR = "."
LOOKBACK_DAYS = 20
READ_COUNT_THRESHOLD = 500
# ============================


STOPWORDS = {
    "我们", "你们", "他们", "这个", "那个", "以及", "如果", "因为", "所以", "就是", "已经", "可以", "进行", "一个", "一些",
    "没有", "更多", "使用", "相关", "内容", "文章", "公众号", "今天", "最近", "目前", "通过", "以下", "其中", "自己",
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Accept-Language": "zh-CN,zh;q=0.9",
}


@dataclass
class Article:
    keyword: str
    title: str
    account: str
    publish_time: Optional[datetime]
    url: str
    content: str
    content_len: int
    read_count: Optional[int]
    is_blue_verified: bool
    collect_status: str
    error_reason: str


def expand_keywords(seed_keyword: str) -> List[str]:
    seed_keyword = seed_keyword.strip()
    if not seed_keyword:
        return []

    expansions = [seed_keyword]
    if "漫剧" in seed_keyword:
        expansions += [
            "漫剧",
            "AI漫剧",
            "AI漫剧 工业化",
            "漫剧 生产流程",
            "漫剧 内容生产",
            "漫画短剧 工业化",
            "短剧 工业化 制作",
        ]

    seen = set()
    out = []
    for kw in expansions:
        if kw not in seen:
            seen.add(kw)
            out.append(kw)
    return out


def tokenize_zh(text: str) -> List[str]:
    tokens = []
    for w in jieba.lcut(text or ""):
        w = w.strip().lower()
        if not w:
            continue
        if w in STOPWORDS:
            continue
        if len(w) < 2:
            continue
        if not re.search(r"[\u4e00-\u9fffA-Za-z0-9]", w):
            continue
        tokens.append(w)
    return tokens


def resolve_to_mp_url(url: str) -> Optional[str]:
    if not url:
        return None

    headers = dict(HEADERS)
    headers["Referer"] = "https://weixin.sogou.com/"

    try:
        r = requests.get(url, headers=headers, timeout=20, allow_redirects=True)
        final_url = r.url
        if "mp.weixin.qq.com" in final_url:
            return final_url.split("#")[0]

        text = r.text or ""
        parts = re.findall(r"url\s*\+=\s*'([^']*)';", text)
        if parts:
            js_url = "".join(parts).replace("@", "")
            if "mp.weixin.qq.com" in js_url:
                return js_url.split("#")[0]
    except Exception:
        return None

    return None


def search_wechat_urls_sogou(keyword: str, pages: int = 3) -> List[str]:
    urls: List[str] = []

    for page in range(1, pages + 1):
        search_url = f"https://weixin.sogou.com/weixin?type=2&query={quote(keyword)}&page={page}"
        try:
            resp = requests.get(search_url, headers=HEADERS, timeout=20)
            html = resp.text
        except Exception:
            continue

        if "请输入验证码" in html or "访问过于频繁" in html:
            break

        soup = BeautifulSoup(html, "lxml")
        nodes = soup.select("ul.news-list li h3 a")

        for a in nodes:
            href = (a.get("href") or "").strip()
            if not href:
                continue
            if href.startswith("/"):
                href = "https://weixin.sogou.com" + href

            mp_url = resolve_to_mp_url(href)
            if mp_url:
                urls.append(mp_url)

        time.sleep(random.uniform(1.0, 2.0))

    seen = set()
    unique = []
    for u in urls:
        if u not in seen:
            seen.add(u)
            unique.append(u)
    return unique


def search_wechat_urls_bing(keyword: str, pages: int = 2) -> List[str]:
    urls: List[str] = []

    for page_idx in range(pages):
        first = page_idx * 10 + 1
        q = f"site:mp.weixin.qq.com {keyword}"
        search_url = f"https://www.bing.com/search?q={quote(q)}&first={first}"

        try:
            resp = requests.get(search_url, headers=HEADERS, timeout=20)
            html = resp.text
        except Exception:
            continue

        soup = BeautifulSoup(html, "lxml")
        nodes = soup.select("li.b_algo h2 a")

        for a in nodes:
            href = (a.get("href") or "").strip()
            mp_url = resolve_to_mp_url(href)
            if mp_url:
                urls.append(mp_url)

        time.sleep(random.uniform(0.8, 1.5))

    seen = set()
    unique = []
    for u in urls:
        if u not in seen:
            seen.add(u)
            unique.append(u)
    return unique


def search_wechat_urls(keyword: str, pages: int = 3) -> List[str]:
    sogou_urls = search_wechat_urls_sogou(keyword, pages=pages)
    if sogou_urls:
        return sogou_urls
    return search_wechat_urls_bing(keyword, pages=max(1, pages // 2))


def parse_publish_time(html: str) -> Optional[datetime]:
    patterns = [
        r'var\s+publish_time\s*=\s*"(\d+)"',
        r'publish_time\s*=\s*["\'](\d+)["\']',
        r'ct\s*=\s*"(\d+)"',
        r'"createTime"\s*:\s*"(\d+)"',
        r'"oriCreateTime"\s*:\s*"(\d+)"',
    ]
    for p in patterns:
        m = re.search(p, html)
        if m:
            try:
                ts = int(m.group(1))
                if ts > 0:
                    return datetime.fromtimestamp(ts)
            except Exception:
                pass
    return None


def parse_account(soup: BeautifulSoup, html: str) -> str:
    js_name = soup.select_one("#js_name")
    if js_name and js_name.get_text(strip=True):
        return js_name.get_text(strip=True)

    m = re.search(r'var\s+nickname\s*=\s*htmlDecode\("([^"]+)"\)', html)
    if m:
        return m.group(1)

    m2 = re.search(r'"profile_nickname"\s*:\s*"([^"]+)"', html)
    if m2:
        return m2.group(1)

    return ""


def parse_is_blue_verified(soup: BeautifulSoup, html: str) -> bool:
    # 公众号文章页常见状态字段：verify_status > 0 通常代表认证账号
    status_patterns = [
        r"verify_status\s*[:=]\s*'?([0-9]+)",
        r"is_phacct_verify\s*[:=]\s*'?([0-9]+)",
    ]
    for p in status_patterns:
        matches = re.findall(p, html)
        for val in matches:
            try:
                if int(val) > 0:
                    return True
            except Exception:
                pass

    verify_type_patterns = [
        r'"verify_type"\s*:\s*"?(-?\d+)"?',
        r'var\s+verify_type\s*=\s*"?(-?\d+)"?',
    ]
    for p in verify_type_patterns:
        matches = re.findall(p, html)
        for val in matches:
            try:
                if int(val) > 0:
                    return True
            except Exception:
                pass

    text_markers = ["微信认证", "已认证", "腾讯认证"]
    whole_text = soup.get_text(" ", strip=True)
    if any(marker in whole_text for marker in text_markers):
        return True

    css_markers = [
        ".icon_verify",
        ".account_meta__verify",
        ".weui-icon-success-no-circle",
        ".profile_meta.verify_meta",
    ]
    for selector in css_markers:
        if soup.select_one(selector):
            return True

    return False


def parse_title(soup: BeautifulSoup) -> str:
    n1 = soup.select_one("#activity-name")
    if n1 and n1.get_text(strip=True):
        return n1.get_text(strip=True)

    n2 = soup.select_one("meta[property='og:title']")
    if n2 and n2.get("content"):
        return n2.get("content").strip()

    n3 = soup.select_one("title")
    if n3 and n3.get_text(strip=True):
        return n3.get_text(strip=True)

    return "（无标题）"


def parse_content(soup: BeautifulSoup) -> str:
    node = soup.select_one("#js_content") or soup.select_one(".rich_media_content")
    if node:
        txt = "\n".join(x.strip() for x in node.stripped_strings if x.strip())
        if len(txt) >= 30:
            return txt

    desc = soup.select_one("meta[property='og:description']")
    if desc and desc.get("content"):
        d = desc.get("content").strip()
        if d:
            return d

    ps = soup.select("p")
    txt = "\n".join(p.get_text(strip=True) for p in ps if p.get_text(strip=True))
    return txt


def _extract_js_var(html: str, name: str) -> Optional[str]:
    patterns = [
        rf"var\s+{re.escape(name)}\s*=\s*\"([^\"]*)\"",
        rf"var\s+{re.escape(name)}\s*=\s*'([^']*)'",
        rf"{re.escape(name)}\s*:\s*\"([^\"]*)\"",
        rf"{re.escape(name)}\s*:\s*'([^']*)'",
    ]
    for p in patterns:
        m = re.search(p, html)
        if m:
            return m.group(1)
    return None


def extract_read_count(url: str, html: str, headers: dict) -> Optional[int]:
    # 1) 直接从页面脚本中尝试提取
    direct_patterns = [
        r'"read_num"\s*:\s*"?([0-9]+)"?',
        r'"readCount"\s*:\s*"?([0-9]+)"?',
        r'"read_count"\s*:\s*"?([0-9]+)"?',
        r"read_num\s*[:=]\s*'?([0-9]+)'?",
        r"readCount\s*[:=]\s*'?([0-9]+)'?",
    ]
    for p in direct_patterns:
        m = re.search(p, html)
        if m:
            try:
                return int(m.group(1))
            except Exception:
                pass

    # 2) 调公众号扩展接口（成功率受风控影响）
    parsed = urlparse(url)
    q = parse_qs(parsed.query)
    biz = q.get("__biz", [None])[0]
    mid = q.get("mid", [None])[0]
    idx = q.get("idx", [None])[0]
    sn = q.get("sn", [None])[0]

    if not all([biz, mid, idx, sn]):
        return None

    params = {
        "__biz": biz,
        "mid": mid,
        "idx": idx,
        "sn": sn,
        "scene": "126",
        "is_only_read": "1",
    }

    optional_keys = ["uin", "key", "pass_ticket", "wxtoken", "appmsg_token"]
    for k in optional_keys:
        v = _extract_js_var(html, k)
        if v:
            params[k] = v

    api = "https://mp.weixin.qq.com/mp/getappmsgext"
    data = {
        "is_only_read": "1",
        "is_temp_url": "0",
        "appmsg_type": "9",
        "reward_uin_count": "0",
    }

    try:
        h = dict(headers)
        h["Referer"] = url
        h["X-Requested-With"] = "XMLHttpRequest"
        r = requests.post(api, headers=h, params=params, data=data, timeout=20)
        j = r.json()
        read_num = ((j.get("appmsgstat") or {}).get("read_num"))
        if read_num is not None:
            return int(read_num)
    except Exception:
        return None

    return None


def crawl_article(url: str, keyword: str) -> Article:
    headers = dict(HEADERS)
    headers["Referer"] = "https://mp.weixin.qq.com/"

    try:
        resp = requests.get(url, headers=headers, timeout=20)
        resp.encoding = "utf-8"
        html = resp.text
    except Exception as e:
        return Article(keyword, "", "", None, url, "", 0, None, False, "failed", f"request_error: {e}")

    soup = BeautifulSoup(html, "lxml")
    title = parse_title(soup)
    account = parse_account(soup, html)
    publish_time = parse_publish_time(html)
    content = parse_content(soup)
    read_count = extract_read_count(url, html, headers)
    is_blue_verified = parse_is_blue_verified(soup, html)

    if len(content) < 20 and not title:
        return Article(keyword, title, account, publish_time, url, content, len(content), read_count, is_blue_verified, "failed", "content_too_short")

    return Article(keyword, title, account, publish_time, url, content, len(content), read_count, is_blue_verified, "ok", "")


def build_summary(df_focus: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    if df_focus.empty:
        rows = [{"item": "总体结论", "value": "未形成可分析样本，建议增加关键词扩展或提高抓取页数。"}]
        return pd.DataFrame(rows), ["未形成可分析样本。"]

    docs = df_focus["content"].fillna("").tolist()
    docs = [d if len(d) <= 3000 else d[:3000] for d in docs]

    if len(docs) == 1:
        one = df_focus.iloc[0]
        lines = [
            f"近{LOOKBACK_DAYS}天仅抓到1篇可用文章，暂无稳定趋势。",
            f"代表文章：《{one['title']}》 {one['url']}",
        ]
        rows = [{"item": "总体结论", "value": lines[0]}, {"item": "代表文章", "value": lines[1]}]
        return pd.DataFrame(rows), lines

    vectorizer = TfidfVectorizer(tokenizer=tokenize_zh, token_pattern=None, max_features=2000, min_df=1)
    X = vectorizer.fit_transform(docs)

    k = min(3, len(docs))
    model = KMeans(n_clusters=k, random_state=42, n_init="auto")
    labels = model.fit_predict(X)

    df_tmp = df_focus.copy()
    df_tmp["topic_id"] = labels
    terms = vectorizer.get_feature_names_out()

    summary_rows = [
        {"item": "样本量", "value": f"{len(df_focus)} 篇（用于观点分析）"},
        {"item": "账号数", "value": f"{df_focus['account'].replace('', '未知账号').nunique()}"},
    ]
    lines = []

    for topic_id in sorted(df_tmp["topic_id"].unique()):
        sub = df_tmp[df_tmp["topic_id"] == topic_id]
        center = model.cluster_centers_[topic_id]
        top_ids = center.argsort()[-8:][::-1]
        kws = [terms[i] for i in top_ids if terms[i].strip()]

        rep = sub.sort_values("content_len", ascending=False).iloc[0]
        item_name = f"观点簇{topic_id}"
        val = f"{len(sub)}篇；关键词：{' / '.join(kws[:6])}；代表：《{rep['title']}》"
        summary_rows.append({"item": item_name, "value": val})
        summary_rows.append({"item": f"观点簇{topic_id}链接", "value": rep["url"]})
        lines.append(f"{item_name}：{val}")
        lines.append(f"链接：{rep['url']}")

    return pd.DataFrame(summary_rows), lines


def build_output_excel_path(output_dir: Path, seed_keyword: str, run_time: datetime) -> Path:
    safe_keyword = re.sub(r"[\\/:*?\"<>|\s]+", "_", (seed_keyword or "关键词").strip())
    safe_keyword = safe_keyword.strip("_") or "关键词"
    date_str = run_time.strftime("%Y%m%d")
    file_name = f"微信公众号抓取_{safe_keyword}_{date_str}.xlsx"
    output_path = output_dir / file_name

    if output_path.exists():
        time_str = run_time.strftime("%H%M%S")
        output_path = output_dir / f"微信公众号抓取_{safe_keyword}_{date_str}_{time_str}.xlsx"

    return output_path


def _build_takeaway(content: str, max_chars: int = 120) -> str:
    text = re.sub(r"\s+", " ", (content or "")).strip()
    if not text:
        return "正文抓取为空，暂无法提炼。"

    parts = [p.strip() for p in re.split(r"[。！？!?；;\n]", text) if p.strip()]
    parts = [p for p in parts if len(p) >= 10]
    if not parts:
        parts = [text]

    takeaway = "；".join(parts[:2]).strip("；")
    if len(takeaway) > max_chars:
        takeaway = takeaway[:max_chars].rstrip() + "..."
    return takeaway


def write_outputs(df_all: pd.DataFrame, df_focus: pd.DataFrame, summary_df: pd.DataFrame, lines: List[str], excel_path: Path) -> None:

    # Sheet1：按用户指定表头输出
    source_df = df_focus.copy()
    col_map = {
        "publish_time": "发布时间",
        "account": "发布人",
        "title": "发布标题",
        "url": "发布连接",
        "read_count": "阅读量",
    }
    for c in col_map:
        if c not in source_df.columns:
            source_df[c] = ""

    sheet1 = source_df[list(col_map.keys())].rename(columns=col_map)
    parsed_time = pd.to_datetime(sheet1["发布时间"], errors="coerce")
    parsed_read = pd.to_numeric(sheet1["阅读量"], errors="coerce")
    sheet1 = (
        sheet1.assign(_ts=parsed_time, _read=parsed_read)
        .sort_values(["_read", "_ts"], ascending=[False, False])
        .drop(columns=["_ts", "_read"])
        .reset_index(drop=True)
    )

    # Sheet2：逐篇 takeaways（仅输出符合条件文章）
    takeaways_rows = []
    if not df_focus.empty:
        tmp = df_focus.copy()
        tmp["_read"] = pd.to_numeric(tmp.get("read_count"), errors="coerce")
        tmp["_ts"] = pd.to_datetime(tmp.get("publish_time"), errors="coerce")
        tmp = tmp.sort_values(["_read", "_ts"], ascending=[False, False])

        for _, r in tmp.iterrows():
            read_num = pd.to_numeric(pd.Series([r.get("read_count")]), errors="coerce").iloc[0]
            if pd.isna(read_num):
                read_val = ""
            else:
                read_val = int(read_num)

            takeaways_rows.append({
                "发布人": str(r.get("account", "") or "未知账号"),
                "发布标题": str(r.get("title", "") or "（无标题）"),
                "阅读量": read_val,
                "蓝标认证": "是" if bool(r.get("is_blue_verified", False)) else "否",
                "takeaways": _build_takeaway(str(r.get("content", ""))),
                "发布连接": str(r.get("url", "") or ""),
            })

    if not takeaways_rows:
        fallback = "未找到符合条件（蓝标认证或阅读量>500）的文章。"
        if lines:
            fallback = f"{fallback} {lines[0]}"
        takeaways_rows = [{
            "发布人": "",
            "发布标题": "",
            "阅读量": "",
            "蓝标认证": "",
            "takeaways": fallback,
            "发布连接": "",
        }]

    sheet2 = pd.DataFrame(takeaways_rows)

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        sheet1.to_excel(writer, sheet_name="抓取文章汇总", index=False)
        sheet2.to_excel(writer, sheet_name="核心信息提取", index=False)


def main() -> None:
    output_dir = Path(OUTPUT_DIR).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    run_time = datetime.now()
    excel_path = build_output_excel_path(output_dir, SEED_KEYWORD, run_time)

    run_log = {
        "seed_keyword": SEED_KEYWORD,
        "expanded_keywords": [],
        "searched_urls": {},
        "total_urls_before_cap": 0,
        "total_urls": 0,
        "success_count": 0,
        "failed_count": 0,
        "focus_count": 0,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    keywords = expand_keywords(SEED_KEYWORD)
    run_log["expanded_keywords"] = keywords

    print(f"[INFO] 输入关键词：{SEED_KEYWORD}")
    print(f"[INFO] 扩展关键词：{keywords}")

    all_pairs: List[Tuple[str, str]] = []
    for kw in keywords:
        urls = search_wechat_urls(kw, pages=max(1, PAGES_PER_KEYWORD))
        run_log["searched_urls"][kw] = len(urls)
        print(f"[INFO] 关键词「{kw}」检索到链接：{len(urls)}")
        all_pairs.extend([(kw, u) for u in urls])

    seen = set()
    unique_pairs = []
    for kw, u in all_pairs:
        if u not in seen:
            seen.add(u)
            unique_pairs.append((kw, u))

    run_log["total_urls_before_cap"] = len(unique_pairs)
    if len(unique_pairs) > max(10, MAX_URLS):
        unique_pairs = unique_pairs[:max(10, MAX_URLS)]
    run_log["total_urls"] = len(unique_pairs)

    print(f"[INFO] 去重后候选链接数：{run_log['total_urls_before_cap']}，本次抓取上限：{len(unique_pairs)}")

    rows: List[Article] = []
    for idx, (kw, u) in enumerate(unique_pairs, start=1):
        rows.append(crawl_article(u, kw))
        if idx % 5 == 0 or idx == len(unique_pairs):
            print(f"[INFO] 抓取进度：{idx}/{len(unique_pairs)}")
        time.sleep(random.uniform(0.8, 1.8))

    df_all = pd.DataFrame([
        {
            "keyword": r.keyword,
            "title": r.title,
            "account": r.account,
            "publish_time": r.publish_time.strftime("%Y-%m-%d %H:%M:%S") if r.publish_time else "",
            "url": r.url,
            "content": r.content,
            "content_len": r.content_len,
            "read_count": r.read_count,
            "is_blue_verified": r.is_blue_verified,
            "collect_status": r.collect_status,
            "error_reason": r.error_reason,
        }
        for r in rows
    ])

    if df_all.empty:
        summary_df = pd.DataFrame([{"item": "总体结论", "value": "未抓取到任何候选链接，请增加 pages 或更换关键词。"}])
        write_outputs(df_all, pd.DataFrame(), summary_df, ["未抓取到任何候选链接。"], excel_path)
        return

    ok_df = df_all[df_all["collect_status"] == "ok"].copy()
    run_log["success_count"] = int(len(ok_df))
    run_log["failed_count"] = int((df_all["collect_status"] != "ok").sum())

    now = datetime.now()
    cutoff = now - timedelta(days=LOOKBACK_DAYS)

    def _parse_time(s: str) -> Optional[datetime]:
        s = (s or "").strip()
        if not s:
            return None
        try:
            return datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
        except Exception:
            return None

    ok_df["_ptime"] = ok_df["publish_time"].apply(_parse_time)
    recent_df = ok_df[ok_df["_ptime"].notna() & (ok_df["_ptime"] >= cutoff)].copy()

    recent_df["_read"] = pd.to_numeric(recent_df.get("read_count"), errors="coerce")
    if "is_blue_verified" in recent_df.columns:
        recent_df["_is_blue"] = recent_df["is_blue_verified"].fillna(False).astype(bool)
    else:
        recent_df["_is_blue"] = False

    focus_df = recent_df[(recent_df["_is_blue"]) | (recent_df["_read"] > READ_COUNT_THRESHOLD)].copy()
    focus_reason = f"仅保留近{LOOKBACK_DAYS}天文章，且满足蓝标认证或阅读量>{READ_COUNT_THRESHOLD}"

    for helper_col in ["_ptime", "_read", "_is_blue"]:
        if helper_col in focus_df.columns:
            focus_df = focus_df.drop(columns=[helper_col])

    run_log["focus_count"] = int(len(focus_df))

    summary_df, lines = build_summary(focus_df)

    read_available = int(pd.to_numeric(focus_df.get("read_count"), errors="coerce").notna().sum()) if not focus_df.empty else 0
    read_coverage = f"{read_available}/{len(focus_df)}"
    if not focus_df.empty and "is_blue_verified" in focus_df.columns:
        blue_count = int(focus_df["is_blue_verified"].fillna(False).sum())
    else:
        blue_count = 0

    summary_df = pd.concat([
        pd.DataFrame([
            {"item": "分析口径", "value": focus_reason},
            {"item": "阅读量可得率", "value": read_coverage},
            {"item": "蓝标账号样本数", "value": str(blue_count)},
        ]),
        summary_df,
    ], ignore_index=True)

    write_outputs(df_all, focus_df, summary_df, lines, excel_path)

    print(f"[INFO] 抓取成功：{run_log['success_count']}，失败：{run_log['failed_count']}")
    print(f"[INFO] 符合条件样本（蓝标认证或阅读量>{READ_COUNT_THRESHOLD}）：{run_log['focus_count']}")
    print(f"[INFO] 单文件输出：{excel_path}")


if __name__ == "__main__":
    main()
