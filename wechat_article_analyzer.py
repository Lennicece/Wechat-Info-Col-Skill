# -*- coding: utf-8 -*-
"""
微信公众号近7天文章抓取 + 观点总结（可交付版）

输出：
1) wechat_articles_report.xlsx
   - articles_last7d: 近7天文章（或兜底可抓取文章）
   - raw_collected: 原始抓取结果（含状态）
   - summary: 自动观点总结
2) wechat_summary.txt: 文字版总结
3) run_log.json: 运行日志（便于排错）

说明：
- 公开网页抓取受反爬影响，不保证 100% 成功。
- 本脚本会输出“抓取状态”列，明确每条链接是否成功抓到正文。
"""

from __future__ import annotations

import argparse
import json
import random
import re
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Optional, Tuple
from urllib.parse import quote

import jieba
import pandas as pd
import requests
from bs4 import BeautifulSoup
from sklearn.cluster import KMeans
from sklearn.feature_extraction.text import TfidfVectorizer


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
    collect_status: str
    error_reason: str


def expand_keywords(seed_keyword: str) -> List[str]:
    """关键词扩展：用户词 + 语义近邻（可自行增减）"""
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

    # 去重保序
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

        # 处理搜狗 /link 的 JS 拼接跳转
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
    """从搜狗微信搜索页提取文章链接（优先）"""
    urls: List[str] = []

    for page in range(1, pages + 1):
        search_url = f"https://weixin.sogou.com/weixin?type=2&query={quote(keyword)}&page={page}"
        try:
            resp = requests.get(search_url, headers=HEADERS, timeout=20)
            html = resp.text
        except Exception:
            continue

        # 反爬验证码
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

    # 去重保序
    seen = set()
    unique = []
    for u in urls:
        if u not in seen:
            seen.add(u)
            unique.append(u)
    return unique


def search_wechat_urls_bing(keyword: str, pages: int = 2) -> List[str]:
    """从 Bing 检索微信公众号文章链接（兜底）"""
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
    """先搜狗，若为空再用 Bing 兜底"""
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


def crawl_article(url: str, keyword: str) -> Article:
    headers = dict(HEADERS)
    headers["Referer"] = "https://mp.weixin.qq.com/"

    try:
        resp = requests.get(url, headers=headers, timeout=20)
        # 微信文章页通常为 UTF-8，强制指定可避免乱码
        resp.encoding = "utf-8"
        html = resp.text
    except Exception as e:
        return Article(
            keyword=keyword,
            title="",
            account="",
            publish_time=None,
            url=url,
            content="",
            content_len=0,
            collect_status="failed",
            error_reason=f"request_error: {e}",
        )

    soup = BeautifulSoup(html, "lxml")
    title = parse_title(soup)
    account = parse_account(soup, html)
    publish_time = parse_publish_time(html)
    content = parse_content(soup)

    if len(content) < 20 and not title:
        return Article(
            keyword=keyword,
            title=title,
            account=account,
            publish_time=publish_time,
            url=url,
            content=content,
            content_len=len(content),
            collect_status="failed",
            error_reason="content_too_short",
        )

    return Article(
        keyword=keyword,
        title=title,
        account=account,
        publish_time=publish_time,
        url=url,
        content=content,
        content_len=len(content),
        collect_status="ok",
        error_reason="",
    )


def build_summary(df_focus: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """生成观点摘要（簇 + 关键词 + 代表文章）"""
    if df_focus.empty:
        rows = [
            {"item": "总体结论", "value": "未形成可分析样本，建议增加关键词扩展或提高抓取页数。"},
        ]
        return pd.DataFrame(rows), ["未形成可分析样本。"]

    docs = df_focus["content"].fillna("").tolist()
    docs = [d if len(d) <= 3000 else d[:3000] for d in docs]

    if len(docs) == 1:
        one = df_focus.iloc[0]
        lines = [
            "近7天仅抓到1篇可用文章，暂无稳定趋势。",
            f"代表文章：《{one['title']}》 {one['url']}",
        ]
        rows = [
            {"item": "总体结论", "value": lines[0]},
            {"item": "代表文章", "value": lines[1]},
        ]
        return pd.DataFrame(rows), lines

    vectorizer = TfidfVectorizer(
        tokenizer=tokenize_zh,
        token_pattern=None,
        max_features=2000,
        min_df=1,
    )
    X = vectorizer.fit_transform(docs)

    k = min(3, len(docs))
    model = KMeans(n_clusters=k, random_state=42, n_init="auto")
    labels = model.fit_predict(X)

    df_tmp = df_focus.copy()
    df_tmp["topic_id"] = labels

    terms = vectorizer.get_feature_names_out()

    summary_rows = []
    lines = []

    summary_rows.append({"item": "样本量", "value": f"{len(df_focus)} 篇（用于观点分析）"})
    summary_rows.append({"item": "账号数", "value": f"{df_focus['account'].replace('', '未知账号').nunique()}"})

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


def main(seed_keyword: str, pages_per_keyword: int, output_dir: Path, max_urls: int) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    run_log = {
        "seed_keyword": seed_keyword,
        "expanded_keywords": [],
        "searched_urls": {},
        "total_urls_before_cap": 0,
        "total_urls": 0,
        "success_count": 0,
        "failed_count": 0,
        "focus_count": 0,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    keywords = expand_keywords(seed_keyword)
    run_log["expanded_keywords"] = keywords

    print(f"[INFO] 输入关键词：{seed_keyword}")
    print(f"[INFO] 扩展关键词：{keywords}")

    all_pairs: List[Tuple[str, str]] = []
    for kw in keywords:
        urls = search_wechat_urls(kw, pages=pages_per_keyword)
        run_log["searched_urls"][kw] = len(urls)
        print(f"[INFO] 关键词「{kw}」检索到链接：{len(urls)}")
        all_pairs.extend([(kw, u) for u in urls])

    # 去重 URL
    seen = set()
    unique_pairs = []
    for kw, u in all_pairs:
        if u not in seen:
            seen.add(u)
            unique_pairs.append((kw, u))

    run_log["total_urls_before_cap"] = len(unique_pairs)
    if len(unique_pairs) > max_urls:
        unique_pairs = unique_pairs[:max_urls]
    run_log["total_urls"] = len(unique_pairs)
    print(f"[INFO] 去重后候选链接数：{run_log['total_urls_before_cap']}，本次抓取上限：{len(unique_pairs)}")

    rows: List[Article] = []
    for idx, (kw, u) in enumerate(unique_pairs, start=1):
        art = crawl_article(u, kw)
        rows.append(art)
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
            "collect_status": r.collect_status,
            "error_reason": r.error_reason,
        }
        for r in rows
    ])

    if df_all.empty:
        # 空跑也给交付文件
        summary_df = pd.DataFrame([
            {"item": "总体结论", "value": "未抓取到任何候选链接，请增加 pages 或更换关键词。"}
        ])
        write_outputs(df_all, pd.DataFrame(), summary_df, ["未抓取到任何候选链接。"], output_dir)
        run_log["success_count"] = 0
        run_log["failed_count"] = 0
        run_log["focus_count"] = 0
        (output_dir / "run_log.json").write_text(json.dumps(run_log, ensure_ascii=False, indent=2), encoding="utf-8")
        return

    ok_df = df_all[df_all["collect_status"] == "ok"].copy()
    run_log["success_count"] = int(len(ok_df))
    run_log["failed_count"] = int((df_all["collect_status"] != "ok").sum())

    # 近7天过滤（如果时间解析失败，则兜底用全部ok）
    now = datetime.now()
    cutoff = now - timedelta(days=7)

    def _parse_time(s: str) -> Optional[datetime]:
        s = (s or "").strip()
        if not s:
            return None
        try:
            return datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
        except Exception:
            return None

    ok_df["_ptime"] = ok_df["publish_time"].apply(_parse_time)
    last7 = ok_df[ok_df["_ptime"].notna() & (ok_df["_ptime"] >= cutoff)].copy()

    if last7.empty:
        # 没法确认近7天时，退回全部 ok，避免“明明抓到了却分析为空”
        focus_df = ok_df.copy()
        focus_reason = "近7天时间解析为空，已使用全部可抓取文章进行分析"
    else:
        focus_df = last7.copy()
        focus_reason = "使用近7天可解析发布时间文章进行分析"

    if "_ptime" in focus_df.columns:
        focus_df = focus_df.drop(columns=["_ptime"])
    if "_ptime" in ok_df.columns:
        ok_df = ok_df.drop(columns=["_ptime"])

    run_log["focus_count"] = int(len(focus_df))

    summary_df, lines = build_summary(focus_df)
    summary_df = pd.concat([
        pd.DataFrame([{"item": "分析口径", "value": focus_reason}]),
        summary_df,
    ], ignore_index=True)

    write_outputs(df_all, focus_df, summary_df, lines, output_dir)

    (output_dir / "run_log.json").write_text(json.dumps(run_log, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[INFO] 抓取成功：{run_log['success_count']}，失败：{run_log['failed_count']}")
    print(f"[INFO] 用于观点分析样本：{run_log['focus_count']}")
    print(f"[INFO] Excel 输出：{output_dir / 'wechat_articles_report.xlsx'}")
    print(f"[INFO] 总结输出：{output_dir / 'wechat_summary.txt'}")


def write_outputs(
    df_all: pd.DataFrame,
    df_focus: pd.DataFrame,
    summary_df: pd.DataFrame,
    lines: List[str],
    output_dir: Path,
) -> None:
    excel_path = output_dir / "wechat_articles_report.xlsx"
    txt_path = output_dir / "wechat_summary.txt"

    # 列顺序统一
    base_cols = [
        "keyword",
        "title",
        "account",
        "publish_time",
        "url",
        "content_len",
        "collect_status",
        "error_reason",
        "content",
    ]

    for df in (df_all, df_focus):
        for c in base_cols:
            if c not in df.columns:
                df[c] = ""

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_focus[base_cols].to_excel(writer, sheet_name="articles_last7d", index=False)
        df_all[base_cols].to_excel(writer, sheet_name="raw_collected", index=False)
        summary_df.to_excel(writer, sheet_name="summary", index=False)

    txt_lines = [
        "微信公众号观点总结（自动生成）",
        f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
    ]
    txt_lines.extend(lines if lines else ["无可用样本。"])

    txt_path.write_text("\n".join(txt_lines), encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--keyword", required=True, help="种子关键词，例如：漫剧工业化")
    parser.add_argument("--pages", type=int, default=3, help="每个关键词检索页数，默认3")
    parser.add_argument(
        "--output-dir",
        default=".",
        help="输出目录，默认当前目录",
    )
    parser.add_argument("--max-urls", type=int, default=60, help="本次最多抓取多少篇候选链接，默认60")
    args = parser.parse_args()

    main(
        seed_keyword=args.keyword,
        pages_per_keyword=max(1, args.pages),
        output_dir=Path(args.output_dir).resolve(),
        max_urls=max(10, args.max_urls),
    )
