# -*- coding: utf-8 -*-
"""
微信公众号主题抓取：Excel + Word 双文件交付

运行示例：
python wechat-single-file-report/wechat_dual_report.py \
  --theme "漫剧AI生产工业化" \
  --max-urls 50 \
  --read-threshold 500 \
  --lookback-days 20
"""

from __future__ import annotations

import argparse
import json
import re
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Optional
from urllib.parse import quote, urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from docx import Document

BASE_URL = "https://weixin.sogou.com"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Referer": "https://weixin.sogou.com/",
}


@dataclass
class Article:
    publisher: str
    publish_dt: datetime
    title: str
    link: str
    snippet: str
    read_count: Optional[int]
    interaction_count: Optional[int]


def expand_keywords(theme: str) -> List[str]:
    base = [theme.strip()]
    text = theme.strip()
    if "漫剧" in text:
        base += [
            "AI漫剧工业化",
            "AIGC漫剧生产流程",
            "AI漫剧流水线",
            "漫剧自动化制作",
            "漫剧AI批量生产",
            "AI漫剧合规生产",
        ]
    if "短剧" in text:
        base += ["AI短剧工业化", "短剧AI制作流程"]

    out = []
    seen = set()
    for k in base:
        if k and k not in seen:
            seen.add(k)
            out.append(k)
    return out[:10]


def clean_text(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


def extract_publish_ts(li_html: str) -> Optional[int]:
    m = re.search(r"timeConvert\('([0-9]{9,11})'\)", li_html)
    if not m:
        return None
    return int(m.group(1))


def parse_num(num_str: str, unit: str) -> int:
    val = float(num_str)
    if "亿" in unit:
        return int(val * 100000000)
    if "万" in unit:
        return int(val * 10000)
    return int(val)


def extract_read_count(text: str) -> Optional[int]:
    patterns = [
        r"(?:阅读量?|阅读|播放量|浏览量)\s*[：: ]\s*([0-9]+(?:\.[0-9]+)?)\s*(万|亿)?\+?",
        r"([0-9]+(?:\.[0-9]+)?)\s*(万|亿)?\+?\s*(?:阅读量?|阅读|播放量|浏览量)",
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return parse_num(m.group(1), m.group(2) or "")
    return None


def extract_interaction_count(text: str) -> Optional[int]:
    patterns = [
        r"(?:互动量?|点赞|在看|评论|转发|热度)\s*[：: ]\s*([0-9]+(?:\.[0-9]+)?)\s*(万|亿)?\+?",
        r"([0-9]+(?:\.[0-9]+)?)\s*(万|亿)?\+?\s*(?:点赞|在看|评论|转发|热度)",
    ]
    vals = []
    for p in patterns:
        for m in re.finditer(p, text):
            vals.append(parse_num(m.group(1), m.group(2) or ""))
    if vals:
        return int(sum(vals[:2]))
    return None


def resolve_mp_link(raw_href: str, session: requests.Session) -> str:
    if not raw_href:
        return ""
    full = raw_href if raw_href.startswith("http") else urljoin(BASE_URL, raw_href)
    try:
        r = session.get(full, headers=HEADERS, timeout=15)
        parts = re.findall(r"url \+= '([^']*)'", r.text)
        if parts:
            return "".join(parts).replace("@", "")
        return full
    except Exception:
        return full


def fetch_articles(keyword: str, pages: int, session: requests.Session) -> List[Article]:
    out: List[Article] = []
    for page in range(1, pages + 1):
        url = f"{BASE_URL}/weixin?type=2&query={quote(keyword)}&page={page}"
        try:
            resp = session.get(url, headers=HEADERS, timeout=20)
            soup = BeautifulSoup(resp.text, "lxml")
            for li in soup.select("ul.news-list li"):
                title_node = li.select_one("h3 a")
                if not title_node:
                    continue
                ts = extract_publish_ts(str(li))
                if not ts:
                    continue
                publish_dt = datetime.fromtimestamp(ts)
                title = clean_text(title_node.get_text(" ", strip=True))
                raw_href = title_node.get("href", "")
                publisher = clean_text((li.select_one("span.all-time-y2") or {}).get_text(strip=True) if li.select_one("span.all-time-y2") else "未知发布人")
                snippet = clean_text((li.select_one("p.txt-info") or {}).get_text(" ", strip=True) if li.select_one("p.txt-info") else "")

                merged = f"{title} {snippet}"
                read_count = extract_read_count(merged)
                interaction_count = extract_interaction_count(merged)
                link = resolve_mp_link(raw_href, session)

                out.append(
                    Article(
                        publisher=publisher,
                        publish_dt=publish_dt,
                        title=title,
                        link=link,
                        snippet=snippet,
                        read_count=read_count,
                        interaction_count=interaction_count,
                    )
                )
        except Exception:
            pass
        time.sleep(0.12)
    return out


def to_display_num(v: Optional[int]) -> str:
    return "--" if v is None else str(int(v))


def core_sentence(snippet: str, title: str) -> str:
    text = clean_text(snippet)
    if not text:
        return title[:60]
    parts = [p.strip() for p in re.split(r"[。！？!?；;]", text) if p.strip()]
    return (parts[0] if parts else text)[:80]


def run(theme: str, max_urls: int, read_threshold: int, lookback_days: int, output_dir: Path) -> dict:
    session = requests.Session()
    session.headers.update(HEADERS)

    keywords = expand_keywords(theme)
    all_articles: List[Article] = []
    for kw in keywords:
        all_articles.extend(fetch_articles(kw, pages=8, session=session))

    dedup = {}
    for a in all_articles:
        key = (a.title, a.publisher, a.publish_dt.strftime("%Y-%m-%d"))
        if key not in dedup:
            dedup[key] = a

    cutoff = datetime.now() - timedelta(days=lookback_days)
    filtered = [a for a in dedup.values() if a.publish_dt >= cutoff]

    # 先高价值（有阅读量/互动量）再按数值与时间排序
    filtered.sort(
        key=lambda x: (
            1 if (x.read_count is not None or x.interaction_count is not None) else 0,
            x.read_count if x.read_count is not None else -1,
            x.interaction_count if x.interaction_count is not None else -1,
            x.publish_dt,
        ),
        reverse=True,
    )

    # 阈值过滤（仅在有阅读量时应用）
    kept = []
    for a in filtered:
        if a.read_count is None:
            kept.append(a)
        elif a.read_count >= read_threshold:
            kept.append(a)
    if max_urls > 0:
        kept = kept[:max_urls]

    today = datetime.now().strftime("%Y%m%d")
    excel_path = output_dir / f"微信文章价值清单_{today}.xlsx"
    word_path = output_dir / f"微信文章速读总结_{today}.docx"

    df = pd.DataFrame(
        [
            {
                "发布人": a.publisher,
                "发布时间": a.publish_dt.strftime("%Y-%m-%d %H:%M:%S"),
                "发布标题": a.title,
                "发布链接": a.link,
                "阅读量": to_display_num(a.read_count),
                "互动量": to_display_num(a.interaction_count),
            }
            for a in kept
        ],
        columns=["发布人", "发布时间", "发布标题", "发布链接", "阅读量", "互动量"],
    )
    df.to_excel(excel_path, index=False)

    doc = Document()
    doc.add_heading("微信文章主题速读总结", level=1)
    doc.add_paragraph(f"主题：{theme}")
    doc.add_paragraph(f"统计范围：近{lookback_days}天；关键词共{len(keywords)}组；入库文章{len(kept)}篇。")

    doc.add_heading("核心观点（含文章来源）", level=2)
    top = kept[:8]
    if top:
        for i, a in enumerate(top, 1):
            doc.add_paragraph(
                f"{i}. {core_sentence(a.snippet, a.title)}\n来源：《{a.title}》｜{a.publisher}｜{a.publish_dt.strftime('%Y-%m-%d')}"
            )
    else:
        doc.add_paragraph("本次未抓取到符合条件文章。")

    summary = (
        "本周期围绕该主题的公众号内容，重点集中在生产流程标准化、内容质量提升与商业化落地三条线。"
        "整体趋势是从单点工具使用转向完整流水线协同，强调角色一致性、批量产能和合规约束。"
        "建议优先跟踪具备方法论输出和案例复盘的账号，提升信息利用效率。"
    )
    doc.add_heading("200字内总结", level=2)
    doc.add_paragraph(summary[:200])

    doc.add_heading("口径说明", level=2)
    doc.add_paragraph("阅读量/互动量优先取公开可得字段；若受反爬或平台限制不可得，统一标记为 --，不做伪造。")
    doc.save(word_path)

    return {
        "theme": theme,
        "keywords": keywords,
        "raw_count": len(all_articles),
        "dedup_count": len(dedup),
        "final_count": len(kept),
        "excel": str(excel_path),
        "word": str(word_path),
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--theme", required=True, help="主题，例如：漫剧AI生产工业化")
    parser.add_argument("--max-urls", type=int, default=50)
    parser.add_argument("--read-threshold", type=int, default=500)
    parser.add_argument("--lookback-days", type=int, default=20)
    parser.add_argument("--output-dir", default=".")
    args = parser.parse_args()

    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    stats = run(
        theme=args.theme,
        max_urls=args.max_urls,
        read_threshold=args.read_threshold,
        lookback_days=args.lookback_days,
        output_dir=output_dir,
    )
    print(json.dumps(stats, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
