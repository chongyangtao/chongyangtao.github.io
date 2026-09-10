from __future__ import annotations

import argparse
import html
import json
import re
from pathlib import Path
from urllib.parse import urlparse

from bs4 import BeautifulSoup
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
ALLOWED_TAGS = {
    "reward-model",
    "context-engineering",
    "multi-party-conversation",
    "generation",
    "information-retrieval",
}
HEADERS = [
    "id", "section", "year", "month", "sort_order", "title", "authors", "authors_html",
    "venue", "venue_html", "paper_url", "tags", "code_url", "code_label", "award", "enabled",
]


def clean(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def section_value(section: str) -> tuple[str, int]:
    match = re.search(r"\d{4}", section)
    return section, int(match.group()) if match else 9999


def extract() -> list[dict]:
    soup = BeautifulSoup((ROOT / "publications.html").read_text(encoding="utf-8"), "html.parser")
    records = []
    order = 0
    for group in soup.select(".publication-year"):
        heading = group.find(["strong", "h2"])
        section = clean(heading.get_text(" ", strip=True)) if heading else "Unsorted"
        year_match = re.search(r"\d{4}", section)
        year = int(year_match.group()) if year_match else ""
        for item in group.select(":scope > .publication-list > .publication-item"):
            title = item.select_one(".publication-title a")
            authors_node = item.select_one(".publication-authors")
            venue_node = item.select_one(".publication-venue")
            links = item.select(".publication-links a")
            if not title or not authors_node or not venue_node:
                continue
            order += 1
            title_text = clean(title.get_text(" ", strip=True))
            paper_url = title.get("href", "").strip()
            authors_html = "".join(str(x) for x in authors_node.contents).strip()
            venue_html = "".join(str(x) for x in venue_node.contents).strip()
            authors = clean(authors_node.get_text(" ", strip=True))
            venue = clean(venue_node.get_text(" ", strip=True))
            code_url = links[0].get("href", "").strip() if links else ""
            code_label = clean(links[0].get_text(" ", strip=True)) if links else ""
            tags = item.get("data-tags", "")
            record = {
                "id": f"pub-{year or 'preprint'}-{order:03d}",
                "section": section,
                "year": year,
                "month": "",
                "sort_order": order,
                "title": title_text,
                "authors": authors,
                "authors_html": authors_html,
                "venue": venue,
                "venue_html": venue_html,
                "paper_url": paper_url,
                "tags": tags,
                "code_url": code_url,
                "code_label": code_label,
                "award": "",
                "enabled": True,
            }
            records.append(record)
    return records


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract publication records into an Excel workbook.")
    parser.add_argument("--output", type=Path, default=ROOT / "data/publications.xlsx")
    args = parser.parse_args()
    records = extract()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Publications"
    sheet.append(HEADERS)
    for record in records:
        sheet.append([record[key] for key in HEADERS])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column in sheet.columns:
        width = min(max(max(len(str(cell.value or "")) for cell in column) + 2, 12), 60)
        sheet.column_dimensions[column[0].column_letter].width = width
    args.output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(args.output)
    print(f"Wrote {len(records)} publications to {args.output}")


if __name__ == "__main__":
    main()
