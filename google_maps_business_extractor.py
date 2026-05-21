#!/usr/bin/env python3
"""
Standalone Google Maps business extractor.

Uses Playwright to search Google Maps, collect listing links from the results
feed, visit each listing, filter unusable rows, and export CSV, JSON, and Excel.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import time
import zipfile
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from urllib.parse import parse_qsl, quote_plus, urlencode, urlsplit, urlunsplit
from xml.sax.saxutils import escape

try:
    from playwright.sync_api import Browser, Page, TimeoutError as PlaywrightTimeoutError
    from playwright.sync_api import sync_playwright
except ImportError:  # pragma: no cover - gives users a clean setup message.
    Browser = None
    Page = None
    PlaywrightTimeoutError = TimeoutError
    sync_playwright = None


SOCIAL_DOMAINS = {
    "instagram.com",
    "facebook.com",
    "fb.com",
    "m.facebook.com",
    "linktr.ee",
    "linktree.com",
    "wa.me",
    "api.whatsapp.com",
    "tiktok.com",
    "youtube.com",
    "youtu.be",
    "twitter.com",
    "x.com",
    "linkedin.com",
    "t.me",
    "telegram.me",
    "threads.net",
}

OUTPUT_COLUMNS = [
    "business_name",
    "address",
    "phone",
    "website",
    "rating",
    "review_count",
    "category",
    "google_maps_link",
]

REQUIRED_FIELDS = [
    "business_name",
    "address",
    "phone",
    "website",
    "rating",
    "review_count",
    "category",
    "google_maps_link",
]


@dataclass(frozen=True)
class Business:
    business_name: str
    address: str
    phone: str
    website: str
    rating: float | None
    review_count: int | None
    category: str
    google_maps_link: str


@dataclass(frozen=True)
class ResultLink:
    name: str
    url: str


def normalize_website(url: str | None) -> str | None:
    """Return normalized domain ('example.com') or None if unusable."""
    if not url or not url.strip():
        return None
    normalized = url.strip().lower()
    normalized = re.sub(r"^https?://", "", normalized)
    normalized = re.sub(r"^www\.", "", normalized)
    normalized = normalized.split("/")[0].split("?")[0].split("#")[0]
    normalized = normalized.strip()
    return normalized or None


def is_social_only(domain: str | None) -> bool:
    """Check if a normalized domain is just a social/profile link."""
    if not domain:
        return True
    return any(domain == social or domain.endswith("." + social) for social in SOCIAL_DOMAINS)


def slugify_query(query: str) -> str:
    slug = re.sub(r"[^a-zA-Z0-9]+", "_", query.strip().lower()).strip("_")
    return slug or "google_maps_results"


def clean_label(label: str | None, prefix: str) -> str:
    if not label:
        return ""
    cleaned = re.sub(rf"^{re.escape(prefix)}\s*:?\s*", "", label.strip(), flags=re.I)
    if cleaned == label.strip() and ":" in cleaned:
        label_prefix, value = cleaned.split(":", 1)
        if len(label_prefix) <= 24:
            cleaned = value
    return cleaned.strip()


def force_english_maps_url(url: str) -> str:
    parts = urlsplit(url)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query["hl"] = "en"
    query["gl"] = "us"
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment))


def parse_float(value: str | None) -> float | None:
    if not value:
        return None
    match = re.search(r"\d+(?:[.,]\d+)?", value)
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", "."))
    except ValueError:
        return None


def parse_int(value: str | None) -> int | None:
    if not value:
        return None
    match = re.search(r"[\d,.]+", value)
    if not match:
        return None
    cleaned = re.sub(r"[^\d]", "", match.group(0))
    if not cleaned:
        return None
    try:
        return int(cleaned)
    except ValueError:
        return None


def require_playwright() -> None:
    if sync_playwright is None:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install -r requirements.txt "
            "and then: playwright install chromium"
        )


def accept_consent_if_present(page: Page) -> None:
    selectors = [
        "button:has-text('Accept all')",
        "button:has-text('I agree')",
        "button:has-text('Accept')",
        "form[action*='consent'] button",
    ]
    for selector in selectors:
        try:
            button = page.locator(selector).first
            if button.count() > 0 and button.is_visible(timeout=1000):
                button.click(timeout=2000)
                page.wait_for_timeout(1500)
                return
        except PlaywrightTimeoutError:
            continue


def search_maps(page: Page, query: str) -> None:
    url = f"https://www.google.com/maps/search/{quote_plus(query)}?hl=en&gl=us"
    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    accept_consent_if_present(page)
    try:
        page.wait_for_selector("div[role='feed'], h1", timeout=20000)
    except PlaywrightTimeoutError:
        page.wait_for_timeout(3000)


def collect_result_links(page: Page, max_results: int, scroll_rounds: int) -> list[ResultLink]:
    links: dict[str, ResultLink] = {}
    stable_rounds = 0
    previous_count = 0

    for _ in range(scroll_rounds):
        listing_links = page.locator("a[href*='/maps/place/']")
        count = listing_links.count()

        for index in range(count):
            anchor = listing_links.nth(index)
            try:
                href = anchor.get_attribute("href", timeout=1000)
                label = anchor.get_attribute("aria-label", timeout=1000) or ""
                title = anchor.get_attribute("title", timeout=1000) or ""
                text = anchor.inner_text(timeout=1000) or ""
            except PlaywrightTimeoutError:
                continue

            name = (label or title or text).strip()
            if not href or not name or "/maps/place/" not in href:
                continue
            if name.lower() in {"website", "directions", "save", "share"}:
                continue
            links.setdefault(href, ResultLink(name=name, url=href))
            if len(links) >= max_results:
                return list(links.values())

        if len(links) == previous_count:
            stable_rounds += 1
        else:
            stable_rounds = 0
        previous_count = len(links)
        if stable_rounds >= 4:
            break

        scroll_results_feed(page)
        page.wait_for_timeout(1200)

    return list(links.values())[:max_results]


def scroll_results_feed(page: Page) -> None:
    feed = page.locator("div[role='feed']").first
    try:
        if feed.count() > 0:
            feed.evaluate("(el) => { el.scrollTop = el.scrollHeight; }", timeout=3000)
            return
    except PlaywrightTimeoutError:
        pass
    page.mouse.wheel(0, 3500)


def text_from_first(page: Page, selectors: Iterable[str], timeout: int = 1200) -> str:
    for selector in selectors:
        locator = page.locator(selector).first
        try:
            if locator.count() > 0:
                text = locator.inner_text(timeout=timeout).strip()
                if text:
                    return text
        except PlaywrightTimeoutError:
            continue
    return ""


def attr_from_first(page: Page, selectors: Iterable[str], attr: str, timeout: int = 1200) -> str:
    for selector in selectors:
        locator = page.locator(selector).first
        try:
            if locator.count() > 0:
                value = locator.get_attribute(attr, timeout=timeout)
                if value and value.strip():
                    return value.strip()
        except PlaywrightTimeoutError:
            continue
    return ""


def aria_value(page: Page, selectors: Iterable[str], prefix: str) -> str:
    label = attr_from_first(page, selectors, "aria-label")
    return clean_label(label, prefix)


def extract_rating(page: Page) -> float | None:
    text = text_from_first(page, ["span.MW4etd", "div.F7nice span[aria-hidden='true']"])
    rating = parse_float(text)
    if rating is not None:
        return rating

    aria = attr_from_first(page, ["div.F7nice"], "aria-label")
    return parse_float(aria)


def extract_review_count(page: Page) -> int | None:
    text = text_from_first(page, ["span.UY7F9", "button[aria-label*='reviews']"])
    count = parse_int(text)
    if count is not None:
        return count

    aria = attr_from_first(page, ["button[aria-label*='reviews']", "span[aria-label*='reviews']"], "aria-label")
    return parse_int(aria)


def extract_business(page: Page, result: ResultLink) -> Business:
    page.goto(force_english_maps_url(result.url), wait_until="domcontentloaded", timeout=60000)
    try:
        page.wait_for_selector("h1, button[data-item-id='address']", timeout=15000)
    except PlaywrightTimeoutError:
        page.wait_for_timeout(2500)

    name = text_from_first(page, ["h1.DUwDvf", "h1"], timeout=3000) or result.name
    address = aria_value(
        page,
        ["button[data-item-id='address']", "button[aria-label^='Address:']"],
        "Address",
    )
    phone = aria_value(
        page,
        ["button[data-item-id^='phone:tel:']", "button[aria-label^='Phone:']"],
        "Phone",
    )
    website = attr_from_first(
        page,
        ["a[data-item-id='authority']", "a[aria-label^='Website:']", "a[href^='http']:has-text('Website')"],
        "href",
    )
    category = text_from_first(page, ["button.DkEaL", "button[jsaction*='category']", "div[jsaction*='category']"])
    rating = extract_rating(page)
    review_count = extract_review_count(page)

    return Business(
        business_name=name,
        address=address,
        phone=phone,
        website=website,
        rating=rating,
        review_count=review_count,
        category=category,
        google_maps_link=page.url,
    )


def scrape_google_maps(
    query: str,
    max_results: int,
    scroll_rounds: int,
    headless: bool,
    slow_mo: int,
) -> list[Business]:
    require_playwright()

    with sync_playwright() as playwright:
        browser: Browser = playwright.chromium.launch(headless=headless, slow_mo=slow_mo)
        context = browser.new_context(
            locale="en-US",
            timezone_id="America/New_York",
            viewport={"width": 1440, "height": 1100},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            extra_http_headers={"Accept-Language": "en-US,en;q=0.9"},
        )
        page = context.new_page()
        try:
            search_maps(page, query)
            result_links = collect_result_links(page, max_results=max_results, scroll_rounds=scroll_rounds)
            print(f"[extract] collected {len(result_links)} result links")

            businesses = []
            for index, result in enumerate(result_links, start=1):
                print(f"[extract] opening {index}/{len(result_links)}: {result.name}")
                try:
                    businesses.append(extract_business(page, result))
                except PlaywrightTimeoutError:
                    continue
                time.sleep(0.4)
            return businesses
        finally:
            context.close()
            browser.close()


def validate_and_filter(rows: Iterable[Business]) -> tuple[list[Business], Counter]:
    exported: list[Business] = []
    skipped: Counter = Counter()
    seen_domains: set[str] = set()

    for row in rows:
        data = asdict(row)
        missing_required = [field for field in REQUIRED_FIELDS if data.get(field) in ("", None)]
        if missing_required:
            skipped[f"missing_{missing_required[0]}"] += 1
            continue

        domain = normalize_website(row.website)
        if not domain:
            skipped["no_website"] += 1
            continue

        if is_social_only(domain):
            skipped["social_only"] += 1
            continue

        if domain in seen_domains:
            skipped["duplicate_website"] += 1
            continue

        seen_domains.add(domain)
        exported.append(row)

    return exported, skipped


def write_csv(rows: list[Business], path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=OUTPUT_COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow(asdict(row))


def write_json(rows: list[Business], path: Path) -> None:
    path.write_text(
        json.dumps([asdict(row) for row in rows], indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def write_excel(rows: list[Business], path: Path) -> None:
    workbook_rows = [OUTPUT_COLUMNS]
    workbook_rows.extend([[asdict(row)[column] for column in OUTPUT_COLUMNS] for row in rows])

    sheet_rows = []
    for row_index, row in enumerate(workbook_rows, start=1):
        cells = []
        for column_index, value in enumerate(row, start=1):
            ref = f"{excel_column(column_index)}{row_index}"
            style = ' s="1"' if row_index == 1 else ""
            if isinstance(value, (int, float)) and value is not None:
                cells.append(f'<c r="{ref}"{style}><v>{value}</v></c>')
            else:
                text = escape("" if value is None else str(value))
                cells.append(f'<c r="{ref}" t="inlineStr"{style}><is><t>{text}</t></is></c>')
        sheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    worksheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <sheetData>{"".join(sheet_rows)}</sheetData>
</worksheet>'''

    files = {
        "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>''',
        "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>''',
        "xl/workbook.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Businesses" sheetId="1" r:id="rId1"/></sheets>
</workbook>''',
        "xl/_rels/workbook.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>''',
        "xl/styles.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font/><font><b/><color rgb="FFFFFFFF"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F2937"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>''',
        "xl/worksheets/sheet1.xml": worksheet_xml,
    }

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for filename, content in files.items():
            archive.writestr(filename, content)


def excel_column(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def export_files(rows: list[Business], output_dir: Path, query: str) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    base = output_dir / slugify_query(query)
    paths = {
        "csv": base.with_suffix(".csv"),
        "json": base.with_suffix(".json"),
        "excel": base.with_suffix(".xlsx"),
    }
    write_csv(rows, paths["csv"])
    write_json(rows, paths["json"])
    write_excel(rows, paths["excel"])
    return paths


def print_summary(raw_count: int, exported_count: int, skipped: Counter, paths: dict[str, Path]) -> None:
    skipped_count = sum(skipped.values())
    print()
    print("=" * 58)
    print(f"Businesses found:    {raw_count}")
    print(f"Businesses exported: {exported_count}")
    print(f"Businesses skipped:  {skipped_count}")
    for reason, count in sorted(skipped.items()):
        print(f"  - {reason}: {count}")
    print()
    print("Output files:")
    for file_type, path in paths.items():
        print(f"  {file_type.upper()}: {path}")
    print("=" * 58)


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract Google Maps businesses to CSV, JSON, and Excel.")
    parser.add_argument("query", help='Search query, e.g. "dentists London"')
    parser.add_argument("--max-results", type=int, default=40, help="Maximum listing pages to open (default: 40)")
    parser.add_argument("--scroll-rounds", type=int, default=18, help="Result-list scroll attempts (default: 18)")
    parser.add_argument("--output-dir", default="output", help="Directory for exported files (default: output)")
    parser.add_argument("--headed", action="store_true", help="Run browser visibly instead of headless")
    parser.add_argument("--slow-mo", type=int, default=0, help="Slow Playwright actions by N milliseconds")
    args = parser.parse_args()

    raw_rows = scrape_google_maps(
        query=args.query,
        max_results=args.max_results,
        scroll_rounds=args.scroll_rounds,
        headless=not args.headed,
        slow_mo=args.slow_mo,
    )
    exported, skipped = validate_and_filter(raw_rows)
    paths = export_files(exported, Path(args.output_dir).resolve(), args.query)
    print_summary(len(raw_rows), len(exported), skipped, paths)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
