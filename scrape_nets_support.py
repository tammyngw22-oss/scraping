"""
Scraper for the NETS "for-you/support" FAQ page.

The page's accordion Q&A content is fully server-rendered (no JS execution
needed), so this fetches the raw HTML and parses it with BeautifulSoup rather
than driving a browser.

Usage:
    python3 scrape_nets_support.py [output_dir]

Writes one raw JSON file per site section (e.g. section-01-top-questions.json)
containing the section title and its Q&A pairs (with product-tag, question,
and answer text). This is raw extracted data for review, not the final
formatted RAG doc output.
"""
import json
import re
import sys
from pathlib import Path
from urllib.request import Request, urlopen

from bs4 import BeautifulSoup

URL = "https://www.nets.com.sg/nets/for-you/support"
SECTION_TITLE_SELECTOR = "p.has-netsblue-color.has-text-color.has-h-4-font-size"


def fetch_html(url):
    req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8")


def html_to_text(el):
    for a in el.find_all("a"):
        href = a.get("href", "")
        a.replace_with(f"{a.get_text(strip=True)} ({href})" if href else a.get_text(strip=True))
    return el.get_text("\n", strip=True)


def extract_sections(html):
    soup = BeautifulSoup(html, "lxml")
    titles = soup.select(SECTION_TITLE_SELECTOR)
    sections = []
    for title_el in titles:
        title = title_el.get_text(strip=True)
        accordion = title_el.find_next("div", class_="kt-accordion-wrap")
        if accordion is None:
            continue
        qa_pairs = []
        for pane in accordion.select("div.kt-accordion-pane"):
            tags = [c for c in pane.get("class", []) if c.startswith("tag-")]
            question_el = pane.select_one(".kt-blocks-accordion-title")
            panel_el = pane.select_one(".kt-accordion-panel-inner")
            if not question_el or not panel_el:
                continue
            qa_pairs.append(
                {
                    "tags": tags,
                    "question": question_el.get_text(strip=True),
                    "answer": html_to_text(panel_el),
                }
            )
        sections.append({"title": title, "qa_pairs": qa_pairs})
    return sections


PHONE_RE = re.compile(r"\b\d{4}[\s-]?\d{4}\b|\b1800[\s-]?\d{3}[\s-]?\d{4}\b")
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")


def flag_contact_info(sections):
    flagged = []
    for s_idx, section in enumerate(sections):
        for q_idx, qa in enumerate(section["qa_pairs"]):
            phones = PHONE_RE.findall(qa["answer"])
            emails = EMAIL_RE.findall(qa["answer"])
            if phones or emails:
                flagged.append(
                    {
                        "section": section["title"],
                        "question": qa["question"],
                        "phones": phones,
                        "emails": emails,
                    }
                )
    return flagged


def main():
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("scraped_data")
    out_dir.mkdir(parents=True, exist_ok=True)

    html = fetch_html(URL)
    sections = extract_sections(html)

    for i, section in enumerate(sections, start=1):
        slug = re.sub(r"[^a-z0-9]+", "-", section["title"].lower()).strip("-")
        out_path = out_dir / f"section-{i:02d}-{slug}.json"
        out_path.write_text(json.dumps(section, indent=2))
        print(f"[{i:02d}] {section['title']}: {len(section['qa_pairs'])} Q&A -> {out_path}")

    flagged = flag_contact_info(sections)
    if flagged:
        print("\nContact info found (needs removal per scraping requirement):")
        for f in flagged:
            print(f"  - [{f['section']}] {f['question']!r}: phones={f['phones']} emails={f['emails']}")


if __name__ == "__main__":
    main()
