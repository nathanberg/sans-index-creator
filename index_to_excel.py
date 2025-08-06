#!/usr/bin/env python3
"""Convert a text index into an Excel spreadsheet compatible with GIAC-Index-Creator.

The input is an index text file produced by ``sans_indexer.py`` or
``index_combiner.py``.  Each line should look like one of the following::

    word: 3, 5, 7
    word: 1(3, 5) | 2(7)

The output Excel file will contain four columns:
``Topic``, ``Description``, ``Page`` and ``Book``.

Example usage::

    python index_to_excel.py index.txt output.xlsx
"""

import argparse
from typing import List

try:
    from openpyxl import Workbook
except ImportError as exc:  # pragma: no cover - dependency missing at runtime
    raise SystemExit("openpyxl is required to run this script") from exc


def parse_line(line: str) -> List[List[str]]:
    """Parse a single index line into [topic, description, page, book] rows."""
    if ":" not in line:
        return []

    # Split on the last colon to allow topics containing colons (e.g. Windows paths)
    topic, raw_pages = line.rsplit(":", 1)
    topic = topic.strip()
    raw_pages = raw_pages.strip()

    rows: List[List[str]] = []

    if "(" in raw_pages:
        # Format like: 1(2, 3) | 2(7)
        for book_chunk in raw_pages.split("|"):
            book_chunk = book_chunk.strip()
            if not book_chunk:
                continue
            if "(" not in book_chunk or ")" not in book_chunk:
                continue
            book, pages = book_chunk.split("(", 1)
            book = book.strip()
            pages = pages.rstrip(")")
            for page in pages.split(","):
                rows.append([topic, None, page.strip(), book])
    else:
        # Format like: 2, 3, 7 (single book)
        for page in raw_pages.split(","):
            rows.append([topic, None, page.strip(), "1"])

    return rows


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_file", help="index text file to convert")
    parser.add_argument("output_file", help="Excel file to create")
    args = parser.parse_args()

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Topic", "Description", "Page", "Book"])

    with open(args.input_file, "r") as f:
        for line in f:
            for row in parse_line(line):
                sheet.append(row)

    workbook.save(args.output_file)
    print(f"Written Excel index to {args.output_file}")


if __name__ == "__main__":
    main()
