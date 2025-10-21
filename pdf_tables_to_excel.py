#!/usr/bin/env python3

import argparse
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pdfplumber
from tqdm import tqdm
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


FINANCIAL_KEYWORDS = [
    # Statements
    "balance sheet",
    "statement of operations",
    "income statement",
    "statement of income",
    "cash flow",
    "cash flows",
    "comprehensive income",
    # Common finance terms
    "revenue",
    "sales",
    "gross profit",
    "gross margin",
    "operating income",
    "operating loss",
    "operating margin",
    "net income",
    "net loss",
    "eps",
    "earnings per share",
    "ebit",
    "ebitda",
    "cogs",
    "cost of goods",
    "operating expenses",
    "r&d",
    "research and development",
    "sg&a",
    "liabilities",
    "assets",
    "equity",
    "stockholders' equity",
    "shareholders' equity",
    "free cash flow",
    "gaap",
    "non-gaap",
]


def cell_has_letters(value: Optional[str]) -> bool:
    if value is None:
        return False
    return bool(re.search(r"[A-Za-z]", str(value)))


def normalize_rows(rows: List[List[Optional[str]]]) -> List[List[Optional[str]]]:
    if not rows:
        return rows
    max_len = max(len(r) for r in rows)
    out: List[List[Optional[str]]] = []
    for r in rows:
        padded = list(r) + [None] * (max_len - len(r))
        cleaned = [str(c).strip() if c is not None else None for c in padded]
        out.append(cleaned)
    return out


def detect_header_row(rows: List[List[Optional[str]]]) -> bool:
    if len(rows) < 2:
        return False
    first_row = rows[0]
    second_row = rows[1]
    first_row_letter_ratio = sum(cell_has_letters(c) for c in first_row) / max(1, len(first_row))
    second_row_letter_ratio = sum(cell_has_letters(c) for c in second_row) / max(1, len(second_row))
    return first_row_letter_ratio >= second_row_letter_ratio


def extract_tables_from_page(page: pdfplumber.page.Page) -> List[List[List[Optional[str]]]]:
    strategies: List[Dict[str, str]] = [
        {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
        {"vertical_strategy": "lines_strict", "horizontal_strategy": "lines_strict"},
        {"vertical_strategy": "text", "horizontal_strategy": "text"},
        {"vertical_strategy": "lines", "horizontal_strategy": "text"},
        {"vertical_strategy": "text", "horizontal_strategy": "lines"},
    ]

    try:
        found = page.find_tables(table_settings=strategies[0])
        tables = [t.extract() for t in found if t is not None]
        if tables:
            return tables
    except Exception:
        pass

    for settings in strategies:
        try:
            tables = page.extract_tables(table_settings=settings)
            if tables:
                return tables
        except Exception:
            continue

    return []


def extract_tables_from_pdf(pdf_path: Path, min_columns: int = 2) -> List[Tuple[int, List[List[Optional[str]]]]]:
    results: List[Tuple[int, List[List[Optional[str]]]]] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            raw_tables = extract_tables_from_page(page)
            for raw in raw_tables:
                norm = normalize_rows(raw)
                # Filter very small/degenerate tables
                if not norm:
                    continue
                if len(norm[0]) < min_columns:
                    continue
                results.append((page_idx, norm))
    return results


def is_financial_table_rows(rows: List[List[Optional[str]]]) -> bool:
    if not rows:
        return False
    sample_lines: List[str] = []
    # columns header + first up to 3 rows
    sample_lines.append(" ".join([str(c) for c in rows[0]]))
    for i in range(1, min(4, len(rows))):
        sample_lines.append(" ".join([str(c) for c in rows[i]]))
    text_sample = "\n".join(sample_lines).lower()
    return any(keyword in text_sample for keyword in FINANCIAL_KEYWORDS)


def collect_pdfs(inputs: List[str]) -> List[Path]:
    pdfs: List[Path] = []
    for inp in inputs:
        p = Path(inp)
        if p.is_file() and p.suffix.lower() == ".pdf":
            pdfs.append(p)
        elif p.is_dir():
            pdfs.extend(sorted(p.glob("**/*.pdf")))
        else:
            matched = list(Path().glob(inp))
            pdfs.extend([m for m in matched if m.suffix.lower() == ".pdf"])
    seen = set()
    unique_pdfs: List[Path] = []
    for p in pdfs:
        if p.resolve() not in seen:
            unique_pdfs.append(p)
            seen.add(p.resolve())
    return unique_pdfs


def sanitize_sheet_name(name: str) -> str:
    invalid = set('[]:*?/\\')
    cleaned = ''.join('_' if c in invalid else c for c in name)
    return cleaned[:31] if len(cleaned) > 31 else cleaned


def make_unique_sheet_name(base: str, used: set) -> str:
    base = sanitize_sheet_name(base)
    name = base
    counter = 1
    while name in used or not name:
        suffix = f"_{counter}"
        limit = 31 - len(suffix)
        name = (base[:limit] if len(base) > limit else base) + suffix
        counter += 1
    used.add(name)
    return name


def autosize_columns(ws) -> None:
    widths: Dict[int, int] = {}
    for row in ws.iter_rows(values_only=True):
        for idx, value in enumerate(row, start=1):
            length = len(str(value)) if value is not None else 0
            widths[idx] = max(widths.get(idx, 0), length)
    for idx, width in widths.items():
        ws.column_dimensions[get_column_letter(idx)].width = min(max(width + 2, 10), 60)


def write_tables_to_excel(
    extracted: List[Tuple[Path, List[Tuple[int, List[List[Optional[str]]]]]]],
    output_path: Path,
    only_financial: bool,
) -> None:
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    used_sheet_names: set = set()
    index_rows: List[Dict[str, object]] = []

    for pdf_path, page_tables in extracted:
        for idx, (page_num, rows) in enumerate(page_tables, start=1):
            if only_financial and not is_financial_table_rows(rows):
                continue
            base_name = f"{pdf_path.stem}_p{page_num}_t{idx}"
            sheet_name = make_unique_sheet_name(base_name, used_sheet_names)
            ws = wb.create_sheet(title=sheet_name)
            for r in rows:
                ws.append([None if (c is None or str(c).strip() == '') else c for c in r])
            autosize_columns(ws)
            index_rows.append(
                {
                    "sheet": sheet_name,
                    "pdf_file": str(pdf_path.name),
                    "page": int(page_num),
                    "rows": int(len(rows)),
                    "cols": int(len(rows[0]) if rows else 0),
                }
            )

    # Index sheet
    ws_idx = wb.create_sheet(title="Tables_Index", index=0)
    ws_idx.append(["sheet", "pdf_file", "page", "rows", "cols"])
    for entry in sorted(index_rows, key=lambda x: (x["pdf_file"], x["page"], x["sheet"])):
        ws_idx.append([entry["sheet"], entry["pdf_file"], entry["page"], entry["rows"], entry["cols"]])
    autosize_columns(ws_idx)

    wb.save(output_path)


def main(argv: Optional[Iterable[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extract all tables from one or more PDFs into an Excel workbook.",
    )
    parser.add_argument(
        "--input", "-i", nargs="+", required=True,
        help="PDF files, directories, or glob patterns (space-separated)",
    )
    parser.add_argument(
        "--output", "-o", required=False,
        help="Output .xlsx path. Default: <single-pdf>-tables.xlsx or combined-tables.xlsx",
    )
    parser.add_argument(
        "--financial-only", action="store_true",
        help="Keep only likely financial tables (heuristic keyword filter)",
    )
    parser.add_argument(
        "--min-columns", type=int, default=2,
        help="Minimum number of columns to keep a table (default: 2)",
    )

    args = parser.parse_args(list(argv) if argv is not None else None)

    pdfs = collect_pdfs(args.input)
    if not pdfs:
        raise SystemExit("No PDFs found for given inputs.")

    output = (
        Path(args.output)
        if args.output
        else (
            Path(pdfs[0].with_name(f"{pdfs[0].stem}-tables.xlsx"))
            if len(pdfs) == 1
            else Path("combined-tables.xlsx")
        )
    )
    if output.suffix.lower() != ".xlsx":
        output = output.with_suffix(".xlsx")

    extracted: List[Tuple[Path, List[Tuple[int, List[List[Optional[str]]]]]]] = []
    for pdf_path in tqdm(pdfs, desc="Extracting tables from PDFs"):
        tables = extract_tables_from_pdf(pdf_path, min_columns=args.min_columns)
        extracted.append((pdf_path, tables))

    write_tables_to_excel(extracted, output, only_financial=args.financial_only)
    print(f"Saved tables to: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
