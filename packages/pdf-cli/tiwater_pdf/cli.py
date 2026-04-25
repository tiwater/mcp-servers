"""PDF CLI for inspection and table extraction."""

import argparse
import base64
import json
import os
import sys
from pathlib import Path

import fitz


def _print_markdown_table(header, rows):
    """Print table in markdown format."""
    if not rows and not header:
        return ""
        
    if not header and rows:
        header = [f"Col {i+1}" for i in range(len(rows[0]))]
        
    def _clean(cell):
        if cell is None:
            return ""
        return str(cell).replace("\n", " ").strip()
        
    clean_header = [_clean(h) for h in header]
    clean_rows = [[_clean(c) for c in row] for row in rows]
    
    # PyMuPDF often includes the header in the first row of extract()
    if clean_header and clean_rows and clean_header == clean_rows[0]:
        clean_rows = clean_rows[1:]
    
    widths = [len(h) for h in clean_header]
    for row in clean_rows:
        for i, cell in enumerate(row):
            if i < len(widths):
                widths[i] = max(widths[i], len(cell))
            else:
                widths.append(len(cell))
                
    while len(clean_header) < len(widths):
        clean_header.append("")
        
    output = []
    output.append("| " + " | ".join(h.ljust(w) for h, w in zip(clean_header, widths)) + " |")
    output.append("|-" + "-|-".join("-" * w for w in widths) + "-|")
    
    for row in clean_rows:
        while len(row) < len(widths):
            row.append("")
        output.append("| " + " | ".join(c.ljust(w) for c, w in zip(row, widths)) + " |")
        
    return "\n".join(output) + "\n"


def _detect_table_title(page, bbox, max_dist: float = 25.0) -> str | None:
    """Detect title/caption for a table based on text blocks above it.
    
    Filters out chart axis labels and legend text by checking:
    - Multi-line text blocks (axis tick labels)
    - Text blocks sitting inside high-vector-density regions (chart areas)
    """
    y0 = bbox[1]
    tx0, tx1 = bbox[0], bbox[2]
    
    blocks = page.get_text("blocks")
    drawings = None  # lazy-load
    candidates = []
    
    for b in blocks:
        if len(b) >= 7 and b[6] == 0:  # text block
            bx0, by0, bx1, by1, btext, bn, btype = b
            btext = btext.strip()
            if not btext:
                continue
                
            # A true title is rarely a massive multi-line block (filters out chart axis ticks)
            if btext.count('\n') >= 3:
                continue
                
            # Check horizontally aligned with the table
            if bx1 > tx0 and bx0 < tx1:
                if by1 <= y0 + 5:  # allow small overlap
                    dist = y0 - by1
                    if dist <= max_dist:
                        # Check if this text block sits inside a chart region
                        # by looking at vector density in the area above it
                        if drawings is None:
                            drawings = page.get_drawings()
                        scan_rect = fitz.Rect(bx0 - 20, by0 - 150, bx1 + 20, by1 + 5)
                        vec_count = 0
                        for d in drawings:
                            if scan_rect.intersects(d['rect']):
                                for item in d['items']:
                                    if item[0] in ('l', 'c'):
                                        vec_count += 1
                                        if vec_count > 100:
                                            break
                                if vec_count > 100:
                                    break
                        if vec_count > 100:
                            continue  # skip — likely a chart axis label
                            
                        candidates.append((dist, btext, b))
                        
    if candidates:
        candidates.sort(key=lambda x: x[0])
        return " ".join(candidates[0][1].split())
        
    return None


def _is_valid_table(page, header, rows, bbox) -> bool:
    """Check if a detected region is a real table vs a chart/graphic.
    
    Filters out charts by combining two signals:
    - Data density: ratio of non-empty cells to total cells
    - Vector density: number of drawing primitives (lines/curves) in the region
    
    A chart typically has high vector density (plot lines, bars, gridlines)
    and very low or zero data density (only legend labels, axis text).
    A real table has moderate-to-high data density and low vector density.
    """
    if not rows:
        return False
        
    def _clean(cell):
        return str(cell).replace("\n", " ").strip() if cell is not None else ""
        
    clean_header = [_clean(h) for h in (header or [])]
    clean_rows = [[_clean(c) for c in row] for row in rows]
    
    if clean_header and clean_rows and clean_header == clean_rows[0]:
        clean_rows = clean_rows[1:]
    
    # Count non-empty data cells
    non_empty = 0
    total = 0
    for row in clean_rows:
        for cell in row:
            total += 1
            if cell:
                non_empty += 1
    
    data_ratio = non_empty / total if total > 0 else 0.0
    
    # Count vector drawing primitives in this region
    drawings = page.get_drawings()
    vector_count = 0
    rect = fitz.Rect(bbox[0] - 2, bbox[1] - 2, bbox[2] + 2, bbox[3] + 2)
    
    for d in drawings:
        if rect.intersects(d['rect']):
            for item in d['items']:
                if item[0] in ('l', 'c'):
                    vector_count += 1
    
    # High vector density indicates a chart/graphic
    if vector_count > 100:
        # With high vectors, require substantial data to be considered a real table.
        # Charts may have a few legend labels (ratio < 0.1) but real tables
        # typically have at least 20% of cells filled.
        if data_ratio < 0.15:
            return False
        
    return True


def _fix_absorbed_title(header: list, rows: list) -> tuple[list, list, str | None]:
    """Fix cases where PyMuPDF absorbs the table title into the header row.
    
    When a title like "Summary" sits visually above a table, PyMuPDF sometimes
    includes it as the header row, making the header mostly empty (e.g.,
    ['', '', '', '', '', 'Summary', '', '', '', '', '']). The real column names
    end up in rows[0].
    
    Returns:
        (fixed_header, fixed_rows, absorbed_title_or_None)
    """
    if not header or not rows:
        return header, rows, None
    
    non_empty = [h for h in header if h and str(h).strip()]
    fill_ratio = len(non_empty) / len(header) if header else 1.0
    
    # If header is mostly empty (<=20% filled), check if first row looks like real headers
    if fill_ratio > 0.2:
        return header, rows, None
    
    if not rows:
        return header, rows, None
        
    first_row = rows[0]
    first_row_non_empty = [c for c in first_row if c and str(c).strip()]
    first_row_ratio = len(first_row_non_empty) / len(first_row) if first_row else 0
    
    # First row should look like a header: mostly non-empty, non-numeric strings
    if first_row_ratio < 0.5:
        return header, rows, None
    
    # Check that first row values look like column names (mostly non-numeric)
    numeric_count = 0
    for c in first_row_non_empty:
        try:
            float(str(c).replace(',', ''))
            numeric_count += 1
        except ValueError:
            pass
    
    if numeric_count > len(first_row_non_empty) * 0.5:
        return header, rows, None  # First row looks like data, not headers
    
    # Promote first row to header, extract the absorbed title
    absorbed_title = " ".join(str(h).strip() for h in non_empty) if non_empty else None
    new_header = first_row
    new_rows = rows[1:]
    
    return new_header, new_rows, absorbed_title


def _table_quality_score(header: list, rows: list) -> float:
    """Score the quality of the extracted table from 0.0 (terrible) to 1.0 (perfect).
    
    Penalizes:
    - Unbalanced parentheses in cells
    - Split tokens (e.g. newline inside what should be a contiguous word)
    - Very high empty cell ratio in data rows
    """
    if not rows:
        return 1.0  # Empty table, nothing structurally wrong
        
    score = 1.0
    cells_checked = 0
    unbalanced_parens = 0
    empty_cells = 0
    
    for row in [header] + rows:
        if not row:
            continue
        for cell in row:
            if cell is None:
                empty_cells += 1
                continue
            
            s = str(cell).strip()
            if not s:
                empty_cells += 1
                continue
                
            cells_checked += 1
            
            # Check for unbalanced parentheses
            if s.count('(') != s.count(')'):
                unbalanced_parens += 1
                
    if cells_checked > 0:
        # Heavily penalize unbalanced parentheses (classic slicing error)
        paren_penalty = (unbalanced_parens / cells_checked) * 2.0
        score -= paren_penalty
        
    # If the vast majority of cells are empty in a large grid, it's often a garbled extraction
    total_cells = len(rows) * (len(header) if header else 1)
    if total_cells > 0:
        empty_ratio = empty_cells / total_cells
        if empty_ratio > 0.8:
            score -= 0.3
            
    return max(0.0, score)


def _render_table_region(doc, page_num: int, table_bbox: tuple) -> bytes:
    """Render a specific table region of a PDF page to a PNG image.
    
    Args:
        doc: The fitz Document
        page_num: 0-indexed page number
        table_bbox: (x0, y0, x1, y1)
        
    Returns:
        PNG image bytes
    """
    page = doc[page_num]
    # Add a small padding around the table bbox
    x0, y0, x1, y1 = table_bbox
    rect = fitz.Rect(max(0, x0 - 20), max(0, y0 - 20), x1 + 20, y1 + 20)
    
    # Render at 2x resolution for better OCR by the LLM
    mat = fitz.Matrix(2, 2)
    pix = page.get_pixmap(matrix=mat, clip=rect)
    return pix.tobytes("png")


def _resolve_llm_client(api_key: str | None = None, base_url: str | None = None):
    """Create an OpenAI-compatible client from explicit args or environment."""
    from openai import OpenAI

    resolved_api_key = (
        api_key
        or os.environ.get("TIWATER_LLM_API_KEY")
        or os.environ.get("OPENAI_API_KEY")
        or os.environ.get("OPENROUTER_API_KEY")
    )
    if not resolved_api_key:
        raise RuntimeError(
            "LLM OCR requires TIWATER_LLM_API_KEY, OPENAI_API_KEY, OPENROUTER_API_KEY, or --api-key"
        )

    resolved_base_url = (
        base_url
        or os.environ.get("TIWATER_LLM_BASE_URL")
        or os.environ.get("OPENAI_BASE_URL")
    )
    if not resolved_base_url and os.environ.get("OPENROUTER_API_KEY") and not os.environ.get("OPENAI_API_KEY"):
        resolved_base_url = "https://openrouter.ai/api/v1"

    if resolved_base_url:
        return OpenAI(base_url=resolved_base_url, api_key=resolved_api_key)
    return OpenAI(api_key=resolved_api_key)


def _llm_extract_table(image_bytes: bytes, api_key: str | None = None, llm_model: str = "google/gemini-2.5-flash") -> tuple[list, list]:
    """Use an LLM (via OpenRouter/OpenAI API) to extract a clean JSON table from an image of a table.
    
    Returns:
        (header, rows)
    """
    try:
        client = _resolve_llm_client(api_key)
    except RuntimeError as error:
        print(f"Warning: {error}", file=sys.stderr)
        return [], []

    b64_image = base64.b64encode(image_bytes).decode('utf-8')
    
    prompt = (
        "You are an expert data extraction assistant. I have provided an image of a table (possibly with a title above it). "
        "Your task is to extract the structured tabular data from this image exactly as it appears. "
        "Do not include the table title. Output a clean JSON structure with 'header' and 'rows'. "
        "If a cell is completely empty in the image, use an empty string. "
        "Ensure column alignment is perfect. Return ONLY valid JSON."
    )
    
    response = client.chat.completions.create(
        model=llm_model,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{b64_image}"
                        }
                    }
                ]
            }
        ],
        temperature=0.0,
    )
    
    try:
        content = response.choices[0].message.content
        data = json.loads(content)
        return data.get("header", []), data.get("rows", [])
    except Exception as e:
        print(f"Warning: Failed to parse LLM JSON response: {e}", file=sys.stderr)
        return [], []


def _render_page_image(doc, page_num: int, zoom: float = 2.5) -> bytes:
    page = doc[page_num]
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    return pix.tobytes("png")


def _parse_page_numbers(value: str | None) -> list[int] | None:
    if not value:
        return None
    pages = []
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        pages.append(int(part))
    return pages


def _extract_json_object(text: str) -> dict:
    text = (text or "").strip()
    if not text:
        raise ValueError("empty LLM response")
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:].strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start >= 0 and end > start:
            return json.loads(text[start:end + 1])
        raise


def llm_ocr(
    pdf_path: Path,
    page_numbers: list[int] | None = None,
    api_key: str | None = None,
    base_url: str | None = None,
    llm_model: str = "gpt-4o-mini",
    zoom: float = 2.5,
) -> dict:
    """Extract page text from scanned PDFs using an OpenAI-compatible vision model."""
    client = _resolve_llm_client(api_key, base_url)
    doc = fitz.open(pdf_path)
    pages = []

    prompt = (
        "Extract the visible text from this PDF page image with high fidelity. "
        "Preserve Chinese and English text, numbers, units, tables, row labels, and reading order. "
        "Use markdown tables when the page contains clear tables. "
        "Do not summarize and do not infer missing values. "
        "Return one JSON object with keys text, tables, warnings. "
        "tables should be an array of markdown table strings or empty if no table is visible."
    )

    try:
        for page_index in range(len(doc)):
            page_number = page_index + 1
            if page_numbers and page_number not in page_numbers:
                continue
            image_bytes = _render_page_image(doc, page_index, zoom=zoom)
            b64_image = base64.b64encode(image_bytes).decode("utf-8")
            response = client.chat.completions.create(
                model=llm_model,
                response_format={"type": "json_object"},
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {"url": f"data:image/png;base64,{b64_image}"},
                            },
                        ],
                    }
                ],
                temperature=0.0,
            )
            content = response.choices[0].message.content or ""
            parsed = _extract_json_object(content)
            pages.append({
                "page": page_number,
                "text": str(parsed.get("text", "")).strip(),
                "tables": parsed.get("tables", []) if isinstance(parsed.get("tables", []), list) else [],
                "warnings": parsed.get("warnings", []) if isinstance(parsed.get("warnings", []), list) else [],
            })
    finally:
        doc.close()

    return {
        "file": str(pdf_path),
        "model": llm_model,
        "pages": pages,
        "page_count": len(pages),
        "text": "\n\n".join(page["text"] for page in pages if page.get("text")),
    }


def _reextract_with_columns(doc, page_num: int, table_bbox: tuple, table_cells: list,
                            ref_cells: list, ref_col_count: int) -> list[list[str]]:
    """Re-extract table content using word positions and reference column boundaries.
    
    When PyMuPDF detects column boundaries slightly wrong on a page, the cell content
    gets garbled (text splits across wrong columns). This function re-extracts by:
    1. Getting all words with their positions from the page
    2. Using column boundaries from a reference table (e.g. the same table on the next page)
    3. Mapping words to the correct cells based on position
    
    Args:
        doc: fitz.Document (must still be open)
        page_num: 0-indexed page number
        table_bbox: (x0, y0, x1, y1) of the table
        table_cells: cell rectangles from the garbled table
        ref_cells: cell rectangles from the reference (clean) table
        ref_col_count: number of columns in the reference table
        
    Returns:
        List of rows, each row a list of cell strings
    """
    page = doc[page_num]
    
    # Derive row boundaries from the garbled table's cells
    row_ys = sorted(set(c[1] for c in table_cells) | set(c[3] for c in table_cells))
    # Derive column boundaries from the reference table's cells
    col_xs = sorted(set(c[0] for c in ref_cells) | set(c[2] for c in ref_cells))
    
    n_rows = len(row_ys) - 1
    n_cols = len(col_xs) - 1
    
    if n_rows <= 0 or n_cols <= 0:
        return []
    
    # Get all words in the table area
    words = page.get_text("words")
    x0, y0, x1, y1 = table_bbox
    
    # Build empty grid
    grid = [[[] for _ in range(n_cols)] for _ in range(n_rows)]
    
    for w in words:
        wx0, wy0, wx1, wy1 = w[:4]
        w_mid_x = (wx0 + wx1) / 2
        w_mid_y = (wy0 + wy1) / 2
        
        # Must be within table bbox (with small tolerance)
        if not (x0 - 5 <= w_mid_x <= x1 + 5 and y0 - 5 <= w_mid_y <= y1 + 5):
            continue
        
        # Find which row
        row_idx = -1
        for ri in range(n_rows):
            if row_ys[ri] - 2 <= w_mid_y <= row_ys[ri + 1] + 2:
                row_idx = ri
                break
        
        # Find which column
        col_idx = -1
        for ci in range(n_cols):
            if col_xs[ci] - 2 <= w_mid_x <= col_xs[ci + 1] + 2:
                col_idx = ci
                break
        
        if row_idx >= 0 and col_idx >= 0:
            grid[row_idx][col_idx].append((wx0, w[4]))  # store x-pos for ordering
    
    # Convert grid to string rows (join words left-to-right within each cell)
    result = []
    for row in grid:
        result.append([
            " ".join(text for _, text in sorted(words_in_cell))
            for words_in_cell in row
        ])
    
    return result


def extract_tables(pdf_path: Path, page_numbers: list[int] | None = None, auto_span: bool = False, llm_fallback: bool = False, api_key: str | None = None, llm_model: str = "google/gemini-2.5-flash") -> dict:
    """Extract tables from PDF pages using PyMuPDF's table detection.

    Args:
        pdf_path: Path to PDF file
        page_numbers: Optional list of page numbers (1-indexed). If None, extracts from all pages.
        auto_span: If True, merge tables with matching headers across consecutive pages
        llm_fallback: If True, use LLM vision to re-extract garbled tables.
        api_key: Optional API key for LLM fallback.
        llm_model: Model string to pass to OpenRouter (default: google/gemini-2.5-flash).

    Returns:
        Dictionary with extracted tables data
    """
    doc = fitz.open(pdf_path)
    all_tables = []

    for page_num in range(len(doc)):
        if page_numbers and (page_num + 1) not in page_numbers:
            continue

        page = doc[page_num]

        for table in page.find_tables():
            title = _detect_table_title(page, table.bbox)
            header = table.header.names
            rows = table.extract()
            
            if not _is_valid_table(page, header, rows, table.bbox):
                continue
            
            # Fix cases where title got absorbed into header row
            header, rows, absorbed_title = _fix_absorbed_title(header, rows)
            if not title and absorbed_title:
                title = absorbed_title
                
            # LLM Fallback for garbled tables
            if llm_fallback:
                q_score = _table_quality_score(header, rows)
                if q_score < 0.95:
                    print(f"Poor extraction quality ({q_score:.2f}) on page {page_num+1}. Falling back to LLM...", file=sys.stderr)
                    img_bytes = _render_table_region(doc, page_num, table.bbox)
                    llm_header, llm_rows = _llm_extract_table(img_bytes, api_key, llm_model)
                    if llm_header or llm_rows:
                        header, rows = llm_header, llm_rows
                
            all_tables.append({
                "page": page_num + 1,
                "title": title,
                "bbox": list(table.bbox),
                "header": header,
                "rows": rows,
                "_cells": list(table.cells),
            })

    total_pages = len(doc)

    if not auto_span:
        for t in all_tables:
            t.pop("_cells", None)
        doc.close()
        return {
            "file": str(pdf_path),
            "total_pages": total_pages,
            "tables_found": len(all_tables),
            "tables": all_tables,
        }

    # Auto-span: merge tables with matching headers on consecutive pages
    spanned_tables = []
    skip_indices = set()

    for i, table in enumerate(all_tables):
        if i in skip_indices:
            continue

        current_header = table["header"]
        span_entries = [table]
        pages = [table["page"]]

        # Look for tables on next pages with matching headers
        j = i + 1
        while j < len(all_tables):
            next_table = all_tables[j]

            if (next_table["page"] == pages[-1] + 1 and
                _headers_match(current_header, next_table["header"])):

                span_entries.append(next_table)
                pages.append(next_table["page"])
                skip_indices.add(j)
                j += 1
            else:
                break

        if len(span_entries) > 1:
            # Find cleanest table as column reference
            def _header_set(t):
                return set(" ".join(str(x).split()).lower() for x in t["header"] if x and str(x).strip())
            
            all_header_sets = [_header_set(e) for e in span_entries]
            
            def _cleanliness_score(idx):
                """Score how 'clean' a table's headers are.
                
                Higher = cleaner. Penalizes:
                - Newlines in header values (indicates garbled multi-line splits)
                - Unbalanced parentheses (indicates column boundary errors)
                Rewards:
                - More non-empty header cells
                - Header values that appear in other tables' headers (consensus)
                """
                entry = span_entries[idx]
                h = entry["header"]
                
                # Count newlines in raw header values (fewer = cleaner)
                newline_penalty = sum(str(x).count('\n') for x in h if x)
                
                # Unbalanced parentheses penalty
                paren_penalty = 0
                for x in h:
                    if x:
                        s = str(x)
                        paren_penalty += abs(s.count('(') - s.count(')'))
                
                # Consensus: how many values appear in other tables
                my_set = all_header_sets[idx]
                consensus = sum(
                    1 for v in my_set 
                    for j, os in enumerate(all_header_sets) 
                    if j != idx and v in os
                )
                
                return consensus * 10 - newline_penalty * 5 - paren_penalty * 3
            
            ref_idx = max(range(len(span_entries)), key=_cleanliness_score)
            ref_table = span_entries[ref_idx]
            ref_cells = ref_table["_cells"]
            ref_col_count = len(ref_table["header"])
            best_header = ref_table["header"]
            ref_header_set = all_header_sets[ref_idx]
            
            # Merge rows, re-extracting garbled pages using clean column boundaries
            merged_rows = []
            for entry_idx, entry in enumerate(span_entries):
                # Detect garbled columns: if < 95% of header values match the reference
                entry_set = all_header_sets[entry_idx]
                shared = entry_set & ref_header_set
                match_ratio = len(shared) / max(len(entry_set), 1)
                needs_reextract = entry is not ref_table and match_ratio < 0.95

                if needs_reextract:
                    llm_success = False
                    if llm_fallback:
                        print(f"Garbled table spanning detected on page {entry['page']}. Falling back to LLM...", file=sys.stderr)
                        img_bytes = _render_table_region(doc, entry["page"] - 1, tuple(entry["bbox"]))
                        llm_h, llm_r = _llm_extract_table(img_bytes, api_key, llm_model)
                        if llm_h or llm_r:
                            merged_rows.extend(llm_r)
                            llm_success = True
                            
                    if not llm_success:
                        # Garbled page — re-extract using clean column boundaries geometrically
                        reextracted = _reextract_with_columns(
                            doc, entry["page"] - 1, tuple(entry["bbox"]),
                            entry["_cells"], ref_cells, ref_col_count
                        )
                        if reextracted:
                            # Re-extraction from table.bbox includes the header as its first row.
                            # We use _fix_absorbed_title to safely separate header from data rows.
                            _, re_rows, _ = _fix_absorbed_title(reextracted[0], reextracted[1:])
                            merged_rows.extend(re_rows)
                        else:
                            merged_rows.extend(entry["rows"])
                else:
                    rows = entry["rows"]
                    if entry is not span_entries[0] and len(rows) > 1 and _headers_match(current_header, rows[0]):
                        merged_rows.extend(rows[1:])
                    else:
                        merged_rows.extend(rows)

            spanned_tables.append({
                "pages": pages,
                "page": pages[0],
                "title": table.get("title"),
                "bbox": table["bbox"],
                "header": best_header,
                "rows": merged_rows,
                "row_count": len(merged_rows),
                "spanned_from": len(span_entries),
            })
        else:
            spanned_tables.append({
                "pages": pages,
                "page": pages[0],
                "title": table.get("title"),
                "bbox": table["bbox"],
                "header": current_header,
                "rows": table["rows"],
                "row_count": len(table["rows"]),
                "spanned_from": 1,
            })

    doc.close()

    return {
        "file": str(pdf_path),
        "total_pages": total_pages,
        "tables_found": len(spanned_tables),
        "tables": spanned_tables,
    }


def find_table_by_name(pdf_path: Path, table_name: str, auto_span: bool = False, llm_fallback: bool = False, api_key: str | None = None, llm_model: str = "google/gemini-2.5-flash") -> dict | None:
    """Find a table by its name/header in the PDF.

    Args:
        pdf_path: Path to PDF file
        table_name: Name of table to find (searches in headers and content)
        auto_span: If True, merge tables with matching headers across consecutive pages
        llm_fallback: If True, use LLM vision to re-extract garbled tables
        api_key: Optional OpenRouter API Key
        llm_model: Target LLM model on OpenRouter

    Returns:
        Dictionary with table data or None if not found
    """
    # Extract all tables
    data = extract_tables(pdf_path, auto_span=auto_span, llm_fallback=llm_fallback, api_key=api_key, llm_model=llm_model)
    
    # Search for the matched table
    for t in data.get("tables", []):
        matches = False
        title = t.get("title")
        
        # Check title first
        if title and table_name.lower() in title.lower():
            matches = True
            
        # Check header
        if not matches:
            header_text = " ".join([h for h in t.get("header", []) if h])
            if table_name.lower() in header_text.lower():
                matches = True
                
        # Check first few rows
        if not matches:
            for row in t.get("rows", [])[:3]:
                row_text = " ".join([str(c) for c in row if c])
                if table_name.lower() in row_text.lower():
                    matches = True
                    break
                    
        if matches:
            t["file"] = str(pdf_path)
            t["table_name"] = table_name
            return t
            
    return None


def inspect(pdf_path: Path) -> dict:
    """Inspect PDF and return metadata.

    Args:
        pdf_path: Path to PDF file

    Returns:
        Dictionary with PDF metadata
    """
    doc = fitz.open(pdf_path)
    page_summaries = []
    total_images = 0
    total_words = 0
    scanned_pages = 0
    for i, page in enumerate(doc):
        image_count = len(page.get_images(full=True))
        word_count = len(page.get_text("words"))
        total_images += image_count
        total_words += word_count
        image_only = image_count > 0 and word_count == 0
        if image_only:
            scanned_pages += 1
        page_summaries.append(
            {
                "page": i + 1,
                "width": page.rect.width,
                "height": page.rect.height,
                "image_count": image_count,
                "word_count": word_count,
                "image_only": image_only,
            }
        )

    metadata = {
        "file": str(pdf_path),
        "pages": len(doc),
        "metadata": doc.metadata,
        "image_count": total_images,
        "word_count": total_words,
        "scanned_page_count": scanned_pages,
        "image_only": scanned_pages == len(doc) and len(doc) > 0,
        "page_sizes": page_summaries,
    }
    doc.close()
    return metadata


def _normalize_header(header: list[str]) -> tuple[str, ...]:
    """Normalize header for comparison by removing whitespace and newlines."""
    return tuple(" ".join(h.split()) if h else "" for h in header)


def _headers_match(header1: list[str], header2: list[str], threshold: float = 0.6) -> bool:
    """Check if two headers match (allowing for minor differences)."""
    norm1 = _normalize_header(header1)
    norm2 = _normalize_header(header2)

    if len(norm1) == 0 or len(norm2) == 0:
        return False

    # Direct match
    if norm1 == norm2:
        return True

    # Fuzzy match - check overlap ratio
    set1 = set(h.lower() for h in norm1 if h)
    set2 = set(h.lower() for h in norm2 if h)

    if not set1 or not set2:
        return False

    intersection = set1 & set2
    union = set1 | set2

    # Also check if key columns match (for better matching of similar tables)
    # If at least 3 key non-empty columns match, consider it a match
    key_match = len(intersection) >= 3 and len(intersection) >= min(len(set1), len(set2)) - 2

    # Same column count + enough intersection: likely same table with garbled columns
    same_col_count = len(norm1) == len(norm2)
    col_overlap = len(intersection) / min(len(set1), len(set2)) if min(len(set1), len(set2)) > 0 else 0
    structural_match = same_col_count and len(intersection) >= 4 and col_overlap >= 0.4
    return key_match or structural_match or (len(intersection) / len(union) >= threshold)


def _load_config() -> dict:
    """Load configuration from ~/.dockit/config.toml or ~/.config/dockit/config.toml if present."""
    import tomllib
    
    config_paths = [
        Path.home() / ".dockit" / "config.toml",
        Path.home() / ".config" / "dockit" / "config.toml",
    ]
    
    for p in config_paths:
        if p.exists():
            try:
                with open(p, "rb") as f:
                    return tomllib.load(f)
            except Exception as e:
                print(f"Warning: Failed to parse config {p}: {e}", file=sys.stderr)
                
    return {}


def main() -> int:
    """Main CLI entry point."""
    config = _load_config()
    llm_config = config.get("llm", {})
    default_api_key = llm_config.get("api_key")
    default_base_url = llm_config.get("base_url")
    default_llm_model = llm_config.get("model", "qwen/qwen3.5-flash-02-23")
    default_ocr_model = (
        llm_config.get("ocr_model")
        or llm_config.get("vision_model")
        or "gpt-4o-mini"
    )

    parser = argparse.ArgumentParser(
        description="dockit-pdf - PDF inspection and table extraction CLI"
    )
    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # inspect command
    inspect_parser = subparsers.add_parser("inspect", help="Inspect PDF metadata")
    inspect_parser.add_argument("input", type=Path, help="PDF file to inspect")
    inspect_parser.add_argument("--json", action="store_true", help="Output as JSON")

    # extract-tables command
    extract_parser = subparsers.add_parser("extract-tables", help="Extract tables from PDF")
    extract_parser.add_argument("input", type=Path, help="PDF file to extract from")
    extract_parser.add_argument("--pages", type=str, help="Page numbers (comma-separated, 1-indexed)")
    extract_parser.add_argument("--auto-span", action="store_true", help="Merge tables spanning multiple pages")
    extract_parser.add_argument("--llm-fallback", action="store_true", help="Use OpenRouter LLM for garbled tables")
    extract_parser.add_argument("--api-key", type=str, default=default_api_key, help="OpenRouter API Key (or set OPENROUTER_API_KEY env var)")
    extract_parser.add_argument("--llm-model", type=str, default=default_llm_model, help="LLM model to use on OpenRouter")
    extract_parser.add_argument("--json", action="store_true", help="Output as JSON")

    # find-table command
    find_parser = subparsers.add_parser("find-table", help="Find table by name")
    find_parser.add_argument("input", type=Path, help="PDF file to search")
    find_parser.add_argument("name", type=str, help="Table name to find")
    find_parser.add_argument("--auto-span", action="store_true", help="Merge tables spanning multiple pages")
    find_parser.add_argument("--llm-fallback", action="store_true", help="Use OpenRouter LLM for garbled tables")
    find_parser.add_argument("--api-key", type=str, default=default_api_key, help="OpenRouter API Key")
    find_parser.add_argument("--llm-model", type=str, default=default_llm_model, help="LLM model to use on OpenRouter")
    find_parser.add_argument("--json", action="store_true", help="Output as JSON")

    # OCR command
    ocr_parser = subparsers.add_parser("ocr", help="Extract scanned PDF text with an OpenAI-compatible vision LLM")
    ocr_parser.add_argument("input", type=Path, help="PDF file to OCR")
    ocr_parser.add_argument("--pages", type=str, help="Page numbers (comma-separated, 1-indexed)")
    ocr_parser.add_argument("--api-key", type=str, default=default_api_key, help="LLM API key")
    ocr_parser.add_argument("--base-url", type=str, default=default_base_url, help="OpenAI-compatible base URL")
    ocr_parser.add_argument("--llm-model", type=str, default=default_ocr_model, help="Vision model to use")
    ocr_parser.add_argument("--zoom", type=float, default=2.5, help="PDF render zoom for page images")
    ocr_parser.add_argument("--json", action="store_true", help="Output as JSON")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    try:
        if args.command == "inspect":
            result = inspect(args.input)
            if args.json:
                print(json.dumps(result, indent=2, ensure_ascii=False))
            else:
                print(f"File: {result['file']}")
                print(f"Pages: {result['pages']}")
                print(f"Metadata: {result['metadata']}")

        elif args.command == "extract-tables":
            pages = None
            if args.pages:
                pages = [int(p.strip()) for p in args.pages.split(",")]
            result = extract_tables(args.input, pages, auto_span=args.auto_span, llm_fallback=args.llm_fallback, api_key=args.api_key, llm_model=args.llm_model)
            if args.json:
                print(json.dumps(result, indent=2, ensure_ascii=False))
            else:
                print(f"File: {result['file']}")
                print(f"Tables found: {result['tables_found']}")
                for table in result["tables"]:
                    title_str = f" (title: '{table['title']}')" if table.get("title") else ""
                    if args.auto_span and table.get("spanned_from", 1) > 1:
                        print(f"\n## Table on Pages {table['pages']}{title_str} (spanned from {table['spanned_from']} tables, {table.get('row_count', len(table['rows']))} rows)")
                    else:
                        print(f"\n## Table on Page {table['page']}{title_str} ({len(table['rows'])} rows)")
                    
                    if table.get("rows") or table.get("header"):
                        print(_print_markdown_table(table.get("header", []), table.get("rows", [])))

        elif args.command == "find-table":
            result = find_table_by_name(args.input, args.name, auto_span=args.auto_span, llm_fallback=args.llm_fallback, api_key=args.api_key)
            if result:
                if args.json:
                    print(json.dumps(result, indent=2, ensure_ascii=False))
                else:
                    pages = result.get("pages", [result["page"]])
                    spanned = result.get("spanned_from", 1)
                    if spanned > 1:
                        print(f"Found '{args.name}' on pages {pages} (spanned from {spanned} tables)")
                    else:
                        print(f"Found '{args.name}' on page {result['page']}")
                    if result.get("title"):
                        print(f"Detected Title: {result['title']}")
                    print(f"Rows: {len(result['rows'])}\n")
                    print(_print_markdown_table(result.get("header", []), result.get("rows", [])))
            else:
                print(f"Table '{args.name}' not found", file=sys.stderr)
                return 1

        elif args.command == "ocr":
            pages = _parse_page_numbers(args.pages)
            result = llm_ocr(
                args.input,
                pages,
                api_key=args.api_key,
                base_url=args.base_url,
                llm_model=args.llm_model,
                zoom=args.zoom,
            )
            if args.json:
                print(json.dumps(result, indent=2, ensure_ascii=False))
            else:
                print(result["text"])

        return 0

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
