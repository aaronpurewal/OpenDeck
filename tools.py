"""
Tool Layer: All Aspose PPTX operations.

Every tool is a standalone function. No global state. Each function receives
the Aspose Presentation object as its first argument and returns a result dict.

Tools never call the LLM. The LLM decides *what* to do; tools do it mechanically.
"""

import math
import unicodedata
import aspose.slides as slides
import aspose.slides.charts as charts
from state import harvest_deck, extract_shape, estimate_char_limit


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize(text: str) -> str:
    """Normalize text for run matching: strip whitespace and normalize unicode."""
    return unicodedata.normalize("NFKC", text.strip())


def _find_shape(slide, shape_name: str):
    """Find a shape by name on a slide."""
    for shape in slide.shapes:
        if shape.name == shape_name:
            return shape
    return None


def _find_layout(prs, layout_name: str):
    """Find a layout slide by name across all masters."""
    for master in prs.masters:
        for layout in master.layout_slides:
            if layout.name == layout_name:
                return layout
    return None


def _find_donor_slide(prs, layout_name: str):
    """
    Find the first existing slide that uses the given layout.

    Returns the slide index, or None if no slide uses this layout.
    """
    for i in range(len(prs.slides)):
        try:
            if (prs.slides[i].layout_slide and
                    prs.slides[i].layout_slide.name == layout_name):
                return i
        except Exception:
            continue
    return None


def _clear_slide_content(slide):
    """
    Clear all text content on a slide while preserving formatting.

    Clears portion text to "" but preserves paragraph/portion structure,
    font formatting, and shape geometry. fill_placeholder will reuse
    existing paragraphs to avoid ghost empty lines.
    """
    for shape in slide.shapes:
        if isinstance(shape, slides.Table):
            for row_idx in range(len(shape.rows)):
                for col_idx in range(len(shape.columns)):
                    cell = shape.rows[row_idx][col_idx]
                    tf = _safe_text_frame(cell)
                    if tf:
                        for para in tf.paragraphs:
                            for portion in para.portions:
                                portion.text = ""
        elif isinstance(shape, charts.Chart):
            pass  # Leave chart data intact
        else:
            tf = _safe_text_frame(shape)
            if tf:
                for para in tf.paragraphs:
                    for portion in para.portions:
                        portion.text = ""


def _normalize_para_format(para, template_para=None):
    """
    Normalize a paragraph's bullet/indent formatting.

    Resets depth to 0 and copies indent/margin from the template paragraph
    (typically the first bullet paragraph from the donor). This prevents
    donor sub-bullet formatting from leaking into fill_placeholder content.
    """
    try:
        pf = para.paragraph_format
        pf.depth = 0
        if template_para:
            tpf = template_para.paragraph_format
            try:
                pf.margin_left = tpf.margin_left
            except Exception:
                pass
            try:
                pf.indent = tpf.indent
            except Exception:
                pass
            # Copy bullet format from template
            try:
                pf.bullet.type = tpf.bullet.type
            except Exception:
                pass
            try:
                pf.bullet.char = tpf.bullet.char
            except Exception:
                pass
            try:
                pf.bullet.height = tpf.bullet.height
            except Exception:
                pass
            # Copy bullet font so the glyph renders in the correct typeface
            try:
                if tpf.bullet.is_bullet_hard_font:
                    pf.bullet.is_bullet_hard_font = True
                    pf.bullet.font.name_ascii = tpf.bullet.font.name_ascii
            except Exception:
                pass
    except Exception:
        pass


def _clear_portion_junk(portion):
    """
    Clear explicit formatting that placeholders often inherit from masters.

    Placeholder default runs frequently have hyperlink-like styling
    (blue fill, underline, bold) from the "Click to add" boilerplate.
    Reset these to NOT_DEFINED so the portion inherits clean theme defaults.
    """
    try:
        pf = portion.portion_format
        pf.font_underline = slides.TextUnderlineType.NONE
        pf.font_bold = slides.NullableBool.NOT_DEFINED
        pf.font_italic = slides.NullableBool.NOT_DEFINED
        # Clear explicit fill color so it inherits from theme
        pf.fill_format.fill_type = slides.FillType.NOT_DEFINED
    except Exception:
        pass


def _safe_text_frame(shape):
    """Get text_frame safely, returns None if not available."""
    try:
        tf = shape.text_frame
        if tf is not None:
            return tf
    except Exception:
        pass
    return None


def _safe_font_height(pf):
    """Get font_height safely, returns 0 if NaN or unavailable."""
    try:
        fh = pf.font_height
        if fh is not None and not math.isnan(fh):
            return fh
    except Exception:
        pass
    return 0


def _truncate_to_fit(text: str, char_limit: int) -> str:
    """
    Truncate text to fit within char_limit.

    Strategy: slide content is typically \\n-separated bullets.
    Drop trailing bullets that don't fit. If a single block of text,
    truncate at the last word boundary and append "...".
    """
    if len(text) <= char_limit:
        return text

    parts = text.split("\n")
    if len(parts) > 1:
        # Multi-bullet: accumulate until we'd exceed the limit
        kept = []
        running = 0
        for part in parts:
            # +1 for the \n separator (except first)
            added = len(part) + (1 if kept else 0)
            if running + added > char_limit:
                break
            kept.append(part)
            running += added
        if kept:
            return "\n".join(kept)

    # Single block or no bullets fit: word-boundary truncation
    cutoff = char_limit - 3  # room for "..."
    if cutoff <= 0:
        return text[:char_limit]
    truncated = text[:cutoff]
    last_space = truncated.rfind(" ")
    if last_space > cutoff * 0.5:
        truncated = truncated[:last_space]
    return truncated + "..."


# ---------------------------------------------------------------------------
# Read Operations
# ---------------------------------------------------------------------------

def get_slide_state(prs, slide_idx: int) -> dict:
    """Targeted read of one slide."""
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    layout_name = "Unknown"
    try:
        if slide.layout_slide:
            layout_name = slide.layout_slide.name
    except Exception:
        pass
    slide_state = {
        "index": slide_idx,
        "layout_name": layout_name,
        "shapes": []
    }
    for shape in slide.shapes:
        shape_state = extract_shape(shape)
        if shape_state:
            slide_state["shapes"].append(shape_state)
    return {"status": "ok", "slide": slide_state}


def get_bounds(prs, slide_idx: int, shape_name: str) -> dict:
    """Get bounding box and char_limit for a specific shape."""
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    shape = _find_shape(slide, shape_name)
    if not shape:
        return {"status": "error", "message": f"Shape '{shape_name}' not found"}
    bounds = {"x": shape.x, "y": shape.y, "w": shape.width, "h": shape.height}
    max_font = 0
    tf = _safe_text_frame(shape)
    if tf:
        for para in tf.paragraphs:
            for portion in para.portions:
                fh = _safe_font_height(portion.portion_format)
                if fh > max_font:
                    max_font = fh
    char_limit = estimate_char_limit(shape.width, shape.height, font_size_pt=max_font)
    return {"status": "ok", "bounds": bounds, "char_limit": char_limit}


def list_layouts(prs) -> dict:
    """List available master layouts in the deck."""
    layouts = []
    for master in prs.masters:
        for layout in master.layout_slides:
            layouts.append({"name": layout.name})
    return {"status": "ok", "layouts": layouts}


# ---------------------------------------------------------------------------
# Structural Write Operations
# ---------------------------------------------------------------------------

def clone_slide(prs, layout_name: str, insert_at: int = None) -> dict:
    """
    Create a new slide from a layout.

    Tries to find an existing "donor" slide that uses the same layout.
    If found, duplicates it and clears its text — preserving designer-set
    shape geometry, font formatting, and paragraph structure.
    Falls back to insert_empty_slide if no existing slide uses this layout.
    """
    layout = _find_layout(prs, layout_name)
    if not layout:
        return {"status": "error",
                "message": f"Layout '{layout_name}' not found"}
    if insert_at is None:
        insert_at = len(prs.slides)

    donor_idx = _find_donor_slide(prs, layout_name)
    if donor_idx is not None:
        prs.slides.insert_clone(insert_at, prs.slides[donor_idx])
        _clear_slide_content(prs.slides[insert_at])
        return {"status": "ok", "slide_idx": insert_at,
                "layout": layout_name, "donor_idx": donor_idx}

    prs.slides.insert_empty_slide(insert_at, layout)
    return {"status": "ok", "slide_idx": insert_at, "layout": layout_name}


def duplicate_slide(prs, source_idx: int, insert_at: int = None) -> dict:
    """Copy an existing slide (content + formatting) to a new position."""
    if source_idx < 0 or source_idx >= len(prs.slides):
        return {"status": "error",
                "message": f"Source index {source_idx} out of range"}
    if insert_at is None:
        insert_at = len(prs.slides)
    source_slide = prs.slides[source_idx]
    prs.slides.insert_clone(insert_at, source_slide)
    return {"status": "ok", "slide_idx": insert_at}


def delete_slides(prs, indices: list[int]) -> dict:
    """Remove slides by index. Indices should be sorted in reverse order."""
    sorted_indices = sorted(indices, reverse=True)
    for idx in sorted_indices:
        if 0 <= idx < len(prs.slides):
            prs.slides.remove_at(idx)
    return {"status": "ok", "deleted_count": len(sorted_indices)}


def reorder_slides(prs, order: list[int]) -> dict:
    """
    Rearrange slides to match new index order.
    order[i] = the current index of the slide that should be at position i.
    """
    if len(order) != len(prs.slides):
        return {"status": "error",
                "message": f"Order length {len(order)} != slide count {len(prs.slides)}"}
    # Collect slide references in new order
    slide_refs = [prs.slides[i] for i in order]
    for target_pos, slide_ref in enumerate(slide_refs):
        current_pos = None
        for i in range(len(prs.slides)):
            if prs.slides[i] == slide_ref:
                current_pos = i
                break
        if current_pos is not None and current_pos != target_pos:
            prs.slides.reorder(target_pos, prs.slides[current_pos])
    return {"status": "ok"}


def save_deck(prs, output_path: str) -> dict:
    """Write the modified presentation to disk."""
    try:
        prs.save(output_path, slides.export.SaveFormat.PPTX)
        return {"status": "ok", "path": output_path}
    except Exception as e:
        return {"status": "error", "message": str(e)}


# ---------------------------------------------------------------------------
# Content — FILL (new/empty slides from cloned layouts)
# ---------------------------------------------------------------------------

def fill_placeholder(prs, slide_idx: int, shape_name: str, text: str) -> dict:
    """
    Write text into a placeholder on a NEWLY CLONED slide.
    Paragraph breaks via "\\n".
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    shape = _find_shape(slide, shape_name)
    tf = _safe_text_frame(shape) if shape else None
    if not tf:
        return {"status": "error",
                "message": f"Shape '{shape_name}' not found or has no text frame on slide {slide_idx}"}

    # Enforce char_limit — LLMs routinely overshoot
    max_font = 0
    for para in tf.paragraphs:
        for portion in para.portions:
            fh = _safe_font_height(portion.portion_format)
            if fh > max_font:
                max_font = fh
    char_limit = estimate_char_limit(shape.width, shape.height, font_size_pt=max_font)
    text = _truncate_to_fit(text, char_limit)

    new_paragraphs = text.split("\n")
    existing_count = tf.paragraphs.count

    # Find the first bullet paragraph (depth 0, with indent) to use as
    # formatting template — ensures all content paragraphs look uniform.
    template_para = None
    for pi in range(min(existing_count, 5)):
        try:
            pf = tf.paragraphs[pi].paragraph_format
            if pf.indent and not math.isnan(pf.indent) and pf.indent != 0:
                template_para = tf.paragraphs[pi]
                break
        except Exception:
            continue

    # Reuse existing paragraphs from the donor slide (preserves font formatting).
    # Normalize bullet/indent so all content paragraphs look uniform.
    for p_idx in range(max(len(new_paragraphs), existing_count)):
        if p_idx < existing_count:
            para = tf.paragraphs[p_idx]
            if p_idx < len(new_paragraphs):
                # Write into existing paragraph
                if para.portions.count > 0:
                    para.portions[0].text = new_paragraphs[p_idx]
                    _clear_portion_junk(para.portions[0])
                    for i in range(1, para.portions.count):
                        para.portions[i].text = ""
                else:
                    portion = slides.Portion()
                    portion.text = new_paragraphs[p_idx]
                    para.portions.add(portion)
                # Normalize formatting: all content paragraphs get uniform style
                if p_idx > 0 and template_para:
                    _normalize_para_format(para, template_para)
            else:
                # Extra donor paragraph — blank it out
                for portion in para.portions:
                    portion.text = ""
        else:
            # More new paragraphs than donor had — add fresh ones
            # Copy paragraph + font formatting from template so they don't
            # render in the shape's large default font.
            new_para = slides.Paragraph()
            portion = slides.Portion()
            portion.text = new_paragraphs[p_idx]
            if template_para and template_para.portions.count > 0:
                tportion = template_para.portions[0]
                try:
                    portion.portion_format.font_height = tportion.portion_format.font_height
                except Exception:
                    pass
                try:
                    portion.portion_format.font_bold = tportion.portion_format.font_bold
                except Exception:
                    pass
                try:
                    portion.portion_format.latin_font = tportion.portion_format.latin_font
                except Exception:
                    pass
                try:
                    portion.portion_format.fill_format.fill_type = tportion.portion_format.fill_format.fill_type
                    if tportion.portion_format.fill_format.fill_type == slides.FillType.SOLID:
                        portion.portion_format.fill_format.solid_fill_color.color = (
                            tportion.portion_format.fill_format.solid_fill_color.color)
                except Exception:
                    pass
            new_para.portions.add(portion)
            tf.paragraphs.add(new_para)
            if template_para:
                _normalize_para_format(new_para, template_para)

    return {"status": "ok", "slide_idx": slide_idx, "shape": shape_name}


def fill_table(prs, slide_idx: int, shape_name: str, rows: list,
               headers: list = None) -> dict:
    """
    Populate a table on a NEWLY CLONED slide.
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    shape = _find_shape(slide, shape_name)
    if not shape or not isinstance(shape, slides.Table):
        return {"status": "error",
                "message": f"Table '{shape_name}' not found on slide {slide_idx}"}

    table = shape
    start_row = 0

    if headers:
        for col_idx, header in enumerate(headers):
            if col_idx >= len(table.columns):
                break
            cell = table.rows[0][col_idx]
            tf = _safe_text_frame(cell)
            if tf and tf.paragraphs.count > 0:
                para = tf.paragraphs[0]
                if para.portions.count > 0:
                    para.portions[0].text = str(header)
                else:
                    portion = slides.Portion()
                    portion.text = str(header)
                    para.portions.add(portion)
        start_row = 1

    for row_idx, row_data in enumerate(rows):
        actual_row = row_idx + start_row
        if actual_row >= len(table.rows):
            break  # Can't add rows dynamically in all Aspose versions
        if not isinstance(row_data, list):
            continue
        for col_idx, cell_value in enumerate(row_data):
            if col_idx >= len(table.columns):
                break
            if cell_value is None:
                continue
            cell = table.rows[actual_row][col_idx]
            tf = _safe_text_frame(cell)
            if tf and tf.paragraphs.count > 0:
                para = tf.paragraphs[0]
                if para.portions.count > 0:
                    para.portions[0].text = str(cell_value)
                else:
                    portion = slides.Portion()
                    portion.text = str(cell_value)
                    para.portions.add(portion)

    return {"status": "ok", "slide_idx": slide_idx, "shape": shape_name}


# ---------------------------------------------------------------------------
# Content — EDIT (surgically modify existing slides)
# ---------------------------------------------------------------------------

def edit_run(prs, slide_idx: int, shape_name: str, para_idx: int,
             run_match: str, new_text: str) -> dict:
    """
    Targeted replacement of a single run's text in an EXISTING shape.
    Formatting is preserved because we only modify .text.
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    shape = _find_shape(slide, shape_name)
    tf = _safe_text_frame(shape) if shape else None
    if not tf:
        return {"status": "error",
                "message": f"Shape '{shape_name}' not found or has no text frame on slide {slide_idx}"}

    if para_idx >= tf.paragraphs.count:
        return {"status": "error",
                "message": f"Paragraph {para_idx} out of range (has {tf.paragraphs.count})"}

    para = tf.paragraphs[para_idx]
    normalized_match = _normalize(run_match)
    for portion in para.portions:
        if _normalize(portion.text) == normalized_match:
            portion.text = new_text
            return {"status": "ok", "slide_idx": slide_idx,
                    "shape": shape_name, "matched": run_match}

    return {"status": "error",
            "message": f"No run matching '{run_match}' in paragraph {para_idx} of '{shape_name}'"}


def edit_paragraph(prs, slide_idx: int, shape_name: str, para_idx: int,
                   new_text: str) -> dict:
    """
    Full rewrite of an entire paragraph in an EXISTING shape.
    Writes all new text into run[0], clears remaining runs.
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    shape = _find_shape(slide, shape_name)
    tf = _safe_text_frame(shape) if shape else None
    if not tf:
        return {"status": "error",
                "message": f"Shape '{shape_name}' not found or has no text frame on slide {slide_idx}"}

    if para_idx >= tf.paragraphs.count:
        return {"status": "error",
                "message": f"Paragraph {para_idx} out of range (has {tf.paragraphs.count})"}

    para = tf.paragraphs[para_idx]
    if para.portions.count == 0:
        return {"status": "error",
                "message": f"Paragraph {para_idx} has no runs"}

    # Write everything into run[0], clear the rest
    para.portions[0].text = new_text
    for i in range(1, para.portions.count):
        para.portions[i].text = ""

    return {"status": "ok", "slide_idx": slide_idx,
            "shape": shape_name, "para_idx": para_idx}


def edit_table_cell(prs, slide_idx: int, shape_name: str, row_idx: int,
                    col_idx: int, new_text: str) -> dict:
    """
    Full rewrite of a single table cell in an EXISTING table.
    Preserves cell formatting by only modifying .text.
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    shape = _find_shape(slide, shape_name)
    if not shape or not isinstance(shape, slides.Table):
        return {"status": "error",
                "message": f"Table '{shape_name}' not found on slide {slide_idx}"}

    table = shape
    if row_idx >= len(table.rows) or col_idx >= len(table.columns):
        return {"status": "error",
                "message": f"Cell [{row_idx},{col_idx}] out of range"}

    cell = table.rows[row_idx][col_idx]

    try:
        if cell.is_merged_cell:
            pass  # Allow writing to merged cells; Aspose handles it
    except Exception:
        pass

    tf = _safe_text_frame(cell)
    if tf and tf.paragraphs.count > 0:
        para = tf.paragraphs[0]
        if para.portions.count > 0:
            para.portions[0].text = str(new_text)
            for i in range(1, para.portions.count):
                para.portions[i].text = ""
        else:
            portion = slides.Portion()
            portion.text = str(new_text)
            para.portions.add(portion)

    return {"status": "ok", "slide_idx": slide_idx,
            "shape": shape_name, "cell": [row_idx, col_idx]}


def edit_table_run(prs, slide_idx: int, shape_name: str, row_idx: int,
                   col_idx: int, para_idx: int, run_match: str,
                   new_text: str) -> dict:
    """
    Targeted replacement of a single run within a table cell.
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    shape = _find_shape(slide, shape_name)
    if not shape or not isinstance(shape, slides.Table):
        return {"status": "error",
                "message": f"Table '{shape_name}' not found on slide {slide_idx}"}

    table = shape
    if row_idx >= len(table.rows) or col_idx >= len(table.columns):
        return {"status": "error",
                "message": f"Cell [{row_idx},{col_idx}] out of range"}

    cell = table.rows[row_idx][col_idx]
    tf = _safe_text_frame(cell)
    if not tf:
        return {"status": "error",
                "message": f"Cell [{row_idx},{col_idx}] has no text frame"}

    if para_idx >= tf.paragraphs.count:
        return {"status": "error",
                "message": f"Paragraph {para_idx} out of range in cell [{row_idx},{col_idx}]"}

    para = tf.paragraphs[para_idx]
    normalized_match = _normalize(run_match)
    for portion in para.portions:
        if _normalize(portion.text) == normalized_match:
            portion.text = new_text
            return {"status": "ok", "slide_idx": slide_idx,
                    "shape": shape_name, "cell": [row_idx, col_idx],
                    "matched": run_match}

    return {"status": "error",
            "message": f"No run matching '{run_match}' in cell [{row_idx},{col_idx}]"}


def update_chart(prs, slide_idx: int, shape_name: str, series: dict) -> dict:
    """
    Update chart data series. series is {"Series Name": [val1, val2, ...], ...}
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    shape = _find_shape(slide, shape_name)
    if not shape or not isinstance(shape, charts.Chart):
        return {"status": "error",
                "message": f"Chart '{shape_name}' not found on slide {slide_idx}"}

    chart = shape
    try:
        chart_data = chart.chart_data
        for chart_series in chart_data.series:
            series_name = chart_series.name if chart_series.name else ""
            if series_name in series:
                new_values = series[series_name]
                for i, val in enumerate(new_values):
                    if i < len(chart_series.data_points):
                        chart_series.data_points[i].value.data = val
        return {"status": "ok", "slide_idx": slide_idx, "shape": shape_name}
    except Exception as e:
        return {"status": "error", "message": f"Chart update failed: {str(e)}"}
