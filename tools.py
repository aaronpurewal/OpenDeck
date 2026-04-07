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
try:
    import aspose.pydrawing as drawing
except ImportError:
    drawing = None
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
                try:
                    tf.column_count = 1
                except Exception:
                    pass
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
                pf.alignment = tpf.alignment
            except Exception:
                pass
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
                    pf.bullet.is_bullet_hard_font = slides.NullableBool.TRUE
                    pf.bullet.font = slides.FontData(tpf.bullet.font.font_name)
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
# Chart / Table Constants
# ---------------------------------------------------------------------------

def _inches(n: float) -> float:
    """Convert inches to Aspose coordinate units (points, 1/72 inch).

    Aspose.Slides for Python takes x/y/w/h in points, then multiplies
    by 12,700 to get EMU during OOXML serialization. NOT raw EMU.
    """
    return n * 72.0


_POSITION_SLOTS = {
    "center":      (_inches(0.8), _inches(1.8), _inches(8.4), _inches(4.8)),
    "left_half":   (_inches(0.5), _inches(1.8), _inches(4.5), _inches(4.8)),
    "right_half":  (_inches(5.2), _inches(1.8), _inches(4.5), _inches(4.8)),
    "bottom_half": (_inches(0.5), _inches(3.6), _inches(9.0), _inches(3.2)),
}

_CHART_TYPE_MAP = {
    "clustered_bar":    charts.ChartType.CLUSTERED_BAR,
    "stacked_bar":      charts.ChartType.STACKED_BAR,
    "line":             charts.ChartType.LINE,
    "pie":              charts.ChartType.PIE,
    "doughnut":         charts.ChartType.DOUGHNUT,
    "clustered_column": charts.ChartType.CLUSTERED_COLUMN,
}

_DATAPOINT_METHOD = {
    "clustered_bar":    "add_data_point_for_bar_series",
    "stacked_bar":      "add_data_point_for_bar_series",
    "clustered_column": "add_data_point_for_bar_series",
    "line":             "add_data_point_for_line_series",
    "pie":              "add_data_point_for_pie_series",
    "doughnut":         "add_data_point_for_doughnut_series",
}


def _get_theme_colors(prs) -> list[str]:
    """Extract accent colors 1-6 from master theme as hex strings."""
    colors = []
    try:
        master = prs.masters[0]
        scheme = master.theme_manager.effective_theme.color_scheme
        for attr in ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]:
            try:
                color = getattr(scheme, attr)
                colors.append(f"#{color.r:02x}{color.g:02x}{color.b:02x}")
            except Exception:
                pass
    except Exception:
        pass
    return colors


def _apply_theme_to_chart(chart, theme_colors: list[str]):
    """Apply theme accent colors to chart series fill."""
    if not drawing:
        return
    for i in range(chart.chart_data.series.count):
        if i >= len(theme_colors):
            break
        try:
            series = chart.chart_data.series[i]
            hex_color = theme_colors[i]
            r = int(hex_color[1:3], 16)
            g = int(hex_color[3:5], 16)
            b = int(hex_color[5:7], 16)
            series.format.fill.fill_type = slides.FillType.SOLID
            series.format.fill.solid_fill_color.color = drawing.Color.from_argb(r, g, b)
        except Exception:
            pass


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

    raw_lines = text.split("\n")
    new_paragraphs = []
    header_flags = []  # None=bullet, "H"=plain header, "HB"=bold header
    for line in raw_lines:
        if line.startswith("[HB] "):
            new_paragraphs.append(line[5:])
            header_flags.append("HB")
        elif line.startswith("[H] "):
            new_paragraphs.append(line[4:])
            header_flags.append("H")
        else:
            new_paragraphs.append(line)
            header_flags.append(None)

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
                # Normalize formatting: headers get no bullet; bullets get uniform style
                hflag = header_flags[p_idx] if p_idx < len(header_flags) else None
                if hflag:
                    try:
                        pf = para.paragraph_format
                        pf.bullet.type = slides.BulletType.NONE
                        pf.margin_left = 0
                        pf.indent = 0
                        if template_para:
                            try:
                                pf.alignment = template_para.paragraph_format.alignment
                            except Exception:
                                pass
                    except Exception:
                        pass
                    if hflag == "HB" and para.portions.count > 0:
                        try:
                            para.portions[0].portion_format.font_bold = slides.NullableBool.TRUE
                        except Exception:
                            pass
                elif p_idx > 0 and template_para:
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
            hflag = header_flags[p_idx] if p_idx < len(header_flags) else None
            if hflag:
                try:
                    pf = new_para.paragraph_format
                    pf.bullet.type = slides.BulletType.NONE
                    pf.margin_left = 0
                    pf.indent = 0
                    if template_para:
                        try:
                            pf.alignment = template_para.paragraph_format.alignment
                        except Exception:
                            pass
                except Exception:
                    pass
                if hflag == "HB":
                    try:
                        portion.portion_format.font_bold = slides.NullableBool.TRUE
                    except Exception:
                        pass
            elif template_para:
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


# ---------------------------------------------------------------------------
# Content — SHAPE GEOMETRY (move, swap, recolor decorations and overlays)
# ---------------------------------------------------------------------------

def move_shape(prs, slide_idx: int, shape_name: str,
               x: float = None, y: float = None,
               dx: float = None, dy: float = None) -> dict:
    """
    Move a shape on a slide. Absolute when x/y provided, relative when dx/dy.

    Coordinates are in points (1/72 inch). Same units Aspose uses for
    shape.x / shape.y. Returns dict with status.
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error",
                "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    shape = _find_shape(slide, shape_name)
    if not shape:
        return {"status": "error",
                "message": f"Shape '{shape_name}' not found on slide {slide_idx}"}
    try:
        old_x, old_y = shape.x, shape.y
    except Exception as e:
        return {"status": "error",
                "message": f"Could not read shape position: {str(e)}"}

    new_x = x if x is not None else (old_x + dx if dx is not None else old_x)
    new_y = y if y is not None else (old_y + dy if dy is not None else old_y)

    try:
        shape.x = new_x
    except Exception:
        pass
    try:
        shape.y = new_y
    except Exception:
        pass
    return {"status": "ok", "slide_idx": slide_idx, "shape": shape_name,
            "from": {"x": old_x, "y": old_y},
            "to": {"x": new_x, "y": new_y}}


def swap_shape_positions(prs, slide_idx: int,
                         shape_name_a: str, shape_name_b: str) -> dict:
    """
    Atomically swap the (x, y) positions of two shapes on a slide.

    Each shape keeps its own width and height; only the top-left point
    is exchanged. This is the cleanest primitive for "these two icons
    trade places" intent.
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error",
                "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    shape_a = _find_shape(slide, shape_name_a)
    shape_b = _find_shape(slide, shape_name_b)
    if not shape_a:
        return {"status": "error",
                "message": f"Shape '{shape_name_a}' not found on slide {slide_idx}"}
    if not shape_b:
        return {"status": "error",
                "message": f"Shape '{shape_name_b}' not found on slide {slide_idx}"}
    try:
        ax, ay = shape_a.x, shape_a.y
        bx, by = shape_b.x, shape_b.y
    except Exception as e:
        return {"status": "error",
                "message": f"Could not read shape positions: {str(e)}"}

    try:
        shape_a.x, shape_a.y = bx, by
    except Exception:
        pass
    try:
        shape_b.x, shape_b.y = ax, ay
    except Exception:
        pass
    return {"status": "ok", "slide_idx": slide_idx,
            "swapped": [shape_name_a, shape_name_b]}


def set_shape_fill(prs, slide_idx: int, shape_name: str,
                   color_hex: str) -> dict:
    """
    Recolor a shape's solid fill (RAG status changes, severity flips).

    color_hex is "#RRGGBB" or "RRGGBB". Returns dict with status.
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error",
                "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    shape = _find_shape(slide, shape_name)
    if not shape:
        return {"status": "error",
                "message": f"Shape '{shape_name}' not found on slide {slide_idx}"}

    hex_str = color_hex.lstrip("#")
    if len(hex_str) != 6:
        return {"status": "error",
                "message": f"Invalid color_hex '{color_hex}', expected #RRGGBB"}
    try:
        r = int(hex_str[0:2], 16)
        g = int(hex_str[2:4], 16)
        b = int(hex_str[4:6], 16)
    except ValueError:
        return {"status": "error",
                "message": f"Invalid hex digits in '{color_hex}'"}

    try:
        fill = shape.fill_format
        fill.fill_type = slides.FillType.SOLID
        if drawing is not None:
            fill.solid_fill_color.color = drawing.Color.from_argb(255, r, g, b)
        else:
            return {"status": "error",
                    "message": "drawing module not available for color setting"}
        return {"status": "ok", "slide_idx": slide_idx, "shape": shape_name,
                "color": f"#{hex_str.lower()}"}
    except Exception as e:
        return {"status": "error",
                "message": f"Failed to set fill: {str(e)}"}


def swap_table_rows(prs, slide_idx: int, shape_name: str,
                    row_idx_a: int, row_idx_b: int) -> dict:
    """
    Atomic high-level row swap for tables with overlay shapes.

    Swaps cell text content between row_idx_a and row_idx_b (per column),
    AND moves any overlay shapes (icons, status dots, harvey balls, logos)
    whose centers fall inside either row's vertical band so they follow
    the content.

    Cell formatting is preserved by editing portion text in place rather
    than replacing the cells.
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error",
                "message": f"Slide index {slide_idx} out of range"}
    slide = prs.slides[slide_idx]
    shape = _find_shape(slide, shape_name)
    if not shape or not isinstance(shape, slides.Table):
        return {"status": "error",
                "message": f"Table '{shape_name}' not found on slide {slide_idx}"}

    table = shape
    n_rows = len(table.rows)
    if row_idx_a < 0 or row_idx_a >= n_rows:
        return {"status": "error",
                "message": f"row_idx_a {row_idx_a} out of range (0..{n_rows - 1})"}
    if row_idx_b < 0 or row_idx_b >= n_rows:
        return {"status": "error",
                "message": f"row_idx_b {row_idx_b} out of range (0..{n_rows - 1})"}
    if row_idx_a == row_idx_b:
        return {"status": "ok", "swapped_cells": 0, "moved_shapes": [],
                "message": "row_idx_a equals row_idx_b, no-op"}

    # --- Step 1: swap cell text per column ---
    swapped_cells = 0
    n_cols = len(table.columns)
    for col_idx in range(n_cols):
        try:
            cell_a = table.rows[row_idx_a][col_idx]
            cell_b = table.rows[row_idx_b][col_idx]
            tf_a = _safe_text_frame(cell_a)
            tf_b = _safe_text_frame(cell_b)
            if not tf_a or not tf_b:
                continue
            text_a = ""
            text_b = ""
            try:
                text_a = tf_a.text or ""
            except Exception:
                pass
            try:
                text_b = tf_b.text or ""
            except Exception:
                pass
            # Write text_b into cell_a, text_a into cell_b
            try:
                tf_a.text = text_b
            except Exception:
                pass
            try:
                tf_b.text = text_a
            except Exception:
                pass
            swapped_cells += 1
        except Exception:
            continue

    # --- Step 2: compute y-bands for both rows ---
    try:
        table_y = shape.y
    except Exception:
        return {"status": "ok", "swapped_cells": swapped_cells,
                "moved_shapes": [],
                "message": "Could not read table y for overlay anchoring"}

    y_cursor = table_y
    row_y = {}  # row_idx -> (y, h)
    try:
        for r in range(n_rows):
            h = table.rows[r].height
            row_y[r] = (y_cursor, h)
            y_cursor += h
    except Exception:
        return {"status": "ok", "swapped_cells": swapped_cells,
                "moved_shapes": [],
                "message": "Could not compute row bands for overlay anchoring"}

    a_y, a_h = row_y[row_idx_a]
    b_y, b_h = row_y[row_idx_b]

    # --- Step 3: collect overlay shapes in either row's band, then apply ---
    # IMPORTANT: collect first, then apply. Otherwise the second pass would
    # re-detect already-moved shapes in their new positions.
    try:
        table_x = shape.x
        table_w = shape.width
    except Exception:
        return {"status": "ok", "swapped_cells": swapped_cells,
                "moved_shapes": []}

    moves = []  # list of (shape_obj, new_x, new_y)
    for s in slide.shapes:
        if s is shape or isinstance(s, slides.Table):
            continue
        try:
            sx, sy = s.x, s.y
            sw, sh = s.width, s.height
        except Exception:
            continue
        cx = sx + sw / 2
        cy = sy + sh / 2
        # Must overlap table horizontally (with small slack)
        if cx < table_x - 36 or cx > table_x + table_w + 36:
            continue
        # Determine which row band the center falls in
        if a_y <= cy <= a_y + a_h:
            # Move from row A to row B
            new_y = sy + (b_y - a_y)
            moves.append((s, sx, new_y))
        elif b_y <= cy <= b_y + b_h:
            new_y = sy + (a_y - b_y)
            moves.append((s, sx, new_y))

    moved_names = []
    for s, nx, ny in moves:
        try:
            s.x = nx
        except Exception:
            pass
        try:
            s.y = ny
        except Exception:
            pass
        try:
            moved_names.append(s.name)
        except Exception:
            pass

    return {"status": "ok", "slide_idx": slide_idx, "shape": shape_name,
            "swapped_cells": swapped_cells, "moved_shapes": moved_names}


# ---------------------------------------------------------------------------
# Content — CREATE (new charts and tables from scratch)
# ---------------------------------------------------------------------------

def create_chart(prs, slide_idx: int, chart_type: str, title: str,
                 categories: list[str], series: list[dict],
                 position: str = "center") -> dict:
    """
    Create a chart on the given slide.

    series = [{"name": "Revenue", "values": [100, 200, 300]}, ...]
    Returns {"status": "ok", "shape_name": "...", "chart_type": "..."} or error dict.
    """
    if chart_type not in _CHART_TYPE_MAP:
        return {"status": "error",
                "message": f"Unknown chart type: {chart_type}. "
                           f"Allowed: {list(_CHART_TYPE_MAP.keys())}"}
    if position not in _POSITION_SLOTS:
        return {"status": "error",
                "message": f"Unknown position: {position}. "
                           f"Allowed: {list(_POSITION_SLOTS.keys())}"}
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide index {slide_idx} out of range"}

    slide = prs.slides[slide_idx]
    ct = _CHART_TYPE_MAP[chart_type]
    x, y, w, h = _POSITION_SLOTS[position]

    try:
        # Use False to avoid stale sample data in the embedded workbook —
        # PowerPoint cross-validates chart XML against the workbook and
        # leftover ghost data from True + clear() causes corruption.
        chart_obj = slide.shapes.add_chart(ct, x, y, w, h, False)

        wb = chart_obj.chart_data.chart_data_workbook
        try:
            wb.clear(0)
        except Exception:
            pass

        # Add categories
        for i, cat_name in enumerate(categories):
            chart_obj.chart_data.categories.add(wb.get_cell(0, i + 1, 0, cat_name))

        # Add series with appropriate datapoint method
        dp_method = _DATAPOINT_METHOD.get(chart_type, "add_data_point_for_bar_series")
        for s_idx, s_data in enumerate(series):
            ser = chart_obj.chart_data.series.add(
                wb.get_cell(0, 0, s_idx + 1, s_data["name"]), ct
            )
            for i, val in enumerate(s_data["values"]):
                cell = wb.get_cell(0, i + 1, s_idx + 1, val)
                getattr(ser.data_points, dp_method)(cell)

        # Set chart title — use add_text_frame_for_overriding only if title
        # is provided, then set overlay=False for standard PowerPoint layout
        if title:
            try:
                chart_obj.has_title = True
                chart_obj.chart_title.add_text_frame_for_overriding(title)
                chart_obj.chart_title.overlay = False
            except Exception:
                pass
        else:
            try:
                chart_obj.has_title = False
            except Exception:
                pass

        # Apply theme colors
        theme_colors = _get_theme_colors(prs)
        if theme_colors:
            _apply_theme_to_chart(chart_obj, theme_colors)

        return {"status": "ok", "shape_name": chart_obj.name,
                "chart_type": chart_type}
    except Exception as e:
        return {"status": "error", "message": f"Chart creation failed: {str(e)}"}


def create_table(prs, slide_idx: int, headers: list[str],
                 rows: list[list[str]], position: str = "center",
                 col_widths: list[float] | None = None) -> dict:
    """
    Create a table on the given slide.

    col_widths: optional list of column widths in inches.
    Returns {"status": "ok", "shape_name": "...", "rows": N, "cols": M} or error dict.
    """
    if position not in _POSITION_SLOTS:
        return {"status": "error",
                "message": f"Unknown position: {position}. "
                           f"Allowed: {list(_POSITION_SLOTS.keys())}"}
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide index {slide_idx} out of range"}

    slide = prs.slides[slide_idx]
    x, y, w, h = _POSITION_SLOTS[position]
    n_cols = len(headers)
    n_rows = len(rows) + 1  # +1 for header row

    if col_widths:
        col_widths_pts = [_inches(cw) for cw in col_widths]
    else:
        col_widths_pts = [w / n_cols] * n_cols

    row_height = min(_inches(0.4), h / n_rows)
    row_heights_pts = [row_height] * n_rows

    try:
        table = slide.shapes.add_table(x, y, col_widths_pts, row_heights_pts)

        # Populate header row with bold formatting
        for col_idx, header_text in enumerate(headers):
            if col_idx >= len(table.columns):
                break
            cell = table.rows[0][col_idx]
            tf = _safe_text_frame(cell)
            if tf and tf.paragraphs.count > 0:
                para = tf.paragraphs[0]
                if para.portions.count > 0:
                    para.portions[0].text = str(header_text)
                    try:
                        para.portions[0].portion_format.font_bold = slides.NullableBool.TRUE
                    except Exception:
                        pass
                else:
                    portion = slides.Portion()
                    portion.text = str(header_text)
                    try:
                        portion.portion_format.font_bold = slides.NullableBool.TRUE
                    except Exception:
                        pass
                    para.portions.add(portion)

            # Apply theme color to header cell background
            theme_colors = _get_theme_colors(prs)
            if theme_colors and drawing:
                try:
                    hex_color = theme_colors[0]
                    r = int(hex_color[1:3], 16)
                    g = int(hex_color[3:5], 16)
                    b = int(hex_color[5:7], 16)
                    cell.fill_format.fill_type = slides.FillType.SOLID
                    cell.fill_format.solid_fill_color.color = drawing.Color.from_argb(r, g, b)
                except Exception:
                    pass

        # Populate data rows
        for row_idx, row_data in enumerate(rows):
            actual_row = row_idx + 1  # skip header
            if actual_row >= len(table.rows):
                break
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

        return {"status": "ok", "shape_name": table.name,
                "rows": n_rows, "cols": n_cols}
    except Exception as e:
        return {"status": "error", "message": f"Table creation failed: {str(e)}"}
