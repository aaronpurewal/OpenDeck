"""
State Layer: Deck harvesting and state utilities.

Responsible for extracting a complete JSON-serializable representation
of a PowerPoint deck via Aspose.Slides. This state dict is the LLM's
"working memory" — it never sees the PPTX file directly.
"""

import math
import aspose.slides as slides
import aspose.slides.charts as charts
from config import CHAR_LIMIT_SAFETY_MARGIN, DEFAULT_FONT_SIZE_PT, DEFAULT_LINE_SPACING


def estimate_char_limit(width_pt: float, height_pt: float,
                        font_size_pt: float = None,
                        line_spacing: float = DEFAULT_LINE_SPACING) -> int:
    """
    Conservative character limit estimate based on shape dimensions.

    Aspose.Slides for Python returns dimensions in points (1 pt = 1/72 inch).
    Font sizes from portion_format.font_height are also in points.
    """
    if not font_size_pt or font_size_pt <= 0 or math.isnan(font_size_pt):
        font_size_pt = DEFAULT_FONT_SIZE_PT

    avg_char_width_pt = font_size_pt * 0.6
    line_height_pt = font_size_pt * line_spacing

    if avg_char_width_pt <= 0 or line_height_pt <= 0:
        return 50

    chars_per_line = int(width_pt / avg_char_width_pt)
    num_lines = int(height_pt / line_height_pt)
    return max(int(chars_per_line * num_lines * CHAR_LIMIT_SAFETY_MARGIN), 1)


def _safe_text_frame(shape):
    """Safely get text_frame from a shape. Returns None if not available."""
    try:
        tf = shape.text_frame
        if tf is not None:
            return tf
    except Exception:
        pass
    return None


def _safe_effective_format(portion_format):
    """
    Get the effective (resolved) portion format via get_effective().

    Resolves the full master → layout → slide inheritance chain so we get
    the actual rendered values, not just locally-set overrides.
    Returns None if get_effective() is unavailable or fails.
    Uses BaseException to catch .NET proxy errors (RuntimeError subclasses).
    """
    try:
        return portion_format.get_effective()
    except BaseException:
        return None


def _safe_font_height(portion_format):
    """Get font_height, returning 0 if NaN or unavailable.

    Tries get_effective() first for resolved inherited values,
    falls back to raw portion_format.
    """
    eff = _safe_effective_format(portion_format)
    if eff is not None:
        try:
            fh = eff.font_height
            if fh is not None and not math.isnan(fh):
                return fh
        except BaseException:
            pass
    try:
        fh = portion_format.font_height
        if fh is not None and not math.isnan(fh):
            return fh
    except BaseException:
        pass
    return 0


def _safe_font_name(portion_format):
    """Get latin font name as string, or None.

    Tries get_effective() first for resolved inherited values.
    """
    eff = _safe_effective_format(portion_format)
    if eff is not None:
        try:
            lf = eff.latin_font
            if lf is not None:
                return str(lf)
        except BaseException:
            pass
    try:
        lf = portion_format.latin_font
        if lf is not None:
            return str(lf)
    except BaseException:
        pass
    return None


def _safe_font_bold(portion_format):
    """Get bold state. Returns True/False/None.

    Tries get_effective() first — effective format returns a direct
    boolean rather than NullableBool, giving resolved inherited values.
    """
    eff = _safe_effective_format(portion_format)
    if eff is not None:
        try:
            return bool(eff.font_bold)
        except BaseException:
            pass
    try:
        val = portion_format.font_bold
        if val == slides.NullableBool.TRUE:
            return True
        elif val == slides.NullableBool.FALSE:
            return False
    except BaseException:
        pass
    return None


def _safe_font_italic(portion_format):
    """Get italic state. Returns True/False/None.

    Tries get_effective() first for resolved inherited values.
    """
    eff = _safe_effective_format(portion_format)
    if eff is not None:
        try:
            return bool(eff.font_italic)
        except BaseException:
            pass
    try:
        val = portion_format.font_italic
        if val == slides.NullableBool.TRUE:
            return True
        elif val == slides.NullableBool.FALSE:
            return False
    except BaseException:
        pass
    return None


def extract_shape(shape) -> dict | None:
    """Extract metadata from a single shape. Returns None for unsupported types."""
    base = {
        "name": shape.name,
        "bounds": {
            "x": shape.x, "y": shape.y,
            "w": shape.width, "h": shape.height
        }
    }

    # --- Text shapes ---
    tf = _safe_text_frame(shape)
    if tf is not None:
        base["type"] = "text"
        try:
            base["text"] = tf.text or ""
        except Exception:
            base["text"] = ""
        base["paragraphs"] = []
        max_font_size = 0

        for para in tf.paragraphs:
            para_info = {"runs": []}
            for portion in para.portions:
                pf = portion.portion_format
                fh = _safe_font_height(pf)
                if fh > max_font_size:
                    max_font_size = fh
                run_info = {
                    "text": portion.text or "",
                    "bold": _safe_font_bold(pf),
                    "italic": _safe_font_italic(pf),
                    "font_size": fh,
                    "font_name": _safe_font_name(pf),
                }
                para_info["runs"].append(run_info)
            base["paragraphs"].append(para_info)

        # Reclassify empty-text autoshapes as decorations.
        # An ellipse, arrow, or callout with no text content is a
        # decoration (RAG dot, severity icon, etc.), not a text shape.
        # Rectangle/text-box autoshapes with empty text are kept as
        # text since they're typically empty content placeholders.
        if not (base["text"] or "").strip():
            try:
                st_name = ""
                st_val = None
                if hasattr(shape, "shape_type"):
                    st_val = shape.shape_type
                    try:
                        st_name = str(st_val.name) if hasattr(st_val, "name") else str(st_val)
                    except Exception:
                        st_name = str(st_val)
                # Shape types that are text containers (stay as "text"):
                _TEXT_CONTAINER_NAMES = {
                    "RECTANGLE", "ROUND_CORNER_RECTANGLE",
                    "TEXT_BOX", "ONE_ROUND_CORNER_RECTANGLE",
                    "TWO_ROUND_CORNER_RECTANGLE",
                    "SNIP_ROUND_RECTANGLE", "TWO_SAMESIDE_ROUND_CORNER_RECTANGLE",
                    "TWO_DIAG_ROUND_CORNER_RECTANGLE",
                }
                if st_name and st_name.upper() not in _TEXT_CONTAINER_NAMES:
                    # This is a decorative shape with no text content
                    base["type"] = "decoration"
                    base["subtype"] = type(shape).__name__
                    base["auto_shape_type"] = st_name
                    try:
                        fill = shape.fill_format
                        if fill is not None and fill.fill_type == slides.FillType.SOLID:
                            c = fill.solid_fill_color.color
                            base["fill_hex"] = f"#{c.r:02x}{c.g:02x}{c.b:02x}"
                    except BaseException:
                        pass
                    base.pop("paragraphs", None)
                    base.pop("text", None)
                    return base
            except Exception:
                pass

        base["char_limit"] = estimate_char_limit(
            shape.width, shape.height, font_size_pt=max_font_size
        )
        return base

    # --- Table shapes ---
    if isinstance(shape, slides.Table):
        try:
            table = shape
            base["type"] = "table"
            base["row_count"] = len(table.rows)
            base["col_count"] = len(table.columns)
            base["rows"] = []
            for row_idx in range(len(table.rows)):
                row = table.rows[row_idx]
                row_data = []
                for col_idx in range(len(table.columns)):
                    cell = row[col_idx]
                    cell_text = ""
                    is_merged = False
                    cell_tf = _safe_text_frame(cell)
                    if cell_tf:
                        try:
                            cell_text = cell_tf.text or ""
                        except Exception:
                            cell_text = ""
                    try:
                        is_merged = cell.is_merged_cell
                    except Exception:
                        pass
                    cell_info = {
                        "text": cell_text,
                        "is_merged": is_merged,
                    }
                    row_data.append(cell_info)
                base["rows"].append(row_data)

            # Compute per-row absolute y-bands for overlay anchoring.
            # Aspose exposes row.height but not row.y, so we compute it
            # cumulatively from table.y.
            base["row_bounds"] = []
            try:
                y_cursor = shape.y
                for row_idx in range(len(table.rows)):
                    h = table.rows[row_idx].height
                    base["row_bounds"].append({"y": y_cursor, "h": h})
                    y_cursor += h
            except Exception:
                pass

            if len(table.rows) > 0 and len(table.columns) > 0:
                base["cell_char_limit"] = estimate_char_limit(
                    table.columns[0].width, table.rows[0].height
                )
            else:
                base["cell_char_limit"] = 50
            return base
        except Exception:
            base["type"] = "table"
            base["row_count"] = 0
            base["col_count"] = 0
            base["rows"] = []
            base["cell_char_limit"] = 50
            return base

    # --- Group shapes (designer-grouped icon+caption etc.) ---
    if isinstance(shape, slides.GroupShape):
        try:
            base["type"] = "group"
            base["children"] = []
            for child in shape.shapes:
                try:
                    base["children"].append(child.name)
                except Exception:
                    pass
            return base
        except Exception:
            base["type"] = "group"
            base["children"] = []
            return base

    # --- Chart shapes ---
    if isinstance(shape, charts.Chart):
        try:
            chart = shape
            base["type"] = "chart"
            base["chart_type"] = str(chart.type) if chart.type else "unknown"
            base["series"] = []
            if chart.chart_data and chart.chart_data.series:
                for series in chart.chart_data.series:
                    series_info = {
                        "name": series.name if series.name else "",
                        "values": []
                    }
                    if series.data_points:
                        for dp in series.data_points:
                            try:
                                series_info["values"].append(dp.value)
                            except Exception:
                                series_info["values"].append(None)
                    base["series"].append(series_info)
            base["categories"] = []
            if chart.chart_data and chart.chart_data.categories:
                for cat in chart.chart_data.categories:
                    try:
                        base["categories"].append(str(cat.label) if cat.label else "")
                    except Exception:
                        base["categories"].append("")
            return base
        except Exception:
            base["type"] = "chart"
            base["chart_type"] = "unknown"
            base["series"] = []
            base["categories"] = []
            return base

    # --- Decoration fall-through (auto shapes, picture frames, ovals,
    #     RAG dots, harvey balls, logos, callouts, freeform shapes) ---
    try:
        base["type"] = "decoration"
        try:
            base["subtype"] = type(shape).__name__
        except Exception:
            base["subtype"] = "unknown"
        try:
            if hasattr(shape, "auto_shape_type"):
                base["auto_shape_type"] = str(shape.auto_shape_type)
        except Exception:
            pass
        # Solid fill color (RAG status badges, severity dots, etc.)
        try:
            fill = shape.fill_format
            if fill is not None and fill.fill_type == slides.FillType.SOLID:
                c = fill.solid_fill_color.color
                base["fill_hex"] = f"#{c.r:02x}{c.g:02x}{c.b:02x}"
        except BaseException:
            pass
        return base
    except Exception:
        return None


def _walk_slide_shapes(slide_shapes, parent_group: str = None) -> list:
    """
    Walk a slide's shape tree and yield extracted shape dicts.

    Recurses into GroupShape children, recording the parent group's name on
    each child so the LLM can decide whether to move the group as one unit
    or address children individually.
    """
    out = []
    for shape in slide_shapes:
        extracted = extract_shape(shape)
        if extracted:
            if parent_group:
                extracted["parent_group"] = parent_group
            out.append(extracted)
        # Recurse into group children
        if isinstance(shape, slides.GroupShape):
            try:
                group_name = shape.name
                out.extend(_walk_slide_shapes(shape.shapes, parent_group=group_name))
            except Exception:
                pass
    return out


def _associate_overlays(slide_shapes: list, slide_w: float = None,
                        slide_h: float = None) -> None:
    """
    Anchor decoration shapes to table rows or text bullet paragraphs.

    Mutates slide_shapes in place. After this runs:
    - Each decoration that overlays a table row gets `anchor` =
      {"kind": "table_row", "shape": <table_name>, "row_idx": N}
    - Each decoration that overlays a text bullet gets `anchor` =
      {"kind": "text_paragraph", "shape": <text_name>, "para_idx": N}
    - Each table with overlays gets `row_overlays` = {row_idx: [names]}
    - Each text shape with overlays gets `para_overlays` = {para_idx: [names]}

    Background candidates (shapes spanning >60% of slide dims) are skipped
    to avoid anchoring section dividers, full-bleed rectangles, or chrome.
    """
    tables = [s for s in slide_shapes if s.get("type") == "table"]
    texts = [s for s in slide_shapes if s.get("type") == "text"]
    decorations = [s for s in slide_shapes if s.get("type") == "decoration"]

    for deco in decorations:
        bounds = deco.get("bounds", {})
        dx, dy = bounds.get("x", 0), bounds.get("y", 0)
        dw, dh = bounds.get("w", 0), bounds.get("h", 0)
        if dw <= 0 or dh <= 0:
            continue

        # Skip background candidates (full-bleed shapes)
        if slide_w and slide_h:
            if dw > 0.6 * slide_w or dh > 0.6 * slide_h:
                continue

        cx, cy = dx + dw / 2, dy + dh / 2

        # First try table row anchoring
        anchored = False
        for table in tables:
            tb = table.get("bounds", {})
            tx, ty = tb.get("x", 0), tb.get("y", 0)
            tw, th = tb.get("w", 0), tb.get("h", 0)
            if tw <= 0 or th <= 0:
                continue
            # Center must be inside the table's bbox (with small horizontal slack)
            if not (tx - 0.5 * 72 <= cx <= tx + tw + 0.5 * 72):
                continue
            if not (ty <= cy <= ty + th):
                continue
            row_bounds = table.get("row_bounds", [])
            for row_idx, rb in enumerate(row_bounds):
                if rb["y"] <= cy <= rb["y"] + rb["h"]:
                    deco["anchor"] = {
                        "kind": "table_row",
                        "shape": table["name"],
                        "row_idx": row_idx,
                    }
                    table.setdefault("row_overlays", {}).setdefault(
                        str(row_idx), []
                    ).append(deco["name"])
                    anchored = True
                    break
            if anchored:
                break

        if anchored:
            continue

        # Fall back to text paragraph anchoring
        for text in texts:
            tb = text.get("bounds", {})
            tx, ty = tb.get("x", 0), tb.get("y", 0)
            tw, th = tb.get("w", 0), tb.get("h", 0)
            if tw <= 0 or th <= 0:
                continue
            # Decoration must be horizontally near (within ~1") and
            # vertically intersect the text shape's bbox
            if cx < tx - 72 or cx > tx + tw + 72:
                continue
            if cy < ty or cy > ty + th:
                continue
            paragraphs = text.get("paragraphs", [])
            if not paragraphs:
                continue
            # Best-effort paragraph anchoring by vertical offset
            para_h = th / max(len(paragraphs), 1)
            offset = cy - ty
            para_idx = min(int(offset / para_h), len(paragraphs) - 1)
            deco["anchor"] = {
                "kind": "text_paragraph",
                "shape": text["name"],
                "para_idx": para_idx,
            }
            text.setdefault("para_overlays", {}).setdefault(
                str(para_idx), []
            ).append(deco["name"])
            break


def harvest_deck(prs: slides.Presentation) -> dict:
    """
    Extract full state from the presentation.

    Returns a JSON-serializable dict containing:
    - slide_count: total number of slides
    - master_layouts: available layouts with placeholder metadata
    - label_list: ordered list of slide labels (position = Aspose index)
    - slides: per-slide shape inventory with text, formatting, bounds, char_limits
    """
    state = {
        "slide_count": len(prs.slides),
        "master_layouts": [],
        "slides": [],
        "label_list": []
    }

    # Walk masters — extract shapes from each layout
    for master in prs.masters:
        for layout in master.layout_slides:
            layout_info = {
                "name": layout.name,
                "shapes": []
            }
            # Iterate layout shapes and find placeholders
            for shape in layout.shapes:
                placeholder = shape.placeholder
                if placeholder is None:
                    continue
                shape_info = {
                    "name": shape.name,
                    "type": "placeholder",
                    "placeholder_idx": placeholder.index,
                    "bounds": {
                        "x": shape.x, "y": shape.y,
                        "w": shape.width, "h": shape.height
                    },
                }
                # Get char_limit using actual font size if available
                max_font = 0
                tf = _safe_text_frame(shape)
                if tf:
                    for para in tf.paragraphs:
                        for portion in para.portions:
                            fh = _safe_font_height(portion.portion_format)
                            if fh > max_font:
                                max_font = fh
                shape_info["char_limit"] = estimate_char_limit(
                    shape.width, shape.height, font_size_pt=max_font
                )
                # Describe the formatting pattern
                if tf and tf.paragraphs.count > 0:
                    shape_info["paragraph_count"] = tf.paragraphs.count
                    try:
                        shape_info["default_text"] = (tf.text or "")[:100]
                    except Exception:
                        shape_info["default_text"] = ""
                layout_info["shapes"].append(shape_info)
            state["master_layouts"].append(layout_info)

    # Build layout → slide label mapping, then attach to each layout
    layout_usage = {}  # layout_name → [slide_labels]
    for i in range(len(prs.slides)):
        try:
            ln = prs.slides[i].layout_slide.name if prs.slides[i].layout_slide else None
        except Exception:
            ln = None
        if ln:
            layout_usage.setdefault(ln, []).append(f"slide_{i}")
    for layout_info in state["master_layouts"]:
        layout_info["used_by"] = layout_usage.get(layout_info["name"], [])

    # Get slide dimensions for background-shape detection
    try:
        slide_w = prs.slide_size.size.width
        slide_h = prs.slide_size.size.height
    except Exception:
        slide_w, slide_h = None, None

    # Walk slides — each gets a stable label
    for i in range(len(prs.slides)):
        slide = prs.slides[i]
        label = f"slide_{i}"
        layout_name = "Unknown"
        try:
            if slide.layout_slide:
                layout_name = slide.layout_slide.name
        except Exception:
            pass
        slide_state = {
            "label": label,
            "index": i,
            "layout_name": layout_name,
            "shapes": []
        }
        state["label_list"].append(label)
        # Walk shapes recursively (handles GroupShape children)
        slide_state["shapes"] = _walk_slide_shapes(slide.shapes)
        # Anchor decorations to tables/text after extraction
        _associate_overlays(slide_state["shapes"], slide_w, slide_h)
        state["slides"].append(slide_state)

    return state


def compact_state(state: dict, max_text_chars: int = 500) -> dict:
    """
    Produce a compact version of the deck state for LLM context.

    The full state can be 200K+ tokens for large decks. The LLM only needs:
    - Slide labels, layout names, shape names
    - Text content (truncated to max_text_chars)
    - char_limits
    - Run-level structure for edit_run targeting (text only, no formatting metadata)
    - Table content (truncated)
    - Chart series names and values

    Drops: bounding boxes, font metadata, bold/italic/font_name/font_size on runs,
    paragraph_count, default_text on layouts.
    """
    compact = {
        "slide_count": state["slide_count"],
        "label_list": state["label_list"],
        "master_layouts": [],
        "slides": []
    }

    # Compact layouts — just name + placeholder names and char_limits
    for layout in state.get("master_layouts", []):
        cl = {"name": layout["name"], "shapes": []}
        for shape in layout.get("shapes", []):
            cl["shapes"].append({
                "name": shape["name"],
                "type": shape.get("type", "placeholder"),
                "char_limit": shape.get("char_limit", 0),
            })
        cl["used_by"] = layout.get("used_by", [])
        compact["master_layouts"].append(cl)

    # Compact slides
    for slide in state.get("slides", []):
        cs = {
            "label": slide["label"],
            "index": slide["index"],
            "layout_name": slide["layout_name"],
            "shapes": []
        }
        for shape in slide.get("shapes", []):
            s_type = shape.get("type", "")

            if s_type == "text":
                text = shape.get("text", "")
                compact_shape = {
                    "name": shape["name"],
                    "type": "text",
                    "char_limit": shape.get("char_limit", 0),
                    "text": text[:max_text_chars] + ("..." if len(text) > max_text_chars else ""),
                }
                # Include run texts for edit_run targeting — just the text, no formatting
                runs_summary = []
                for pi, para in enumerate(shape.get("paragraphs", [])):
                    run_texts = [r["text"] for r in para.get("runs", []) if r.get("text")]
                    if run_texts:
                        runs_summary.append({"p": pi, "runs": run_texts})
                if runs_summary:
                    compact_shape["paragraphs"] = runs_summary
                if shape.get("para_overlays"):
                    compact_shape["para_overlays"] = shape["para_overlays"]
                if shape.get("parent_group"):
                    compact_shape["parent_group"] = shape["parent_group"]
                cs["shapes"].append(compact_shape)

            elif s_type == "table":
                compact_shape = {
                    "name": shape["name"],
                    "type": "table",
                    "row_count": shape.get("row_count", 0),
                    "col_count": shape.get("col_count", 0),
                    "cell_char_limit": shape.get("cell_char_limit", 50),
                }
                # Include first few rows of data for context
                rows = shape.get("rows", [])
                compact_rows = []
                for row in rows[:8]:  # Cap at 8 rows
                    compact_row = []
                    for cell in row:
                        cell_text = cell.get("text", "") if isinstance(cell, dict) else str(cell)
                        compact_row.append(cell_text[:60])
                    compact_rows.append(compact_row)
                if compact_rows:
                    compact_shape["rows"] = compact_rows
                if len(rows) > 8:
                    compact_shape["total_rows"] = len(rows)
                if shape.get("row_overlays"):
                    compact_shape["row_overlays"] = shape["row_overlays"]
                cs["shapes"].append(compact_shape)

            elif s_type == "chart":
                compact_shape = {
                    "name": shape["name"],
                    "type": "chart",
                    "chart_type": shape.get("chart_type", "unknown"),
                    "series": shape.get("series", []),
                    "categories": shape.get("categories", []),
                }
                cs["shapes"].append(compact_shape)

            elif s_type == "decoration":
                compact_shape = {
                    "name": shape["name"],
                    "type": "decoration",
                }
                if shape.get("subtype"):
                    compact_shape["subtype"] = shape["subtype"]
                if shape.get("auto_shape_type"):
                    compact_shape["auto_shape_type"] = shape["auto_shape_type"]
                if shape.get("fill_hex"):
                    compact_shape["fill_hex"] = shape["fill_hex"]
                if shape.get("anchor"):
                    compact_shape["anchor"] = shape["anchor"]
                if shape.get("parent_group"):
                    compact_shape["parent_group"] = shape["parent_group"]
                cs["shapes"].append(compact_shape)

            elif s_type == "group":
                compact_shape = {
                    "name": shape["name"],
                    "type": "group",
                    "children": shape.get("children", []),
                }
                cs["shapes"].append(compact_shape)

            else:
                cs["shapes"].append({
                    "name": shape["name"],
                    "type": s_type,
                })

        compact["slides"].append(cs)

    return compact
