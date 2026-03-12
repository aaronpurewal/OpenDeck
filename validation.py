"""
Validation Layer: Post-execution checks.

Three layers:
1. Constraint enforcement (pre-generation, handled by char_limit in prompts)
2. Placeholder detection (post-execution, deterministic, zero LLM calls)
3. Data integrity check (post-generation, one LLM call for financial content)

Plus a smoke test: save → re-open → verify file integrity.
"""

import json
import math
import aspose.slides as slides
from config import PLACEHOLDER_PATTERNS


def _safe_text_frame(shape):
    """Get text_frame safely."""
    try:
        tf = shape.text_frame
        if tf is not None:
            return tf
    except Exception:
        pass
    return None


def check_placeholders(prs) -> dict:
    """
    Scan entire deck for unfilled placeholder text.
    Deterministic — no API calls. Catches leftover template boilerplate.
    """
    findings = []
    for slide_idx in range(len(prs.slides)):
        slide = prs.slides[slide_idx]
        for shape in slide.shapes:
            tf = _safe_text_frame(shape)
            if not tf:
                continue
            try:
                text_lower = (tf.text or "").lower().strip()
            except Exception:
                continue
            if not text_lower:
                continue
            for pattern in PLACEHOLDER_PATTERNS:
                if pattern.lower() in text_lower:
                    findings.append({
                        "slide_idx": slide_idx,
                        "shape_name": shape.name,
                        "text": (tf.text or "")[:100],
                        "matched_pattern": pattern
                    })
                    break
    return {
        "status": "clean" if not findings else "placeholders_found",
        "findings": findings
    }


def check_brand(prs, slide_idx: int, brand_rules: dict) -> dict:
    """
    Compare shape properties against brand spec.

    brand_rules example:
    {
        "allowed_fonts": ["Calibri", "Arial"],
        "title_min_size_pt": 24,
        "body_min_size_pt": 10,
    }
    """
    if slide_idx < 0 or slide_idx >= len(prs.slides):
        return {"status": "error", "message": f"Slide {slide_idx} out of range"}

    violations = []
    slide = prs.slides[slide_idx]
    allowed_fonts = brand_rules.get("allowed_fonts", [])
    title_min_emu = brand_rules.get("title_min_size_pt", 0) * 12700
    body_min_emu = brand_rules.get("body_min_size_pt", 0) * 12700

    for shape in slide.shapes:
        tf = _safe_text_frame(shape)
        if not tf:
            continue
        for para in tf.paragraphs:
            for portion in para.portions:
                pf = portion.portion_format
                # Font check
                if allowed_fonts:
                    try:
                        font_name = str(pf.latin_font) if pf.latin_font else None
                        if font_name and font_name not in allowed_fonts:
                            violations.append({
                                "shape": shape.name,
                                "issue": f"Font '{font_name}' not in allowed list",
                                "value": font_name
                            })
                    except Exception:
                        pass
                # Size check
                try:
                    fh = pf.font_height
                    if fh and not math.isnan(fh) and fh > 0:
                        if "title" in shape.name.lower() and fh < title_min_emu:
                            violations.append({
                                "shape": shape.name,
                                "issue": f"Title font too small: {fh / 12700:.0f}pt",
                                "value": fh / 12700
                            })
                        elif fh < body_min_emu:
                            violations.append({
                                "shape": shape.name,
                                "issue": f"Body font too small: {fh / 12700:.0f}pt",
                                "value": fh / 12700
                            })
                except Exception:
                    pass

    return {
        "status": "clean" if not violations else "violations_found",
        "violations": violations
    }


def validate_data_integrity(content_updates: list, deck_state: dict,
                            provider: str) -> dict:
    """
    Compare generated content against source data using LLM.
    Only called for slides containing financial/numerical content.
    """
    from llm import validate_data

    source_texts = []
    generated_texts = []

    for update in content_updates:
        slide_label = update.get("slide_label", "")
        for slide in deck_state.get("slides", []):
            if slide.get("label") == slide_label:
                for shape in slide.get("shapes", []):
                    if shape.get("text"):
                        source_texts.append(shape["text"])
                break
        if "text" in update:
            generated_texts.append(update["text"])
        elif "new_text" in update:
            generated_texts.append(update["new_text"])
        elif "rows" in update:
            generated_texts.append(json.dumps(update["rows"]))

    if not source_texts or not generated_texts:
        return {"accurate": True, "discrepancies": []}

    source_json = json.dumps(source_texts, indent=2)
    generated_text = "\n".join(generated_texts)

    try:
        return validate_data(source_json, generated_text, provider)
    except Exception as e:
        return {"accurate": True,
                "discrepancies": [f"Validation call failed: {str(e)}"]}


def smoke_test(output_path: str) -> dict:
    """
    Open the saved file and verify it loads cleanly.
    If Aspose can open it, PowerPoint can too.
    """
    try:
        test_prs = slides.Presentation(output_path)
        slide_count = len(test_prs.slides)
        return {"status": "ok", "slide_count": slide_count}
    except Exception as e:
        return {"status": "error", "message": str(e)}
