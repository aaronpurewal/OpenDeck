"""
Streamlit UI: OpenDeck.

Two-column layout:
  Left:  Upload + deck preview (slide thumbnails)
  Right: Chat + plan review + execution progress + download
"""

import streamlit as st
import streamlit.components.v1 as components
import json
import os
import base64
import time
import tempfile
from io import BytesIO

from pipeline import step1_harvest, step2_plan, step3_execute
from config import LLM_PROVIDER, DEFAULT_OUTPUT_DIR, LOCAL_API_BASE

# --- Page Config ---
st.set_page_config(
    page_title="OpenDeck",
    page_icon="🔬",
    layout="wide"
)

# --- CSS Injection ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

/* Global */
.stApp {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    background: #FAFAFA;
}
html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}
h1, h2, h3, h4, h5, h6 {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    letter-spacing: -0.03em;
}

/* Hide Streamlit default header/toolbar */
[data-testid="stHeader"] {
    display: none !important;
}
[data-testid="stToolbar"] {
    display: none !important;
}
header[data-testid="stHeader"] {
    background: transparent !important;
    height: 0 !important;
}

/* Lock viewport — no outer scroll */
html, body {
    overflow: hidden !important;
    height: 100vh !important;
}
.stApp {
    height: 100vh !important;
    overflow: hidden !important;
}
[data-testid="stAppViewContainer"] {
    overflow: hidden !important;
    height: 100vh !important;
}
[data-testid="stMain"] {
    overflow: hidden !important;
}

/* Main block container: tight padding, viewport height */
.block-container {
    padding-top: 0.75rem !important;
    padding-bottom: 0.5rem !important;
    max-width: 1280px;
    height: 100vh !important;
    overflow: hidden !important;
}

/* Scrollable containers (st.container with height) */
[data-testid="stVerticalBlockBorderWrapper"] [data-testid="stVerticalBlock"]:has(> div > [data-testid="stVerticalBlock"]) {
    scrollbar-width: thin;
}
[data-testid="stVerticalBlockBorderWrapper"] {
    border: none !important;
}

/* Thin custom scrollbar */
::-webkit-scrollbar {
    width: 6px;
    height: 6px;
}
::-webkit-scrollbar-thumb {
    background: #D4D4D8;
    border-radius: 3px;
}
::-webkit-scrollbar-track {
    background: transparent;
}

/* Hide sidebar entirely */
[data-testid="stSidebar"] {
    display: none !important;
}
[data-testid="collapsedControl"] {
    display: none !important;
}
/* Buttons — rounded, modern */
.stButton > button {
    font-family: 'Inter', sans-serif;
    font-weight: 600;
    font-size: 13px;
    letter-spacing: 0.01em;
    border-radius: 10px;
    padding: 10px 24px;
    transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #6366F1 0%, #8B5CF6 100%);
    border: none;
    color: white;
    box-shadow: 0 2px 8px rgba(99,102,241,0.25);
}
.stButton > button[kind="primary"]:hover {
    box-shadow: 0 4px 16px rgba(99,102,241,0.35);
    transform: translateY(-1px);
}
.stButton > button[kind="secondary"] {
    background: white;
    border: 1px solid #E4E4E7;
    color: #18181B;
}
.stButton > button[kind="secondary"]:hover {
    border-color: #6366F1;
    color: #6366F1;
    background: #F5F3FF;
}

/* File Uploader — clean card */
[data-testid="stFileUploader"] {
    border: 2px dashed #D4D4D8;
    border-radius: 12px;
    padding: 20px;
    background: #FFFFFF;
    transition: all 0.2s ease;
}
[data-testid="stFileUploader"]:hover {
    border-color: #6366F1;
    background: #FAFAFE;
}

/* Chat Messages */
[data-testid="stChatMessage"] {
    border-radius: 12px;
    padding: 14px 18px;
    margin: 6px 0;
    border: 1px solid #F4F4F5;
}

/* Progress Bar — gradient */
.stProgress > div > div > div {
    background: linear-gradient(90deg, #6366F1 0%, #8B5CF6 50%, #A78BFA 100%) !important;
    border-radius: 6px;
}
.stProgress > div > div {
    background-color: #E4E4E7 !important;
    border-radius: 6px;
    height: 6px !important;
}

/* Metrics — glass card */
[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid #F4F4F5;
    border-radius: 12px;
    padding: 20px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.02);
}
[data-testid="stMetricValue"] {
    font-weight: 700 !important;
}

/* Expander */
[data-testid="stExpander"] {
    border: 1px solid #F4F4F5;
    border-radius: 12px;
    background: #FFFFFF;
}

/* Text Area (Plan Editor) */
.stTextArea textarea {
    font-family: 'SF Mono', 'JetBrains Mono', 'Fira Code', monospace !important;
    font-size: 13px !important;
    background: #18181B !important;
    color: #E4E4E7 !important;
    border: 1px solid #3F3F46 !important;
    border-radius: 10px !important;
    line-height: 1.6 !important;
}

/* Dividers */
hr {
    border-color: #F4F4F5 !important;
}

/* Download button */
.stDownloadButton > button {
    background: linear-gradient(135deg, #059669 0%, #10B981 100%) !important;
    color: white !important;
    border: none !important;
    font-weight: 600;
    border-radius: 10px !important;
    box-shadow: 0 2px 8px rgba(5,150,105,0.2);
}
.stDownloadButton > button:hover {
    box-shadow: 0 4px 16px rgba(5,150,105,0.3) !important;
    transform: translateY(-1px);
}

/* Alert boxes */
.stAlert {
    border-radius: 10px;
    font-size: 14px;
}

/* Animations */
@keyframes fadeIn {
    from { opacity: 0; transform: translateY(8px); }
    to { opacity: 1; transform: translateY(0); }
}
@keyframes pulse {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.4; }
}
@keyframes shimmer {
    0% { background-position: -200% 0; }
    100% { background-position: 200% 0; }
}
@keyframes gradientShift {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}
</style>
""", unsafe_allow_html=True)


# --- Session State Init ---
if "phase" not in st.session_state:
    st.session_state.phase = "upload"
if "prs" not in st.session_state:
    st.session_state.prs = None
if "deck_state" not in st.session_state:
    st.session_state.deck_state = None
if "plan" not in st.session_state:
    st.session_state.plan = None
if "execution_log" not in st.session_state:
    st.session_state.execution_log = None
if "provider" not in st.session_state:
    st.session_state.provider = LLM_PROVIDER
if "output_path" not in st.session_state:
    st.session_state.output_path = None
if "input_path" not in st.session_state:
    st.session_state.input_path = None
if "messages" not in st.session_state:
    st.session_state.messages = []
if "auto_approve" not in st.session_state:
    st.session_state.auto_approve = False
if "timer_start" not in st.session_state:
    st.session_state.timer_start = None
if "timer_paused_at" not in st.session_state:
    st.session_state.timer_paused_at = None
if "timer_paused_total" not in st.session_state:
    st.session_state.timer_paused_total = 0.0
if "timer_final" not in st.session_state:
    st.session_state.timer_final = None


# ---------------------------------------------------------------------------
# UI Helpers
# ---------------------------------------------------------------------------

def _card(html: str, accent: str = None, glass: bool = False) -> str:
    """Wrap HTML content in a styled card."""
    border = f"border-left: 3px solid {accent};" if accent else ""
    bg = "background: rgba(255,255,255,0.7); backdrop-filter: blur(12px);" if glass else "background:#fff;"
    return (f'<div style="{bg} border:1px solid #F4F4F5; border-radius:12px;'
            f'padding:20px 24px; margin:10px 0; box-shadow:0 1px 3px rgba(0,0,0,0.04),'
            f'0 1px 2px rgba(0,0,0,0.02); {border} animation:fadeIn 0.3s ease;">'
            f'{html}</div>')


def _pill(text: str, color: str = "#6366F1", bg: str = None) -> str:
    """Render a colored pill badge."""
    if bg is None:
        bg = color + "12"
    return (f'<span style="background:{bg}; color:{color}; padding:3px 10px;'
            f'border-radius:20px; font-size:11px; font-weight:600;'
            f'letter-spacing:0.02em;">{text}</span>')


def _section_label(text: str):
    """Render a section label."""
    st.markdown(f'<p style="font-size:12px; font-weight:600; color:#71717A;'
                f'letter-spacing:0.06em; text-transform:uppercase;'
                f'margin-bottom:8px;">{text}</p>', unsafe_allow_html=True)


def _render_phase_indicator(current: str):
    """Render a modern phase stepper."""
    display = [("upload", "Upload"), ("planning", "Plan"),
               ("review", "Review"), ("executing", "Execute"), ("done", "Done")]
    current_display = "review" if current == "editing" else current
    order = [p[0] for p in display]
    current_idx = order.index(current_display) if current_display in order else 0

    items = []
    for i, (key, label) in enumerate(display):
        if i < current_idx:
            # Completed
            items.append(f"""<div style="display:flex; align-items:center; gap:6px;">
                <span style="display:inline-flex; align-items:center; justify-content:center;
                    width:24px; height:24px; border-radius:50%;
                    background:linear-gradient(135deg, #059669, #10B981);
                    color:white; font-size:12px; font-weight:700;">&#10003;</span>
                <span style="font-size:12px; font-weight:600; color:#059669;">{label}</span>
            </div>""")
        elif i == current_idx:
            # Current — gradient ring
            items.append(f"""<div style="display:flex; align-items:center; gap:6px;">
                <span style="display:inline-flex; align-items:center; justify-content:center;
                    width:24px; height:24px; border-radius:50%;
                    background:linear-gradient(135deg, #6366F1, #8B5CF6);
                    color:white; font-size:12px; font-weight:700;">{i + 1}</span>
                <span style="font-size:12px; font-weight:600; color:#6366F1;">{label}</span>
            </div>""")
        else:
            items.append(f"""<div style="display:flex; align-items:center; gap:6px;">
                <span style="display:inline-flex; align-items:center; justify-content:center;
                    width:24px; height:24px; border-radius:50%; background:#F4F4F5;
                    color:#A1A1AA; font-size:12px; font-weight:600;">{i + 1}</span>
                <span style="font-size:12px; font-weight:500; color:#A1A1AA;">{label}</span>
            </div>""")

    sep = '<span style="color:#D4D4D8; margin:0 6px; font-size:14px;">&#8250;</span>'
    html = f"""<div style="display:flex; align-items:center; gap:4px;
        padding:12px 0; margin-bottom:8px;">{sep.join(items)}</div>"""
    st.markdown(html, unsafe_allow_html=True)


def _render_plan_display(plan: dict):
    """Render the plan as structured visual cards."""
    # Reasoning
    reasoning = plan.get("reasoning", "")
    if reasoning:
        st.markdown(_card(
            f'<div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">'
            f'<span style="font-size:16px;">&#128161;</span>'
            f'<span style="font-size:12px; font-weight:700; color:#6366F1;'
            f'text-transform:uppercase; letter-spacing:0.06em;">Strategy</span></div>'
            f'<p style="font-size:14px; color:#3F3F46; margin:0; line-height:1.7;">{reasoning}</p>',
            accent="#6366F1"
        ), unsafe_allow_html=True)

    # Structural changes
    structural = plan.get("structural_changes", [])
    if structural:
        steps_html = ""
        for i, step in enumerate(structural):
            action = step.get("action", "")
            args = step.get("args", {})
            label = step.get("label", "")

            icons = {"clone_slide": "&#43;", "delete_slides": "&#8722;",
                     "reorder_slides": "&#8645;", "duplicate_slide": "&#9741;"}
            colors = {"clone_slide": "#059669", "delete_slides": "#EF4444",
                      "reorder_slides": "#F59E0B", "duplicate_slide": "#6366F1"}
            icon = icons.get(action, "&#8226;")
            color = colors.get(action, "#71717A")

            detail_parts = []
            if label:
                detail_parts.append(_pill(label, "#6366F1", "#EEF2FF"))
            if "layout_name" in args:
                detail_parts.append(f'<span style="color:#71717A; font-size:12px;">'
                    f'layout: {args["layout_name"]}</span>')
            if "labels" in args:
                detail_parts.append(_pill(f'{len(args["labels"])} slides', "#71717A", "#F4F4F5"))
            details = " ".join(detail_parts)

            steps_html += f"""<div style="display:flex; align-items:center; gap:14px;
                padding:12px 0; {'border-top:1px solid #F4F4F5;' if i > 0 else ''}">
                <span style="display:inline-flex; align-items:center; justify-content:center;
                    width:32px; height:32px; border-radius:10px; background:{color}10;
                    color:{color}; font-size:16px; font-weight:700; flex-shrink:0;">{icon}</span>
                <div>
                    <span style="font-size:13px; font-weight:600; color:#18181B;">{action}</span>
                    <div style="margin-top:4px;">{details}</div>
                </div>
            </div>"""

        st.markdown(_card(
            f'<p style="font-size:12px; font-weight:700; color:#71717A;'
            f'text-transform:uppercase; letter-spacing:0.06em;'
            f'margin:0 0 12px 0;">Structural Changes</p>{steps_html}'
        ), unsafe_allow_html=True)

    # Content manifest
    manifest = plan.get("content_manifest", [])
    if manifest:
        rows_html = ""
        for i, item in enumerate(manifest):
            action = item.get("action", "")
            shape = item.get("shape_name", "")
            slide = item.get("slide_label", "")
            bg = "#FAFAFA" if i % 2 == 0 else "#FFFFFF"

            action_colors = {
                "fill_placeholder": "#059669", "fill_table": "#059669",
                "edit_run": "#6366F1", "edit_paragraph": "#6366F1",
                "edit_table_cell": "#6366F1", "edit_table_run": "#6366F1",
                "update_chart": "#8B5CF6", "create_chart": "#F59E0B",
                "create_table": "#F59E0B",
            }
            ac = action_colors.get(action, "#71717A")

            rows_html += f"""<tr style="background:{bg};">
                <td style="padding:10px 14px; font-size:13px; font-weight:500; color:#18181B;">{slide}</td>
                <td style="padding:10px 14px; font-size:13px; color:#52525B;">{shape}</td>
                <td style="padding:10px 14px;">{_pill(action, ac)}</td>
            </tr>"""

        count = _pill(f'{len(manifest)} shapes', "#6366F1", "#EEF2FF")

        st.markdown(_card(
            f'<div style="display:flex; align-items:center; justify-content:space-between;'
            f'margin-bottom:14px;">'
            f'<span style="font-size:12px; font-weight:700; color:#71717A;'
            f'text-transform:uppercase; letter-spacing:0.06em;">Content Manifest</span>'
            f'{count}</div>'
            f'<div style="border-radius:10px; overflow:hidden; border:1px solid #F4F4F5;">'
            f'<table style="width:100%; border-collapse:collapse;">'
            f'<tr style="background:#18181B;">'
            f'<th style="padding:10px 14px; text-align:left; color:#A1A1AA; font-size:11px;'
            f'font-weight:600; text-transform:uppercase; letter-spacing:0.06em;">Slide</th>'
            f'<th style="padding:10px 14px; text-align:left; color:#A1A1AA; font-size:11px;'
            f'font-weight:600; text-transform:uppercase; letter-spacing:0.06em;">Shape</th>'
            f'<th style="padding:10px 14px; text-align:left; color:#A1A1AA; font-size:11px;'
            f'font-weight:600; text-transform:uppercase; letter-spacing:0.06em;">Action</th>'
            f'</tr>{rows_html}</table></div>'
        ), unsafe_allow_html=True)


def render_slide_thumbnails(prs):
    """Render slide thumbnails using Aspose."""
    thumbnails = []
    try:
        import aspose.slides as slides
        from PIL import Image
        for i in range(len(prs.slides)):
            slide = prs.slides[i]
            bitmap = slide.get_thumbnail(0.25, 0.25)
            img_path = os.path.join(tempfile.gettempdir(), f"thumb_{i}.png")
            bitmap.save(img_path)
            thumbnails.append(img_path)
    except Exception:
        pass
    return thumbnails


def _auto_download(file_path: str, file_name: str):
    """Trigger an automatic browser download via injected JS."""
    with open(file_path, "rb") as f:
        data = base64.b64encode(f.read()).decode()
    mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    html = (
        f'<a id="auto-dl" href="data:{mime};base64,{data}" '
        f'download="{file_name}" style="display:none"></a>'
        f'<script>document.getElementById("auto-dl").click();</script>'
    )
    components.html(html, height=0)


def _timer_elapsed() -> float:
    """Calculate elapsed seconds excluding paused time."""
    if st.session_state.timer_start is None:
        return 0.0
    now = time.time()
    elapsed = now - st.session_state.timer_start - st.session_state.timer_paused_total
    if st.session_state.timer_paused_at:
        elapsed -= (now - st.session_state.timer_paused_at)
    return max(0.0, elapsed)


def _render_stopwatch(running: bool):
    """Render a styled stopwatch pill."""
    elapsed = _timer_elapsed()
    if running:
        bg = "#FEF3C7"
        color = "#D97706"
        border = "1px solid #FDE68A"
        label_text = ""
        dot = ('<span style="display:inline-block; width:6px; height:6px; border-radius:50%;'
               'background:#D97706; margin-right:8px; animation:pulse 1s infinite;"></span>')
    else:
        bg = "#F4F4F5"
        color = "#71717A"
        border = "1px solid #E4E4E7"
        label_text = " paused"
        dot = ""

    html = f"""
    <div id="sw-badge" style="display:inline-flex; align-items:center; background:{bg};
        color:{color}; padding:6px 14px; border-radius:20px; border:{border};
        font-family:'SF Mono','JetBrains Mono','Fira Code',monospace;
        font-size:13px; font-weight:600;">
        {dot}<span id="sw-val">{elapsed:.1f}s{label_text}</span>
    </div>
    <script>
    (function() {{
        var offset = {elapsed};
        var running = {'true' if running else 'false'};
        var start = performance.now();
        var suffix = '{label_text}';
        function tick() {{
            if (!running) return;
            var now = performance.now();
            var el = document.getElementById('sw-val');
            if (el) el.innerText = (offset + (now - start) / 1000).toFixed(1) + 's' + suffix;
            requestAnimationFrame(tick);
        }}
        tick();
    }})();
    </script>
    """
    components.html(html, height=38)


def _timer_start():
    st.session_state.timer_start = time.time()
    st.session_state.timer_paused_at = None
    st.session_state.timer_paused_total = 0.0
    st.session_state.timer_final = None


def _timer_pause():
    if st.session_state.timer_paused_at is None:
        st.session_state.timer_paused_at = time.time()


def _timer_resume():
    if st.session_state.timer_paused_at is not None:
        st.session_state.timer_paused_total += time.time() - st.session_state.timer_paused_at
        st.session_state.timer_paused_at = None


def _timer_stop():
    st.session_state.timer_final = _timer_elapsed()


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

st.markdown("""
<div style="position:relative; padding:36px 24px 32px 24px;
    background: radial-gradient(ellipse at top, #1E1B4B 0%, #18181B 60%, #0A0A0B 100%);
    border-radius:16px; overflow:hidden; text-align:center;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1), 0 8px 24px rgba(99,102,241,0.08);">
    <div style="position:absolute; top:0; left:50%; transform:translateX(-50%);
        width:60%; height:1px;
        background:linear-gradient(90deg, transparent, #6366F1, transparent);"></div>
    <div style="position:absolute; top:-40px; left:50%; transform:translateX(-50%);
        width:300px; height:80px; border-radius:50%;
        background:radial-gradient(ellipse, rgba(99,102,241,0.25), transparent 70%);
        filter: blur(20px);"></div>
    <div style="position:relative; display:inline-flex; align-items:center; gap:14px;
        margin-bottom:12px;">
        <div style="width:42px; height:42px; border-radius:12px;
            background:linear-gradient(135deg, #6366F1, #8B5CF6);
            display:inline-flex; align-items:center; justify-content:center;
            box-shadow: 0 4px 16px rgba(99,102,241,0.4),
                        inset 0 1px 0 rgba(255,255,255,0.2);">
            <span style="color:white; font-size:20px;">&#9638;</span>
        </div>
        <span style="color:#FAFAFA; font-size:28px; font-weight:800;
            letter-spacing:-0.04em;
            background: linear-gradient(135deg, #FAFAFA 0%, #C7D2FE 100%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            background-clip: text;">OpenDeck</span>
        <span style="background:rgba(99,102,241,0.15); color:#A5B4FC;
            border:1px solid rgba(99,102,241,0.3);
            padding:4px 10px; border-radius:20px; font-size:10px; font-weight:600;
            letter-spacing:0.05em;">v1.0</span>
    </div>
    <p style="color:#A1A1AA; font-size:14px; margin:0; line-height:1.5;
        max-width:520px; display:inline-block;">
        Upload a deck, describe what you want changed, and download the result.
    </p>
</div>
""", unsafe_allow_html=True)

st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Top bar: Model selector + deck stats
# ---------------------------------------------------------------------------

_PROVIDER_OPTIONS = ["openai", "anthropic", "local"]
_PROVIDER_LABELS = {
    "openai": "OpenAI (GPT)",
    "anthropic": "Anthropic (Claude)",
    "local": "Local (LM Studio)",
}

_topbar_cols = st.columns([2, 1, 1, 1])

with _topbar_cols[0]:
    st.markdown('<p style="font-size:11px; font-weight:600; color:#71717A; '
                'text-transform:uppercase; letter-spacing:0.06em; '
                'margin:0 0 4px 0;">Model Selector</p>',
                unsafe_allow_html=True)
    provider = st.selectbox(
        "Model Selector",
        _PROVIDER_OPTIONS,
        index=_PROVIDER_OPTIONS.index(st.session_state.provider)
              if st.session_state.provider in _PROVIDER_OPTIONS else 1,
        format_func=lambda x: _PROVIDER_LABELS.get(x, x),
        label_visibility="collapsed"
    )
    st.session_state.provider = provider

with _topbar_cols[1]:
    if provider == "local":
        try:
            import urllib.request
            urllib.request.urlopen(LOCAL_API_BASE.replace("/v1", ""), timeout=1)
            st.markdown('<div style="display:flex; align-items:center; gap:6px; '
                        'padding:8px 12px; background:#ECFDF5; border:1px solid #A7F3D0; '
                        'border-radius:8px;">'
                        '<span style="width:6px; height:6px; border-radius:50%; '
                        'background:#10B981; display:inline-block;"></span>'
                        '<span style="font-size:12px; color:#059669; font-weight:500;">'
                        'Connected</span></div>', unsafe_allow_html=True)
        except Exception:
            st.markdown('<div style="display:flex; align-items:center; gap:6px; '
                        'padding:8px 12px; background:#FEF3C7; border:1px solid #FDE68A; '
                        'border-radius:8px;">'
                        '<span style="width:6px; height:6px; border-radius:50%; '
                        'background:#F59E0B; display:inline-block;"></span>'
                        '<span style="font-size:12px; color:#D97706; font-weight:500;">'
                        'Not reachable</span></div>', unsafe_allow_html=True)

if st.session_state.deck_state:
    with _topbar_cols[2]:
        st.markdown(f'<div style="padding:8px 12px; background:#F4F4F5; '
                    f'border:1px solid #E4E4E7; border-radius:8px; text-align:center;">'
                    f'<span style="font-size:11px; color:#71717A; text-transform:uppercase; '
                    f'letter-spacing:0.05em; font-weight:600;">Slides</span> '
                    f'<span style="font-size:14px; color:#18181B; font-weight:700; '
                    f'margin-left:6px;">{st.session_state.deck_state.get("slide_count", 0)}</span>'
                    f'</div>', unsafe_allow_html=True)
    with _topbar_cols[3]:
        layouts = len(st.session_state.deck_state.get("master_layouts", []))
        st.markdown(f'<div style="padding:8px 12px; background:#F4F4F5; '
                    f'border:1px solid #E4E4E7; border-radius:8px; text-align:center;">'
                    f'<span style="font-size:11px; color:#71717A; text-transform:uppercase; '
                    f'letter-spacing:0.05em; font-weight:600;">Layouts</span> '
                    f'<span style="font-size:14px; color:#18181B; font-weight:700; '
                    f'margin-left:6px;">{layouts}</span>'
                    f'</div>', unsafe_allow_html=True)

st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Main Layout
# ---------------------------------------------------------------------------

_MAIN_HEIGHT = 520  # Fits in 768+ viewports after header/topbar/breadcrumb

left_col, right_col = st.columns([1, 2])

# === LEFT COLUMN: Upload + Preview ===
with left_col:
    _section_label("Your Deck")

    uploaded_file = st.file_uploader(
        "Upload a .pptx file",
        type=["pptx"],
        key="file_uploader",
        label_visibility="collapsed"
    )

    if uploaded_file and st.session_state.phase == "upload":
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        with st.spinner("Reading your deck..."):
            try:
                prs, deck_state = step1_harvest(tmp_path)
                st.session_state.prs = prs
                st.session_state.deck_state = deck_state
                st.session_state.input_path = tmp_path
                st.session_state.phase = "planning"
                st.session_state.messages = []
                st.rerun()
            except Exception as e:
                st.error(f"Failed to load deck: {str(e)}")

    # Scrollable thumbnail pane
    with st.container(height=_MAIN_HEIGHT, border=False):
        if st.session_state.prs:
            thumbnails = render_slide_thumbnails(st.session_state.prs)
            if thumbnails:
                for i, thumb_path in enumerate(thumbnails):
                    st.markdown(f"""<div style="position:relative; margin-bottom:4px;">
                        <span style="position:absolute; top:6px; left:6px; z-index:10;
                            background:linear-gradient(135deg, #6366F1, #8B5CF6);
                            color:white; width:22px; height:22px; border-radius:6px;
                            display:flex; align-items:center; justify-content:center;
                            font-size:11px; font-weight:700; box-shadow:0 2px 4px rgba(0,0,0,0.15);"
                        >{i + 1}</span>
                    </div>""", unsafe_allow_html=True)
                    st.image(thumb_path, use_container_width=True)
            else:
                if st.session_state.deck_state:
                    for slide in st.session_state.deck_state.get("slides", []):
                        layout = slide.get("layout_name", "Unknown")
                        shapes = len(slide.get("shapes", []))
                        st.text(f"  {slide['label']}: {layout} ({shapes} shapes)")

    # Download button (outside scrollable area so it's always visible)
    if st.session_state.output_path and os.path.exists(st.session_state.output_path):
        with open(st.session_state.output_path, "rb") as f:
            st.download_button(
                "Download Result",
                data=f.read(),
                file_name="result.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True
            )


# === RIGHT COLUMN: Chat + Plan Review + Execution ===
with right_col:
    _render_phase_indicator(st.session_state.phase)

    # Scrollable content container — everything phase-specific goes inside
    right_body = st.container(height=_MAIN_HEIGHT, border=False)
    with right_body:
        # Show chat history
        for msg in st.session_state.messages:
            with st.chat_message(msg["role"]):
                if msg.get("is_plan"):
                    _render_plan_display(msg["content"])
                elif msg.get("is_log"):
                    log = msg["content"].get("log", [])
                    ok = sum(1 for l in log if l.get("status") == "ok")
                    err = sum(1 for l in log if l.get("status") == "error")
                    st.markdown(f"**{ok} operations succeeded, {err} failed**")
                else:
                    st.markdown(msg["content"])

    # --- PLANNING PHASE ---
    if st.session_state.phase == "planning":
      with right_body:
        st.session_state.auto_approve = st.toggle(
            "Auto-approve plan",
            value=st.session_state.auto_approve,
            help="Skip the review step and execute immediately"
        )

        user_input = st.chat_input("What should I change in this deck?")
        if user_input:
            _timer_start()
            st.session_state.messages.append({"role": "user", "content": user_input})
            with st.chat_message("user"):
                st.markdown(user_input)

            _render_stopwatch(running=True)

            provider_name = _PROVIDER_LABELS.get(st.session_state.provider, st.session_state.provider)
            st.markdown(f"""<div style="display:flex; align-items:center; gap:8px; margin:10px 0;">
                <span style="display:inline-block; width:6px; height:6px; border-radius:50%;
                    background:linear-gradient(135deg, #6366F1, #8B5CF6);
                    animation:pulse 1s infinite;"></span>
                <span style="font-size:13px; color:#71717A;">Planning with</span>
                {_pill(provider_name, "#6366F1", "#EEF2FF")}
            </div>""", unsafe_allow_html=True)

            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    plan = step2_plan(
                        st.session_state.deck_state,
                        user_input,
                        st.session_state.provider
                    )

                if plan is None:
                    st.error("Couldn't generate a plan. Please try again.")
                else:
                    st.session_state.plan = plan
                    if st.session_state.auto_approve:
                        st.session_state.phase = "executing"
                    else:
                        st.session_state.phase = "review"

                    _render_plan_display(plan)

                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": plan,
                        "is_plan": True
                    })
                    st.rerun()

    # --- REVIEW PHASE ---
    if st.session_state.phase == "review":
      with right_body:
        _timer_pause()
        _render_stopwatch(running=False)
        st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

        col_approve, col_edit, col_reset = st.columns([2, 1, 1])
        with col_approve:
            if st.button("Approve & Execute", type="primary", use_container_width=True):
                _timer_resume()
                st.session_state.phase = "executing"
                st.rerun()
        with col_edit:
            if st.button("Edit Plan", use_container_width=True):
                st.session_state.phase = "editing"
                st.rerun()
        with col_reset:
            if st.button("Start Over", use_container_width=True):
                st.session_state.phase = "planning"
                st.session_state.plan = None
                st.session_state.messages = []
                st.rerun()

    # --- EDIT PLAN PHASE ---
    if st.session_state.phase == "editing":
      with right_body:
        st.divider()
        st.markdown(_card(
            '<div style="display:flex; align-items:center; gap:8px;">'
            '<span style="font-size:14px;">&#9888;</span>'
            '<span style="color:#D97706; font-size:13px; font-weight:500;">'
            'Manual edits bypass AI validation. Make sure your JSON is valid.</span></div>',
            accent="#F59E0B"
        ), unsafe_allow_html=True)

        _section_label("Plan Editor")
        edited_plan_str = st.text_area(
            "Edit Plan JSON",
            value=json.dumps(st.session_state.plan, indent=2),
            height=400,
            label_visibility="collapsed"
        )
        col_save, col_cancel = st.columns(2)
        with col_save:
            if st.button("Save & Execute", type="primary", use_container_width=True):
                try:
                    st.session_state.plan = json.loads(edited_plan_str)
                    _timer_resume()
                    st.session_state.phase = "executing"
                    st.rerun()
                except json.JSONDecodeError as e:
                    st.error(f"Invalid JSON: {e}")
        with col_cancel:
            if st.button("Cancel", use_container_width=True):
                st.session_state.phase = "review"
                st.rerun()

    # --- EXECUTING PHASE ---
    if st.session_state.phase == "executing":
      with right_body:
        st.divider()
        _render_stopwatch(running=True)

        # Step indicator
        st.markdown("""<div style="display:flex; gap:20px; margin:14px 0 18px 0;
            align-items:center;">
            <div style="display:flex; align-items:center; gap:6px;">
                <span style="width:8px; height:8px; border-radius:50%;
                    background:linear-gradient(135deg, #6366F1, #8B5CF6);
                    animation:pulse 1s infinite;"></span>
                <span style="font-size:12px; font-weight:600; color:#6366F1;">Structure</span>
            </div>
            <span style="color:#D4D4D8; font-size:16px;">&#8594;</span>
            <span style="font-size:12px; font-weight:500; color:#A1A1AA;">Content</span>
            <span style="color:#D4D4D8; font-size:16px;">&#8594;</span>
            <span style="font-size:12px; font-weight:500; color:#A1A1AA;">Validate</span>
        </div>""", unsafe_allow_html=True)

        progress = st.progress(0)
        status_text = st.empty()

        status_text.text("Running structural changes...")
        progress.progress(10)
        status_text.text("Generating content...")
        progress.progress(30)

        os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)
        output_path = os.path.join(DEFAULT_OUTPUT_DIR, "result.pptx")

        result = step3_execute(
            plan=st.session_state.plan,
            deck_state=st.session_state.deck_state,
            prs=st.session_state.prs,
            provider=st.session_state.provider,
            output_path=output_path
        )

        progress.progress(90)
        status_text.text("Validating...")

        if result["status"] == "complete":
            progress.progress(100)
            _timer_stop()
            st.session_state.output_path = result["output_path"]
            st.session_state.execution_log = result
            st.session_state.phase = "done"
            status_text.text("")

            st.session_state.messages.append({
                "role": "assistant",
                "content": result,
                "is_log": True
            })
            st.rerun()
        else:
            progress.progress(100)
            st.error(f"Something went wrong: {result.get('message', 'Unknown error')}")
            if result.get("detail"):
                st.error(result['detail'])
            if result.get("log"):
                with st.expander("See what happened"):
                    for entry in result["log"]:
                        status = entry.get("status", "")
                        action = entry.get("action", "")
                        msg = entry.get("message", "")
                        if status == "error":
                            st.error(f"{action}: {msg}")
                        else:
                            st.text(f"{action}: {status}")
            with st.expander("Plan that was sent"):
                st.json(st.session_state.plan)
            st.session_state.phase = "review"

    # --- DONE PHASE ---
    if st.session_state.phase == "done":
      with right_body:
        st.divider()

        final = st.session_state.timer_final
        timer_html = ""
        if final is not None:
            timer_html = (
                f'<div style="margin-top:20px;">'
                f'<span style="font-family:SF Mono,JetBrains Mono,Fira Code,monospace;'
                f'font-size:40px; font-weight:800; color:#18181B;'
                f'letter-spacing:-0.03em;">{final:.1f}s</span>'
                f'<p style="font-size:11px; color:#A1A1AA; font-weight:500;'
                f'text-transform:uppercase; letter-spacing:0.06em;'
                f'margin:6px 0 0 0;">Total time</p></div>'
            )

        st.markdown(_card(
            '<div style="text-align:center; padding:16px 0;">'
            '<div style="display:inline-flex; align-items:center; justify-content:center;'
            'width:56px; height:56px; border-radius:16px;'
            'background:linear-gradient(135deg, #059669, #10B981);'
            'margin-bottom:16px; box-shadow:0 4px 12px rgba(5,150,105,0.2);">'
            '<span style="color:white; font-size:28px; font-weight:700;">&#10003;</span></div>'
            '<h3 style="margin:0; font-size:22px; font-weight:700; color:#18181B;'
            'letter-spacing:-0.02em;">Done! Your deck is ready.</h3>'
            '<p style="font-size:14px; color:#71717A; margin:8px 0 0 0;">'
            'The modified file has been downloaded automatically.</p>'
            f'{timer_html}</div>'
        ), unsafe_allow_html=True)

        # Auto-download
        if st.session_state.output_path and os.path.exists(st.session_state.output_path):
            _auto_download(st.session_state.output_path, "result.pptx")

        # Metrics
        if st.session_state.execution_log:
            log = st.session_state.execution_log.get("log", [])
            ok_count = sum(1 for l in log if l.get("status") == "ok")
            err_count = sum(1 for l in log if l.get("status") == "error")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Operations", ok_count + err_count)
            with col2:
                st.metric("Succeeded", ok_count)
            with col3:
                st.metric("Failed", err_count)

            if st.session_state.execution_log.get("data_warnings"):
                st.markdown(_card(
                    '<div style="display:flex; align-items:center; gap:8px; margin-bottom:8px;">'
                    '<span style="font-size:14px;">&#9888;</span>'
                    '<span style="font-size:12px; font-weight:700; color:#D97706;'
                    'text-transform:uppercase; letter-spacing:0.06em;">'
                    'Data Warnings</span></div>'
                    + "".join(f'<p style="font-size:13px; color:#3F3F46; margin:4px 0;">{w}</p>'
                              for w in st.session_state.execution_log["data_warnings"]),
                    accent="#F59E0B"
                ), unsafe_allow_html=True)

            with st.expander("Execution details"):
                st.json(log)

        st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)
        col_new, col_upload = st.columns(2)
        with col_new:
            if st.button("New Instruction", type="primary", use_container_width=True):
                st.session_state.phase = "planning"
                st.session_state.plan = None
                st.session_state.execution_log = None
                st.session_state.output_path = None
                st.session_state.timer_start = None
                st.session_state.timer_final = None
                st.rerun()
        with col_upload:
            if st.button("Upload New Deck", use_container_width=True):
                for key in ["prs", "deck_state", "plan", "execution_log",
                           "output_path", "input_path", "messages"]:
                    if key in st.session_state:
                        del st.session_state[key]
                st.session_state.timer_start = None
                st.session_state.timer_final = None
                st.session_state.phase = "upload"
                st.rerun()

    # --- UPLOAD PHASE (no deck loaded) ---
    if st.session_state.phase == "upload" and not st.session_state.prs:
      with right_body:
        st.markdown("""
        <div style="text-align:center; padding:60px 20px; animation:fadeIn 0.5s ease;">
            <div style="display:inline-flex; align-items:center; justify-content:center;
                width:64px; height:64px; border-radius:16px;
                background:linear-gradient(135deg, #6366F1, #8B5CF6);
                margin-bottom:20px; box-shadow:0 4px 16px rgba(99,102,241,0.2);">
                <span style="color:white; font-size:28px;">&#9638;</span>
            </div>
            <h2 style="font-size:24px; font-weight:700; color:#18181B;
                letter-spacing:-0.03em; margin:0 0 8px 0;">
                Edit any PowerPoint with AI</h2>
            <p style="font-size:15px; color:#71717A; margin:0 0 36px 0; max-width:400px;
                display:inline-block; line-height:1.6;">
                Upload your deck, describe what you want changed, review the plan,
                and download the result. All formatting preserved.</p>
            <div style="display:flex; justify-content:center; gap:12px; flex-wrap:wrap;">
                <div style="background:white; border:1px solid #F4F4F5; border-radius:12px;
                    padding:18px 22px; text-align:center; min-width:150px;
                    box-shadow:0 1px 3px rgba(0,0,0,0.04);">
                    <div style="font-size:20px; margin-bottom:8px;">&#127919;</div>
                    <p style="font-size:13px; font-weight:600; color:#18181B; margin:0 0 4px 0;">
                        AI-Powered Plans</p>
                    <p style="font-size:12px; color:#A1A1AA; margin:0; line-height:1.4;">
                        Review before anything runs</p>
                </div>
                <div style="background:white; border:1px solid #F4F4F5; border-radius:12px;
                    padding:18px 22px; text-align:center; min-width:150px;
                    box-shadow:0 1px 3px rgba(0,0,0,0.04);">
                    <div style="font-size:20px; margin-bottom:8px;">&#9999;</div>
                    <p style="font-size:13px; font-weight:600; color:#18181B; margin:0 0 4px 0;">
                        Surgical Precision</p>
                    <p style="font-size:12px; color:#A1A1AA; margin:0; line-height:1.4;">
                        Only changes what you ask</p>
                </div>
                <div style="background:white; border:1px solid #F4F4F5; border-radius:12px;
                    padding:18px 22px; text-align:center; min-width:150px;
                    box-shadow:0 1px 3px rgba(0,0,0,0.04);">
                    <div style="font-size:20px; margin-bottom:8px;">&#128274;</div>
                    <p style="font-size:13px; font-weight:600; color:#18181B; margin:0 0 4px 0;">
                        Brand-Safe</p>
                    <p style="font-size:12px; color:#A1A1AA; margin:0; line-height:1.4;">
                        Preserves all formatting</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
