"""
Streamlit UI: Surgical Slide Engine.

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
    page_title="Surgical Slide Engine",
    page_icon="🔬",
    layout="wide"
)

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


def render_slide_thumbnails(prs):
    """Render slide thumbnails using Aspose."""
    thumbnails = []
    try:
        import aspose.slides as slides
        from PIL import Image
        for i in range(len(prs.slides)):
            slide = prs.slides[i]
            # Generate thumbnail at 1/4 scale
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
    """Render a JS-based stopwatch that ticks in the browser."""
    elapsed = _timer_elapsed()
    if running:
        color = "#ff6b35"
        label = ""
    else:
        color = "#888"
        label = " (paused)"
    html = f"""
    <div id="sw" style="font-family:monospace;font-size:1.1em;color:{color};padding:4px 0;">
        ⏱ {elapsed:.1f}s{label}
    </div>
    <script>
    (function() {{
        var offset = {elapsed};
        var running = {'true' if running else 'false'};
        var start = performance.now();
        function tick() {{
            if (!running) return;
            var now = performance.now();
            var el = document.getElementById('sw');
            if (el) el.innerText = '⏱ ' + (offset + (now - start) / 1000).toFixed(1) + 's';
            requestAnimationFrame(tick);
        }}
        tick();
    }})();
    </script>
    """
    components.html(html, height=36)


def _timer_start():
    """Start/reset the stopwatch."""
    st.session_state.timer_start = time.time()
    st.session_state.timer_paused_at = None
    st.session_state.timer_paused_total = 0.0
    st.session_state.timer_final = None


def _timer_pause():
    """Pause the stopwatch."""
    if st.session_state.timer_paused_at is None:
        st.session_state.timer_paused_at = time.time()


def _timer_resume():
    """Resume the stopwatch after pause."""
    if st.session_state.timer_paused_at is not None:
        st.session_state.timer_paused_total += time.time() - st.session_state.timer_paused_at
        st.session_state.timer_paused_at = None


def _timer_stop():
    """Stop the stopwatch and record final time."""
    st.session_state.timer_final = _timer_elapsed()


# --- Header ---
st.title("Surgical Slide Engine")
st.caption("Upload a branded deck. Give an instruction. Get a clean result.")

# --- Sidebar: Provider Selection ---
_PROVIDER_OPTIONS = ["openai", "anthropic", "local"]
_PROVIDER_LABELS = {
    "openai": "OpenAI (GPT)",
    "anthropic": "Anthropic (Claude)",
    "local": "Local (Qwen 3.5 35B)",
}

with st.sidebar:
    st.header("Settings")
    provider = st.selectbox(
        "LLM Provider",
        _PROVIDER_OPTIONS,
        index=_PROVIDER_OPTIONS.index(st.session_state.provider)
              if st.session_state.provider in _PROVIDER_OPTIONS else 1,
        format_func=lambda x: _PROVIDER_LABELS.get(x, x)
    )
    st.session_state.provider = provider

    if provider == "local":
        st.caption(f"Server: {LOCAL_API_BASE}")
        # Quick connectivity check
        try:
            import urllib.request
            urllib.request.urlopen(LOCAL_API_BASE.replace("/v1", ""), timeout=1)
            st.success("Connected", icon="✅")
        except Exception:
            st.warning("Server not reachable. Start LM Studio and enable the local server.", icon="⚠️")

    if st.session_state.deck_state:
        st.divider()
        st.metric("Slides", st.session_state.deck_state.get("slide_count", 0))
        layouts = st.session_state.deck_state.get("master_layouts", [])
        st.metric("Layouts", len(layouts))

    st.divider()
    st.caption("v1.0 — Plan-then-Execute Architecture")

# --- Main Layout ---
left_col, right_col = st.columns([1, 2])

# === LEFT COLUMN: Upload + Preview ===
with left_col:
    st.subheader("Deck")

    uploaded_file = st.file_uploader(
        "Upload a .pptx file",
        type=["pptx"],
        key="file_uploader"
    )

    if uploaded_file and st.session_state.phase == "upload":
        # Save uploaded file to temp
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        with st.spinner("Harvesting deck state..."):
            try:
                prs, deck_state = step1_harvest(tmp_path)
                st.session_state.prs = prs
                st.session_state.deck_state = deck_state
                st.session_state.input_path = tmp_path
                st.session_state.phase = "planning"
                st.session_state.messages = []
                st.success(f"Loaded {deck_state['slide_count']} slides")
                st.rerun()
            except Exception as e:
                st.error(f"Failed to load deck: {str(e)}")

    # Show slide thumbnails if deck is loaded
    if st.session_state.prs:
        thumbnails = render_slide_thumbnails(st.session_state.prs)
        if thumbnails:
            for i, thumb_path in enumerate(thumbnails):
                st.image(thumb_path, caption=f"Slide {i + 1}", use_container_width=True)
        else:
            # Fallback: show slide list
            if st.session_state.deck_state:
                for slide in st.session_state.deck_state.get("slides", []):
                    layout = slide.get("layout_name", "Unknown")
                    shapes = len(slide.get("shapes", []))
                    st.text(f"  {slide['label']}: {layout} ({shapes} shapes)")

    # Download button
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
    st.subheader("Instructions")

    # Show chat history
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            if msg.get("is_plan"):
                st.json(msg["content"])
            elif msg.get("is_log"):
                st.json(msg["content"])
            else:
                st.markdown(msg["content"])

    # --- PLANNING PHASE ---
    if st.session_state.phase == "planning":
        st.session_state.auto_approve = st.checkbox(
            "Auto-approve plan (skip review step)",
            value=st.session_state.auto_approve,
            help="When checked, the plan will be executed immediately without review."
        )
        user_input = st.chat_input("What should I do with this deck?")
        if user_input:
            _timer_start()
            st.session_state.messages.append({"role": "user", "content": user_input})
            with st.chat_message("user"):
                st.markdown(user_input)

            _render_stopwatch(running=True)
            with st.chat_message("assistant"):
                with st.spinner("Generating plan (Pass 1)..."):
                    plan = step2_plan(
                        st.session_state.deck_state,
                        user_input,
                        st.session_state.provider
                    )

                if plan is None:
                    st.error("Failed to generate plan. Please try again.")
                else:
                    st.session_state.plan = plan
                    if st.session_state.auto_approve:
                        st.session_state.phase = "executing"
                    else:
                        st.session_state.phase = "review"

                    # Show reasoning
                    if "reasoning" in plan:
                        st.markdown(f"**Approach:** {plan['reasoning']}")

                    # Show structural changes
                    structural = plan.get("structural_changes", [])
                    if structural:
                        st.markdown("**Structural Changes:**")
                        for i, step in enumerate(structural):
                            action = step.get("action", "")
                            args = step.get("args", {})
                            label = step.get("label", "")
                            desc = f"{i + 1}. `{action}`"
                            if label:
                                desc += f" → {label}"
                            if "layout_name" in args:
                                desc += f" (layout: {args['layout_name']})"
                            if "labels" in args:
                                desc += f" ({len(args['labels'])} slides)"
                            st.markdown(desc)

                    # Show content manifest
                    manifest = plan.get("content_manifest", [])
                    if manifest:
                        st.markdown(f"**Content to generate:** {len(manifest)} shapes")
                        for item in manifest:
                            action = item.get("action", "")
                            shape = item.get("shape_name", "")
                            slide = item.get("slide_label", "")
                            st.text(f"  {action} → {slide}/{shape}")

                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": plan,
                        "is_plan": True
                    })
                    st.rerun()

    # --- REVIEW PHASE ---
    if st.session_state.phase == "review":
        _timer_pause()
        _render_stopwatch(running=False)
        st.divider()
        col_approve, col_edit, col_reset = st.columns(3)

        with col_approve:
            if st.button("Approve Plan", type="primary", use_container_width=True):
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
        st.divider()
        edited_plan_str = st.text_area(
            "Edit Plan JSON",
            value=json.dumps(st.session_state.plan, indent=2),
            height=400
        )
        col_save, col_cancel = st.columns(2)
        with col_save:
            if st.button("Save & Approve", type="primary", use_container_width=True):
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
        st.divider()
        _render_stopwatch(running=True)
        progress = st.progress(0)
        status_text = st.empty()

        status_text.text("Executing structural changes...")
        progress.progress(10)

        status_text.text("Generating content (Pass 2)...")
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
        status_text.text("Validating output...")

        if result["status"] == "complete":
            progress.progress(100)
            _timer_stop()
            st.session_state.output_path = result["output_path"]
            st.session_state.execution_log = result
            st.session_state.phase = "done"

            # Show results
            status_text.text("Done!")
            st.success(f"Output saved: {result['output_path']}")

            # Show execution log
            log = result.get("log", [])
            ok_count = sum(1 for l in log if l.get("status") == "ok")
            err_count = sum(1 for l in log if l.get("status") == "error")
            st.markdown(f"**Operations:** {ok_count} succeeded, {err_count} failed")

            if result.get("data_warnings"):
                st.warning("Data integrity warnings:")
                for w in result["data_warnings"]:
                    st.text(f"  ⚠ {w}")

            st.session_state.messages.append({
                "role": "assistant",
                "content": result,
                "is_log": True
            })
            st.rerun()
        else:
            progress.progress(100)
            st.error(f"Execution failed: {result.get('message', 'Unknown error')}")
            if result.get("detail"):
                st.error(f"Detail: {result['detail']}")
            if result.get("log"):
                st.subheader("Execution Log")
                for entry in result["log"]:
                    status = entry.get("status", "")
                    action = entry.get("action", "")
                    msg = entry.get("message", "")
                    if status == "error":
                        st.error(f"{action}: {msg}")
                    else:
                        st.success(f"{action}: {status}")
                st.json(result["log"])
            # Also show the plan that was sent
            with st.expander("Plan that was executed"):
                st.json(st.session_state.plan)
            st.session_state.phase = "review"

    # --- DONE PHASE ---
    if st.session_state.phase == "done":
        st.divider()
        final = st.session_state.timer_final
        if final is not None:
            st.success(f"Deck transformation complete! ⏱ {final:.1f}s")
        else:
            st.success("Deck transformation complete!")

        # Auto-download the result file
        if st.session_state.output_path and os.path.exists(st.session_state.output_path):
            _auto_download(st.session_state.output_path, "result.pptx")

        if st.session_state.execution_log:
            log = st.session_state.execution_log.get("log", [])
            ok_count = sum(1 for l in log if l.get("status") == "ok")
            err_count = sum(1 for l in log if l.get("status") == "error")
            col1, col2, col3 = st.columns(3)
            col1.metric("Operations", ok_count + err_count)
            col2.metric("Succeeded", ok_count)
            col3.metric("Failed", err_count)

            with st.expander("Execution Log"):
                st.json(log)

        if st.button("New Instruction", use_container_width=True):
            st.session_state.phase = "planning"
            st.session_state.plan = None
            st.session_state.execution_log = None
            st.session_state.output_path = None
            st.session_state.timer_start = None
            st.session_state.timer_final = None
            st.rerun()

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
        st.info("Upload a .pptx file to get started.")
