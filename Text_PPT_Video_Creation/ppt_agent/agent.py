"""
agent.py - Gradio UI and CLI for the AI PPT Generation Agent.

Run (web UI):  python -m ppt_agent.agent
Run (CLI):     python -m ppt_agent.agent --topic "Your Topic" --slides 10 --theme "Dark Pro"
"""

import argparse
import os
import socket
import sys
import uuid
from pathlib import Path

# ── Make sure parent dir is on path ─────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent.parent))

import gradio as gr
from dotenv import load_dotenv

from .llm      import get_client, generate_outline, LAYOUT_DESCRIPTIONS
from .tts import generate_all_audio, get_voice_choices
from .builder  import build_pptx
from .themes   import THEMES
from .exporter import sync_and_export_video

load_dotenv()

OUTPUT_DIR = Path(__file__).parent / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

# ── Pipeline ─────────────────────────────────────────────────────────────────

def run_pipeline(
    topic: str,
    n_slides: int,
    theme_name: str,
    allowed_layouts: list,
    tts_enabled: bool,
    voice_name: str,
    api_key: str,
    progress: gr.Progress = gr.Progress(track_tqdm=True),
):
    """Main pipeline: LLM → TTS → PPTX builder."""
    logs = []

    def log(msg: str):
        logs.append(msg)
        return "\n".join(logs)

    # Validation
    if not topic.strip():
        raise gr.Error("Please enter a topic.")
    if not allowed_layouts:
        raise gr.Error("Select at least one slide layout.")
    
    if not api_key.strip():
        raise gr.Error("Please enter your Gemini API key.")

    # Always ensure title + closing are available
    for required in ("title_slide", "closing_slide"):
        if required not in allowed_layouts:
            allowed_layouts.append(required)

    yield log("Connecting to Gemini AI ..."), None, gr.update(interactive=False), gr.update(visible=False)

    client = get_client(api_key.strip() if api_key else "", provider="gemini")

    yield log(f"Generating {n_slides}-slide outline on: {topic} ..."), None, gr.update(interactive=False), gr.update(visible=False)

    try:
        slides_data = generate_outline(
            client, 
            topic, 
            n_slides, 
            allowed_layouts, 
            theme_name, 
            provider="gemini",
            log_cb=log
        )
    except Exception as e:
        raise gr.Error(f"Gemini error: {e}")

    yield log(f"Outline ready - {len(slides_data)} slides generated."), None, gr.update(interactive=False), gr.update(visible=False)

    # TTS
    session_id = uuid.uuid4().hex[:8]
    audio_dir  = OUTPUT_DIR / f"audio_{session_id}"
    audio_paths = [None] * len(slides_data)

    if tts_enabled:
        yield log("Generating narration audio (this may take ~30s) ..."), None, gr.update(interactive=False), gr.update(visible=False)
        notes = [s.get("speaker_notes", "") for s in slides_data]
        try:
            audio_paths = generate_all_audio(notes, voice_name, str(audio_dir))
            n_audio = sum(1 for p in audio_paths if p)
            yield log(f"Audio ready - {n_audio} clips generated."), None, gr.update(interactive=False), gr.update(visible=False)
        except Exception as e:
            yield log(f"TTS failed ({e}), continuing without audio ..."), None, gr.update(interactive=False), gr.update(visible=False)
            audio_paths = [None] * len(slides_data)
    else:
        yield log("TTS disabled - skipping audio generation."), None, gr.update(interactive=False), gr.update(visible=False)

    # Build PPTX
    out_path = str(OUTPUT_DIR / f"presentation_{session_id}.pptx")
    yield log("Building PowerPoint slides ..."), None, gr.update(interactive=False), gr.update(visible=False)

    try:
        for i, slide_data in enumerate(slides_data):
            yield log(f"Slide {i+1}/{len(slides_data)}: {slide_data.get('title','')[:50]}"), None, gr.update(interactive=False), gr.update(visible=False)

        build_pptx(slides_data, audio_paths, theme_name, out_path)
    except Exception as e:
        raise gr.Error(f"PPTX build error: {e}")

    yield log("Done! Presentation saved. Click 'Export to Video' to generate an MP4."), out_path, gr.update(interactive=True), gr.update(visible=False)


def run_export_video(pptx_path: str):
    """Export the current PPTX to MP4 using PowerPoint COM."""
    logs = []
    def log(msg):
        logs.append(msg)
        return "\n".join(logs)

    if not pptx_path:
        raise gr.Error("Generate a presentation first before exporting to video.")

    pptx_path = pptx_path.strip()
    if not os.path.exists(pptx_path):
        raise gr.Error(f"PPTX file not found: {pptx_path}")

    mp4_path = pptx_path.replace(".pptx", ".mp4")

    try:
        yield log("Starting PowerPoint for video export..."), None
        sync_and_export_video(pptx_path, mp4_path, log_cb=log)
        yield log(f"Video ready!"), mp4_path
    except Exception as e:
        raise gr.Error(f"Video export failed: {e}")


# ── Gradio UI ─────────────────────────────────────────────────────────────────

LAYOUT_CHOICES = list(LAYOUT_DESCRIPTIONS.keys())
THEME_CHOICES  = ["Auto-Select (AI Pick)", "Random Theme"] + list(THEMES.keys())
def _default_voice() -> str:
    return "Aria (US Female)"

CSS = """
body { font-family: 'Inter', sans-serif; }
.gradio-container { max-width: 1100px !important; margin: auto; }
.title-block { text-align: center; padding: 2rem 0 1.5rem; background: linear-gradient(180deg, rgba(99,102,241,0.08) 0%, rgba(236,72,153,0.04) 100%); border-radius: 12px; border: 1px solid rgba(99,102,241,0.1); }
.title-block h1 { font-size: 2.6rem; font-weight: 800; letter-spacing: -0.5px;
    background: linear-gradient(135deg, #6366f1, #ec4899);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 0.5rem; }
.title-block p  { color: #64748b; font-size: 1.1rem; margin-top: 0.6rem; font-weight: 500; letter-spacing: 0.3px; }
.log-box textarea { font-family: monospace !important; font-size: 0.82rem !important; }
.download-btn { background: linear-gradient(135deg,#6366f1,#ec4899) !important;
    color: #fff !important; font-weight: 700 !important; }
"""

UI_THEME = gr.themes.Base(
    primary_hue=gr.themes.colors.indigo,
    secondary_hue=gr.themes.colors.pink,
    font=[gr.themes.GoogleFont("Inter"), "sans-serif"],
)


def build_ui():
    with gr.Blocks(title="AI PPT Agent") as demo:

        gr.HTML("""
        <div class="title-block">
          <h1>✨ AI Presentation Studio</h1>
          <p>Create Professional Presentations Instantly with AI-Powered Content & Narration</p>
        </div>
        """)

        with gr.Row():
            # ── LEFT: Inputs ──────────────────────────────────────────────
            with gr.Column(scale=5):
                gr.Markdown("### Presentation Settings")

                gemini_key = gr.Textbox(
                    label="Gemini API Key",
                    placeholder="AIza…",
                    type="password",
                    value=os.getenv("GEMINI_API_KEY", ""),
                )
                
                topic = gr.Textbox(
                    label="Topic / Title",
                    placeholder="e.g. 'The Future of Artificial Intelligence' or 'Climate Change Solutions'",
                    lines=2,
                )
                with gr.Row():
                    n_slides = gr.Slider(
                        label="Number of Slides",
                        minimum=5, maximum=20, step=1, value=10,
                    )
                    theme_dd = gr.Dropdown(
                        label="Color Theme",
                        choices=THEME_CHOICES,
                        value="Auto-Select (AI Pick)",
                    )

                gr.Markdown("### Slide Layout Types")
                layouts = gr.CheckboxGroup(
                    label="Allowed layouts (first + last slides are always title & closing)",
                    choices=LAYOUT_CHOICES,
                    value=LAYOUT_CHOICES,  # all selected by default
                )

                gr.Markdown("### Narration / TTS")
                with gr.Row():
                    tts_on = gr.Checkbox(label="Enable voice narration", value=True)
                    voice_dd = gr.Dropdown(
                        label="Voice",
                        choices=get_voice_choices(),
                        value=_default_voice(),
                    )

                gen_btn = gr.Button(
                    "Generate Presentation",
                    variant="primary",
                    elem_classes=["download-btn"],
                    size="lg",
                )

            # ── RIGHT: Output ─────────────────────────────────────────────
            with gr.Column(scale=5):
                gr.Markdown("### Progress and Output")
                log_box = gr.Textbox(
                    label="Pipeline Log",
                    lines=12,
                    interactive=False,
                    elem_classes=["log-box"],
                    placeholder="Logs will appear here after you click Generate …",
                )
                file_out = gr.File(
                    label="Download Your PPTX",
                    interactive=False,
                    file_types=[".pptx"],
                )

                # Hidden state to hold the current pptx path for video export
                current_pptx = gr.State(value=None)

                export_btn = gr.Button(
                    "Export to Video (MP4) — requires Microsoft PowerPoint",
                    variant="secondary",
                    interactive=False,
                )
                video_out = gr.File(
                    label="Download MP4 Video",
                    interactive=False,
                    file_types=[".mp4"],
                    visible=False,
                )

                gr.Markdown("""
                **Tips:**
                - **LLM Provider:** Powered by Google Gemini.
                - Be specific in your topic for better content
                - Narration audio is embedded and synced per slide
                - Export to MP4 uses PowerPoint — Microsoft Office must be installed
                """)

        # ── Event handlers ─────────────────────────────────────────────────
        def _on_generate(*args):
            last_log, last_file, last_btn, last_vid = None, None, None, None
            for log_val, file_val, btn_val, vid_val in run_pipeline(*args):
                last_log, last_file, last_btn, last_vid = log_val, file_val, btn_val, vid_val
                yield log_val, file_val, file_val, btn_val, vid_val
        
        gen_btn.click(
            fn=_on_generate,
            inputs=[topic, n_slides, theme_dd, layouts, tts_on, voice_dd, gemini_key],
            outputs=[log_box, file_out, current_pptx, export_btn, video_out],
        )

        def _on_export(pptx_state):
            # pptx_state is the gr.File value (could be a dict with 'name' key or a string path)
            if isinstance(pptx_state, dict):
                pptx_path = pptx_state.get("name", "")
            else:
                pptx_path = pptx_state or ""
            last_log, last_vid = None, None
            for log_val, vid_val in run_export_video(pptx_path):
                last_log, last_vid = log_val, vid_val
                yield log_val, gr.update(value=vid_val, visible=vid_val is not None)

        export_btn.click(
            fn=_on_export,
            inputs=[current_pptx],
            outputs=[log_box, video_out],
        )

    return demo


def run_cli(
    topic: str,
    n_slides: int,
    theme_name: str,
    allowed_layouts: list[str],
    tts_enabled: bool,
    voice_name: str,
    api_key: str,
    export_video: bool,
) -> str:
    """Run the pipeline from the command line. Returns path to output PPTX."""
    if not topic.strip():
        raise SystemExit("Error: --topic is required.")
    if not api_key.strip():
        api_key = os.getenv("GEMINI_API_KEY", "") or os.getenv("GOOGLE_API_KEY", "")
    if not api_key.strip():
        raise SystemExit("Error: set GEMINI_API_KEY in .env or pass --api-key.")

    for required in ("title_slide", "closing_slide"):
        if required not in allowed_layouts:
            allowed_layouts = list(allowed_layouts) + [required]

    def log(msg: str):
        print(msg)

    log(f"Connecting to Gemini …")
    client = get_client(api_key.strip())

    log(f"Generating {n_slides}-slide outline: {topic}")
    slides_data = generate_outline(
        client, topic, n_slides, allowed_layouts, theme_name, log_cb=log
    )
    log(f"Outline ready — {len(slides_data)} slides")

    session_id = uuid.uuid4().hex[:8]
    audio_dir = OUTPUT_DIR / f"audio_{session_id}"
    audio_paths = [None] * len(slides_data)

    if tts_enabled:
        log("Generating narration audio …")
        notes = [s.get("speaker_notes", "") for s in slides_data]
        try:
            audio_paths = generate_all_audio(notes, voice_name, str(audio_dir))
            log(f"Audio ready — {sum(1 for p in audio_paths if p)} clips")
        except Exception as e:
            log(f"TTS failed ({e}), continuing without audio")
            audio_paths = [None] * len(slides_data)
    else:
        log("TTS disabled")

    out_path = str(OUTPUT_DIR / f"presentation_{session_id}.pptx")
    log("Building PowerPoint …")
    build_pptx(slides_data, audio_paths, theme_name, out_path, progress_cb=log)
    log(f"PPTX saved: {out_path}")

    if export_video:
        mp4_path = out_path.replace(".pptx", ".mp4")
        log("Exporting MP4 (requires Microsoft PowerPoint) …")
        sync_and_export_video(out_path, mp4_path, log_cb=log)
        log(f"MP4 saved: {mp4_path}")

    return out_path


def _parse_cli_args(argv: list[str] | None = None):
    parser = argparse.ArgumentParser(
        description="Generate an AI PowerPoint presentation from a topic.",
    )
    parser.add_argument("--topic", "-t", help="Presentation topic / title")
    parser.add_argument("--slides", "-n", type=int, default=10, help="Number of slides (5–20)")
    parser.add_argument(
        "--theme",
        default="Auto-Select (AI Pick)",
        choices=THEME_CHOICES,
        help="Color theme name",
    )
    parser.add_argument(
        "--layouts",
        nargs="*",
        default=LAYOUT_CHOICES,
        help="Allowed slide layouts (space-separated)",
    )
    parser.add_argument("--api-key", help="Gemini API key (or use GEMINI_API_KEY in .env)")
    parser.add_argument("--no-tts", action="store_true", help="Skip voice narration")
    parser.add_argument(
        "--voice",
        default="Aria (US Female)",
        choices=get_voice_choices(),
        help="TTS voice when narration is enabled",
    )
    parser.add_argument(
        "--export-video",
        action="store_true",
        help="Also export MP4 via PowerPoint (Windows + Office required)",
    )
    return parser.parse_args(argv)


def _find_free_port(start: int = 7861, end: int = 7875) -> int:
    """Return the first available port in the given range."""
    for port in range(start, end + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("0.0.0.0", port))
                return port
            except OSError:
                continue
    raise OSError(f"Cannot find empty port in range: {start}-{end}")


if __name__ == "__main__":
    args = _parse_cli_args()
    if args.topic:
        run_cli(
            topic=args.topic,
            n_slides=max(5, min(20, args.slides)),
            theme_name=args.theme,
            allowed_layouts=args.layouts,
            tts_enabled=not args.no_tts,
            voice_name=args.voice,
            api_key=args.api_key or "",
            export_video=args.export_video,
        )
    else:
        demo = build_ui()
        env_port = os.getenv("GRADIO_SERVER_PORT")
        if env_port:
            try:
                port = int(env_port)
            except ValueError:
                raise ValueError("GRADIO_SERVER_PORT must be an integer.")
        else:
            port = _find_free_port(7861, 7875)

        demo.launch(
            server_name="0.0.0.0",
            server_port=port,
            share=False,
            inbrowser=True,
            css=CSS,
            theme=UI_THEME,
        )
