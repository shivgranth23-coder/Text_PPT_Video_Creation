"""
tts.py - Text-to-Speech narration (edge-tts, cloned voice via XTTS, gTTS fallback).
"""

import asyncio
import os
from pathlib import Path

# Edge-TTS preset voices
EDGE_VOICES = {
    "Aria (US Female)":   "en-US-AriaNeural",
    "Guy (US Male)":      "en-US-GuyNeural",
    "Jenny (US Female)":  "en-US-JennyNeural",
    "Neerja (IN Female)": "en-IN-NeerjaNeural",
    "Prabhat (IN Male)":  "en-IN-PrabhatNeural",
    "Sonia (UK Female)":  "en-GB-SoniaNeural",
    "Ryan (UK Male)":     "en-GB-RyanNeural",
}

# Back-compat alias used by agent.py
VOICES = dict(EDGE_VOICES)


def get_voice_choices() -> list[str]:
    """All voices for the dropdown."""
    return list(EDGE_VOICES.keys())


# ── edge-tts (primary for presets) ──────────────────────────────────────────

async def _edge_generate(text: str, voice: str, path: str):
    import edge_tts

    max_retries = 3
    for attempt in range(max_retries):
        try:
            communicate = edge_tts.Communicate(text, voice)
            await communicate.save(path)
            if os.path.exists(path) and os.path.getsize(path) > 0:
                return
            raise Exception("Empty file generated")
        except Exception as e:
            print(f"  Attempt {attempt+1} failed for TTS: {e}")
            if attempt < max_retries - 1:
                await asyncio.sleep(2)
            else:
                raise Exception(f"Failed after {max_retries} attempts: {e}") from e


async def _edge_generate_all(notes: list[str], voice: str, out_dir: str) -> list[str | None]:
    tasks = []
    paths = []
    for i, note in enumerate(notes):
        path = os.path.join(out_dir, f"slide_{i+1}_audio.mp3")
        paths.append(path)
        if note and note.strip():
            tasks.append(_edge_generate(note.strip(), voice, path))
        else:
            tasks.append(asyncio.sleep(0))
    await asyncio.gather(*tasks, return_exceptions=True)
    result = []
    for i, p in enumerate(paths):
        if notes[i] and notes[i].strip() and os.path.exists(p) and os.path.getsize(p) > 0:
            result.append(p)
        else:
            result.append(None)
    return result





# ── gTTS (fallback) ─────────────────────────────────────────────────────────

def _gtts_generate_all(notes: list[str], out_dir: str) -> list[str | None]:
    from gtts import gTTS

    result = []
    for i, note in enumerate(notes):
        path = os.path.join(out_dir, f"slide_{i+1}_audio.mp3")
        if note and note.strip():
            try:
                tts = gTTS(text=note.strip(), lang="en", slow=False)
                tts.save(path)
                result.append(path)
            except Exception:
                result.append(None)
        else:
            result.append(None)
    return result


# ── Public API ───────────────────────────────────────────────────────────────

def generate_all_audio(
    notes: list[str],
    voice_name: str,
    out_dir: str,
    progress_cb=None,
) -> list[str | None]:
    """Generate one .mp3 per slide from speaker notes."""
    Path(out_dir).mkdir(parents=True, exist_ok=True)



    voice_id = EDGE_VOICES.get(voice_name, "en-US-AriaNeural")
    if progress_cb:
        progress_cb("Generating narration audio …")

    try:
        import edge_tts  # noqa: F401
        return asyncio.run(_edge_generate_all(notes, voice_id, out_dir))
    except ImportError:
        if progress_cb:
            progress_cb("edge-tts not found - falling back to gTTS …")
        return _gtts_generate_all(notes, out_dir)
