"""
exporter.py - Export PPTX to MP4 using PowerPoint COM automation.

Strategy:
  1. Open the PPTX via PowerPoint COM.
  2. For each slide, find the embedded audio shape, set it to auto-play,
     and set the slide advance time to match the audio duration.
  3. Call CreateVideo() to export a synchronized MP4.
"""

import os
import sys
import subprocess
import time


def _ensure_pywin32():
    """Import pywin32; auto-install with the same Python that runs this app if missing."""
    try:
        import win32com.client
        import pythoncom
        return win32com.client, pythoncom
    except ImportError:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "pywin32"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.STDOUT,
        )
        import win32com.client
        import pythoncom
        return win32com.client, pythoncom


def sync_and_export_video(ppt_path: str, mp4_path: str, log_cb=None) -> str:
    """
    Open the PPTX with PowerPoint COM, synchronize audio timings, and export to MP4.

    Args:
        ppt_path: Absolute path to the .pptx file.
        mp4_path: Absolute path to write the .mp4 file.
        log_cb:   Optional callable(str) for progress messages.

    Returns:
        The mp4_path on success.

    Raises:
        RuntimeError on failure.
    """
    try:
        win32com, pythoncom = _ensure_pywin32()
    except Exception as e:
        raise RuntimeError(
            "pywin32 is required for video export (Windows + PowerPoint).\n"
            f"Install with: {sys.executable} -m pip install pywin32\n"
            f"Details: {e}"
        )

    ppt_path = os.path.abspath(ppt_path)
    mp4_path = os.path.abspath(mp4_path)

    def log(msg):
        if log_cb:
            log_cb(msg)
        print(msg)

    if not os.path.exists(ppt_path):
        raise RuntimeError(f"PPTX file not found: {ppt_path}")

    log("Starting PowerPoint...")
    # Must call CoInitialize on any new thread that uses COM (e.g. Gradio worker threads)
    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.Dispatch("PowerPoint.Application")
        # Ensure PowerPoint is visible to allow video rendering to succeed
        try:
            ppt_app.Visible = True
        except Exception:
            pass
        ppt_app.DisplayAlerts = 1  # Suppress repair dialogs
    except Exception as e:
        pythoncom.CoUninitialize()
        raise RuntimeError(f"Could not launch PowerPoint: {e}")

    try:
        log(f"Opening: {os.path.basename(ppt_path)}")
        # Open with window to ensure proper video rendering context
        presentation = ppt_app.Presentations.Open(ppt_path, WithWindow=True)
    except Exception as e:
        try:
            ppt_app.Quit()
        except Exception:
            pass
        raise RuntimeError(f"Could not open presentation: {e}")

    try:
        log("Synchronizing audio timings with slides...")
        for slide in presentation.Slides:
            slide_duration = 5.0  # default seconds if no audio found
            audio_found = False

            for shape in slide.Shapes:
                # Type 16 = msoMedia
                if shape.Type == 16:
                    log(f"  Slide {slide.SlideIndex}: audio found, setting auto-play...")
                    try:
                        # Make audio auto-play when slide is shown
                        shape.AnimationSettings.PlaySettings.PlayOnEntry = True
                        shape.AnimationSettings.PlaySettings.HideWhileNotPlaying = True

                        # Get audio duration in milliseconds
                        length_ms = shape.MediaFormat.Length
                        if length_ms > 0:
                            slide_duration = (length_ms / 1000.0) + 1.0  # 1s buffer
                            audio_found = True
                    except Exception as inner_e:
                        log(f"  Warning: could not configure shape: {inner_e}")

            # Set slide to advance automatically after audio finishes
            slide.SlideShowTransition.AdvanceOnTime = True
            slide.SlideShowTransition.AdvanceTime = slide_duration
            if not audio_found:
                log(f"  Slide {slide.SlideIndex}: no audio, will advance after {slide_duration}s")

        log("Exporting to MP4 (this may take several minutes)...")
        presentation.CreateVideo(mp4_path, True, 5, 1080, 30, 85)

        # Wait a moment for export to register and start
        time.sleep(2)

        # Poll until done — status: 3=Done, 4=Failed
        dots = 0
        while True:
            try:
                status = presentation.CreateVideoStatus
            except Exception as com_e:
                log(f"  Warning (COM status check): {com_e}. Retrying in 2 seconds...")
                time.sleep(2)
                continue

            if status in (3, 4):
                break
            dots = (dots + 1) % 4
            log("Exporting" + "." * (dots + 1))
            time.sleep(3)

        # Get final status with retry
        try:
            status = presentation.CreateVideoStatus
        except Exception:
            status = 3  # Assume success if COM query fails at the end but video exists
            if not os.path.exists(mp4_path) or os.path.getsize(mp4_path) == 0:
                status = 4

        if status == 3:
            # Save the updated presentation so timings persist (optional, catch read-only error)
            try:
                presentation.Save()
            except Exception as save_e:
                log(f"  Note: Could not save slide timing updates back to PPTX: {save_e}")
            log(f"Video export complete: {mp4_path}")
            return mp4_path
        else:
            raise RuntimeError(f"PowerPoint video export failed (status={status}).")

    except Exception as e:
        raise RuntimeError(str(e))
    finally:
        try:
            presentation.Close()
        except Exception:
            pass
        try:
            if ppt_app.Presentations.Count == 0:
                ppt_app.Quit()
        except Exception:
            pass
        # Always uninitialize COM for this thread
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
