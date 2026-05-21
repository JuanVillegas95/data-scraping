"""
Reusable transcription microservice.
Usage:
    from transcriber import transcribe_url
    text = await transcribe_url("https://www.tiktok.com/@user/video/123")

Works with any URL yt-dlp supports: TikTok, Instagram, Reddit, Twitter/X, YouTube, etc.
"""

import asyncio
import tempfile
from pathlib import Path

_model = None


def _load_model():
    global _model
    if _model is None:
        from faster_whisper import WhisperModel
        print("[transcriber] Loading Whisper model (once per session)...")
        # base = good balance of speed/accuracy on CPU; bump to "small" for better accuracy
        _model = WhisperModel("base", device="cpu", compute_type="int8")
        print("[transcriber] Model ready.")
    return _model


async def transcribe_url(url: str, cookies_file: str | None = None) -> str | None:
    """
    Download audio from `url`, transcribe in English, delete all temp files.
    Returns the transcript string, or None if download or transcription failed.
    """
    with tempfile.TemporaryDirectory(prefix="transcriber_") as tmpdir:
        tmp = Path(tmpdir)
        audio_out = tmp / "audio.%(ext)s"

        # ── 1. download audio ──────────────────────────────────────────────
        cmd = [
            "yt-dlp",
            "--extract-audio",
            "--audio-format", "mp3",
            "--audio-quality", "5",        # mid quality — enough for speech
            "--output", str(audio_out),
            "--no-playlist",
            "--quiet",
            "--no-warnings",
            "--socket-timeout", "30",
        ]
        if cookies_file:
            cmd += ["--cookies", cookies_file]
        cmd.append(url)

        proc = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.DEVNULL,
            stderr=asyncio.subprocess.PIPE,
        )
        _, stderr = await proc.communicate()

        if proc.returncode != 0:
            err = stderr.decode(errors="replace")[:300]
            print(f"  [transcriber] yt-dlp failed ({url}): {err}")
            return None

        audio_files = list(tmp.glob("audio.*"))
        if not audio_files:
            print(f"  [transcriber] No audio file found after download: {url}")
            return None

        audio_path = str(audio_files[0])

        # ── 2. transcribe ──────────────────────────────────────────────────
        def _transcribe():
            model = _load_model()
            segments, _ = model.transcribe(
                audio_path,
                language="en",
                beam_size=1,               # faster on CPU
                vad_filter=True,           # skip silence
            )
            return " ".join(seg.text.strip() for seg in segments).strip()

        try:
            text = await asyncio.to_thread(_transcribe)
        except Exception as e:
            print(f"  [transcriber] Whisper failed ({url}): {e}")
            return None

        # ── 3. temp dir and all its contents auto-deleted here ─────────────
        return text or None
