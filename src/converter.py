import hashlib
import os
import subprocess
import tempfile

from docx import Document
from pydub import AudioSegment
from rich.console import Console
from rich.progress import BarColumn, MofNCompleteColumn, Progress, TextColumn
from rich.table import Table

console = Console()

# Voices: Daniel (en_GB) for headings, Kathy (en_US) for body
VOICE_HEADING = "Jamie"
VOICE_BODY = "Serena"

# Silence durations in milliseconds
PAUSE_AFTER_HEADING_MS = 900
PAUSE_AFTER_PARAGRAPH_MS = 500

CACHE_DIR = os.path.join(tempfile.gettempdir(), "docx_to_mp3", "tts")


def _is_heading(paragraph):
    style = paragraph.style.name
    return style.startswith("Heading") or style in ("Title", "Subtitle")


def _cache_path(text, voice):
    key = hashlib.sha256(f"{voice}:{text}".encode()).hexdigest()[:16]
    return os.path.join(CACHE_DIR, f"{key}.mp3")


def _say_to_mp3(text, voice, output_path):
    with tempfile.NamedTemporaryFile(suffix=".aiff", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        subprocess.run(
            ["say", "-v", voice, "-o", tmp_path, "--", text],
            check=True,
            capture_output=True,
        )
        AudioSegment.from_file(tmp_path, format="aiff").export(
            output_path, format="mp3"
        )
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


def _para_to_segment(para):
    text = para.text.strip()
    is_heading = _is_heading(para)
    voice = VOICE_HEADING if is_heading else VOICE_BODY
    pause_ms = (
        PAUSE_AFTER_HEADING_MS if is_heading else PAUSE_AFTER_PARAGRAPH_MS
    )

    cached = _cache_path(text, voice)
    if not os.path.exists(cached):
        os.makedirs(CACHE_DIR, exist_ok=True)
        _say_to_mp3(text, voice, cached)

    return AudioSegment.from_mp3(cached) + AudioSegment.silent(
        duration=pause_ms
    )


def _build_audio(paragraphs):
    combined = AudioSegment.empty()
    with Progress(
        TextColumn("[bold cyan]{task.description}"),
        BarColumn(),
        MofNCompleteColumn(),
        console=console,
        transient=False,
    ) as progress:
        task = progress.add_task("Converting", total=len(paragraphs))
        for para in paragraphs:
            text = para.text.strip()
            is_heading = _is_heading(para)
            style = "bold yellow" if is_heading else "white"
            preview = text[:80] + "…" if len(text) > 80 else text
            progress.console.print(f"  [{style}]{preview}[/{style}]")
            combined += _para_to_segment(para)
            progress.advance(task)
    return combined


def _print_summary(mp3_path, combined):
    size_kb = os.path.getsize(mp3_path) / 1024
    duration_s = len(combined) / 1000
    minutes, seconds = divmod(int(duration_s), 60)
    kbps = (size_kb * 8) / duration_s if duration_s > 0 else 0

    table = Table(show_header=False, box=None, padding=(0, 2))
    table.add_column(style="bold green")
    table.add_column(style="white")
    table.add_row("Saved", mp3_path)
    table.add_row("Size", f"{size_kb:.1f} KB")
    table.add_row("Duration", f"{minutes}:{seconds:02d}")
    table.add_row("Bitrate", f"{kbps:.0f} kbps")
    console.print(table)


def convert(docx_path: str, mp3_path: str) -> None:
    """Convert a .docx file to an .mp3 using macOS TTS."""
    paragraphs = [p for p in Document(docx_path).paragraphs if p.text.strip()]
    combined = _build_audio(paragraphs)
    os.makedirs(os.path.dirname(os.path.abspath(mp3_path)), exist_ok=True)
    combined.export(mp3_path, format="mp3")
    _print_summary(mp3_path, combined)
