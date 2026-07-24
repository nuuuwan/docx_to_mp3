import hashlib
import os
import re
import shutil
import subprocess
import tempfile

from docx import Document
from pydub import AudioSegment
from rich.console import Console
from rich.progress import BarColumn, MofNCompleteColumn, Progress, TextColumn
from rich.table import Table

console = Console()

VOICE_HEADING = "Jamie"
VOICE_BODY = "Serena"

PAUSE_BEFORE_HEADING_MS = 500
PAUSE_AFTER_HEADING_MS = 500

PAUSE_AFTER_PARAGRAPH_MS = 250

MAX_CHUNK_MS = 60 * 60 * 1000

CACHE_DIR = os.path.join(tempfile.gettempdir(), "docx_to_mp3", "tts")
FILE_CACHE_DIR = os.path.join(
    tempfile.gettempdir(), "docx_to_mp3", "file_cache"
)

# Paragraph styles that represent metadata / navigation rather than content
_SKIP_STYLE_PREFIXES = ("toc", "index", "bibliography", "caption")
_SKIP_STYLES_EXACT = {
    "TOC Heading",
    "Caption",
    "Header",
    "Footer",
}


def _is_heading(paragraph):
    style = paragraph.style.name
    return style.startswith("Heading") or style in ("Title", "Subtitle")


def _is_metadata(paragraph):
    """Return True for TOC entries, captions, headers, footers, etc."""
    name = paragraph.style.name
    if name in _SKIP_STYLES_EXACT:
        return True
    lower = name.lower()
    return any(lower.startswith(prefix) for prefix in _SKIP_STYLE_PREFIXES)


def _cache_path(text, voice):
    key = hashlib.sha256(f"{voice}:{text}".encode()).hexdigest()[:16]
    return os.path.join(CACHE_DIR, f"{key}.mp3")


def _file_cache_dir(paragraphs):
    """Return the cache directory for the assembled output of this exact set of paragraphs."""
    parts = []
    for para in paragraphs:
        text = para.text.strip()
        voice = VOICE_HEADING if _is_heading(para) else VOICE_BODY
        parts.append(f"{voice}:{text}")
    key = hashlib.sha256("\n".join(parts).encode()).hexdigest()[:16]
    return os.path.join(FILE_CACHE_DIR, key)


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


PAUSE_AFTER_SEPARATOR_MS = 2000

# Texts that should produce silence instead of being read aloud
_SEPARATOR_TEXTS = {"---", "...", "\u2026", "***", "* * *", "- - -"}


def _is_separator(text):
    return text in _SEPARATOR_TEXTS or set(text).issubset(set("-*~ \t"))


def _para_to_segment(para):
    text = para.text.strip()
    if _is_separator(text):
        return AudioSegment.silent(duration=PAUSE_AFTER_SEPARATOR_MS)

    is_heading = _is_heading(para)
    voice = VOICE_HEADING if is_heading else VOICE_BODY

    pause_before_ms = PAUSE_BEFORE_HEADING_MS if is_heading else 0

    pause_after_ms = (
        PAUSE_AFTER_HEADING_MS if is_heading else PAUSE_AFTER_PARAGRAPH_MS
    )

    cached = _cache_path(text, voice)
    if not os.path.exists(cached):
        os.makedirs(CACHE_DIR, exist_ok=True)
        _say_to_mp3(text, voice, cached)

    return (
        AudioSegment.silent(duration=pause_before_ms)
        + AudioSegment.from_mp3(cached)
        + AudioSegment.silent(duration=pause_after_ms)
    )


def _slugify(text, max_len=40):
    """Convert heading text to a safe filename fragment."""
    slug = re.sub(r"[^\w\s-]", "", text.lower())
    slug = re.sub(r"[\s_-]+", "_", slug).strip("_")
    return slug[:max_len] if slug else "part"


def _flush_chunk(audio, label, index, output_dir):
    """Export a completed chunk to disk immediately and return its path."""
    slug = _slugify(label) if label else f"part_{index}"
    filename = f"{index:03d}_{slug}.mp3"
    out_path = os.path.join(output_dir, filename)
    audio.export(out_path, format="mp3")

    size_mb = os.path.getsize(out_path) / (1024 * 1024)
    dur_s = len(audio) / 1000
    mins, secs = divmod(int(dur_s), 60)
    console.print(
        f"  [green]{filename}[/green]  "
        f"[white]{mins}:{secs:02d}  {size_mb:.2f} MB[/white]"
    )
    return out_path


def _print_summary(output_dir, paths, total_ms, paragraphs):
    total_size_mb = sum(os.path.getsize(p) / (1024 * 1024) for p in paths)
    mins, secs = divmod(int(total_ms / 1000), 60)
    all_text = " ".join(p.text.strip() for p in paragraphs)
    word_count = len(all_text.split())

    table = Table(show_header=False, box=None, padding=(0, 2))
    table.add_column(style="bold green")
    table.add_column(style="white")
    table.add_row("Output folder", output_dir)
    table.add_row("Files", str(len(paths)))
    table.add_row("Total duration", f"{mins}:{secs:02d}")
    table.add_row("Total size", f"{total_size_mb:.2f} MB")
    table.add_row("Paragraphs", f"{len(paragraphs):,}")
    table.add_row("Words", f"{word_count:,}")
    console.print(table)


def convert(docx_path: str, output_dir: str) -> None:
    """Convert a .docx file to a folder of .mp3 files split by heading and duration.

    Each chunk is written to disk the moment its duration cap is reached,
    so memory stays bounded and files appear incrementally.
    """
    paragraphs = [
        p
        for p in Document(docx_path).paragraphs
        if p.text.strip() and not _is_metadata(p)
    ]
    os.makedirs(output_dir, exist_ok=True)

    # File-level cache: if all paragraphs are unchanged the assembled chunks
    # are already stored; copy them to output_dir and return early.
    file_cache = _file_cache_dir(paragraphs)
    cached_chunks = (
        sorted([f for f in os.listdir(file_cache) if f.endswith(".mp3")])
        if os.path.isdir(file_cache)
        else []
    )
    if cached_chunks:
        console.print(
            "[bold cyan]File unchanged — copying cached output.[/bold cyan]"
        )
        cached_paths = []
        for name in cached_chunks:
            src = os.path.join(file_cache, name)
            dst = os.path.join(output_dir, name)
            shutil.copy2(src, dst)
            cached_paths.append(dst)
        total_cached_ms = sum(
            len(AudioSegment.from_mp3(p)) for p in cached_paths
        )
        _print_summary(output_dir, cached_paths, total_cached_ms, paragraphs)
        return

    paths = []
    chunk_index = 1
    current_label = ""
    current_audio = AudioSegment.empty()
    current_ms = 0
    total_ms = 0

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
            style_tag = "bold yellow" if is_heading else "white"
            preview = text[:80] + "\u2026" if len(text) > 80 else text
            progress.console.print(f"  [{style_tag}]{preview}[/{style_tag}]")

            seg = _para_to_segment(para)
            seg_ms = len(seg)

            # Flush at a heading boundary or when the cap would be exceeded
            should_split = (is_heading and current_ms > 0) or (
                current_ms + seg_ms > MAX_CHUNK_MS and current_ms > 0
            )

            if should_split:
                total_ms += current_ms
                paths.append(
                    _flush_chunk(
                        current_audio, current_label, chunk_index, output_dir
                    )
                )
                chunk_index += 1
                current_audio = AudioSegment.empty()
                current_ms = 0
                if is_heading:
                    current_label = text

            if is_heading and current_ms == 0:
                current_label = text

            current_audio += seg
            current_ms += seg_ms
            progress.advance(task)

    # Flush the final chunk
    if current_ms > 0:
        total_ms += current_ms
        paths.append(
            _flush_chunk(
                current_audio, current_label, chunk_index, output_dir
            )
        )

    # Populate the file-level cache with the freshly generated chunks.
    os.makedirs(file_cache, exist_ok=True)
    for p in paths:
        shutil.copy2(p, os.path.join(file_cache, os.path.basename(p)))

    _print_summary(output_dir, paths, total_ms, paragraphs)
