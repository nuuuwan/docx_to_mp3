import hashlib
import os
import subprocess
import tempfile
from abc import ABC
from functools import cache, cached_property

from docx import Document
from pydub import AudioSegment


class TextToSpeechCache:
    @staticmethod
    def get_mp3(text: str, voice="Alex") -> str:
        h = hashlib.sha256(f"{voice}:{text}".encode()).hexdigest()[:16]
        DIR_TTS_CACHE = os.path.join(
            tempfile.gettempdir(), "docx_to_mp3", "tts_cache"
        )
        os.makedirs(DIR_TTS_CACHE, exist_ok=True)

        temp_path = os.path.join(DIR_TTS_CACHE, f"{h}.mp3")
        subprocess.run(
            ["say", "-v", voice, "-o", temp_path, "--", text],
            check=True,
            capture_output=True,
        )
        print(f'"{text}" -> {temp_path}')
        return temp_path


class Audible(ABC):
    def get_audio(self) -> AudioSegment:
        raise NotImplementedError()


class Heading(Audible):
    def __init__(self, text: str):
        self.text = text

    def get_audio(self) -> AudioSegment:
        temp_file_path = TextToSpeechCache.get_mp3(self.text, voice="Jamie")
        return AudioSegment.from_mp3(temp_file_path)


class Paragraph(Audible):
    def __init__(self, text: str):
        self.text = text

    def get_audio(self) -> AudioSegment:
        temp_file_path = TextToSpeechCache.get_mp3(self.text, voice="Serena")
        return AudioSegment.from_mp3(temp_file_path)


class Chapter(Audible):
    def __init__(
        self, i_chapter: int, title: str, paragraphs: list[Paragraph]
    ):
        self.i_chapter = i_chapter
        self.title = title
        self.paragraphs = paragraphs

    def get_audio(self) -> AudioSegment:
        audio = Heading(self.title).get_audio()
        for para in self.paragraphs:
            audio += para.get_audio()
        return audio


class DocxFile(Audible):
    def __init__(self, file_path: str):
        self.file_path = file_path

    @cached_property
    def chapters(self):
        doc = Document(self.file_path)
        chapters = []
        current_chapter_title = None
        current_paragraphs = []

        for para in doc.paragraphs:
            if para.text.strip() == "":
                continue
            if para.style.name.startswith("Heading"):
                if current_chapter_title is not None:
                    chapters.append(
                        Chapter(
                            len(chapters) + 1,
                            current_chapter_title,
                            current_paragraphs,
                        )
                    )
                current_chapter_title = para.text
                current_paragraphs = []
            else:
                current_paragraphs.append(Paragraph(para.text))

        if current_chapter_title is not None:
            chapters.append(
                Chapter(
                    len(chapters) + 1,
                    current_chapter_title,
                    current_paragraphs,
                )
            )

        return chapters

    def get_audio(self) -> AudioSegment:
        audio = AudioSegment.silent(duration=0)
        for chapter in self.chapters:
            audio += chapter.get_audio()
        return audio

    def build_audio(self):
        output_dir = os.path.splitext(self.file_path)[0] + ".audio"
        os.makedirs(output_dir, exist_ok=True)

        # chapters
        output_chapters_dir = os.path.join(output_dir, "chapters")
        os.makedirs(output_chapters_dir, exist_ok=True)

        for chapter in self.chapters:
            chapter_audio = chapter.get_audio()
            chapter_file_path = os.path.join(
                output_chapters_dir, f"chapter-{chapter.i_chapter:02d}.mp3"
            )
            chapter_audio.export(chapter_file_path, format="mp3")
            print(
                f"Chapter {chapter.i_chapter} exported to {chapter_file_path}"
            )

        # docx
        audio = self.get_audio()
        output_file_path = os.path.join(output_dir, "docx.mp3")
        audio.export(output_file_path, format="mp3")
        print(f"Audio exported to {output_file_path}")
