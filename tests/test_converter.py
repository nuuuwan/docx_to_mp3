import os
import shutil
import unittest

from src.converter import CACHE_DIR, convert

TEST_DOCX = os.path.join(os.path.dirname(__file__), "test.docx")
TEST_OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "test_output")


class TestConverter(unittest.TestCase):
    def setUp(self):
        if os.path.exists(CACHE_DIR):
            shutil.rmtree(CACHE_DIR)
        if os.path.exists(TEST_OUTPUT_DIR):
            shutil.rmtree(TEST_OUTPUT_DIR)

    def tearDown(self):
        if os.path.exists(TEST_OUTPUT_DIR):
            shutil.rmtree(TEST_OUTPUT_DIR)

    def test_convert_produces_mp3_folder(self):
        convert(TEST_DOCX, TEST_OUTPUT_DIR)
        self.assertTrue(
            os.path.isdir(TEST_OUTPUT_DIR), "Output folder was not created"
        )
        mp3_files = [
            f for f in os.listdir(TEST_OUTPUT_DIR) if f.endswith(".mp3")
        ]
        self.assertGreater(len(mp3_files), 0, "No MP3 files were created")
        for mp3 in mp3_files:
            path = os.path.join(TEST_OUTPUT_DIR, mp3)
            self.assertGreater(os.path.getsize(path), 0, f"{mp3} is empty")
