import os
import unittest

from src.converter import convert

TEST_DOCX = os.path.join(os.path.dirname(__file__), "test.docx")
TEST_MP3 = os.path.join(os.path.dirname(__file__), "test.mp3")


class TestConverter(unittest.TestCase):
    def test_convert_produces_mp3(self):
        convert(TEST_DOCX, TEST_MP3)
        self.assertTrue(os.path.exists(TEST_MP3), "MP3 file was not created")
        self.assertGreater(os.path.getsize(TEST_MP3), 0, "MP3 file is empty")
