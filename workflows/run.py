import os
import sys

from src.converter import convert

# Allow running from project root
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))


def main():
    if len(sys.argv) < 2:
        print(
            "Usage: python workflows/run.py <path/to/file.docx> [output.mp3]"
        )
        sys.exit(1)

    docx_path = sys.argv[1]
    if len(sys.argv) >= 3:
        mp3_path = sys.argv[2]
    else:
        mp3_path = os.path.splitext(docx_path)[0] + ".mp3"

    convert(docx_path, mp3_path)


if __name__ == "__main__":
    main()
