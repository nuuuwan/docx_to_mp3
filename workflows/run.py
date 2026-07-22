import os
import sys

from src.converter import convert

# Allow running from project root
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))


def main():
    if len(sys.argv) < 2:
        print(
            "Usage: python workflows/run.py <path/to/file.docx> [output_dir]"
        )
        sys.exit(1)

    docx_path = sys.argv[1]
    if len(sys.argv) >= 3:
        output_dir = sys.argv[2]
    else:
        output_dir = os.path.splitext(docx_path)[0] + "_mp3"

    convert(docx_path, output_dir)


if __name__ == "__main__":
    main()
