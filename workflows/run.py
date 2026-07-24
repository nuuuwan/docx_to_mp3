import os
import sys

from DocxFile import DocxFile

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(
            "Usage: python workflows/run.py <path/to/file.docx> [output_dir]"
        )
        sys.exit(1)

    docx_path = sys.argv[1]
    DocxFile(docx_path).build_audio()
