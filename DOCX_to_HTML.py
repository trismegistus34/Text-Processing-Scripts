from docx.api import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from typing import IO

document = Document("source.docx")

for p in document.paragraphs:
    if p.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
        with open('output.html', 'a') as f:  # type: IO[str]
            print("<p style=\"text-align:right;\">", end="", file=f)
    elif p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
        with open('output.html', 'a') as f:  # type: IO[str]
            print("<p style=\"text-align:center;\">", end="", file=f)
    else:
        with open('output.html', 'a') as f:  # type: IO[str]
            print("<p>", end="", file=f)
    for run in p.runs:
        if run.italic:
            with open('output.html', 'a') as f:  # type: IO[str]
                print("<i>", end="", file=f)
        if run.bold:
            with open('output.html', 'a') as f:  # type: IO[str]
                print("<b>", end="", file=f)
        if run.underline:
            with open('output.html', 'a') as f:  # type: IO[str]
                print("<u>", end="", file=f)
        with open('output.html', 'a') as f:  # type: IO[str]
            print(run.text, end="", file=f)
        if run.italic:
            with open('output.html', 'a') as f:  # type: IO[str]
                print("</i>", end="", file=f)
        if run.bold:
            with open('output.html', 'a') as f:  # type: IO[str]
                print("</b>", end="", file=f)
        if run.underline:
            with open('output.html', 'a') as f:  # type: IO[str]
                print("</u>", end="", file=f)
    with open('output.html', 'a') as f:  # type: IO[str]
        print("</p>", file=f)

