from docx.api import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from typing import IO

document = Document("source.docx")

separator = "***" #Can change this to whatever section separator applies to document
i = 0 #Variable to check if our paragraph comes after a section separator, so we can remove text-indent

for p in document.paragraphs:
    hasStyle = 0;
    with open('output.html', 'a') as f:  # type: IO[str]
        print("<p", end="", file=f)
    if i == 0:
        with open('output.html', 'a') as f:  # type: IO[str]
            print(" style=\"text-indent:0;\">", end="", file=f)
        i = 1
        hasStyle = 1 #variable to check if the paragraph tag has styles or not, so we can add the closing quotes if necessary.
    else:
        if p.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
            with open('output.html', 'a') as f:  # type: IO[str]
                print(" style=\"text-align:right;>", end="", file=f)
            hasStyle = 1
        elif p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            with open('output.html', 'a') as f:  # type: IO[str]
                print(" style=\"text-align:center;", end="", file=f)
            hasStyle = 1
        if hasStyle == 1:
            with open('output.html', 'a') as f:  # type: IO[str]
                print("\"", end="", file=f)
        with open('output.html', 'a') as f:  # type: IO[str]
            print(">", end="", file=f)
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
        if run.underline:
            with open('output.html', 'a') as f:  # type: IO[str]
                print("</u>", end="", file=f)
        if run.bold:
            with open('output.html', 'a') as f:  # type: IO[str]
                print("</b>", end="", file=f)
        if run.italic:
            with open('output.html', 'a') as f:  # type: IO[str]
                print("</i>", end="", file=f)

    with open('output.html', 'a') as f:  # type: IO[str]
        print("</p>", file=f)
    if p.text == separator:
        i = 0
