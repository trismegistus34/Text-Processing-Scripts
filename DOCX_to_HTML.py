from docx.api import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

document = Document("source.docx")

separator = "***" #Can change this to whatever section separator applies to document
i = 0 #Variable to check if our paragraph comes after a section separator, so we can remove text-indent

with open('output.html', 'w', encoding='utf-8') as f:
    f.write(
        '<!DOCTYPE html>\n<html>\n<head>\n<meta charset="UTF-8">\n<title>Converted Document</title>\n</head>\n<body>\n')

    for p in document.paragraphs:
        hasStyle = 0
        print("<p", end="", file=f)
        if i == 0:
            print(" style=\"text-indent:0;\">", end="", file=f)
            i = 1
            hasStyle = 1 #variable to check if the paragraph tag has styles or not, so we can add the closing quotes if necessary.
        else:
            if p.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                print(" style=\"text-align:right;>", end="", file=f)
                hasStyle = 1
            elif p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                print(" style=\"text-align:center;", end="", file=f)
                hasStyle = 1
            if hasStyle == 1:
                print("\"", end="", file=f)
            print(">", end="", file=f)
        for run in p.runs:
            if run.italic:
                print("<i>", end="", file=f)
            if run.bold:
                print("<b>", end="", file=f)
            if run.underline:
                print("<u>", end="", file=f)
            print(run.text, end="", file=f)
            if run.underline:
                print("</u>", end="", file=f)
            if run.bold:
                print("</b>", end="", file=f)
            if run.italic:
                print("</i>", end="", file=f)
        print("</p>", file=f)
        if p.text == separator:
            i = 0
        if p.style.name.startswith("Heading") or p.style.name == "Subchapter" or p.style.name == "Title":
            i = 0
