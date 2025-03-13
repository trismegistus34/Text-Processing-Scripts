# Text-Processing-Scripts
Simple Python scripts for processing text files, mainly to convert them to HTML. Making them to use for my personal website, and also for the sake of learning Python. Made using Pycharm and python-docx (https://github.com/python-openxml/python-docx).

## DOCX_to_HTML
Takes a simply-formatted word document and converts it into HTML. Keeps bolded, italicized and underlined words. Takes a section separator (by default "***") and after each line containing said separator, the following paragraph has the text-indent:0 property applied to it. Same for paragraphs which follow headings. Currently doesn't interact with images.
