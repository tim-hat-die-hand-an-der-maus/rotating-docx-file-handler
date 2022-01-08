import os
import sys
from typing import Tuple
from xml.etree.ElementTree import ElementTree

from markdown import Markdown

from src.rotatingdocxfilehandler import HandlerDocument


LEAF_TAGS = ["h1", "h2", "h3", "h4", "h5", "h6", "em", "strong", "code", "a"]


def parse_leaf(tag, parent, document):
    if tag.tag == 'em' and tag.text:
        parent.add_run(tag.text).italic = True
    elif tag.tag == 'strong' and tag.text:
        parent.add_run(tag.text).bold = True
    # headings
    elif tag.tag.startswith("h") and tag.tag[1].isdigit():
        level = int(tag.tag[1])
        # apply -1 since `0` is `Title`
        document.add_heading(tag.text, level=level - 1)
    elif tag.tag == "code":
        document.add_paragraph(tag.text, style="Quote")
    if tag.tail:
        tail = tag.tail
        parent.add_run(tail)


def paragraph_serializer(document: HandlerDocument, ptag) -> Tuple[int, HandlerDocument]:
    paragraph = document.add_paragraph()

    c = 0
    for tag in ptag.iter():
        if tag.tag in LEAF_TAGS:
            parse_leaf(tag, paragraph, document)
        else:
            print(f"found non-leaf tag in <p>: {tag.tag}")
            break

        c += 1

    return c, document


def docx_serializer(element):
    root = ElementTree(element).getroot()
    document = HandlerDocument("test.docx")

    c = 0
    for tag in root.iter():
        print(f"{tag.tag} '{tag.text}' '{tag.tail}'")
        if c > 0:
            print("c > 0")
            c -= 1
            continue
        if tag.tag == "p":
            print("<p>")
            cmod, document = paragraph_serializer(document, tag)
            c += cmod
        elif tag.tag in LEAF_TAGS:
            print(f"leaf tag <{tag.tag}> in docx")
            paragraph = document.add_paragraph()
            parse_leaf(tag, paragraph, document)
        elif tag.text is not None:
            print(f"add '{tag.text}'")
            document.add_paragraph(tag.text)
        else:
            print("=== <other>")

    return document


def main():
    filename = "test.docx"
    if os.path.exists(filename):
        os.remove(filename)

    with open("test.md") as f:
        source = f.read()
    md_object = Markdown()
    md_object.serializer = docx_serializer
    md_object.stripTopLevelTags = False
    md_object.postprocessors = []

    # noinspection PyTypeChecker
    # overwritten by returning HandlerDocument from out docx_serializer and emptying the postprocessors
    doc: HandlerDocument = md_object.convert(source)
    doc.save(filename)


if __name__ == "__main__":
    main()
