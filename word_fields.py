"""Parse Word fields once, retaining boundaries, editable instructions, and visible text.

The returned index is a snapshot of the XML tree. It handles nested fields and fields
spanning paragraphs; bookmarks and proofing markers do not change field boundaries.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def instruction_text(elements):
    """Read both complex-field instrText nodes and simple-field instr attributes."""
    return "".join(
        element.get(W + "instr", "") if element.tag == W + "fldSimple" else element.text or ""
        for element in elements
    )


def replace_instruction(elements, text):
    """Rewrite only this field's instruction, leaving nested fields untouched."""
    for index, element in enumerate(elements):
        value = text if index == 0 else ""
        if element.tag == W + "fldSimple":
            element.set(W + "instr", value)
        else:
            element.text = value


@dataclass(eq=False)
class WordField:
    start: etree._Element
    paragraph: etree._Element | None
    paragraph_index: int
    parent: WordField | None
    elements: list = field(default_factory=list)
    display: list[str] = field(default_factory=list)
    separated: bool = False
    end: etree._Element | None = None

    @property
    def instruction(self):
        return instruction_text(self.elements)


@dataclass
class FieldIndex:
    fields: list[WordField] = field(default_factory=list)
    paragraphs: list = field(default_factory=list)
    outside: dict = field(default_factory=dict)
    protected: set = field(default_factory=set)
    begins: int = 0
    ends: int = 0
    structure_errors: int = 0

    def collected(self, with_display=True):
        return [
            (f.elements, "".join(f.display) if with_display else "", f.paragraph_index)
            for f in self.fields
            if f.end is not None
        ]

    def nested_count(self, instruction_marker):
        """Count matching fields contained inside another matching field."""
        count = 0
        for current in self.fields:
            if instruction_marker not in current.instruction:
                continue
            parent = current.parent
            while parent is not None:
                if instruction_marker in parent.instruction:
                    count += 1
                    break
                parent = parent.parent
        return count


def parse_fields(root):
    """Index fields in XML order, or reuse an index while its tree is unchanged.

    Each XML node is visited once, including paragraphs inside text boxes. A text
    box is a separate Word story, so its field stack cannot close a body field.
    Protected nodes include incomplete fields: conversion must never treat their
    cached display as plain text, even when verification reports a malformed field.
    """
    if isinstance(root, FieldIndex):
        return root
    result = FieldIndex()
    stack = []
    stories = []
    paragraphs = []
    paragraph_numbers = {}
    for event, element in etree.iterwalk(root, events=("start", "end")):
        tag = element.tag
        if event == "end":
            if tag == W + "fldSimple":
                while stack and stack[-1].start is not element:
                    result.structure_errors += 1
                    stack.pop()
                if stack:
                    stack.pop().end = element
            elif tag == W + "p":
                paragraphs.pop()
            elif tag == W + "txbxContent":
                result.structure_errors += len(stack)
                stack = stories.pop()
            continue

        if tag == W + "txbxContent":
            stories.append(stack)
            stack = []
        if tag == W + "p":
            paragraph_numbers[element] = len(result.paragraphs)
            result.paragraphs.append(element)
            result.outside[element] = []
            paragraphs.append(element)
        paragraph = paragraphs[-1] if paragraphs else None

        if stack or tag in (W + "fldChar", W + "fldSimple"):
            result.protected.add(element)
        if tag == W + "fldSimple" or (
            tag == W + "fldChar" and element.get(W + "fldCharType") == "begin"
        ):
            current = WordField(
                element,
                paragraph,
                paragraph_numbers.get(paragraph, -1),
                stack[-1] if stack else None,
            )
            if tag == W + "fldSimple":
                current.elements.append(element)
                current.separated = True
            else:
                result.begins += 1
            result.fields.append(current)
            stack.append(current)
        elif tag == W + "fldChar":
            kind = element.get(W + "fldCharType")
            if kind == "end":
                result.ends += 1
                if stack and stack[-1].start.tag != W + "fldSimple":
                    stack.pop().end = element
                else:
                    result.structure_errors += 1
            elif kind == "separate" and stack and not stack[-1].separated:
                stack[-1].separated = True
            else:
                result.structure_errors += 1
        elif tag == W + "instrText" and stack and not stack[-1].separated:
            stack[-1].elements.append(element)
        elif tag == W + "t":
            if not stack and paragraph is not None:
                result.outside[paragraph].append(element.text or "")
            else:
                # A nested result belongs to its own field and any enclosing result,
                # but never to an enclosing instruction that has not separated yet.
                for current in reversed(stack):
                    if not current.separated:
                        break
                    current.display.append(element.text or "")
    result.structure_errors += len(stack)
    result.outside = {p: "".join(parts) for p, parts in result.outside.items()}
    return result
