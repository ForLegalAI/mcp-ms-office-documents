"""Build Word content controls in tests, as Word writes them (no Word needed)."""
from docx.oxml import parse_xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_NS = f'xmlns:w="{W}" xmlns:w14="{W14}"'


def _pr(tag, alias, sdt_id, extra):
    tag_xml = f'<w:tag w:val="{tag}"/>' if tag is not None else ""
    alias_xml = f'<w:alias w:val="{alias}"/>' if alias else ""
    return f'<w:sdtPr>{alias_xml}{tag_xml}<w:id w:val="{sdt_id}"/>{extra}</w:sdtPr>'


def checkbox(tag=None, checked=False, sdt_id=1, alias=None):
    glyph = "☒" if checked else "☐"
    box = (f'<w14:checkbox><w14:checked w14:val="{1 if checked else 0}"/>'
           '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>'
           '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>')
    return parse_xml(
        f'<w:sdt {_NS}>{_pr(tag, alias, sdt_id, box)}<w:sdtContent><w:r><w:rPr>'
        '<w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr>'
        f'<w:t>{glyph}</w:t></w:r></w:sdtContent></w:sdt>')


def dropdown(tag=None, items=(("basic", "Basic plan"), ("pro", "Pro plan")),
             sdt_id=2, alias=None, prompt="Choose an item."):
    list_xml = "".join(f'<w:listItem w:displayText="{d}" w:value="{v}"/>' for v, d in items)
    extra = ('<w:showingPlcHdr/>'
             f'<w:dropDownList>{list_xml}</w:dropDownList>')
    return parse_xml(
        f'<w:sdt {_NS}>{_pr(tag, alias, sdt_id, extra)}<w:sdtContent><w:r><w:rPr>'
        f'<w:rStyle w:val="PlaceholderText"/><w:b/></w:rPr><w:t>{prompt}</w:t></w:r>'
        '</w:sdtContent></w:sdt>')


def text(tag=None, sdt_id=3, alias=None, multiline=False, prompt="Click here.",
         binding=False, block=False):
    extra = ('<w:showingPlcHdr/>'
             + ('<w:dataBinding w:xpath="/root/x" w:storeItemID="{00000000-0000-0000-0000-000000000000}"/>'
                if binding else "")
             + ('<w:text w:multiLine="1"/>' if multiline else '<w:text/>'))
    run = (f'<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/><w:i/></w:rPr>'
           f'<w:t>{prompt}</w:t></w:r>')
    content = f'<w:p>{run}</w:p>' if block else run
    return parse_xml(
        f'<w:sdt {_NS}>{_pr(tag, alias, sdt_id, extra)}'
        f'<w:sdtContent>{content}</w:sdtContent></w:sdt>')


def rich_text(tag=None, sdt_id=4):
    return parse_xml(
        f'<w:sdt {_NS}>{_pr(tag, None, sdt_id, "")}<w:sdtContent><w:r><w:t>rich</w:t>'
        '</w:r></w:sdtContent></w:sdt>')


def add_inline(paragraph, sdt):
    """Put a run-level control at the end of *paragraph*; return it."""
    paragraph._p.append(sdt)
    return sdt


def add_block(doc, sdt):
    """Put a block-level control in the body, before the section properties."""
    doc.element.body.insert(len(doc.element.body) - 1, sdt)
    return sdt


def checked(sdt):
    """``(ticked, glyph)`` of a check box."""
    state = sdt.find(f".//{{{W14}}}checked").get(f"{{{W14}}}val") == "1"
    return state, sdt.find(f".//{{{W}}}t").text


def content_text(sdt):
    """The visible text of a control (line breaks as ``\\n``)."""
    out = []
    for el in sdt.find(f"{{{W}}}sdtContent").iter():
        if el.tag == f"{{{W}}}t":
            out.append(el.text or "")
        elif el.tag == f"{{{W}}}br":
            out.append("\n")
    return "".join(out)
