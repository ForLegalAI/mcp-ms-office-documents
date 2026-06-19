"""Full markdown content processor (handles empty lines, soft breaks, blocks).
This module contains process_markdown_content and process_markdown_block which
orchestrate all block-level and inline parsing into a python-docx Document.
"""
import logging
from .patterns import (
    HEADING_PATTERN,
    PAGE_BREAK_PATTERN,
    HORIZONTAL_LINE_PATTERN,
    IMAGE_PATTERN,
    TABLE_LINE_PATTERN,
    ORDERED_LIST_PATTERN,
    UNORDERED_LIST_PATTERN,
    STYLE_DIRECTIVE_PATTERN,
)
from .inline_formatting import parse_inline_formatting
from .block_elements import (
    parse_table,
    add_table_to_doc,
    process_list_items,
    add_horizontal_line,
    add_image_to_doc,
    detect_alignment,
    process_alignment_block,
)
from .style_map import (
    DEFAULT_STYLE_MAP,
    apply_style,
    add_mapped_heading,
    apply_style_to_block_element,
)
logger = logging.getLogger(__name__)
def process_markdown_content(doc, content, return_elements=False,
                             style_map=DEFAULT_STYLE_MAP):
    """Process full markdown content with all features: spacing, soft breaks, blocks.
    This is the single source of truth for converting a markdown string into
    document elements. Both the base tool and dynamic template placeholder
    replacement use this function.
    Args:
        doc: The python-docx Document instance.
        content: Raw markdown text (may contain newlines).
        return_elements: If True, created elements are detached from the doc body
            and returned (for reinsertion at a specific position).
    Returns:
        List of XML elements if return_elements is True, otherwise an empty list.
    """
    lines = content.split('\n')
    n = len(lines)
    i = 0
    all_elements = []
    while i < n:
        line = lines[i]
        # --- Empty line handling (preserve spacing) ---
        if not line.strip():
            empty_line_count = 1
            i += 1
            while i < n and not lines[i].strip():
                empty_line_count += 1
                i += 1
            if empty_line_count >= 2:
                for _ in range(empty_line_count - 1):
                    p = doc.add_paragraph()
                    if return_elements:
                        all_elements.append(p._p)
                        doc._body._body.remove(p._p)
            continue
        # --- Soft line breaks (trailing two spaces) ---
        if line.endswith('  '):
            paragraph_lines = []
            while i < n:
                current_line = lines[i]
                if not current_line.strip():
                    break
                paragraph_lines.append(current_line)
                i += 1
                if not current_line.endswith('  '):
                    break
            full_text = '  \n'.join(paragraph_lines)
            first_line = paragraph_lines[0].strip()
            if first_line.startswith('#'):
                stripped_hashes = first_line.lstrip('#')
                level = len(first_line) - len(stripped_hashes)
                heading = add_mapped_heading(doc, min(level, 6), style_map)
                parse_inline_formatting(stripped_hashes.strip(), heading)
                elem = heading._p
            elif first_line.startswith('>'):
                quote_text = full_text[1:].strip()
                quote_para = doc.add_paragraph()
                apply_style(quote_para, style_map.quote)
                parse_inline_formatting(quote_text, quote_para)
                elem = quote_para._p
            else:
                para = doc.add_paragraph()
                parse_inline_formatting(full_text, para)
                elem = para._p
            if return_elements:
                all_elements.append(elem)
                doc._body._body.remove(elem)
            continue
        # --- All other block elements: delegate to block processor ---
        i, block_elems = process_markdown_block(doc, lines, i,
                                                return_element=return_elements,
                                                style_map=style_map)
        if return_elements:
            all_elements.extend(block_elems)
    return all_elements
def process_markdown_block(doc, lines, start_idx, return_element=True,
                           style_map=DEFAULT_STYLE_MAP):
    """Process a single markdown block element and return created XML elements.
    Returns:
        Tuple of (next_index, list_of_elements).
    """
    line = lines[start_idx]
    stripped = line.strip()
    elements = []
    def _collect(element):
        """If return_element, detach *element* from body and collect it."""
        if return_element:
            elements.append(element)
            doc._body._body.remove(element)
    try:
        # Heading
        heading_match = HEADING_PATTERN.match(stripped)
        if heading_match:
            level = len(heading_match.group(1))
            heading = add_mapped_heading(doc, min(level, 6), style_map)
            parse_inline_formatting(heading_match.group(2), heading)
            _collect(heading._p)
            return start_idx + 1, elements
        # Table (lines starting with |)
        if TABLE_LINE_PATTERN.match(stripped):
            table_data, col_alignments, next_idx = parse_table(lines, start_idx)
            if table_data:
                # Check preceding lines for table directives
                borderless = False
                col_widths = None
                for lookback in range(1, 3):  # check up to 2 lines back
                    if start_idx - lookback < 0:
                        break
                    prev = lines[start_idx - lookback].strip().lower()
                    if prev in ('<!-- borderless -->', '<!--borderless-->'):
                        borderless = True
                    elif prev.startswith('<!-- widths:') and prev.endswith('-->'):
                        # Parse <!-- widths: 30 70 --> or <!-- widths: 20 50 30 -->
                        inner = prev[len('<!-- widths:'):-len('-->')].strip()
                        try:
                            col_widths = [float(v) for v in inner.split()]
                        except ValueError:
                            col_widths = None
                word_table = add_table_to_doc(table_data, doc,
                                             col_alignments=col_alignments,
                                             borderless=borderless,
                                             col_widths=col_widths,
                                             table_style=style_map.table)
                if word_table is not None:
                    _collect(word_table._tbl)
                return next_idx, elements
        # Page break (---)
        if PAGE_BREAK_PATTERN.match(stripped):
            doc.add_page_break()
            _collect(doc.paragraphs[-1]._p)
            return start_idx + 1, elements
        # Horizontal line (***)
        if HORIZONTAL_LINE_PATTERN.match(stripped):
            _collect(add_horizontal_line(doc)._p)
            return start_idx + 1, elements
        # Image (![alt](url))
        img_match = IMAGE_PATTERN.match(stripped)
        if img_match:
            body = doc._body._body
            existing_children = list(body) if return_element else None
            add_image_to_doc(doc, img_match.group(2), img_match.group(1))
            if return_element:
                for element in list(body)[len(existing_children):]:
                    elements.append(element)
                    body.remove(element)
            return start_idx + 1, elements
        # Alignment (inline or block-open)
        align_result = detect_alignment(stripped)
        if align_result is not None:
            inner, alignment = align_result
            if inner is not None:
                para = doc.add_paragraph()
                para.alignment = alignment
                parse_inline_formatting(inner, para)
                _collect(para._p)
                return start_idx + 1, elements
            else:
                idx, block_elems = process_alignment_block(
                    lines, start_idx + 1, doc, alignment, return_elements=return_element
                )
                if return_element and block_elems:
                    elements.extend(block_elems)
                return idx, elements
        # Ordered list
        if ORDERED_LIST_PATTERN.match(stripped):
            return process_list_items(
                lines, start_idx, doc, is_ordered=True, level=0, return_elements=return_element,
                number_styles=style_map.list_number, bullet_styles=style_map.list_bullet,
            )
        # Unordered list
        if UNORDERED_LIST_PATTERN.match(stripped):
            return process_list_items(
                lines, start_idx, doc, is_ordered=False, level=0, return_elements=return_element,
                number_styles=style_map.list_number, bullet_styles=style_map.list_bullet,
            )
        # Blockquote (> text)
        if stripped.startswith('>'):
            quote_text = stripped[1:].strip()
            quote_para = doc.add_paragraph()
            apply_style(quote_para, style_map.quote)
            parse_inline_formatting(quote_text, quote_para)
            _collect(quote_para._p)
            return start_idx + 1, elements
        # Style directive: <!-- style: Name --> applies Name to the next block.
        style_directive = STYLE_DIRECTIVE_PATTERN.match(stripped)
        if style_directive:
            style_name = style_directive.group(1).strip()
            # Skip the directive line and any blank lines to reach the next block.
            next_idx = start_idx + 1
            while next_idx < len(lines) and not lines[next_idx].strip():
                next_idx += 1
            if next_idx >= len(lines):
                return next_idx, elements  # nothing follows; directive is a no-op
            body = doc._body._body
            # Snapshot existing children (holding the proxies alive so lxml keeps a
            # stable identity for them); new elements are inserted before the trailing
            # <w:sectPr>, so positional slicing would be unreliable.
            existing = None if return_element else list(body)
            new_idx, block_elems = process_markdown_block(
                doc, lines, next_idx, return_element=return_element, style_map=style_map
            )
            if return_element:
                produced = block_elems
            else:
                produced = [el for el in list(body) if el not in existing]
            for el in produced:
                apply_style_to_block_element(doc, el, style_name)
            if return_element:
                elements.extend(block_elems)
            return new_idx, elements
        # HTML comment directives (e.g. <!-- borderless -->) — skip silently
        if stripped.startswith('<!--') and stripped.endswith('-->'):
            return start_idx + 1, elements
        # Regular paragraph
        para = doc.add_paragraph()
        parse_inline_formatting(stripped, para)
        _collect(para._p)
        return start_idx + 1, elements
    except Exception as e:
        logger.error("Failed to process markdown block at line %d: %s", start_idx, e, exc_info=True)
        return start_idx + 1, elements
