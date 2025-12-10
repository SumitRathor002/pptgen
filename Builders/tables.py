import re
import os
import base64
import math
import requests
from io import BytesIO
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE
import copy

from Builders.shapes import add_bg_shape, add_image_element, add_shape_element
from Builders.charts import add_chart_element
from Builders.texts import add_list_paragraphs

from Builders.utils import SubElement, pixels_to_emu, safe_float, get_font_size_pt, is_bold, parse_color, get_para_alignment, get_vertical_alignment, TEXT_HTML_TAGS
from Builders.utils import parse_border_radius, is_uniform_border, make_rounded_image
from Builders.extractors import extract_list_info_from_hierarchy
from Builders.logger import logger

CELL_BORDER_COLOR = "808080"  # Gray color for cell borders
CELL_BORDER_WIDTH = '9525'  # 0.75pt in EMU
TABLE_CELL_HEIGHT_REDUCTION_FACTOR = 1


def _set_cell_border(cell, border_color="000000", border_width='12700'):
    """
    Ensure the given pptx table cell has explicit border XML nodes for all four sides.

    This function manipulates the cell's tcPr XML to remove prior border nodes with
    the same tags then inserts new a:lnL/a:lnR/a:lnT/a:lnB nodes with the specified
    color and width to produce consistent borders across PPTX renderers.

    Parameters:
        cell: pptx.table._Cell instance to modify.
        border_color (str): Hex color (without '#') for border strokes.
        border_width (str): Width in EMU string expected by pptx XML.

    Returns:
        The modified cell instance.
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    for lines in ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB']:
        # Every time before a node is inserted, the nodes with the same tag should be removed.
        tag = lines.split(":")[-1]
        for e in tcPr.getchildren():
            if tag in str(e.tag):
                tcPr.remove(e)
        # end
        
        # Create new border element
        ln = SubElement(tcPr, lines, w=border_width, cap='flat', cmpd='sng', algn='ctr')
        solidFill = SubElement(ln, 'a:solidFill')
        SubElement(solidFill, 'a:srgbClr', val=border_color)
        SubElement(ln, 'a:prstDash', val='solid')
        SubElement(ln, 'a:round')
        SubElement(ln, 'a:headEnd', type='none', w='med', len='med')
        SubElement(ln, 'a:tailEnd', type='none', w='med', len='med')
    
    return cell


def set_cell_border_enhanced(cell, side, width_px, color_rgb, style='solid', default_border_color=CELL_BORDER_COLOR):
    """Apply border to individual side of a cell - used for non-uniform borders"""

    def is_valid_hex_color(color):
        hex_pattern = r'^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$'
        return bool(re.match(hex_pattern, color))

    if width_px <= 0 or not color_rgb:
        return

    # Convert color to hex string
    color_hex = str(color_rgb).replace("#", "").lower()
    if not is_valid_hex_color(color_hex):
        color_hex = default_border_color

    # Convert width to EMU (1pt = 12700 EMU, 1px ≈ 0.75pt)
    width_emu = str(int(width_px * 12700))
    
    # Map side to line element
    side_tag_map = {
        'left': 'a:lnL',
        'right': 'a:lnR',
        'top': 'a:lnT',
        'bottom': 'a:lnB'
    }
    ln_tag = side_tag_map.get(side.lower())
    if not ln_tag:
        return

    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    # Remove existing border if present
    tag = ln_tag.split(":")[-1]
    for e in tcPr.getchildren():
        if tag in str(e.tag):
            tcPr.remove(e)
    
    # Create new border
    ln = SubElement(tcPr, ln_tag, w=width_emu, cap='flat', cmpd='sng', algn='ctr')
    solidFill = SubElement(ln, 'a:solidFill')
    srgbClr = SubElement(solidFill, 'a:srgbClr', val=color_hex)
    
    # Set dash style based on border style
    dash_val = 'solid'
    if style == 'dashed':
        dash_val = 'dash'
    elif style == 'dotted':
        dash_val = 'sysDot'
    
    prstDash = SubElement(ln, 'a:prstDash', val=dash_val)
    round_ = SubElement(ln, 'a:round')
    headEnd = SubElement(ln, 'a:headEnd', type='none', w='med', len='med')
    tailEnd = SubElement(ln, 'a:tailEnd', type='none', w='med', len='med')


def set_cell_border(cell, side, width_px, color_rgb, style='solid'):
    """Apply a border to a table cell side using enhanced method"""
    set_cell_border_enhanced(cell, side, width_px, color_rgb, style)


def set_appropriate_column_width(table, table_info, cols, table_width):
    """
    Assign reasonable pptx column widths based on the first rendered row cell widths.

    Parameters:
        table: pptx.table.Table object whose columns will be sized.
        table_info (dict): Table metadata including rows with cell rects and colSpans.
        cols (int): Total number of columns in the table.
        table_width (float): Total table width in pixels for distributing remaining space.

    Notes:
        - Uses per-cell widths and colSpan to set column.width in EMU.
        - Ensures a minimum column width to avoid zero-width columns.
    """
    
    cols_added = set()
    used_width = 0
    for c_idx in range(cols):
        cells_col = []
        # need to track prev row span else cidx might access wrong cells.
        prev_row_span = None
        for row in table_info['rows']:
            cells = row.get('cells', [])
            if c_idx >= len(cells):
                continue
            
            if prev_row_span and (prev_row_span - 1):
                prev_row_span -= 1
                continue
            
            cell = cells[c_idx]
            width = cell.get('rect', {}).get('width') or 8
            col_span = cell.get('colSpan') or 1
            row_span = cell.get('rowSpan') or 1
            prev_row_span = row_span
            cells_col.append((col_span, width))
        
        if not cells_col:
            continue
            
        cells_col.sort(key=lambda x: (x[0], -x[1]))      
        calculated_width = cells_col[0][1]
        table.columns[c_idx].width = pixels_to_emu(max(8, calculated_width))
        cols_added.add(c_idx)
        used_width += calculated_width

    if cols_added:
        remaining_cols = cols - max(cols_added) + 1
        per_cell_width = (table_width - used_width) // remaining_cols if remaining_cols > 0 else 0

        for c_idx in range(max(cols_added) + 1, cols):
            if c_idx in cols_added:
                continue
            table.columns[c_idx].width = pixels_to_emu(max(8, per_cell_width))

    # --- Column width calculation ---
    # current_col = 0
    # row_cells = first_row['cells']
    # total_assigned = 0.0
    # for cell_data in row_cells:
    #     cell_rect = cell_data.get('rect', {})
    #     cell_w = cell_rect.get('width', 0)
    #     col_span = max(1, cell_data.get('colSpan', 1))
    #     per_col_w = cell_w / col_span if col_span > 1 else cell_w
    #     for i in range(col_span):
    #         if current_col + i < cols:
    #             table.columns[current_col + i].width = pixels_to_emu(max(8, per_col_w))
    #             total_assigned += per_col_w
    #     current_col += col_span
    # # If any remaining columns distribute remaining space
    # remaining_cols = cols - current_col
    # if remaining_cols > 0:
    #     remaining_w = max(0, width - total_assigned)
    #     per_remaining = remaining_w / remaining_cols if remaining_cols else 0
    #     for i in range(remaining_cols):
    #         idx = current_col + i
    #         if idx < cols:
    #             table.columns[idx].width = pixels_to_emu(max(8, per_remaining))


def set_appropriate_row_height(table_info, table, rows):
    """
    Set row heights on the pptx table using CSS-specified heights or rendered heights.

    Parameters:
        table_info (dict): Table metadata including rows with styles/rects.
        table: pptx.table.Table object whose rows will be sized.
        rows (int): Number of rows in the pptx table.

    Returns:
        list: List of computed row heights in pixels.

    Notes:
        - Prefers explicit CSS heights when provided; otherwise uses extracted rendered height.
        - Clamps heights to sensible min/max values to keep layout stable.
    """
    # --- row heights with consistent sizing ---
    row_heights = []
    for row_data in table_info['rows']:
        r_index = row_data['index']
        if r_index >= rows:
            continue
        
        # Get CSS height from row data
        row_styles = row_data.get('styles', {})
        css_height = row_styles.get('height', '')
        
        specified_height = safe_float(css_height, None)
        # Get rendered height from rect
        rendered_height = row_data.get('rect', {}).get('height', 25)
        
        # iterate all cells and find the maximum height of cell with lowest rowspan
        row_cells = []
        for cell_data in row_data.get('cells', []):
            height = cell_data.get('rect', {}).get('height', 20) 
            row_span = cell_data.get('rowSpan') or 1
            row_cells.append((row_span, height))

        # sort the cell with rowspan asc, height descending
        if row_cells:
            row_cells.sort(key=lambda x: (x[0], -x[1]))
            max_content_height = row_cells[0][1]
        else:
            max_content_height = 0
            # html_content = cell_data.get('htmlContent', [])
            # for content_elem in html_content:
            #     if content_elem.get('type') == 'br':
            #         continue
            #     elem_rect = content_elem.get('rect', {})
            #     elem_height = elem_rect.get('height', 0)
            #     # Add relative y position to get total height needed
            #     total_height = elem_rect.get('y', 0) + elem_height
            #     if total_height > max_content_height:
            #         max_content_height = total_height
            
        # Use CSS specified height as primary for consistency with HTML if not found use rendered height
        final_height = specified_height if specified_height else rendered_height
        # If there are nested elements taller than the current height, adjust
        if max_content_height > 0:
            final_height = max(final_height, max_content_height)

        table.rows[r_index].height = pixels_to_emu(final_height) 
        row_heights.append(final_height)
    return row_heights


def apply_default_borders(rows, cols, table, reference_id=0, default_border_color=CELL_BORDER_COLOR, default_border_width=CELL_BORDER_WIDTH): 
    """
    Add a default border to every cell in the table.

    Parameters:
        rows (int): Number of rows.
        cols (int): Number of columns.
        table: pptx.table.Table instance.
        reference_id: Logging reference id.
        default_border_color (str): Hex color string for borders.
        default_border_width (str): Width value in EMU string for XML.

    Notes:
        - Uses _set_cell_border which manipulates XML to ensure consistent borders.
        - Errors in applying a border to a given cell are logged at debug level.
    """
    # Apply default borders to ALL table cells
    for r in range(rows):
        for c in range(cols):
            try:
                pptx_cell = table.cell(r, c)
                _set_cell_border(pptx_cell, border_color=default_border_color, border_width=default_border_width)
            except Exception as e:
                logger.error(f"Error applying border to cell ({r}, {c}): {e}", reference_id=reference_id)


def _extract_dimension_value(value, fallback):
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        parsed = safe_float(value, fallback)
        if parsed:
            return parsed
    return fallback


def prepare_cell_element_for_rendering(element, cell_x, cell_y, cell_width, cell_height, slide_width, slide_height):
    if not element or element.get('type') == 'br':
        return None

    rect = element.get('rect', {})
    width = _extract_dimension_value(element.get('width'), _extract_dimension_value(rect.get('width'), cell_width))
    height = _extract_dimension_value(element.get('height'), _extract_dimension_value(rect.get('height'), cell_height))
    width = max(1, min(width if width else cell_width, cell_width))
    height = max(1, min(height if height else cell_height, cell_height))
    element['width'] = width
    element['height'] = height

    abs_x = safe_float(element.get('x'), None)
    if abs_x is None:
        abs_x = cell_x + safe_float(rect.get('x', 0.0))
    abs_y = safe_float(element.get('y'), None)
    if abs_y is None:
        abs_y = cell_y + safe_float(rect.get('y', 0.0))

    max_left = cell_x + max(0, cell_width - width)
    max_top = cell_y + max(0, cell_height - height)
    abs_x = max(cell_x, min(abs_x, max_left))
    abs_y = max(cell_y, min(abs_y, max_top))
    abs_x = max(0, min(abs_x, slide_width - width))
    abs_y = max(0, min(abs_y, slide_height - height))
    element['x'] = abs_x
    element['y'] = abs_y

    inline_group = element.get('inlineGroup')
    if inline_group:
        group_rect = inline_group.setdefault('groupRect', {})
        group_rect.update({'x': abs_x, 'y': abs_y, 'width': width, 'height': height})

    shape_info = element.get('shapeInfo')
    if shape_info is not None:
        shape_rect = shape_info.get('rect')
        if shape_rect:
            shape_rect['x'] = safe_float(shape_rect.get('x'), abs_x)
            shape_rect['y'] = safe_float(shape_rect.get('y'), abs_y)
            shape_rect['width'] = safe_float(shape_rect.get('width'), width)
            shape_rect['height'] = safe_float(shape_rect.get('height'), height)
        else:
            shape_info['rect'] = {'x': abs_x, 'y': abs_y, 'width': width, 'height': height}

    chart_info = element.get('chartInfo')
    if chart_info is not None:
        chart_info.setdefault('width', width)
        chart_info.setdefault('height', height)

    media_info = element.get('mediaInfo')
    if media_info is not None:
        media_info.setdefault('naturalWidth', width)
        media_info.setdefault('naturalHeight', height)

    for pseudo_key in ('pseudoBefore', 'pseudoAfter'):
        if element.get(pseudo_key):
            element[pseudo_key] = prepare_cell_element_for_rendering(
                element[pseudo_key], cell_x, cell_y, cell_width, cell_height, slide_width, slide_height
            )

    prepared_children = []
    for child in element.get('children', []):
        prepared_child = prepare_cell_element_for_rendering(
            child, cell_x, cell_y, cell_width, cell_height, slide_width, slide_height
        )
        if prepared_child:
            prepared_children.append(prepared_child)
    element['children'] = prepared_children

    return element

def clone_cell_content_elements(children):
    return [copy.deepcopy(child) for child in (children or []) if child]

def flatten_cell_content(children):
    """
    Flatten nested cell content to extract text elements (p, ul, ol, span, etc.)
    from wrapper containers like div.
    
    This handles structures like: td > div > [p, p, ul] by extracting the actual
    content elements from intermediate wrappers.
    """
    if not children:
        return []
    
    flattened = []
    for child in children:
        child_type = child.get('type', '')
        
        # These are actual content elements we want to keep
        if child_type in ['p', 'ul', 'ol', 'span', 'strong', 'em', 'b', 'i', 'u', 'br', 'a', 'img', 'canvas']:
            flattened.append(copy.deepcopy(child))
            # For ul/ol, also extract listInfo
            if child_type in ['ul', 'ol']:
                child_copy = flattened[-1]
                if not child_copy.get('listInfo'):
                    child_copy['listInfo'] = extract_list_info_from_hierarchy(child)
        # div is a wrapper - look inside for content
        elif child_type == 'div':
            nested_children = child.get('children', [])
            if nested_children:
                # Recursively flatten nested content
                flattened.extend(flatten_cell_content(nested_children))
            elif child.get('text'):
                # div with direct text (treat as paragraph-like)
                flattened.append(copy.deepcopy(child))
        # Other elements with children might contain content
        elif child.get('children'):
            flattened.extend(flatten_cell_content(child.get('children', [])))
        elif child.get('text'):
            # Element with direct text
            flattened.append(copy.deepcopy(child))
    
    return flattened

def process_table_cell_content(slide, html_content, cell_x, cell_y, cell_width, cell_height, slide_width, slide_height, reference_id=0, img_dict={}, is_table_cell=False):
    """Process HTML content within a table cell and render elements.
    
    Args:
        slide: pptx Slide object.
        html_content: List of content elements to render.
        cell_x: Absolute X position of the cell.
        cell_y: Absolute Y position of the cell (based on cumulative row heights).
        cell_width: Width of the cell in pixels.
        cell_height: Height of the cell in pixels.
        slide_width: Slide width in pixels.
        slide_height: Slide height in pixels.
        reference_id: Logging identifier.
        img_dict: Cache of prefetched images.
        is_table_cell: Whether this is table cell content.
    """
    if not html_content:
        return False

    has_renderable_content = False
    for content_elem in html_content:
        elem_type = content_elem.get('type')
        if elem_type == 'br':
            continue

        # Skip links in table cells - handled as cell text with hyperlinks
        if elem_type == 'a' and is_table_cell:
            continue

        # Get element dimensions
        media_info = content_elem.get('mediaInfo', {})
        elem_width = content_elem.get('width') or content_elem.get('rect', {}).get('width') or media_info.get('currentWidth', 10)
        elem_height = content_elem.get('height') or content_elem.get('rect', {}).get('height') or media_info.get('currentHeight', 10)
        elem_width = max(1, elem_width)
        elem_height = max(1, elem_height)
        
        # Calculate element position
        elem_x = content_elem.get('x')
        elem_y = content_elem.get('y')
        
        if elem_x is not None and elem_y is not None:
            abs_x = elem_x
            abs_y = elem_y
        else:
            # Use relative positioning within cell
            elem_rect = content_elem.get('rect', {})
            rel_x = elem_rect.get('x', 0)
            rel_y = elem_rect.get('y', 0)
            abs_x = cell_x + rel_x
            abs_y = cell_y + rel_y
        
        # Ensure within slide bounds
        abs_x = max(0, min(abs_x, slide_width - elem_width))
        abs_y = max(0, min(abs_y, slide_height - elem_height))
        content_elem['x'] = abs_x
        content_elem['y'] = abs_y
        content_elem['width'] = elem_width
        content_elem['height'] = elem_height
        
        if elem_type == 'img':
            add_image_element(slide, content_elem, slide_width, slide_height, img_dict=img_dict, reference_id=reference_id)
            has_renderable_content = True
        elif elem_type == 'canvas' and content_elem.get('chartInfo'):
            add_chart_element(slide, content_elem, slide_width, slide_height, reference_id=reference_id)
            has_renderable_content = True
        elif content_elem.get('shapeInfo'):
            add_shape_element(slide, content_elem)
            has_renderable_content = True
    return has_renderable_content


def apply_cell_css_borders(cell, cell_styles, reference_id=0):
    """Apply CSS-specified borders to a table cell with uniform border optimization"""
    # Get border properties for all sides
    sides = ['Top', 'Right', 'Bottom', 'Left']
    border_props = {}

    for side in sides:
        width_key = f'border{side}Width'
        style_key = f'border{side}Style'
        color_key = f'border{side}Color'

        width_str = cell_styles.get(width_key, '0px')
        style_val = cell_styles.get(style_key, 'none')
        border_color = cell_styles.get(color_key)
        border_color = parse_color(border_color)
        # Parse width (convert from px to numeric)
        width_px = safe_float(width_str.replace('px', '')) if width_str else 0

        border_props[side.lower()] = {
            'width': width_px,
            'style': style_val,
            'color': border_color
        }

    # Apply borders individually for non-uniform cases
    side_names = ['top', 'right', 'bottom', 'left']

    for side_name in side_names:
        props = border_props[side_name]
        width_px = props['width']
        style_val = props['style']
        border_color = props['color']

        # Apply border if width > 0 and style is not 'none'
        if not border_color or width_px < 0 or style_val in ['none', 'hidden']:
            continue

        set_cell_border_enhanced(cell, side_name, width_px, border_color, style_val, CELL_BORDER_COLOR)


def process_table_cell_text(text_frame, html_content):
    """
    Process HTML content within a table cell and render text elements.
    
    Parameters:
        text_frame: pptx TextFrame object to populate.
        html_content: List of HTML content elements from the cell.
    """
    idx = 0
    for _, element in enumerate(html_content):
        text = ''
        href = ''
        italic = False
        styles = element.get('styles', {})
        
        if element.get('type') in ['ol', 'ul']:
            list_info = element.get('listInfo', {})
            if list_info:
                add_list_paragraphs(
                    text_frame=text_frame, list_info=list_info, level=0, element=element
                )
            idx += 1
            continue
       
        if element.get('type') == 'a' and element.get('linkInfo', {}).get('href'):
            text = element['linkInfo'].get('text') or "(link)"
            href = element.get('linkInfo', {}).get('href')
                    
        if element.get('type') in ['p', 'div'] and element.get('text'):
            text = element['text']

        if element.get('type') == 'em' and element.get('text'):
            text = element['text']
            italic = True

        if not text:
            continue
        
        p = text_frame.paragraphs[-1]
        if idx:
            p = text_frame.add_paragraph()
        
        run = p.add_run()
        run.text = text
        font = run.font
        font_size_px = safe_float(styles.get('fontSize', '12').replace('px', ''))
        font.name = styles.get('fontFamily', 'Arial').split(',')[0].strip('"\'')
        font.size = Pt(max(6, get_font_size_pt(font_size_px)))
        font.bold = is_bold(styles.get('fontWeight'))
        font.italic = str(styles.get('fontStyle')) == 'italic' or italic
        link_color = parse_color(styles.get('color'))
        if link_color:
            font.color.rgb = link_color

        text_decoration = str(styles.get('textDecoration', '')) or 'underline'
        font.underline = 'underline' in text_decoration
        if href:
            run.hyperlink.address = href
        idx += 1


def handle_cell_inline_group(cell_data, text_frame, text_align, cell_styles, existing_para):
    """
    Render inlineGroup markup inside a table cell's text_frame.

    Iterates inlineElements, creates runs and applies per-run styling (size, family, bold/italic, color).
    Trims trailing spaces from the last run.

    Returns:
        str: Concatenated plain text of the processed inline elements (without trailing whitespace).
    """
    first = True
    cell_text = ""
    p = existing_para
    for inline_element in cell_data['inlineGroup']['inlineElements']:
        if inline_element.get('type') == 'br':
            p = text_frame.add_paragraph()
            p.alignment = get_para_alignment(text_align)
            first = True
            continue
        t = inline_element.get('text', '')
        if first:
            t = t.lstrip()
        if not t.strip():
            continue
        first = False
        run = p.add_run()
        run.text = t
        cell_text += t
        inline_styles = inline_element.get('styles', {})
        font = run.font
        fs_px = safe_float(inline_styles.get('fontSize', '16'))
        font.size = Pt(max(6, get_font_size_pt(fs_px)))
        font.name = inline_styles.get('fontFamily', 'Arial').split(',')[0].strip('"\'')
        font.bold = inline_styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
        font.italic = inline_styles.get('fontStyle', 'normal') == 'italic'
        clr = parse_color(inline_styles.get('color') or cell_styles.get('color'))
        if clr:
            font.color.rgb = clr
    # Trim trailing spaces
    if text_frame.paragraphs and text_frame.paragraphs[-1].runs:
        last_run = text_frame.paragraphs[-1].runs[-1]
        last_run.text = last_run.text.rstrip()
        
    return cell_text


def get_cell(table, cell_data, row_index, cell_index, rows, cols, reference_id=0):
    """
    Retrieve a pptx cell object for the given cell_data and apply merges if needed.

    Handles colSpan and rowSpan by merging the base cell with the endpoint cell.
    Reapplies borders on merged cells. Exceptions during merge are logged.

    Parameters:
        table: pptx.table.Table instance.
        cell_data (dict): Metadata for the cell including cellIndex, colSpan, rowSpan.
        row_index (int): Row index within table.
        cell_index (int): Column index within table.
        rows (int): Total rows.
        cols (int): Total columns.
        reference_id: Logging identifier.

    Returns:
        pptx.table._Cell: The (possibly merged) cell object for writing content.
    """
    # Merge spans
    pptx_cell = table.cell(row_index, cell_index)
    cell_styles = cell_data.get('styles', {})
    col_span = max(1, cell_data.get('colSpan', 1))
    row_span = max(1, cell_data.get('rowSpan', 1))
    apply_cell_css_borders(pptx_cell, cell_styles, reference_id)
    if col_span > 1 or row_span > 1:
        try:
            end_row = min(rows - 1, row_index + row_span - 1)
            end_col = min(cols - 1, cell_index + col_span - 1)
            merged_cell = pptx_cell.merge(table.cell(end_row, end_col))
            if merged_cell is not None:
                pptx_cell = merged_cell
        except Exception as e:
            logger.error(f"Error merging cell ({row_index}, {cell_index}): {e}", reference_id=reference_id)
    
    return pptx_cell


def fill_table(slide, slide_width, slide_height, table, table_info, x, y, rows, cols, row_heights, reference_id=0, img_dict={}):
    """
    Populate table cells with background, padding, and text content.

    - Applies row-level background if present.
    - Ensures text_frame margins and vertical alignment per cell styles.
    - Renders inlineGroup content or plain text, applies font sizing and color,
      and uses a non-breaking space for empty cells to maintain row heights.

    Parameters:
        slide: pptx Slide object.
        slide_width (int): Slide width in pixels.
        slide_height (int): Slide height in pixels.
        table: pptx.table.Table object.
        table_info (dict): Structure describing rows, cells and styles.
        x (float): Table x position.
        y (float): Table y position.
        rows (int): Number of rows in table.
        cols (int): Number of columns in table.
        row_heights (list): List of row heights.
        reference_id: Logging id for context.
        img_dict (dict): Cache of prefetched images.
    """
    rect = table_info.get('rect', {})
    table_top_y = rect.get('y', 0)
    
    # Track occupied cells for rowspan/colspan handling
    occupied_cells = set()

    # Track height adjustments for rows without images
    adjusted_heights = 0
    
    # Pre-calculate cumulative row Y offsets for consistent vertical positioning
    cumulative_row_y = [0.0]
    for rh in row_heights:
        cumulative_row_y.append(cumulative_row_y[-1] + rh)
    
    for idx, row_data in enumerate(table_info['rows']):
        row_index = row_data['index']
        if row_index >= rows:
            continue
        table_cell_html_content = []
        current_col = 0
        for cell_data in row_data.get('cells', []):
            # Skip occupied columns in this row
            while current_col < cols and (row_index, current_col) in occupied_cells:
                current_col += 1

            if current_col >= cols:
                break

            col_span = max(1, cell_data.get('colSpan', 1))
            row_span = max(1, cell_data.get('rowSpan', 1))

            # Mark all cells occupied by this cell's rowspan/colspan
            for r in range(row_index, min(rows, row_index + row_span)):
                for c in range(current_col, min(cols, current_col + col_span)):
                    occupied_cells.add((r, c))

            pptx_cell = get_cell(table, cell_data, row_index, current_col, rows, cols, reference_id)
            if pptx_cell is None:
                continue
            
            text_frame = pptx_cell.text_frame
            text_frame.word_wrap = True
            cell_styles = cell_data.get('styles', {})
            
            # Calculate absolute cell position for nested elements
            cell_rect = cell_data.get('rect', {})
            abs_cell_x = x + (cell_rect.get('x', 0) - rect.get('x', 0))
            abs_cell_y = y + cumulative_row_y[row_index] if row_index < len(cumulative_row_y) else y + (cell_rect.get('y', 0) - table_top_y)
            cell_width = cell_rect.get('width', 50)
            cell_height = row_heights[row_index] if row_index < len(row_heights) else cell_rect.get('height', 20)
            
            html_content = cell_data.get('htmlContent') or []

            # Collect non-text content (images, charts) for processing after table creation
            chart_image_content = [elem for elem in html_content if elem.get('type') not in TEXT_HTML_TAGS]
            if chart_image_content:
                table_cell_html_content.append({
                    'html_content': chart_image_content,
                    'cell_x': abs_cell_x,
                    'cell_y': abs_cell_y,
                    'cell_width': cell_width,
                    'cell_height': cell_height
                })

            process_table_cell_text(text_frame, html_content)

            vertical_align = (cell_styles.get('verticalAlign') or '').strip().lower()
            pptx_cell.vertical_anchor = get_vertical_alignment(vertical_align)
        
            # Enhanced padding with consistent values
            padding_left = safe_float(cell_styles.get('paddingLeft'))
            padding_right = safe_float(cell_styles.get('paddingRight'))
            padding_top = safe_float(cell_styles.get('paddingTop', '4px'))
            padding_bottom = safe_float(cell_styles.get('paddingBottom', '4px'))
            
            if padding_left is not None:
                pptx_cell.margin_left = pixels_to_emu(padding_left)
            
            if padding_right is not None:
                pptx_cell.margin_right = pixels_to_emu(padding_right)
            
            if padding_top is not None:
                pptx_cell.margin_top = pixels_to_emu(padding_top)
            
            if padding_bottom is not None:
                pptx_cell.margin_bottom = pixels_to_emu(padding_bottom)
            
            p = text_frame.paragraphs[-1]
            text_align = cell_styles.get('textAlign')
            p.alignment = get_para_alignment(text_align)
            # Background
            row_bg_color = parse_color(row_data.get('styles', {}).get('backgroundColor'))
            bg_color_cell = parse_color(cell_styles.get('backgroundColor'))
            background_color = bg_color_cell or row_bg_color
            if background_color:
                pptx_cell.fill.solid()
                pptx_cell.fill.fore_color.rgb = background_color
            else:
                pptx_cell.fill.background()

            # Handle regular cell text (no link)
            cell_text = (cell_data.get('text') or '').strip()
            has_only_br = html_content and all(
                elem.get('type') == 'br' for elem in html_content
            )
            
            if (not html_content or has_only_br) and cell_text:        
                
                # Get cell text content
                if cell_data.get('inlineGroup') and cell_data['inlineGroup'].get('inlineElements'):
                    # Handle inline formatted content
                    handle_cell_inline_group(cell_data, text_frame, text_align, cell_styles, p)
                else:
                    run = p.add_run()
                    run.text = cell_text
                    font = run.font
                    fs_px = safe_float(cell_styles.get('fontSize', '10'))
                    font.size = Pt(max(6, get_font_size_pt(fs_px)))
                    font.name = cell_styles.get('fontFamily', 'Arial').split(',')[0].strip('"\'')
                    font.bold = cell_styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
                    font.italic = cell_styles.get('fontStyle', 'normal') == 'italic'
                    clr = parse_color(cell_styles.get('color'))
                    if clr:
                        font.color.rgb = clr

                p.space_before = Pt(0)
                p.space_after = Pt(0)
                p.line_spacing = 1.0

            current_col += col_span
        
        if len(table_cell_html_content) == 0:
            table.rows[idx].height = math.floor(table.rows[idx].height * TABLE_CELL_HEIGHT_REDUCTION_FACTOR)     
            
            # since we are reducing the row heights it will impact the below rows which needs to adjusted upwards
            # if we collect all height that has been lost and substract it from x-coordinate of below rows  
            if row_heights and idx < len(row_heights):
                adjusted_heights += row_heights[idx] - (row_heights[idx] * TABLE_CELL_HEIGHT_REDUCTION_FACTOR)
        
        for content_data in table_cell_html_content:
            process_table_cell_content(
                slide,
                content_data['html_content'],
                content_data['cell_x'],
                content_data['cell_y'] - adjusted_heights,
                content_data['cell_width'],
                content_data['cell_height'],
                slide_width,
                slide_height,
                reference_id=reference_id,
                img_dict=img_dict,
                is_table_cell=True
            )


def add_table_content(slide, content_data, content_type, cell_x, cell_y, cell_width, cell_height, slide_width, slide_height, reference_id=0, img_dict={}):
    """Add content (images, links, or charts) within a table cell with proper positioning"""

    if content_type == 'table_image':
        img_src = content_data.get('src', '')
        if not img_src:
            return

        # Get dimensions from top-level element first, then rect, then mediaInfo
        media_info = content_data.get('mediaInfo', {})
        img_width = content_data.get('width') or content_data.get('rect', {}).get('width') or media_info.get('currentWidth', 20)
        img_height = content_data.get('height') or content_data.get('rect', {}).get('height') or media_info.get('currentHeight', 20)
        img_width = max(1, img_width)
        img_height = max(1, img_height)

        # Get position - use element x/y directly if available (absolute coords from Extract.js)
        elem_x = content_data.get('x')
        elem_y = content_data.get('y')
        
        if elem_x is not None and elem_y is not None:
            # Use absolute positioning from Extract.js extraction
            abs_x = elem_x
            abs_y = elem_y
        else:
            # Fall back to relative positioning within cell
            img_rect = content_data.get('rect', {})
            rel_x = img_rect.get('x', 0)
            # Center image vertically in cell
            vertical_center_offset = (cell_height - img_height) / 2
            abs_x = cell_x + rel_x
            abs_y = cell_y + vertical_center_offset
            # Ensure image stays within cell bounds
            abs_x = max(cell_x, min(abs_x, cell_x + cell_width - img_width))
            abs_y = max(cell_y, min(abs_y, cell_y + cell_height - img_height))

        # Ensure image stays within slide bounds
        abs_x = max(0, min(abs_x, slide_width - img_width))
        abs_y = max(0, min(abs_y, slide_height - img_height))

        styles = content_data.get('styles', {})
        natural_width = content_data.get('naturalWidth', img_width)

        border_radius_str = styles.get('borderRadius', '0px')
        radius_ratio = parse_border_radius(border_radius_str, img_width, img_height)
        radius_display = radius_ratio * min(img_width, img_height)
        has_radius = radius_display > 0

        temp_path = img_dict.get(img_src) if img_dict else None
        if not temp_path:
            if img_src.startswith('data:'):
                _, data = img_src.split(',', 1)
                img_data = base64.b64decode(data)
                temp_path = BytesIO(img_data)

            elif img_src.startswith('http'):
                try:
                    response = requests.get(img_src, timeout=10)
                    if response.status_code == 200:
                        temp_path = BytesIO(response.content)
                except Exception as e:
                    logger.error(f"Failed to download table image: {img_src} - {e}", reference_id=reference_id)
                    return

            elif os.path.exists(img_src):
                with open(img_src, 'rb') as f:
                    temp_path = BytesIO(f.read())

        if not temp_path:
            logger.error(f"Image not found: {img_src}", reference_id=reference_id)
            return

        image_to_add = temp_path
        if has_radius:
            scale_x = natural_width / img_width if img_width > 0 else 1
            radius_natural = int(radius_display * scale_x)
            temp_rounded = make_rounded_image(temp_path, radius_natural)
            image_to_add = temp_rounded

        # Add image with calculated position
        picture = slide.shapes.add_picture(
            image_to_add,
            pixels_to_emu(abs_x), pixels_to_emu(abs_y),
            pixels_to_emu(img_width), pixels_to_emu(img_height)
        )
        picture.shadow.inherit = False

        # Handle borders if present
        has_border = is_uniform_border(styles)
        if not has_border:
            return

        border_width = safe_float(styles.get('borderTopWidth', '0px'))
        border_color = parse_color(styles.get('borderTopColor'))
        if not border_color or border_width < 0:
            return

        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if has_radius else MSO_SHAPE.RECTANGLE
        border_shape = slide.shapes.add_shape(
            shape_type,
            pixels_to_emu(abs_x), pixels_to_emu(abs_y),
            pixels_to_emu(img_width), pixels_to_emu(img_height)
        )
        if has_radius:
            border_shape.adjustments[0] = radius_ratio
        border_shape.fill.background()
        border_shape.line.width = Pt(border_width)
        border_shape.line.color.rgb = border_color
        border_shape.shadow.inherit = False

        # Move border shape behind picture
        sp = border_shape._sp
        parent = sp.getparent()
        parent.remove(sp)
        pic_sp = picture._sp
        idx = list(parent).index(pic_sp)
        parent.insert(idx, sp)


    elif content_type == 'table_link':
        # For links, we'll return the link data to be processed within the cell text
        # This function will now just return the link data for cell processing
        return {
            'href': content_data.get('href', ''),
            'text': content_data.get('text', '').strip(),
            'styles': content_data.get('styles', {})
        }

    elif content_type == 'table_chart':
        chart_data = content_data.get('chartData')
        if not chart_data:
            return

        # Calculate absolute position within the slide
        chart_rect = content_data.get('rect', {})
        rel_x = chart_rect.get('x', 0)
        chart_width = max(10, chart_rect.get('width', 200))
        chart_height = max(10, chart_rect.get('height', 150))

        # Center chart vertically in cell
        vertical_center_offset = (cell_height - chart_height) / 2

        # Position relative to cell with vertical centering
        abs_x = cell_x + rel_x
        abs_y = cell_y + vertical_center_offset

        # Ensure chart stays within cell bounds
        abs_x = max(cell_x, min(abs_x, cell_x + cell_width - chart_width))
        abs_y = max(cell_y, min(abs_y, cell_y + cell_height - chart_height))

        # Ensure chart stays within slide bounds
        abs_x = max(0, min(abs_x, slide_width - chart_width))
        abs_y = max(0, min(abs_y, slide_height - chart_height))

        # Create a temporary element structure for the chart
        chart_element = {
            'chartInfo': {
                'chartData': chart_data,
                'width': chart_width,
                'height': chart_height
            },
            'x': abs_x,
            'y': abs_y,
            'width': chart_width,
            'height': chart_height
        }

        # Use the existing add_chart_element function
        add_chart_element(slide, chart_element, slide_width, slide_height, reference_id)


def add_table_element(slide, element, slide_width, slide_height, reference_id=0, parent_has_shadow=False, img_dict={}):
    """
    Create and add a pptx table shape to the slide using tableInfo metadata.

    Steps:
      - Compute rect and clamp to slide bounds.
      - Optionally render a background/border container using add_bg_shape.
      - Create the pptx table, set style/widths/heights, apply default borders,
        populate cells and return the created table shape.

    Parameters:
        slide: pptx.Slide to which table will be added.
        element (dict): Element metadata containing tableInfo.
        slide_width (int): Slide w# th in pixels.
        slide_height (int): Slide height in pixels.
        reference_id: Logging reference id.
        parent_has_shadow (bool): If parent already has shadow, child shadow is skipped.
        img_dict (dict): Optional cache mapping image src to BytesIO or path.

    Returns:
        pptx shape: The pptx shape that contains the table (or None if no table).
    """
    table_info = element.get('tableInfo', {})
    if not table_info.get('rows'):
        return None
        
    rect = table_info.get('rect', {})
    x = safe_float(rect.get('x', element.get('x', 0)))
    y = safe_float(rect.get('y', element.get('y', 0)))
    width = safe_float(rect.get('width', element.get('width', 100)))
    height = safe_float(rect.get('height', element.get('height', 100)))
    
    styles = table_info.get('styles', {})
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none' and not parent_has_shadow
    bg_color = parse_color(styles.get('backgroundColor'))
    border_radius_str = styles.get('borderRadius', '0px')
    border_radius = parse_border_radius(border_radius_str, width, height)
    has_radius = border_radius > 0
    has_border = is_uniform_border(styles)
   
    # Clamp to slide bounds (keep existing safeguard)
    if x + width > slide_width:
        width = max(10, slide_width - x)
    if y + height > slide_height:
        height = max(10, slide_height - y)

    if bg_color or has_border or has_radius or has_shadow:
        add_bg_shape(slide, styles, x, y, width, height, reference_id=reference_id)

    rows = table_info['rowCount']
    cols = table_info['columnCount']

    if not (rows and cols):
        logger.error(f"Invalid rows: {rows}, cols: {cols}", reference_id=reference_id)
        return

    if not table_info.get('rows'):
        return

    table_shape = slide.shapes.add_table(
        rows, cols,
        pixels_to_emu(x), pixels_to_emu(y),
        pixels_to_emu(width), pixels_to_emu(height)
    )
    table = table_shape.table

    # Set table style to ensure borders render and disable default formatting
    tbl = table._graphic_frame._graphicFrame.graphicData.tbl
    style_id = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'
    tbl[0][-1].text = style_id
    table.first_row = False  # Disable header row default styling


    # first_row = next((r for r in table_info['rows'] if r.get('cells')), None)
    # if not first_row:
    #     return
    
    set_appropriate_column_width(table, table_info, cols, width)
    row_heights = set_appropriate_row_height(table_info, table, rows)
    # apply_default_borders(rows, cols, table, reference_id)
    fill_table(slide, slide_width, slide_height, table, table_info, x, y, rows, cols, row_heights, reference_id, img_dict=img_dict)
    
    return table_shape
