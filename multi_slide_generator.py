import re
import os
import base64
import json
import math
import requests
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE, MSO_COLOR_TYPE
import requests
from PIL import Image, ImageDraw
import os
import re
import math
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
import json
import re
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LEGEND_POSITION
from PIL import Image
import os

# Font scaling factor to match HTML rendering more closely
FONT_SCALE_FACTOR = 0.88  # Reduce font sizes by 12% to better match HTML

def SubElement(parent, tagname, **kwargs):
    element = OxmlElement(tagname)
    element.attrib.update(kwargs)
    parent.append(element)
    return element


def makeParaBulletPointed(para, bullet_char="•", bullet_font="Arial", bullet_size_pt=None):
    """Apply bullet formatting to a paragraph.
    Works for both slide text boxes and table cells.
    
    Args:
        para: Paragraph object to apply bullets to
        bullet_char: Character to use as bullet (default: •)
        bullet_font: Font for bullet character (default: Arial)
        bullet_size_pt: Size of bullet in points (optional, uses text size if None)
    """
    try:
        pPr = para._p.get_or_add_pPr()
        
        # Set margins and indentation (in EMUs: 914400 EMUs = 1 inch)
        # marL: left margin, indent: first line indent (negative for hanging indent)
        pPr.set('marL', '228600')  # ~0.25 inch left margin
        pPr.set('indent', '-228600')  # Hanging indent
        
        # Remove any existing bullet formatting
        for child in list(pPr):
            if 'bu' in child.tag.lower():
                pPr.remove(child)
        
        # Add bullet font with proper namespace
        buFont = OxmlElement('a:buFont')
        buFont.set('typeface', bullet_font)
        buFont.set('charset', '0')
        pPr.append(buFont)
        
        # Add bullet character
        buChar = OxmlElement('a:buChar')
        buChar.set('char', bullet_char)
        pPr.append(buChar)
        
        # Optionally set bullet size relative to text
        if bullet_size_pt:
            buSzPts = OxmlElement('a:buSzPts')
            buSzPts.set('val', str(int(bullet_size_pt * 100)))  # Size in 1/100th of a point
            pPr.append(buSzPts)
        
        return True
    except Exception as e:
        print(f"Error applying bullet formatting: {e}")
        return False


def _set_cell_border(cell, border_color="808080", border_width='12700'):
    """Apply uniform border to all 4 sides of a cell - more reliable for identical borders"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    for lines in ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB']:
        # Remove existing border elements with same tag
        tag = lines.split(":")[-1]
        for e in tcPr.getchildren():
            if tag in str(e.tag):
                tcPr.remove(e)
        
        # Create new border element
        ln = SubElement(tcPr, lines, w=border_width, cap='flat', cmpd='sng', algn='ctr')
        solidFill = SubElement(ln, 'a:solidFill')
        srgbClr = SubElement(solidFill, 'a:srgbClr', val=border_color)
        prstDash = SubElement(ln, 'a:prstDash', val='solid')
        round_ = SubElement(ln, 'a:round')
        headEnd = SubElement(ln, 'a:headEnd', type='none', w='med', len='med')
        tailEnd = SubElement(ln, 'a:tailEnd', type='none', w='med', len='med')
    
    return cell


def set_cell_border_enhanced(cell, side, width_px, color_rgb, style='solid'):
    """Apply border to individual side of a cell - used for non-uniform borders"""
    if width_px <= 0 or not color_rgb:
        return

    # Convert color to hex string
    try:
        color_hex = f'{color_rgb.r:02X}{color_rgb.g:02X}{color_rgb.b:02X}'
    except:
        color_hex = "808080"  # Default gray
    
    # Convert width to EMU (1pt = 12700 EMU, 1px ≈ 0.75pt)
    width_emu = str(int(width_px * 0.75 * 12700))
    
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

def safe_int(value, default=0):
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return default


def safe_float(value, default=0.0):
    """Safely convert value to float with better error handling"""
    try:
        if isinstance(value, str):
            value = value.replace('px', '').replace('%', '').replace('em', '').replace('rem', '')
        return float(value) if value else default
    except (ValueError, TypeError):
        return default
    

def pixels_to_emu(pixels):
    return math.floor(pixels * 9525)


def px_to_pt(px):
    """Convert pixels to points"""
    return px * 0.75


def get_font_size_pt(font_size_px):
    if font_size_px <= 0:
        return 12
    # Apply scaling factor to reduce font size slightly for better HTML matching
    scaled_size = max(6, int(px_to_pt(font_size_px) * FONT_SCALE_FACTOR))
    return scaled_size


def parse_border_radius(radius_str, shape_width_px, shape_height_px):
    """Parse border radius from CSS string"""
    if not radius_str or radius_str == '0px':
        return 0
    try:
        if isinstance(radius_str, str):
            radius_px = safe_float(radius_str.replace('px', '').replace('%', ''))
        else:
            radius_px = safe_float(radius_str)
        min_dimension = min(shape_width_px, shape_height_px)
        if min_dimension > 0:
            return min(0.5, radius_px / min_dimension)
        return 0
    except (ValueError, TypeError):
        return 0


def parse_color(color_str):
    """Enhanced color parsing with better RGB extraction"""
    if not color_str or color_str in ['transparent', 'rgba(0, 0, 0, 0)', 'none', 'initial', 'inherit']:
        return None
   
    # Handle linear-gradient
    if color_str.startswith('linear-gradient'):
        # For gradients, take the first color as approximation
        match = re.search(r'rgb[a]?\(\d+,\s*\d+,\s*\d+', color_str)
        if match:
            return parse_color(match.group())
   
    # Handle RGB/RGBA
    if color_str.startswith(('rgb', 'rgba')):
        parts = re.findall(r'[\d.]+', color_str)
        if len(parts) >= 3:
            return RGBColor(min(255, int(float(parts[0]))), min(255, int(float(parts[1]))), min(255, int(float(parts[2]))))
   
    # Handle hex colors
    elif color_str.startswith('#'):
        color_str = color_str.lstrip('#')
        if len(color_str) == 6:
            return RGBColor(int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16))
        elif len(color_str) == 3:
            return RGBColor(int(color_str[0]*2, 16), int(color_str[1]*2, 16), int(color_str[2]*2, 16))
   
    # Named colors
    named_colors = {
        'black': RGBColor(0, 0, 0), 'white': RGBColor(255, 255, 255),
        'red': RGBColor(255, 0, 0), 'green': RGBColor(0, 128, 0),
        'blue': RGBColor(0, 0, 255), 'yellow': RGBColor(255, 255, 0),
        'gray': RGBColor(128, 128, 128), 'grey': RGBColor(128, 128, 128),
        'silver': RGBColor(192, 192, 192), 'maroon': RGBColor(128, 0, 0),
        'olive': RGBColor(128, 128, 0), 'lime': RGBColor(0, 255, 0),
        'aqua': RGBColor(0, 255, 255), 'teal': RGBColor(0, 128, 128),
        'navy': RGBColor(0, 0, 128), 'fuchsia': RGBColor(255, 0, 255),
        'purple': RGBColor(128, 0, 128)
    }
   
    return named_colors.get(color_str.lower())


def is_uniform_border(styles):
    widths = [safe_float(styles.get(f'border{side}Width', '0px')) for side in ['Top', 'Right', 'Bottom', 'Left']]
    styles_list = [styles.get(f'border{side}Style', 'none') for side in ['Top', 'Right', 'Bottom', 'Left']]
    colors = [styles.get(f'border{side}Color', '') for side in ['Top', 'Right', 'Bottom', 'Left']]
    return len(set(widths)) == 1 and len(set(styles_list)) == 1 and len(set(colors)) == 1 and widths[0] > 0


def has_any_border(styles):
    """Enhanced border detection"""
    for side in ['Top', 'Right', 'Bottom', 'Left']:
        width_key = f'border{side}Width'
        style_key = f'border{side}Style'
        color_key = f'border{side}Color'
       
        width = safe_float(styles.get(width_key, '0px'))
        style = styles.get(style_key, 'none')
        color = parse_color(styles.get(color_key, ''))
       
        if width > 0 and style not in ['none', 'hidden'] and color is not None:
            return True
    return False


def get_border_info(styles):
    """Enhanced border information extraction"""
    border_info = {}
    for side in ['Top', 'Right', 'Bottom', 'Left']:
        width_key = f'border{side}Width'
        style_key = f'border{side}Style'
        color_key = f'border{side}Color'
       
        width = safe_float(styles.get(width_key, '0px'))
        style = styles.get(style_key, 'none')
        color = parse_color(styles.get(color_key, ''))
       
        border_info[side.lower()] = {
            'width': width,
            'style': style,
            'color': color,
            'has_border': width > 0 and style not in ['none', 'hidden'] and color is not None
        }
    return border_info


def make_rounded_image(input_path, output_path, radius):
    im = Image.open(input_path).convert("RGBA")
    mask = Image.new("L", im.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle((0, 0) + im.size, radius=radius, fill=255)
    im.putalpha(mask)
    im.save(output_path, "PNG")


def create_precise_border_shapes(slide, x, y, width, height, border_info, border_radius=0):
    """Create precise border shapes with accurate positioning"""
    shapes_created = []
    # Always check and draw all four borders if needed
    for side, info in border_info.items():
        if not info['has_border']:
            continue
        border_width = info['width']
        border_color = info['color']
        border_style = info['style']
        # Calculate precise line positions
        if side == 'top':
            line_x = x
            line_y = y
            line_width = width
            line_height = border_width
        elif side == 'right':
            line_x = x + width - border_width
            line_y = y
            line_width = border_width
            line_height = height
        elif side == 'bottom':
            line_x = x
            line_y = y + height - border_width
            line_width = width
            line_height = border_width
        elif side == 'left':
            line_x = x
            line_y = y
            line_width = border_width
            line_height = height
        try:
            # Always use rectangle for side borders, rounded only for full shape
            shape_type = MSO_SHAPE.RECTANGLE
            border_shape = slide.shapes.add_shape(
                shape_type,
                pixels_to_emu(line_x), pixels_to_emu(line_y),
                pixels_to_emu(line_width), pixels_to_emu(line_height)
            )
            border_shape.fill.solid()
            border_shape.fill.fore_color.rgb = border_color
            border_shape.line.fill.background()
            border_shape.shadow.inherit = False
            if border_style == 'dashed':
                border_shape.line.dash_style = MSO_LINE.DASH
            elif border_style == 'dotted':
                border_shape.line.dash_style = MSO_LINE.ROUND_DOT
            shapes_created.append(border_shape)
        except Exception as e:
            print(f"Error creating {side} border: {e}")
    return shapes_created


def add_bg_shape(slide, styles, x, y, width, height):
    """Enhanced background shape creation with precise positioning"""
    bg_color = parse_color(styles.get('backgroundColor'))
    border_radius_str = styles.get('borderRadius', '0px')
    border_radius = parse_border_radius(border_radius_str, width, height)
    has_radius = border_radius > 0
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none'
    has_uniform_border = is_uniform_border(styles)
    has_any_border_sides = has_any_border(styles)
    shapes_created = []
    # Create main background shape
    if bg_color or has_uniform_border or has_radius or has_shadow:
        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if has_radius else MSO_SHAPE.RECTANGLE
        try:
            bg_shape = slide.shapes.add_shape(
                shape_type,
                pixels_to_emu(x), pixels_to_emu(y),
                pixels_to_emu(width), pixels_to_emu(height)
            )
            if has_radius:
                bg_shape.adjustments[0] = border_radius
            if bg_color:
                bg_shape.fill.solid()
                bg_shape.fill.fore_color.rgb = bg_color
            else:
                bg_shape.fill.background()
            bg_shape.shadow.inherit = False
            # Handle uniform borders
            if has_uniform_border:
                border_width = safe_float(styles.get('borderTopWidth', '1px'))
                border_color = parse_color(styles.get('borderTopColor', 'black'))
                border_style = styles.get('borderTopStyle', 'solid')
                if border_color:
                    bg_shape.line.width = Pt(max(0.5, border_width))
                    bg_shape.line.color.rgb = border_color
                    if border_style == 'dashed':
                        bg_shape.line.dash_style = MSO_LINE.DASH
                    elif border_style == 'dotted':
                        bg_shape.line.dash_style = MSO_LINE.ROUND_DOT
                else:
                    bg_shape.line.fill.background()
            else:
                bg_shape.line.fill.background()
            if has_shadow:
                apply_shadow(bg_shape, box_shadow)
            shapes_created.append(bg_shape)
        except Exception as e:
            print(f"Error adding bg shape: {e}")
    # Always handle non-uniform borders for all sides
    if has_any_border_sides and not has_uniform_border:
        border_info = get_border_info(styles)
        border_shapes = create_precise_border_shapes(
            slide, x, y, width, height, border_info,
            border_radius * min(width, height) if has_radius else 0
        )
        shapes_created.extend(border_shapes)
    return shapes_created


def apply_shadow(shape, box_shadow_str):
    if box_shadow_str == 'none':
        return
    parts = box_shadow_str.split()
    if len(parts) < 3:
        return
    offset_x_px = safe_float(parts[0])
    offset_y_px = safe_float(parts[1])
    blur_px = safe_float(parts[2])
    spread_px = safe_float(parts[3]) if len(parts) > 3 else 0
    color_str = parts[4] if len(parts) > 4 else parts[3]
    if offset_x_px == 0 and offset_y_px == 0 and blur_px == 0 and spread_px == 0:
        return
    color = parse_color(color_str)
    alpha = 1.0
    if 'rgba' in color_str:
        color_parts = re.findall(r'\d+', color_str)
        if len(color_parts) == 4:
            alpha = float(color_parts[3]) / 255
    if not color:
        return
    distance_px = math.sqrt(offset_x_px**2 + offset_y_px**2)
    direction = math.degrees(math.atan2(offset_y_px, offset_x_px)) if distance_px > 0 else 0
    shape.shadow.inherit = False
    shape.shadow.blur = Pt(px_to_pt(blur_px))
    shape.shadow.distance = Pt(px_to_pt(distance_px))
    shape.shadow.angle = direction
    shape.shadow.color.type = MSO_COLOR_TYPE.RGB
    shape.shadow.color.rgb = color
    shape.shadow.transparency = 1 - alpha


def add_inline_group_element(slide, element, slide_width, slide_height, parent_has_shadow=False):
    inline_group = element.get('inlineGroup')
    if not inline_group:
        return
    inline_elements = inline_group.get('inlineElements', [])
    if not inline_elements:
        return
    
    has_content = any(elem.get('text', '').strip() for elem in inline_elements)
    if not has_content:
        return
    
    group_rect = inline_group.get('groupRect', {})
    x = max(0, min(group_rect.get('x', 0), slide_width - 10))
    y = max(0, min(group_rect.get('y', 0), slide_height - 10))
    width = max(10, min(group_rect.get('width', 100), slide_width - x))
    height = max(10, min(group_rect.get('height', 20), slide_height - y))
    styles = inline_group.get('styles', {})
    
    try:
        # Only add text content if there are inline elements
        if has_content:
            # Add regular background shapes
            if parse_color(styles.get('backgroundColor')) or has_any_border(styles):
                add_bg_shape(slide, styles, x, y, width, height)
            
            textbox = slide.shapes.add_textbox(
                pixels_to_emu(x), pixels_to_emu(y),
                pixels_to_emu(width), pixels_to_emu(height)
            )
            text_frame = textbox.text_frame
            text_frame.word_wrap = True
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            text_frame.margin_left = pixels_to_emu(safe_float(styles.get('paddingLeft', '0px').replace('px', '')))
            text_frame.margin_right = pixels_to_emu(safe_float(styles.get('paddingRight', '0px').replace('px', '')))
            text_frame.margin_top = pixels_to_emu(safe_float(styles.get('paddingTop', '0px').replace('px', '')))
            text_frame.margin_bottom = pixels_to_emu(safe_float(styles.get('paddingBottom', '0px').replace('px', '')))
            textbox.fill.background()
            textbox.line.fill.background()
            textbox.shadow.inherit = False
            text_frame.clear()
            p = text_frame.paragraphs[0]
            first = True
            text_align = styles.get('textAlign', 'left')
            alignment = PP_ALIGN.CENTER if text_align == 'center' else PP_ALIGN.RIGHT if text_align == 'right' else PP_ALIGN.LEFT
            p.alignment = alignment
            
            for inline_element in inline_elements:
                if inline_element.get('type') == 'br':
                    if p is not None:
                        p = text_frame.add_paragraph()
                        p.alignment = alignment
                    first = True
                    continue
                element_text = inline_element.get('text', '')
                if first:
                    element_text = element_text.lstrip()
                if not element_text.strip():
                    continue
                first = False
                run = p.add_run()
                run.text = element_text
                inline_styles = inline_element.get('styles', {})
                font = run.font
                font_size_px = safe_float(inline_styles.get('fontSize', '16').replace('px', ''))
                font.name = inline_styles.get('fontFamily', 'Segoe UI').split(',')[0].strip('"\'')
                font.size = Pt(get_font_size_pt(font_size_px))
                font.bold = inline_styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
                font.italic = inline_styles.get('fontStyle', 'normal') == 'italic'
                color = parse_color(inline_styles.get('color'))
                if color:
                    font.color.rgb = color
            
            # Trim trailing spaces from the last run in the last paragraph
            if text_frame.paragraphs and text_frame.paragraphs[-1].runs:
                last_run = text_frame.paragraphs[-1].runs[-1]
                last_run.text = last_run.text.rstrip()
                
    except Exception as e:
        print(f"Failed to add inline group element: {e}")


def add_overlay_element(slide, element, slide_width, slide_height):
    """Render overlay elements as shapes using their CSS."""
    x = element.get('x', 0)
    y = element.get('y', 0)
    width = max(1, element.get('width', 100))
    height = max(1, element.get('height', 10))
    styles = element.get('styles', {})
    element_class = element.get('className', '')
    
    # For overlay elements, use the rect coordinates if available
    rect = element.get('rect', {})
    if rect:
        x = rect.get('x', x)
        y = rect.get('y', y)
        width = max(1, rect.get('width', width))
        height = max(1, rect.get('height', height))
    
    x = max(0, min(x, slide_width - width))
    y = max(0, min(y, slide_height - height))
    
    bg_color = parse_color(styles.get('backgroundColor'))
    border_radius_str = styles.get('borderRadius', '0px')
    border_radius = parse_border_radius(border_radius_str, width, height)
    has_radius = border_radius > 0
    has_border = is_uniform_border(styles)
    has_any_border_sides = has_any_border(styles)
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none'
    
    if bg_color or has_border or has_any_border_sides or has_radius or has_shadow:
        add_bg_shape(slide, styles, x, y, width, height)
    
    text = element.get('text', '').strip()
    if text:
        textbox = slide.shapes.add_textbox(
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        text_frame.clear()
        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = text
        font = run.font
        font_size_px = safe_float(styles.get('fontSize', '12'))
        font.name = styles.get('fontFamily', 'Arial').split(',')[0].strip('"\'')
        font.size = Pt(max(6, get_font_size_pt(font_size_px)))
        font.bold = styles.get('fontWeight', '400') in ['bold', '700', '800', '900']
        font.italic = styles.get('fontStyle') == 'italic'
        color = parse_color(styles.get('color', 'black'))
        if color:
            font.color.rgb = color
        textbox.fill.background()
        textbox.line.fill.background()
        textbox.shadow.inherit = False


def get_bullet_style_for_list(element, list_info):
    """
    Determine bullet char, color, font, and size for a list based on className or CSS styles.
    Returns (bullet_char, bullet_color, bullet_font, bullet_size_pt)
    """
    class_name = element.get('className', '')
    styles = element.get('styles', {})
    bullet_char = '\u2022'
    bullet_color = RGBColor(40, 167, 69)
    bullet_font = 'Arial'
    bullet_size_pt = 12

    # Try to infer from class or parent context
    if 'top-points' in class_name:
        bullet_color = RGBColor(40, 167, 69)
        bullet_font = 'Arial'
        bullet_size_pt = 14
        bullet_char = '●'
    elif 'category-points' in class_name:
        bullet_color = RGBColor(40, 167, 69)
        bullet_font = 'Arial'
        bullet_size_pt = 12
        bullet_char = '●'
    # Try to get from styles if present
    color_str = styles.get('color')
    if color_str:
        parsed = parse_color(color_str)
        if parsed:
            bullet_color = parsed
    font_size = safe_float(styles.get('fontSize', '12'))
    if font_size:
        bullet_size_pt = get_font_size_pt(font_size)
    font_family = styles.get('fontFamily')
    if font_family:
        bullet_font = font_family.split(',')[0].strip('"\'')
    # Try to get bullet char from ::before content if available
    if 'listInfo' in element and element['listInfo'].get('bulletChar'):
        bullet_char = element['listInfo']['bulletChar']
    return bullet_char, bullet_color, bullet_font, bullet_size_pt

def add_list_paragraphs(text_frame, list_info, level=0, counters=None, element=None, bullet_style_override=None):
    if counters is None:
        counters = {}
    list_type = list_info.get('type')
    is_ordered = list_type == 'ol'
    list_styles = list_info.get('listStyles', {})
    list_style_type = list_styles.get('listStyleType', 'disc' if not is_ordered else 'decimal')
    if is_ordered:
        counter_key = f'ol_{level}'
        counters[counter_key] = list_info.get('start', 1) - 1
    items = list_info.get('items', [])
    avg_space_px = 0
    if len(items) > 1:
        spaces = []
        for i in range(len(items) - 1):
            item_bottom = items[i]['rect']['y'] + items[i]['rect']['height']
            next_top = items[i+1]['rect']['y']
            space_px = next_top - item_bottom
            if space_px > 0:
                spaces.append(space_px)
        avg_space_px = sum(spaces) / len(spaces) if spaces else 0
    space_after_pt = px_to_pt(avg_space_px)
    line_height_str = list_info.get('styles', {}).get('lineHeight', 'normal')
    if line_height_str == 'normal':
        line_spacing = 1.15
    else:
        try:
            line_spacing = float(line_height_str)
        except ValueError:
            line_spacing = 1.15

    # Determine bullet style for this list
    if bullet_style_override:
        bullet_char, bullet_color, bullet_font, bullet_size_pt = bullet_style_override
    else:
        bullet_char, bullet_color, bullet_font, bullet_size_pt = get_bullet_style_for_list(element or {}, list_info)

    first_item = True
    for item in items:
        p = None
        item_styles = item.get('styles', {})
        default_font_size_px = safe_float(item_styles.get('fontSize', '16').replace('px', ''))
        default_font_size_pt = get_font_size_pt(default_font_size_px)
        default_font_name = item_styles.get('fontFamily', 'Segoe UI').split(',')[0].strip('"\'')
        default_color = parse_color(item_styles.get('color'))

        if item.get('inlineGroup') and item['inlineGroup'].get('inlineElements'):
            inline_elements = item['inlineGroup']['inlineElements']
            first = True
            for inline_element in inline_elements:
                if inline_element.get('type') == 'br':
                    if p is not None:
                        p = text_frame.add_paragraph()
                        # Apply bullet formatting with proper size
                        makeParaBulletPointed(p, bullet_char, bullet_font, bullet_size_pt)
                        p.level = level
                        p.space_after = Pt(space_after_pt)
                        p.line_spacing = line_spacing
                    first = True
                    continue
                element_text = inline_element.get('text', '')
                if p is None:
                    if first_item and level == 0:
                        p = text_frame.paragraphs[0]
                    else:
                        p = text_frame.add_paragraph()
                    # Apply bullet formatting with proper size
                    makeParaBulletPointed(p, bullet_char, bullet_font, bullet_size_pt)
                    p.level = level
                    p.space_after = Pt(space_after_pt)
                    p.line_spacing = line_spacing
                
                if first:
                    element_text = element_text.lstrip()
                if not element_text.strip():
                    continue
                first = False
                run = p.add_run()
                run.text = element_text
                inline_styles = inline_element.get('styles', {})
                font = run.font
                font_size_px = safe_float(inline_styles.get('fontSize', '16').replace('px', ''))
                font.name = inline_styles.get('fontFamily', 'Segoe UI').split(',')[0].strip('"\'')
                font.size = Pt(get_font_size_pt(font_size_px))
                font.bold = inline_styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
                font.italic = inline_styles.get('fontStyle', 'normal') == 'italic'
                color = parse_color(inline_styles.get('color'))
                if color:
                    font.color.rgb = color
            if text_frame.paragraphs and text_frame.paragraphs[-1].runs:
                last_run = text_frame.paragraphs[-1].runs[-1]
                last_run.text = last_run.text.rstrip()
        else:
            if first_item and level == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            # Apply bullet formatting with proper size
            makeParaBulletPointed(p, bullet_char, bullet_font, bullet_size_pt)
            p.level = level
            p.space_after = Pt(space_after_pt)
            p.line_spacing = line_spacing

            run = p.add_run()
            run.text = item.get('text', '').strip()
            font = run.font
            font.name = default_font_name
            font.size = Pt(default_font_size_pt)
            font.bold = item_styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
            font.italic = item_styles.get('fontStyle', 'normal') == 'italic'
            color = parse_color(item_styles.get('color'))
            if color:
                font.color.rgb = color

        first_item = False
        if item.get('nestedList'):
            add_list_paragraphs(text_frame, item['nestedList'], level + 1, counters, element, bullet_style_override)

def add_list_element(slide, element, slide_width, slide_height, parent_has_shadow=False):
    list_info = element.get('listInfo', {})
    if not list_info.get('items'):
        return
    rect = list_info.get('rect', {})
    x = safe_int(rect.get('x', 0))
    y = safe_int(rect.get('y', 0))
    width = safe_int(rect.get('width', element.get('width', 100)))
    height = safe_int(rect.get('height', element.get('height', 100)))
    styles = element.get('styles', {})
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none' and not parent_has_shadow
    bg_color = parse_color(styles.get('backgroundColor'))
    border_radius_str = styles.get('borderRadius', '0px')
    border_radius = parse_border_radius(border_radius_str, width, height)
    has_radius = border_radius > 0
    has_border = is_uniform_border(styles)
    has_any_border_sides = has_any_border(styles)
    try:
        if bg_color or has_border or has_any_border_sides or has_radius or has_shadow:
            add_bg_shape(slide, styles, x, y, width, height)
        textbox = slide.shapes.add_textbox(
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.TOP
        text_frame.margin_left = pixels_to_emu(safe_float(styles.get('paddingLeft', '0px').replace('px', '')))
        text_frame.margin_right = pixels_to_emu(safe_float(styles.get('paddingRight', '0px').replace('px', '')))
        text_frame.margin_top = pixels_to_emu(safe_float(styles.get('paddingTop', '0px').replace('px', '')))
        text_frame.margin_bottom = pixels_to_emu(safe_float(styles.get('paddingBottom', '0px').replace('px', '')))
        textbox.fill.background()
        textbox.line.fill.background()
        textbox.shadow.inherit = False
        text_frame.clear()
        bullet_style = get_bullet_style_for_list(element, list_info)
        add_list_paragraphs(text_frame, list_info, element=element, bullet_style_override=bullet_style)
    except Exception as e:
        print(f"Failed to add list: {e}")


def add_link_element(slide, element, slide_width, slide_height, parent_has_shadow=False):
    """Add a hyperlink element to the slide (handles both direct and linkInfo style)"""
    # Support both direct and linkInfo structure
    link_info = element.get('linkInfo', {})
    link_href = link_info.get('href') or element.get('href', '')
    link_text = link_info.get('text') or element.get('text', '').strip()
    if not link_text:
        return

    x = element.get('x', 0)
    y = element.get('y', 0)
    width = max(1, element.get('width', 100))
    height = max(1, element.get('height', 20))
    x = max(0, min(x, slide_width - width))
    y = max(0, min(y, slide_height - height))
    styles = element.get('styles', {})

    try:
        textbox = slide.shapes.add_textbox(
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        text_frame.clear()
        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = link_text
        font = run.font
        font_size_px = safe_float(styles.get('fontSize', '12').replace('px', ''))
        font.name = styles.get('fontFamily', 'Arial').split(',')[0].strip('"\'')
        font.size = Pt(max(6, get_font_size_pt(font_size_px)))
        font.bold = styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
        font.italic = styles.get('fontStyle', 'normal') == 'italic'
        link_color = parse_color(styles.get('color', '#0066cc'))
        if link_color:
            font.color.rgb = link_color
        else:
            font.color.rgb = RGBColor(0, 102, 204)
        text_decoration = styles.get('textDecoration', 'underline')
        if 'underline' in text_decoration:
            font.underline = True
        if link_href:
            try:
                hyperlink = run.hyperlink
                hyperlink.address = link_href
            except Exception as hyperlink_error:
                print(f"Failed to add hyperlink functionality: {hyperlink_error}")
        text_align = styles.get('textAlign', 'left')
        if text_align == 'center':
            p.alignment = PP_ALIGN.CENTER
        elif text_align == 'right':
            p.alignment = PP_ALIGN.RIGHT
        else:
            p.alignment = PP_ALIGN.LEFT
        textbox.fill.background()
        textbox.line.fill.background()
        textbox.shadow.inherit = False
    except Exception as e:
        print(f"Failed to add link element: {e}")

def process_table_cell_content(slide, html_content, cell_x, cell_y, cell_width, cell_height, slide_width, slide_height, is_table_cell=False):
    """Process HTML content within a table cell and render elements"""
    if not html_content:
        return False
    
    has_renderable_content = False
    for content_elem in html_content:
        elem_type = content_elem.get('type')
        if elem_type == 'br':
            continue
        
        # Skip links in table cells - they'll be handled as cell text with hyperlinks
        if elem_type == 'a' and is_table_cell:
            continue
            
        elem_rect = content_elem.get('rect', {})
        rel_x = elem_rect.get('x', 0)
        rel_y = elem_rect.get('y', 0)
        elem_width = max(1, elem_rect.get('width', 10))
        elem_height = max(1, elem_rect.get('height', 10))
        abs_x = cell_x + rel_x
        abs_y = cell_y + rel_y
        abs_x = max(cell_x, min(abs_x, cell_x + cell_width - elem_width))
        abs_y = max(cell_y, min(abs_y, cell_y + cell_height - elem_height))
        abs_x = max(0, min(abs_x, slide_width - elem_width))
        abs_y = max(0, min(abs_y, slide_height - elem_height))
        content_elem['x'] = abs_x
        content_elem['y'] = abs_y
        content_elem['width'] = elem_width
        content_elem['height'] = elem_height
        if elem_type == 'img':
            add_image_element(slide, content_elem, slide_width, slide_height)
            has_renderable_content = True
        elif elem_type == 'canvas' and content_elem.get('chartInfo'):
            add_chart_element(slide, content_elem, slide_width, slide_height)
            has_renderable_content = True
        elif elem_type == 'a' and not is_table_cell:
            add_link_element(slide, content_elem, slide_width, slide_height)
            has_renderable_content = True
        elif elem_type in ['ul', 'ol']:
            add_list_element(slide, content_elem, slide_width, slide_height)
            has_renderable_content = True
        elif elem_type == 'div':
            if content_elem.get('text', '').strip():
                add_text_element(slide, content_elem, slide_width, slide_height)
                has_renderable_content = True
            elif (has_any_border(content_elem.get('styles', {})) or
                  parse_color(content_elem.get('styles', {}).get('backgroundColor'))):
                styles = content_elem.get('styles', {})
                add_bg_shape(slide, styles, abs_x, abs_y, elem_width, elem_height)
                has_renderable_content = True
        elif elem_type == 'span':
            if content_elem.get('text', '').strip():
                add_text_element(slide, content_elem, slide_width, slide_height)
                has_renderable_content = True
        elif content_elem.get('shapeInfo'):
            add_shape_element(slide, content_elem, slide_width, slide_height)
            has_renderable_content = True
    return has_renderable_content


def set_cell_border(cell, side, width_px, color_rgb, style='solid'):
    """Apply a border to a table cell side using enhanced method"""
    set_cell_border_enhanced(cell, side, width_px, color_rgb, style)

def apply_cell_css_borders(cell, cell_styles):
    """Apply CSS-specified borders to a table cell with uniform border optimization"""
    
    # Default border color (gray)
    DEFAULT_BORDER_COLOR = '808080'
    
    # Get border properties for all sides
    sides = ['Top', 'Right', 'Bottom', 'Left']
    border_props = {}
    
    for side in sides:
        width_key = f'border{side}Width'
        style_key = f'border{side}Style'
        color_key = f'border{side}Color'
        
        width_str = cell_styles.get(width_key, '0px')
        style_val = cell_styles.get(style_key, 'none')
        color_str = cell_styles.get(color_key, '')
        
        # Parse width (convert from px to numeric)
        width_px = safe_float(width_str.replace('px', '')) if width_str else 0
        
        # Use default color instead of parsing CSS color
        border_color = DEFAULT_BORDER_COLOR
        
        border_props[side.lower()] = {
            'width': width_px,
            'style': style_val,
            'color': border_color
        }
    
    # Check if all borders are uniform (same width, style, and color)
    first_side = border_props['top']
    is_uniform = all(
        border_props[side]['width'] == first_side['width'] and
        border_props[side]['style'] == first_side['style'] and
        border_props[side]['color'] == first_side['color']
        for side in ['top', 'right', 'bottom', 'left']
    )
    
    # Check if border should be applied (width > 0 and style not 'none')
    has_border = (
        first_side['width'] > 0 and 
        first_side['style'] not in ['none', 'hidden']
    )
    
    if is_uniform and has_border:
        # Use optimized uniform border function
        width_px = first_side['width']
        border_color = first_side['color']
        
        # Convert width to EMU (1pt = 12700 EMU, 1px ≈ 0.75pt)
        width_emu = str(int(width_px * 0.75 * 12700))
        
        _set_cell_border(cell, border_color=border_color, border_width=width_emu)
    else:
        # Apply borders individually for non-uniform cases
        side_names = ['top', 'right', 'bottom', 'left']
        
        for side_name in side_names:
            props = border_props[side_name]
            width_px = props['width']
            style_val = props['style']
            border_color = props['color']
            
            # Apply border if width > 0 and style is not 'none'
            if width_px > 0 and style_val not in ['none', 'hidden']:
                # Convert hex string to RGBColor for consistency
                try:
                    color_int = int(border_color, 16)
                    r = (color_int >> 16) & 0xFF
                    g = (color_int >> 8) & 0xFF
                    b = color_int & 0xFF
                    color_rgb = RGBColor(r, g, b)
                    set_cell_border_enhanced(cell, side_name, width_px, color_rgb, style_val)
                except:
                    # Fallback if color conversion fails
                    pass


def add_table_content(slide, content_data, content_type, cell_x, cell_y, cell_width, cell_height, slide_width, slide_height):
    """Add content (images, links, or charts) within a table cell with proper positioning"""
    try:
        if content_type == 'table_image':
            img_src = content_data.get('src', '')
            if not img_src:
                return
            
            # Calculate absolute position within the slide
            img_rect = content_data.get('rect', {})
            rel_x = img_rect.get('x', 0)
            rel_y = img_rect.get('y', 0)
            img_width = max(1, img_rect.get('width', 20))
            img_height = max(1, img_rect.get('height', 20))
            
            # Center image vertically in cell
            vertical_center_offset = (cell_height - img_height) / 2
            
            # Position relative to cell with vertical centering
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
            natural_height = content_data.get('naturalHeight', img_height)
            
            border_radius_str = styles.get('borderRadius', '0px')
            radius_ratio = parse_border_radius(border_radius_str, img_width, img_height)
            radius_display = radius_ratio * min(img_width, img_height)
            has_radius = radius_display > 0
            
            temp_path = None
            if img_src.startswith('data:'):
                _, data = img_src.split(',', 1)
                img_data_bytes = base64.b64decode(data)
                temp_path = 'temp_table_image.png'
                with open(temp_path, 'wb') as f:
                    f.write(img_data_bytes)
            elif img_src.startswith('http'):
                response = requests.get(img_src, timeout=10)
                if response.status_code == 200:
                    temp_path = 'temp_table_image.png'
                    with open(temp_path, 'wb') as f:
                        f.write(response.content)
                else:
                    print(f"Failed to download table image: {img_src}")
                    return
            else:
                if os.path.exists(img_src):
                    temp_path = img_src
                else:
                    print(f"Table image file not found: {img_src}")
                    return
            
            # Verify image before adding
            with Image.open(temp_path) as img:
                img.verify()
            
            image_to_add = temp_path
            if has_radius:
                scale_x = natural_width / img_width if img_width > 0 else 1
                radius_natural = int(radius_display * scale_x)
                temp_rounded = 'temp_table_rounded.png'
                make_rounded_image(temp_path, temp_rounded, radius_natural)
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
            if has_border:
                border_width = safe_float(styles.get('borderTopWidth', '0px'))
                border_color = parse_color(styles.get('borderTopColor'))
                if border_color and border_width > 0:
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
            
            # Clean up temporary files
            if image_to_add != temp_path and os.path.exists(image_to_add):
                os.remove(image_to_add)
            if temp_path != img_src and os.path.exists(temp_path):
                os.remove(temp_path)

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
            rel_y = chart_rect.get('y', 0)
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
            add_chart_element(slide, chart_element, slide_width, slide_height)
                    
    except Exception as e:
        print(f"Failed to add table content ({content_type}): {e}")
        return None

def add_table_element(slide, element, slide_width, slide_height, parent_has_shadow=False):
    table_info = element.get('tableInfo', {})
    if not table_info.get('rows'):
        return None
        
    rect = table_info.get('rect', {})
    x = safe_float(rect.get('x', element.get('x', 0)))
    y = safe_float(rect.get('y', element.get('y', 0)))
    width = safe_float(rect.get('width', element.get('width', 100)))
    height = safe_float(rect.get('height', element.get('height', 100)))
    
    element_class = element.get('className', '')
    styles = table_info.get('styles', {})
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none' and not parent_has_shadow
    bg_color = parse_color(styles.get('backgroundColor'))
    border_radius_str = styles.get('borderRadius', '0px')
    border_radius = parse_border_radius(border_radius_str, width, height)
    has_radius = border_radius > 0
    has_border = is_uniform_border(styles)
    has_any_border_sides = has_any_border(styles)
    try:
        # Clamp to slide bounds (keep existing safeguard)
        if x + width > slide_width:
            width = max(10, slide_width - x)
        if y + height > slide_height:
            height = max(10, slide_height - y)

        if bg_color or has_border or has_radius or has_shadow:
            add_bg_shape(slide, styles, x, y, width, height)

        rows = table_info['rowCount']
        cols = table_info['columnCount']
        table_shape = slide.shapes.add_table(
            rows, cols,
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )
        table = table_shape.table

        # Set table style to ensure borders render and disable default formatting
        try:
            tbl = table._graphic_frame._graphicFrame.graphicData.tbl
            style_id = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'
            tbl[0][-1].text = style_id
            table.first_row = False  # Disable header row default styling
        except Exception:
            pass

        # --- Column width calculation ---
        if table_info['rows']:
            first_row = next((r for r in table_info['rows'] if r.get('cells')), None)
            if first_row:
                current_col = 0
                row_cells = first_row['cells']
                total_assigned = 0.0
                for cell_data in row_cells:
                    cell_rect = cell_data.get('rect', {})
                    cell_w = cell_rect.get('width', 0)
                    col_span = max(1, cell_data.get('colSpan', 1))
                    per_col_w = cell_w / col_span if col_span > 1 else cell_w
                    for i in range(col_span):
                        if current_col + i < cols:
                            table.columns[current_col + i].width = pixels_to_emu(max(8, per_col_w))
                            total_assigned += per_col_w
                    current_col += col_span
                # If any remaining columns distribute remaining space
                remaining_cols = cols - current_col
                if remaining_cols > 0:
                    remaining_w = max(0, width - total_assigned)
                    per_remaining = remaining_w / remaining_cols if remaining_cols else 0
                    for i in range(remaining_cols):
                        idx = current_col + i
                        if idx < cols:
                            table.columns[idx].width = pixels_to_emu(max(8, per_remaining))

        # --- row heights with consistent sizing and content height consideration ---
        for row_data in table_info['rows']:
            r_index = row_data['index']
            if r_index >= rows:
                continue
            
            # Get CSS height from row data
            row_styles = row_data.get('styles', {})
            css_height = row_styles.get('height', '')
            
            # Parse CSS height (e.g., "25px" -> 25)
            if css_height and 'px' in css_height:
                try:
                    specified_height = float(css_height.replace('px', ''))
                except:
                    specified_height = None
            else:
                specified_height = None
            
            # Get rendered height from rect
            rendered_height = row_data.get('rect', {}).get('height', 25)
            
            # Calculate maximum content height in this row by checking htmlContent
            max_content_height = 0
            for cell_data in row_data.get('cells', []):
                html_content = cell_data.get('htmlContent', [])
                for content_elem in html_content:
                    if content_elem.get('type') == 'br':
                        continue
                    elem_rect = content_elem.get('rect', {})
                    elem_height = elem_rect.get('height', 0)
                    # Add relative y position to get total height needed
                    total_height = elem_rect.get('y', 0) + elem_height
                    if total_height > max_content_height:
                        max_content_height = total_height
            
            # Choose the most appropriate height considering content
            if specified_height:
                # Use CSS specified height as primary for consistency with HTML
                final_height = specified_height
            else:
                # Fall back to rendered height
                final_height = rendered_height
            
            # If there are nested elements taller than the current height, adjust
            if max_content_height > 0:
                # Add padding for content (8px top + 8px bottom = 16px total)
                required_height_for_content = max_content_height + 16
                final_height = max(final_height, required_height_for_content)
            
            final_height = max(20, min(final_height, 300))  # Increased max height for charts
            
            table.rows[r_index].height = pixels_to_emu(final_height)

        # Create a simple grid to track occupied cells for rowspan
        occupied_cells = set()
        
        # Track elements to add after table creation
        table_cell_html_content = []
        
        # Populate & style cells
        for row_data in table_info['rows']:
            row_index = row_data['index']
            if row_index >= rows:
                continue
                
            current_col = 0
            for cell_data in row_data.get('cells', []):
                # Skip occupied columns in this row
                while current_col < cols and (row_index, current_col) in occupied_cells:
                    current_col += 1
                
                if current_col >= cols:
                    break
                    
                col_span = max(1, cell_data.get('colSpan', 1))
                row_span = max(1, cell_data.get('rowSpan', 1))
                
                # Mark all cells that will be occupied by this cell's rowspan/colspan
                for r in range(row_index, min(rows, row_index + row_span)):
                    for c in range(current_col, min(cols, current_col + col_span)):
                        occupied_cells.add((r, c))
                
                pptx_cell = table.cell(row_index, current_col)

                # Handle merges with rowspan and colspan
                if col_span > 1 or row_span > 1:
                    try:
                        end_row = min(rows - 1, row_index + row_span - 1)
                        end_col = min(cols - 1, current_col + col_span - 1)
                        
                        # Perform the merge
                        pptx_cell = pptx_cell.merge(table.cell(end_row, end_col))
                    except Exception as e:
                        print(f"Error merging cell ({row_index}, {current_col}) with span {row_span}x{col_span}: {e}")

                # Apply CSS-specified borders to this cell
                cell_styles = cell_data.get('styles', {})
                apply_cell_css_borders(pptx_cell, cell_styles)

                # Calculate absolute cell position for nested elements
                cell_rect = cell_data.get('rect', {})
                abs_cell_x = x + (cell_rect.get('x', 0) - rect.get('x', 0))
                abs_cell_y = y + (cell_rect.get('y', 0) - rect.get('y', 0))
                cell_width = cell_rect.get('width', 50)
                cell_height = cell_rect.get('height', 20)
                
                # Check if cell has HTML content
                html_content = cell_data.get('htmlContent', [])
                has_html_content = bool(html_content)
                
                # Check if cell has link element
                has_link = False
                link_element = None
                if has_html_content:
                    for content_elem in html_content:
                        if content_elem.get('type') == 'a':
                            has_link = True
                            link_element = content_elem
                            break
                
                # Store non-link HTML content for processing after table creation
                if has_html_content:
                    # Filter out link elements - they'll be handled as cell text
                    non_link_content = [elem for elem in html_content if elem.get('type') != 'a']
                    if non_link_content:
                        table_cell_html_content.append({
                            'html_content': non_link_content,
                            'cell_x': abs_cell_x,
                            'cell_y': abs_cell_y,
                            'cell_width': cell_width,
                            'cell_height': cell_height
                        })

                if pptx_cell is None:
                    print(f"Warning: pptx_cell is None for row {row_index}, cell {current_col}")
                    continue

                text_frame = pptx_cell.text_frame
                text_frame.word_wrap = True

                vertical_align = (cell_styles.get('verticalAlign') or '').strip().lower()
                if vertical_align == 'top':
                    pptx_cell.vertical_anchor = MSO_ANCHOR.TOP
                elif vertical_align == 'bottom':
                    pptx_cell.vertical_anchor = MSO_ANCHOR.BOTTOM
                else:
                    pptx_cell.vertical_anchor = MSO_ANCHOR.MIDDLE

                padding_left = safe_float(cell_styles.get('paddingLeft', '6px'))
                padding_right = safe_float(cell_styles.get('paddingRight', '6px'))
                padding_top = safe_float(cell_styles.get('paddingTop', '4px'))
                padding_bottom = safe_float(cell_styles.get('paddingBottom', '4px'))
                text_frame.margin_left = pixels_to_emu(padding_left)
                text_frame.margin_right = pixels_to_emu(padding_right)
                text_frame.margin_top = pixels_to_emu(padding_top)
                text_frame.margin_bottom = pixels_to_emu(padding_bottom)

                row_bg_color = parse_color(row_data.get('styles', {}).get('backgroundColor'))
                bg_color_cell = parse_color(cell_styles.get('backgroundColor'))
                if bg_color_cell:
                    pptx_cell.fill.solid()
                    pptx_cell.fill.fore_color.rgb = bg_color_cell
                elif row_bg_color:
                    pptx_cell.fill.solid()
                    pptx_cell.fill.fore_color.rgb = row_bg_color
                else:
                    pptx_cell.fill.background()

                # Handle link as cell text with hyperlink
                if has_link and link_element:
                    link_info = link_element.get('linkInfo', {})
                    link_href = link_info.get('href') or link_element.get('href', '')
                    link_text = link_info.get('text') or link_element.get('text', '').strip()
                    link_styles = link_element.get('styles', {})
                    
                    if link_text:
                        text_frame.clear()
                        p = text_frame.paragraphs[0]
                        text_align = cell_styles.get('textAlign', 'left')
                        p.alignment = PP_ALIGN.CENTER if text_align == 'center' else PP_ALIGN.RIGHT if text_align == 'right' else PP_ALIGN.LEFT
                        
                        run = p.add_run()
                        run.text = link_text
                        
                        # Apply link formatting
                        font = run.font
                        font_size_px = safe_float(link_styles.get('fontSize', '12').replace('px', ''))
                        font.size = Pt(max(6, get_font_size_pt(font_size_px)))
                        font.name = link_styles.get('fontFamily', 'Arial').split(',')[0].strip('"\'')
                        font.bold = link_styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
                        font.italic = link_styles.get('fontStyle', 'normal') == 'italic'
                        
                        # Apply link color
                        link_color = parse_color(link_styles.get('color', '#0066cc'))
                        if link_color:
                            font.color.rgb = link_color
                        else:
                            font.color.rgb = RGBColor(0, 102, 204)
                        
                        # Add underline
                        text_decoration = link_styles.get('textDecoration', 'underline')
                        if 'underline' in text_decoration:
                            font.underline = True
                        
                        # Add hyperlink functionality
                        if link_href:
                            try:
                                hyperlink = run.hyperlink
                                hyperlink.address = link_href
                            except Exception as hyperlink_error:
                                print(f"Failed to add hyperlink to table cell: {hyperlink_error}")
                        
                        p.space_before = Pt(0)
                        p.space_after = Pt(0)
                        p.line_spacing = 1.0
                else:
                    # Handle regular cell text (no link)
                    cell_text = (cell_data.get('text') or '').strip()
                    has_only_br = has_html_content and all(
                        elem.get('type') == 'br' for elem in html_content
                    )
                    
                    if (not has_html_content or has_only_br) and cell_text:
                        text_frame.clear()
                        p = text_frame.paragraphs[0]
                        text_align = cell_styles.get('textAlign', 'left')
                        p.alignment = PP_ALIGN.CENTER if text_align == 'center' else PP_ALIGN.RIGHT if text_align == 'right' else PP_ALIGN.LEFT

                        if cell_data.get('inlineGroup') and cell_data['inlineGroup'].get('inlineElements'):
                            first = True
                            for inline_element in cell_data['inlineGroup']['inlineElements']:
                                if inline_element.get('type') == 'br':
                                    p = text_frame.add_paragraph()
                                    p.alignment = PP_ALIGN.CENTER if text_align == 'center' else PP_ALIGN.RIGHT if text_align == 'right' else PP_ALIGN.LEFT
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
                            if text_frame.paragraphs and text_frame.paragraphs[-1].runs:
                                last_run = text_frame.paragraphs[-1].runs[-1]
                                last_run.text = last_run.text.rstrip()
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
                    else:
                        text_frame.clear()
                
                current_col += col_span
        
        # Process non-link HTML content for all cells after table is created
        for content_data in table_cell_html_content:
            process_table_cell_content(
                slide,
                content_data['html_content'],
                content_data['cell_x'],
                content_data['cell_y'],
                content_data['cell_width'],
                content_data['cell_height'],
                slide_width,
                slide_height,
                is_table_cell=True
            )
        
        return table_shape
    except Exception as e:
        print(f"Failed to add table: {e}")
        return None


def add_image_element(slide, element, slide_width, slide_height, parent_has_shadow=False):
    media_info = element.get('mediaInfo', {})
    img_src = media_info.get('src', '')
    styles = element.get('styles', {})
    if not img_src:
        return
    # Use precise positioning from extraction
    x = element.get('x', 0)
    y = element.get('y', 0)
    width = max(1, element.get('width', 100))
    height = max(1, element.get('height', 100))
   
    # Ensure coordinates are within slide bounds
    x = max(0, min(x, slide_width - width))
    y = max(0, min(y, slide_height - height))
   
    natural_width = media_info.get('naturalWidth', width)
    border_radius_str = styles.get('borderRadius', '0px')
    radius_ratio = parse_border_radius(border_radius_str, width, height)
    radius_display = radius_ratio * min(width, height)
    has_radius = radius_display > 0
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none' and not parent_has_shadow
    has_border = is_uniform_border(styles)
   
    try:
        temp_path = None
        if img_src.startswith('data:'):
            _, data = img_src.split(',', 1)
            img_data = base64.b64decode(data)
            temp_path = 'temp_image.png'
            with open(temp_path, 'wb') as f:
                f.write(img_data)
        elif img_src.startswith('http'):
            response = requests.get(img_src, timeout=10)
            if response.status_code == 200:
                temp_path = 'temp_image.png'
                with open(temp_path, 'wb') as f:
                    f.write(response.content)
            else:
                print(f"Failed to download image: {img_src}")
                return
        else:
            if os.path.exists(img_src):
                temp_path = img_src
            else:
                print(f"Image file not found: {img_src}")
                return
       
        with Image.open(temp_path) as img:
            img.verify()
       
        image_to_add = temp_path
        if has_radius:
            scale_x = natural_width / width if width > 0 else 1
            radius_natural = int(radius_display * scale_x)
            temp_rounded = 'temp_rounded.png'
            make_rounded_image(temp_path, temp_rounded, radius_natural)
            image_to_add = temp_rounded
       
        # Add image with precise positioning
        picture = slide.shapes.add_picture(
            image_to_add,
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )
        picture.shadow.inherit = False
       
        # Handle borders and shadows
        if has_border or has_radius or has_shadow:
            shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if has_radius else MSO_SHAPE.RECTANGLE
            border_shape = slide.shapes.add_shape(
                shape_type,
                pixels_to_emu(x), pixels_to_emu(y),
                pixels_to_emu(width), pixels_to_emu(height)
            )
            if has_radius:
                border_shape.adjustments[0] = radius_ratio
            border_shape.fill.background()
            border_shape.shadow.inherit = False
            if has_shadow:
                apply_shadow(border_shape, box_shadow)
            if has_border:
                border_width = safe_float(styles.get('borderTopWidth', '0px'))
                border_color = parse_color(styles.get('borderTopColor'))
                if border_color:
                    border_shape.line.width = Pt(border_width)
                    border_shape.line.color.rgb = border_color
                    border_style = styles.get('borderTopStyle', 'solid')
                    if border_style == 'dashed':
                        border_shape.line.dash_style = MSO_LINE.DASH
                    elif border_style == 'dotted':
                        border_shape.line.dash_style = MSO_LINE.ROUND_DOT
            else:
                border_shape.line.fill.background()
            # Move border shape behind picture
            sp = border_shape._sp
            parent = sp.getparent()
            parent.remove(sp)
            pic_sp = picture._sp
            idx = list(parent).index(pic_sp)
            parent.insert(idx, sp)
        elif has_shadow:
            apply_shadow(picture, box_shadow)
       
        # Clean up temporary files
        if image_to_add != temp_path and os.path.exists(image_to_add):
            os.remove(image_to_add)
        if temp_path != img_src and os.path.exists(temp_path):
            os.remove(temp_path)
    except Exception as e:
        print(f"Failed to add image: {e}")


def add_text_element(slide, element, slide_width, slide_height, parent_has_shadow=False):
    """Enhanced text element creation with precise positioning"""
    text = element.get('text', '').strip()
    if not text:
        return
   
    # Use precise positioning
    x = element.get('x', 0)
    y = element.get('y', 0)
    width = max(1, element.get('width', 100))
    height = max(1, element.get('height', 100))
   
    # Ensure coordinates are within slide bounds
    x = max(0, min(x, slide_width - width))
    y = max(0, min(y, slide_height - height))
   
    styles = element.get('styles', {})
    element_class = element.get('className', '')

    if ('table-heading' in element_class or 'heading' in element_class or 
        'pie-heading' in element_class):
        padding_left = safe_float(styles.get('paddingLeft', '0px'))
        
        # For headings with ::before pseudo elements, adjust position
        if padding_left > 0:
            x += max(0, padding_left - 8)   # Smaller offset to prevent too much gap
            width = max(1, width - padding_left + 8)  # Adjust width accordingly
    
    box_shadow = styles.get('boxShadow', 'none')
    has_shadow = box_shadow != 'none' and not parent_has_shadow
    bg_color = parse_color(styles.get('backgroundColor'))
    border_radius_str = styles.get('borderRadius', '0px')
    border_radius = parse_border_radius(border_radius_str, width, height)
    has_radius = border_radius > 0
    has_border = is_uniform_border(styles)
   
    has_any_border_sides = has_any_border(styles)
   
    try:
        # Add background/border shapes first
        if bg_color or has_border or has_any_border_sides or has_radius or has_shadow:
            add_bg_shape(slide, styles, x, y, width, height)
       
        # Create text box with precise positioning
        textbox = slide.shapes.add_textbox(
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )
       
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
       
        # Set vertical alignment based on flex properties or default to middle for short text
        display = styles.get('display', 'block')
        align_items = styles.get('alignItems', 'stretch')
        justify_content = styles.get('justifyContent', 'flex-start')
       
        if display == 'flex' and align_items == 'center':
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        elif height < 50: # For small elements like company names
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        else:
            text_frame.vertical_anchor = MSO_ANCHOR.TOP
       
        # margin calculation for headings with pseudo elements
        margin_left = safe_float(styles.get('paddingLeft', '0px'))
        margin_right = safe_float(styles.get('paddingRight', '0px'))
        margin_top = safe_float(styles.get('paddingTop', '0px'))
        margin_bottom = safe_float(styles.get('paddingBottom', '0px'))
        
        # Special adjustment for headings with pseudo elements
        if ('table-heading' in element_class or 'pie-heading' in element_class):
            # Reduce left margin to account for positioning adjustment
            margin_left = max(0, margin_left - 12)
        
        text_frame.margin_left = pixels_to_emu(margin_left)
        text_frame.margin_right = pixels_to_emu(margin_right)
        text_frame.margin_top = pixels_to_emu(margin_top)
        text_frame.margin_bottom = pixels_to_emu(margin_bottom)
       
        # Remove textbox styling
        textbox.fill.background()
        textbox.line.fill.background()
        textbox.shadow.inherit = False
       
        # Clear and set text
        text_frame.clear()
        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = text
       
        # Apply text formatting
        font = run.font
        font_size_px = safe_float(styles.get('fontSize', '12'))
        font.name = styles.get('fontFamily', 'Arial').split(',')[0].strip('"\'')
        font.size = Pt(max(6, get_font_size_pt(font_size_px)))
        font.bold = styles.get('fontWeight', '400') in ['bold', '700', '800', '900']
        font.italic = styles.get('fontStyle') == 'italic'
       
        # Apply text color
        color = parse_color(styles.get('color', 'black'))
        if color:
            font.color.rgb = color
       
        # Apply text alignment based on CSS text-align and flex properties
        text_align = styles.get('textAlign', 'left')
        if text_align == 'center' or (display == 'flex' and justify_content == 'center'):
            p.alignment = PP_ALIGN.CENTER
        elif text_align == 'right' or (display == 'flex' and justify_content == 'flex-end'):
            p.alignment = PP_ALIGN.RIGHT
        else:
            p.alignment = PP_ALIGN.LEFT
           
    except Exception as e:
        print(f"Failed to add text: {e}")


def parse_linear_gradient(gradient_str, total_width):
    """Parse linear-gradient color stops (hex, rgb, rgba) into solid segments.

    Supports patterns like:
      linear-gradient(to right, rgb(37,29,95) 0%, rgb(37,29,95) 12%, rgb(217,217,217) 12%, rgb(217,217,217) 100%)
      linear-gradient(to right, #251d5f 0%, #251d5f 12%, #D9D9D9 12%, #D9D9D9 100%)
    Returns list of (x, width, RGBColor)
    """
    if not isinstance(gradient_str, str) or 'linear-gradient' not in gradient_str:
        return []

    try:
        # Match both hex and rgb/rgba colors with percentage stops
        pattern = r'((?:#[0-9a-fA-F]{3,6})|(?:rgb[a]?\(\s*\d+\s*,\s*\d+\s*,\s*\d+(?:\s*,\s*[\d.]+\s*)?\)))\s+([\d.]+)%'
        stops = re.findall(pattern, gradient_str)
        if not stops or len(stops) < 2:
            return []

        # Build segments from consecutive stop pairs
        segments = []
        for i in range(len(stops) - 1):
            color_text, start_pct = stops[i]
            _, end_pct = stops[i + 1]
           
            start_pct = float(start_pct)
            end_pct = float(end_pct)
            if end_pct <= start_pct:
                continue
            seg_x = (start_pct / 100.0) * total_width
            seg_w = ((end_pct - start_pct) / 100.0) * total_width
            color = parse_color(color_text)
            if color and seg_w > 0:
                segments.append((seg_x, seg_w, color))

        # If last stop color not represented (should not happen with pair iteration), ignore.

        return segments
    except Exception as e:
        print(f"Gradient parse failed: {e}")
        return []


def add_pseudo_element(slide, element, slide_width, slide_height):
    x = safe_float(element.get('x', 0))
    y = safe_float(element.get('y', 0))
    width = max(1, safe_float(element.get('width', 0)))
    height = max(1, safe_float(element.get('height', 0)))
    styles = element.get('styles', {})
    text = element.get('text', '')
    pseudo_type = element.get('pseudoType', '')
    parent_class = element.get('parentClassName', '')
    parent_tag = element.get('parentTagName', '')
    x = max(0, min(x, slide_width - width))
    y = max(0, min(y, slide_height - height))
    background_full = styles.get('background', '')
    has_gradient = isinstance(background_full, str) and 'linear-gradient' in background_full

    if has_gradient:
        segments = parse_linear_gradient(background_full, width)
        if segments:
            for seg_x, seg_w, color in segments:
                if seg_w > 0 and color:
                    try:
                        shape = slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            pixels_to_emu(x + seg_x),
                           
                            pixels_to_emu(y),
                            pixels_to_emu(seg_w),
                            pixels_to_emu(height)
                        )
                        shape.fill.solid()
                        shape.fill.fore_color.rgb = color
                        shape.line.fill.background()
                        shape.shadow.inherit = False
                    except Exception as e:
                        print(f"Error creating gradient segment: {e}")
        else:
            add_bg_shape(slide, styles, x, y, width, height)
    else:
        bg_color = parse_color(styles.get('backgroundColor'))
        if bg_color or has_gradient:
            add_bg_shape(slide, styles, x, y, width, height)

    if (
        pseudo_type == '::before'
        and parent_tag == 'li'
        and text
        and 'parentText' in element
        and element['parentText'].strip()
    ):
        try:
            textbox = slide.shapes.add_textbox(
                pixels_to_emu(x), pixels_to_emu(y),
                pixels_to_emu(width), pixels_to_emu(height)
            )
            text_frame = textbox.text_frame
            text_frame.word_wrap = True
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            text_frame.margin_left = 0
            text_frame.margin_right = 0
            text_frame.margin_top = 0
            text_frame.margin_bottom = 0
            text_frame.clear()
            p = text_frame.paragraphs[0]

            run_bullet = p.add_run()
            run_bullet.text = text + " "
            font_bullet = run_bullet.font
            font_size_px = safe_float(styles.get('fontSize', '12'))
            font_bullet.name = styles.get('fontFamily', 'Meiryo').split(',')[0].strip('"\'')
            font_bullet.size = Pt(max(6, get_font_size_pt(font_size_px)))
            font_bullet.bold = styles.get('fontWeight', '400') in ['bold', '700', '800', '900']
            font_bullet.italic = styles.get('fontStyle', 'italic')
            color = parse_color(styles.get('color', 'black'))
            if color:
                font_bullet.color.rgb = color

            run_text = p.add_run()
            run_text.text = element['parentText'].strip()
            parent_styles = element.get('parentStyles', {})
            font_text = run_text.font
            font_size_px2 = safe_float(parent_styles.get('fontSize', styles.get('fontSize', '12')))
            font_text.name = parent_styles.get('fontFamily', styles.get('fontFamily', 'Meiryo')).split(',')[0].strip('"\'')
            font_text.size = Pt(max(6, get_font_size_pt(font_size_px2)))
            font_text.bold = parent_styles.get('fontWeight', '400') in ['bold', '700', '800', '900']
            font_text.italic = parent_styles.get('fontStyle', 'normal') == 'italic'
            color2 = parse_color(parent_styles.get('color', styles.get('color', 'black')))
            if color2:
                font_text.color.rgb = color2
            text_align = parent_styles.get('textAlign', styles.get('textAlign', 'left'))
            p.alignment = PP_ALIGN.CENTER if text_align == 'center' else PP_ALIGN.RIGHT if text_align == 'right' else PP_ALIGN.LEFT
            textbox.fill.background()
            textbox.line.fill.background()
            textbox.shadow.inherit = False

            element['_li_bullet_text_rendered'] = True
        except Exception as e:
            print(f"Failed to add bullet+text pseudo element: {e}")
        return

    # Fallback: normal pseudo element rendering
    if text and text not in ['""', "''", 'none']:
        try:
            textbox = slide.shapes.add_textbox(
                pixels_to_emu(x), pixels_to_emu(y),
                pixels_to_emu(width), pixels_to_emu(height)
            )
            text_frame = textbox.text_frame
            text_frame.word_wrap = True
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            text_frame.margin_left = 0
            text_frame.margin_right = 0
            text_frame.margin_top = 0
            text_frame.margin_bottom = 0
            text_frame.clear()
            p = text_frame.paragraphs[0]
            run = p.add_run()
            run.text = text
            font = run.font
            font_size_px = safe_float(styles.get('fontSize', '12'))
            font.name = styles.get('fontFamily', 'Meiryo').split(',')[0].strip('"\'')
            font.size = Pt(max(6, get_font_size_pt(font_size_px)))
            font.bold = styles.get('fontWeight', '400') in ['bold', '700', '800', '900']
            font.italic = styles.get('fontStyle', 'italic')
            color = parse_color(styles.get('color', 'black'))
            if color:
                font.color.rgb = color
            text_align = styles.get('textAlign', 'left')
            p.alignment = PP_ALIGN.CENTER if text_align == 'center' else PP_ALIGN.RIGHT if text_align == 'right' else PP_ALIGN.LEFT
            textbox.fill.background()
            textbox.line.fill.background()
            textbox.shadow.inherit = False
        except Exception as e:
            print(f"Failed to add pseudo element text: {e}")


def parse_datalabel_formatter(formatter_str, labels, values, index):
    """
    Parses a Chart.js datalabel formatter string and returns the formatted label.
    Handles different formatter structures for labels, values, or both.
    """
    if not formatter_str or "function" not in formatter_str:
        return str(values[index]) if index < len(values) else ""

    label = labels[index] if index < len(labels) else ""
    value = values[index] if index < len(values) else ""

    # Extract the return statement from the function string
    return_statement_match = re.search(r'return\s+(.*?);', formatter_str)
    if not return_statement_match:
        return str(value)

    return_expression = return_statement_match.group(1).strip()
    
    # Build the formatted string based on the expression
    parts = [part.strip() for part in return_expression.split('+')]
    result_parts = []
    for part in parts:
        part = part.strip('"\'') # Remove string quotes
        if 'context.chart.data.labels[context.dataIndex]' in part:
            result_parts.append(str(label))
        elif 'value' in part:
            result_parts.append(str(value))
        elif part == '\\n': # Handle newlines
            result_parts.append('\n')
        else: # Handle static strings like '%'
            result_parts.append(part)
            
    return "".join(result_parts)

def apply_datalabel_positioning(data_labels, anchor, align, offset=0, chart_type_str="pie"):
    """Apply Chart.js style positioning to PowerPoint data labels"""
    try:
        if chart_type_str == "pie":
            if anchor == "end" and align == "start":
                # Outside the slice, at the edge
                data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
            elif anchor == "center" or (not anchor and not align):
                # Center of slice (default)
                data_labels.position = XL_LABEL_POSITION.CENTER
            elif anchor == "start":
                # Inside, near center
                data_labels.position = XL_LABEL_POSITION.INSIDE_END
            else:
                # Default to outside for better readability
                data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        elif chart_type_str == "bar":
            if align == "top" or anchor == "end":
                data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
            elif align == "center" or anchor == "center":
                data_labels.position = XL_LABEL_POSITION.CENTER
            else:
                data_labels.position = XL_LABEL_POSITION.INSIDE_END
        else:
            if anchor == "end" or align == "top":
                data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
            elif anchor == "center" or align == "center":
                data_labels.position = XL_LABEL_POSITION.CENTER
            else:
                data_labels.position = XL_LABEL_POSITION.INSIDE_END
    except Exception as e:
        print(f"Error applying data label positioning: {e}")

def add_chart_element(slide, element, slide_width, slide_height):
    """Enhanced chart rendering with accurate configuration extraction"""
    chart_config = None
    
    if element.get('chartInfo'):
        chart_info = element['chartInfo']
        chart_data = chart_info.get('chartData')
        
        if chart_data:
            if isinstance(chart_data, dict):
                # Data is already a parsed dictionary
                chart_config = chart_data
            elif isinstance(chart_data, str):
                # Data is a string, needs parsing
                try:
                    chart_config = json.loads(chart_data)
                except json.JSONDecodeError as e:
                    print(f"Error parsing chart data string: {e}")
                    return
            else:
                # Unsupported format
                return
    
    # Fallback for older format
    if not chart_config:
        chart_config = element.get('chartConfig')

    if not chart_config:
        return
    
    # Use element dimensions directly
    x = max(0, min(element.get('x', 0), slide_width - 10))
    y = max(0, min(element.get('y', 0), slide_height - 10))
    width = max(10, min(element.get('width', 400), slide_width - x))
    height = max(10, min(element.get('height', 300), slide_height - y))
    
    # Apply scaling factor if canvas dimensions differ from CSS dimensions
    if element.get('chartInfo'):
        chart_info = element['chartInfo']
        canvas_width = chart_info.get('width', width)
        canvas_height = chart_info.get('height', height)
        
        # Use canvas dimensions if they're more accurate
        if canvas_width > 0 and canvas_height > 0:
            width = min(canvas_width, slide_width - x)
            height = min(canvas_height, slide_height - y)
    
    try:
        chart_type_str = chart_config.get('type')
        options = chart_config.get('options', {})
        
        # Check if all data values are negative for bar charts
        all_negative = False
        has_negative_values = False
        if chart_type_str == 'bar':
            datasets = chart_config.get('data', {}).get('datasets', [])
            if datasets:
                all_values = []
                for dataset in datasets:
                    data_values = dataset.get('data', [])
                    all_values.extend([val for val in data_values if isinstance(val, (int, float))])
                
                if all_values:
                    has_negative_values = any(val < 0 for val in all_values)
                    all_negative = all(val < 0 for val in all_values)
        
        # Map chart types with better handling
        if chart_type_str == 'pie':
            chart_type = XL_CHART_TYPE.PIE
        elif chart_type_str == 'doughnut':
            # Doughnut charts in PowerPoint are rendered as pie charts with a hole size
            chart_type = XL_CHART_TYPE.DOUGHNUT
        elif chart_type_str == 'line':
            datasets = chart_config.get('data', {}).get('datasets', [])
            has_fill = any(dataset.get('fill') is not None and dataset.get('fill') != False for dataset in datasets)
            if has_fill:
                chart_type = XL_CHART_TYPE.AREA_STACKED if any(dataset.get('fill') == '-1' for dataset in datasets) else XL_CHART_TYPE.AREA
            else:
                point_radius = datasets[0].get('pointRadius', 0) if datasets else 0
                chart_type = XL_CHART_TYPE.LINE_MARKERS if point_radius > 0 else XL_CHART_TYPE.LINE
        elif chart_type_str == 'bar':
            index_axis = options.get('indexAxis', 'x')
            chart_type = XL_CHART_TYPE.BAR_CLUSTERED if index_axis == 'y' else XL_CHART_TYPE.COLUMN_CLUSTERED
        else:
            return
        
        chart_data = CategoryChartData()
        data = chart_config.get('data', {})
        labels = data.get('labels', [])
        
        # Enhanced label handling for multi-line labels
        processed_labels = []
        for label in labels:
            if isinstance(label, str):
                # Handle \n in labels by replacing with actual line breaks
                processed_label = label.replace('\\n', '\n')
                processed_labels.append(processed_label)
            else:
                processed_labels.append(str(label))
        
        # Reverse category order for horizontal bar charts to match HTML rendering
        is_horizontal_bar = chart_type == XL_CHART_TYPE.BAR_CLUSTERED
        if is_horizontal_bar:
            processed_labels = list(reversed(processed_labels))
        
        # For doughnut/pie charts without labels, use empty strings for categories
        if chart_type_str in ['pie', 'doughnut'] and not processed_labels:
            datasets = data.get('datasets', [])
            if datasets and datasets[0].get('data'):
                processed_labels = [''] * len(datasets[0]['data'])
        
        chart_data.categories = processed_labels
        
        # Enhanced dataset handling to prevent label display issues
        datasets = data.get('datasets', [])
        for i, dataset in enumerate(datasets):
            # Get the series label but don't use it if it should be hidden
            series_label = dataset.get('label', f'Series {i+1}')
            series_data = dataset.get('data', [])
            
            # Reverse data order for horizontal bar charts to match category order
            if is_horizontal_bar:
                series_data = list(reversed(series_data))
            
            # For single series charts where legend is explicitly disabled, use empty label
            plugins = options.get('plugins', {})
            legend_opts = plugins.get('legend', {})
            legend_display = legend_opts.get('display')
            
            if legend_display is False and len(datasets) == 1:
                # Use empty string for series label to prevent it showing up as text
                chart_data.add_series('', tuple(series_data))
            else:
                chart_data.add_series(series_label, tuple(series_data))
        
        graphic_frame = slide.shapes.add_chart(
            chart_type,
            pixels_to_emu(x),
            pixels_to_emu(y),
            pixels_to_emu(width),
            pixels_to_emu(height),
            chart_data
        )
        chart = graphic_frame.chart

        # Configure doughnut hole size (cutout percentage)
        if chart_type_str == 'doughnut':
            try:
                cutout_str = options.get('cutout', '50%')
                # Parse cutout percentage (e.g., "70%" -> 70)
                if isinstance(cutout_str, str) and cutout_str.endswith('%'):
                    cutout_pct = int(cutout_str.rstrip('%'))
                else:
                    cutout_pct = int(cutout_str)
                
                # PowerPoint uses hole size (0-90), where Chart.js uses cutout percentage
                # Convert: hole_size = cutout_pct (clamped to 10-90 for safety)
                hole_size = max(10, min(90, cutout_pct))
                
                # Access the doughnut chart's hole size property
                plot = chart.plots[0]
                plot.hole_size = hole_size
            except Exception as e:
                print(f"Error setting doughnut hole size: {e}")

        # Enhanced Legend handling with accurate font configuration
        plugins = options.get('plugins', {})
        legend_opts = plugins.get('legend', {})
        
        # Fix legend display logic - properly handle explicit display settings
        legend_display = legend_opts.get('display')
        
        if legend_display is False:
            chart.has_legend = False
            # Also ensure series names don't show up anywhere else
            try:
                # Hide series names from chart title area
                if hasattr(chart, 'chart_title'):
                    chart.chart_title.has_text_frame = False
                
                # Make sure no series labels appear on the plot area
                for series in chart.series:
                    try:
                        series.name = ''
                    except (AttributeError, TypeError):
                        # Some series types don't allow name setting, ignore silently
                        pass
                    
            except Exception:
                # Ignore any errors in hiding series labels
                pass
        
        elif legend_display is True:
            chart.has_legend = True
            
            position = legend_opts.get('position', 'top')
            if position == 'bottom':
                chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            elif position == 'left':
                chart.legend.position = XL_LEGEND_POSITION.LEFT
            elif position == 'right':
                chart.legend.position = XL_LEGEND_POSITION.RIGHT
            else: # top
                chart.legend.position = XL_LEGEND_POSITION.TOP
            
            chart.legend.include_in_layout = False
            
            # Improved legend font configuration with error handling and scaling
            try:
                legend_labels = legend_opts.get('labels', {})
                legend_font = legend_labels.get('font', {})
                legend_color_str = legend_labels.get('color')
                
                if legend_font.get('size'):
                    # Apply font scaling factor to legend font size
                    original_size = int(legend_font['size'])
                    scaled_size = max(6, int(original_size * FONT_SCALE_FACTOR))
                    chart.legend.font.size = Pt(scaled_size)
                else:
                    chart.legend.font.size = Pt(int(10 * FONT_SCALE_FACTOR))  # Default size with scaling
                
                if legend_font.get('family'):
                    # Clean font family name and handle special fonts
                    font_family = legend_font['family'].strip()
                    if font_family in ['Meiryo UI', 'Meiryo']:
                        # Use a more compatible font for PowerPoint
                        chart.legend.font.name = 'Arial'
                    else:
                        chart.legend.font.name = font_family
                else:
                    chart.legend.font.name = 'Arial'
                
                if legend_color_str:
                    legend_color = parse_color(legend_color_str)
                    if legend_color:
                        chart.legend.font.color.rgb = legend_color
            except Exception as e:
                print(f"Error configuring legend font: {e}")
                # Set safe defaults with scaling
                chart.legend.font.size = Pt(int(10 * FONT_SCALE_FACTOR))
                chart.legend.font.name = 'Arial'
        else:
            # Default behavior when display is not specified
            datasets = data.get('datasets', [])
            if len(datasets) > 1:
                chart.has_legend = True
            else:
                chart.has_legend = False
                # Hide series names for single series without legend
                try:
                    for series in chart.series:
                        try:
                            series.name = ''
                        except (AttributeError, TypeError):
                            # Some series types don't allow name setting, ignore silently
                            pass
                except Exception:
                    # Ignore any errors in hiding single series labels
                    pass

        # Enhanced Data Labels with proper font settings from options and scaling
        datalabel_opts = plugins.get('datalabels', {})
        has_datalabels = datalabel_opts.get('display', False)
        
        if has_datalabels:
            plot = chart.plots[0]
            plot.has_data_labels = True
            data_labels = plot.data_labels
            
            # Get font settings from datalabels options - apply scaling to exact values from JSON
            font_opts = datalabel_opts.get('font', {})
            label_color_str = datalabel_opts.get('color', '#333333')
            label_color = parse_color(label_color_str)
            original_font_size = font_opts.get('size', 8)  # Use exact size from JSON
            # Apply font scaling factor to datalabels font size
            scaled_font_size = max(6, int(original_font_size * FONT_SCALE_FACTOR))
            font_family = font_opts.get('family', 'Arial')
            font_weight = font_opts.get('weight', 'normal')

            # Clean font family for data labels too
            if font_family in ['Meiryo UI', 'Meiryo']:
                font_family = 'Arial'

            # Apply font settings to data labels using scaled values
            data_labels.font.bold = font_weight in ['bold', '600', '700', '800', '900']
            data_labels.font.size = Pt(scaled_font_size)  # Use scaled font size
            data_labels.font.name = font_family
            if label_color:
                data_labels.font.color.rgb = label_color
            
            apply_datalabel_positioning(
                data_labels,
                datalabel_opts.get('anchor', 'center'),
                datalabel_opts.get('align', 'center'),
                datalabel_opts.get('offset', 0),
                chart_type_str
            )

            formatter_str = datalabel_opts.get('formatter')
            if formatter_str:
                data_labels.show_category_name = False
                data_labels.show_value = False
                data_labels.show_percentage = False
                
                # Apply formatting to individual points with consistent scaled font settings
                for series in chart.series:
                    for i, point in enumerate(series.points):
                        point.has_data_label = True
                        data_label = point.data_label
                        values = series.values
                        formatted_text = parse_datalabel_formatter(formatter_str, labels, values, i)
                        data_label.text_frame.text = formatted_text
                        
                        # Apply exact scaled font settings to each point's data label
                        if data_label.text_frame.paragraphs:
                            for paragraph in data_label.text_frame.paragraphs:
                                paragraph.font.size = Pt(scaled_font_size)  # Use scaled size
                                paragraph.font.bold = font_weight in ['bold', '600', '700', '800', '900']
                                paragraph.font.name = font_family
                                if label_color:
                                    paragraph.font.color.rgb = label_color
                        # Also apply to individual runs in paragraphs
                        for paragraph in data_label.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(scaled_font_size)  # Use scaled size
                                run.font.bold = font_weight in ['bold', '600', '700', '800', '900']
                                run.font.name = font_family
                                if label_color:
                                    run.font.color.rgb = label_color
            else:
                # For non-custom formatters, still apply the exact scaled font settings
                data_labels.show_category_name = False
                data_labels.show_value = True
                data_labels.show_percentage = False
        else:
            # Explicitly disable data labels when not configured
            try:
                plot = chart.plots[0]
                plot.has_data_labels = False
            except Exception:
                pass

        # Disable chart title to prevent series labels from appearing there
        try:
            if hasattr(chart, 'chart_title'):
                chart.chart_title.has_text_frame = False
            # Alternative method to disable title
            chart.has_title = False
        except Exception as e:
            print(f"Error disabling chart title: {e}")

        if chart_type_str != 'pie':
            scales = options.get('scales', {})
            is_horizontal_bar = chart_type == XL_CHART_TYPE.BAR_CLUSTERED
            
            # CORRECTED: For horizontal bars, x is value axis and y is category axis
            # For vertical charts, x is category axis and y is value axis
            if is_horizontal_bar:
                value_scale = scales.get('x', {})
                category_scale = scales.get('y', {})
                value_axis = chart.value_axis
                category_axis = chart.category_axis
            else:
                value_scale = scales.get('y', {})
                category_scale = scales.get('x', {})
                value_axis = chart.value_axis
                category_axis = chart.category_axis

            # Configure value axis
            try:
                # Handle special case for all-negative bar charts
                if all_negative and chart_type_str == 'bar':
                    # For all-negative data, remove max and use only min
                    val_min = value_scale.get('min')
                    if val_min is not None:
                        value_axis.minimum_scale = float(val_min)
                    
                    # Don't set maximum scale for all-negative charts
                    # Let PowerPoint auto-scale the maximum
                    
                    # Don't force beginAtZero for all-negative charts
                else:
                    # Normal handling for mixed or positive data
                    # Handle beginAtZero
                    if value_scale.get('beginAtZero', True):
                        value_axis.minimum_scale = 0.0
                    
                    # Handle min/max values
                    val_min = value_scale.get('min')
                    val_max = value_scale.get('max')
                    suggested_max = value_scale.get('suggestedMax')
                    
                    if val_min is not None:
                        value_axis.minimum_scale = float(val_min)
                    
                    if val_max is not None:
                        value_axis.maximum_scale = float(val_max)
                    elif suggested_max is not None:
                        value_axis.maximum_scale = float(suggested_max)
                
                # Handle step size
                ticks = value_scale.get('ticks', {})
                step_size = ticks.get('stepSize')
                if step_size:
                    value_axis.major_unit = float(step_size)
                
                # Handle value axis visibility and styling
                if value_scale.get('display', True) == False:
                    value_axis.visible = False
                    value_axis.has_major_gridlines = False
                else:
                    # Configure tick labels if axis is visible
                    if value_axis.tick_labels:
                        tick_font = ticks.get('font', {})
                        if tick_font.get('size'):
                            # Apply font scaling factor to value axis tick labels
                            original_tick_size = tick_font['size']
                            scaled_tick_size = max(6, int(original_tick_size * FONT_SCALE_FACTOR))
                            value_axis.tick_labels.font.size = Pt(scaled_tick_size)
                        
                        # Handle font family for axis labels
                        if tick_font.get('family'):
                            font_family = tick_font['family'].strip()
                            if font_family in ['Meiryo UI', 'Meiryo']:
                                value_axis.tick_labels.font.name = 'Arial'
                            else:
                                value_axis.tick_labels.font.name = font_family
                        
                        tick_color = parse_color(ticks.get('color', '#888888'))
                        if tick_color:
                            value_axis.tick_labels.font.color.rgb = tick_color
                
                # Hide axis line if border is not displayed
                if value_scale.get('border', {}).get('display') is False:
                    value_axis.format.line.fill.background()
                
                # Remove tick marks for both horizontal and vertical bar charts
                if chart_type in [XL_CHART_TYPE.BAR_CLUSTERED, XL_CHART_TYPE.COLUMN_CLUSTERED]:
                    try:
                        axis_element = value_axis._element
                        major_tick = axis_element.find(qn('c:majorTickMark'))
                        if major_tick is not None:
                            major_tick.set('val', 'none')
                        else:
                            SubElement(axis_element, 'c:majorTickMark', val='none')
                        
                        minor_tick = axis_element.find(qn('c:minorTickMark'))
                        if minor_tick is not None:
                            minor_tick.set('val', 'none')
                        else:
                            SubElement(axis_element, 'c:minorTickMark', val='none')
                    except Exception as tick_error:
                        print(f"Error removing value axis tick marks: {tick_error}")

            except Exception as e:
                print(f"Error configuring value axis: {e}")
            
            # Configure category axis
            try:
                cat_ticks = category_scale.get('ticks', {})
                
                # Category axis visibility
                if category_scale.get('display', True) == False:
                    category_axis.visible = False
                else:
                    category_axis.visible = True
                    
                    # Remove tick marks for both horizontal and vertical bar charts on category axis
                    if chart_type in [XL_CHART_TYPE.BAR_CLUSTERED, XL_CHART_TYPE.COLUMN_CLUSTERED]:
                        try:
                            axis_element = category_axis._element
                            major_tick = axis_element.find(qn('c:majorTickMark'))
                            if major_tick is not None:
                                major_tick.set('val', 'none')
                            else:
                                SubElement(axis_element, 'c:majorTickMark', val='none')
                        
                            minor_tick = axis_element.find(qn('c:minorTickMark'))
                            if minor_tick is not None:
                                minor_tick.set('val', 'none')
                            else:
                                SubElement(axis_element, 'c:minorTickMark', val='none')
                        except Exception as tick_error:
                            print(f"Error removing category axis tick marks: {tick_error}")
                    
                    if category_axis.tick_labels:
                        tick_font = cat_ticks.get('font', {})
                        if tick_font.get('size'):
                            original_cat_size = tick_font['size']
                            scaled_cat_size = max(6, int(original_cat_size * FONT_SCALE_FACTOR))
                            category_axis.tick_labels.font.size = Pt(scaled_cat_size)
                        else:
                            category_axis.tick_labels.font.size = Pt(int(8 * FONT_SCALE_FACTOR))
                        
                        if tick_font.get('family'):
                            font_family = tick_font['family'].strip()
                            if font_family in ['Meiryo UI', 'Meiryo']:
                                category_axis.tick_labels.font.name = 'Arial'
                            else:
                                category_axis.tick_labels.font.name = font_family
                        
                        tick_color = parse_color(cat_ticks.get('color', '#666666'))
                        if tick_color:
                            category_axis.tick_labels.font.color.rgb = tick_color
                        
                        # Handle label rotation
                        max_rotation = cat_ticks.get('maxRotation', 0)
                        min_rotation = cat_ticks.get('minRotation', 0)
                        if max_rotation == 0 and min_rotation == 0:
                            category_axis.tick_labels.orientation = 0
                        elif max_rotation > 0:
                            category_axis.tick_labels.orientation = max_rotation
                        
                        # Position labels for horizontal bars with negative values
                        if is_horizontal_bar:
                            try:
                                axis_element = category_axis._element
                                
                                if has_negative_values:
                                    tick_lbl_pos = axis_element.find(qn('c:tickLblPos'))
                                    if tick_lbl_pos is not None:
                                        tick_lbl_pos.set('val', 'low')
                                    else:
                                        SubElement(axis_element, 'c:tickLblPos', val='low')
                                    
                                    crosses = axis_element.find(qn('c:crosses'))
                                    if crosses is not None:
                                        crosses.set('val', 'autoZero')
                                    else:
                                        SubElement(axis_element, 'c:crosses', val='autoZero')
                                else:
                                    tick_lbl_pos = axis_element.find(qn('c:tickLblPos'))
                                    if tick_lbl_pos is not None:
                                        tick_lbl_pos.set('val', 'low')
                                    else:
                                        SubElement(axis_element, 'c:tickLblPos', val='low')
                            except Exception as tick_pos_error:
                                print(f"Error setting tick label position for horizontal bars: {tick_pos_error}")
                        
                        # Position labels for vertical bars (COLUMN_CLUSTERED) with negative values
                        elif chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED:
                            try:
                                axis_element = category_axis._element
                                
                                if has_negative_values:
                                    # Move labels to top when there are negative values
                                    tick_lbl_pos = axis_element.find(qn('c:tickLblPos'))
                                    if tick_lbl_pos is not None:
                                        tick_lbl_pos.set('val', 'high')
                                    else:
                                        SubElement(axis_element, 'c:tickLblPos', val='high')
                                    
                                    crosses = axis_element.find(qn('c:crosses'))
                                    if crosses is not None:
                                        crosses.set('val', 'autoZero')
                                    else:
                                        SubElement(axis_element, 'c:crosses', val='autoZero')
                                else:
                                    # Keep labels at bottom for all positive values
                                    tick_lbl_pos = axis_element.find(qn('c:tickLblPos'))
                                    if tick_lbl_pos is not None:
                                        tick_lbl_pos.set('val', 'low')
                                    else:
                                        SubElement(axis_element, 'c:tickLblPos', val='low')
                            except Exception as tick_pos_error:
                                print(f"Error setting tick label position for vertical bars: {tick_pos_error}")
                
                # Handle axis line display
                border_config = category_scale.get('border', {})
                if border_config.get('display') is False:
                    category_axis.format.line.fill.background()
                elif border_config.get('display') is True:
                    border_color = parse_color(border_config.get('color', '#666666'))
                    if border_color:
                        category_axis.format.line.color.rgb = border_color
                        
            except Exception as e:
                print(f"Error configuring category axis: {e}")

            # Configure gridlines (corrected for horizontal bars)
            try:
                if is_horizontal_bar:
                    # For horizontal bars: x-axis (value) has gridlines, y-axis (category) typically doesn't
                    value_grid = value_scale.get('grid', {})
                    category_grid = category_scale.get('grid', {})
                else:
                    # For vertical charts: y-axis (value) has gridlines, x-axis (category) typically doesn't
                    value_grid = value_scale.get('grid', {})
                    category_grid = category_scale.get('grid', {})

                # Value axis gridlines
                if value_grid.get('display', False):
                    value_axis.has_major_gridlines = True
                    grid_color_str = value_grid.get('color')
                    if grid_color_str:
                        grid_color = parse_color(grid_color_str)
                        if grid_color:
                            value_axis.major_gridlines.format.line.color.rgb = grid_color
                else:
                    value_axis.has_major_gridlines = False

                # Category axis gridlines
                if category_grid.get('display', False):
                    category_axis.has_major_gridlines = True
                    grid_color_str = category_grid.get('color')
                    if grid_color_str:
                        grid_color = parse_color(grid_color_str)
                        if grid_color:
                            category_axis.major_gridlines.format.line.color.rgb = grid_color
                else:
                    category_axis.has_major_gridlines = False
                        
            except Exception as e:
                print(f"Error configuring gridlines: {e}")
            
            # Configure bar gap and overlap for better bar width control
            if chart_type in [XL_CHART_TYPE.BAR_CLUSTERED, XL_CHART_TYPE.COLUMN_CLUSTERED]:
                try:
                    plot = chart.plots[0]
                    
                    # Extract categoryPercentage and barPercentage from datasets
                    datasets = data.get('datasets', [])
                    category_percentage = 0.8  # Chart.js default
                    
                    if datasets and len(datasets) > 0:
                        # Get from first dataset (typically all datasets share these values)
                        category_percentage = datasets[0].get('categoryPercentage', 0.8)
                    
                    # Convert Chart.js percentages to PowerPoint gap_width
                    # Chart.js categoryPercentage controls the space used by the category group
                    # gap_width in PowerPoint is the percentage of space BETWEEN categories
                    # Formula: gap_width = (1 - categoryPercentage) / categoryPercentage * 100
                    if category_percentage > 0:
                        gap_width = int((1 - category_percentage) / category_percentage * 100)
                        # Clamp to reasonable range (PowerPoint accepts 0-500)
                        gap_width = max(0, min(500, gap_width))
                        plot.gap_width = gap_width
                    
                    # Convert barPercentage to overlap for clustered charts
                    # For single series, overlap doesn't matter
                    # For multiple series, barPercentage controls bar width within the category group
                    if len(datasets) > 1:
                        # overlap in PowerPoint: negative = gap, 0 = touching, positive = overlapping
                        # We'll set to 0 for now as Chart.js barPercentage is more about individual bar width
                        plot.overlap = 0
                    else:
                        plot.overlap = 0
                        
                except Exception as e:
                    print(f"Error configuring bar width: {e}")

        # Enhanced color application for series/points
        try:
            datasets_original = data.get('datasets', [])  # Use original order for colors
            for i, series in enumerate(chart.series):
                if i < len(datasets_original):
                    dataset = datasets_original[i]
                    
                    # Line/Area chart colors
                    if chart_type in [XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS, XL_CHART_TYPE.AREA, XL_CHART_TYPE.AREA_STACKED]:
                        border_color = parse_color(dataset.get('borderColor'))
                        if border_color:
                            series.format.line.color.rgb = border_color
                        
                        # Enhanced area fill handling based on dataset configuration
                        if chart_type in [XL_CHART_TYPE.AREA, XL_CHART_TYPE.AREA_STACKED]:
                            bg_color_str = dataset.get('backgroundColor', '')
                            
                            # Check if backgroundColor is transparent or not set for area fill
                            if bg_color_str and bg_color_str.lower() not in ['transparent', 'none']:
                                bg_color = parse_color(bg_color_str)
                                if bg_color:
                                    series.format.fill.solid()
                                    series.format.fill.fore_color.rgb = bg_color
                            else:
                                # If backgroundColor is transparent or not set, don't fill
                                series.format.fill.background()

                    # Bar/Column/Pie/Doughnut chart colors
                    elif chart_type in [XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.BAR_CLUSTERED, XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT]:
                        background_colors = dataset.get('backgroundColor', [])
                        
                        # Reverse colors for horizontal bars to match reversed data
                        if is_horizontal_bar:
                            background_colors = list(reversed(background_colors))
                        
                        for j, point in enumerate(series.points):
                            if j < len(background_colors):
                                color = parse_color(background_colors[j])
                                if color:
                                    point.format.fill.solid()
                                    point.format.fill.fore_color.rgb = color
                            
                            # Disable "invert if negative" for bar/column charts to prevent automatic color inversion
                            if chart_type in [XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.BAR_CLUSTERED]:
                                try:
                                    _ = SubElement(point.format.element, 'c:invertIfNegative', val='0')
                                except Exception as e:
                                    print(f"Error disabling invertIfNegative for point {j}: {e}")
        except Exception as e:
            print(f"Error applying series colors: {e}")

    except Exception as e:
        print(f"Failed to add chart: {e}")
        return
    

def add_shape_element(slide, element, slide_width, slide_height):
    """Enhanced shape rendering with proper text support and expanded clip-path mapping"""
    shape_info = element.get('shapeInfo')
    if not shape_info:
        return

    rect = shape_info.get('rect', element)
    x = rect.get('x', 0)
    y = rect.get('y', 0)
    width = max(1, rect.get('width', 10))
    height = max(1, rect.get('height', 10))

    styles = shape_info.get('styles', element.get('styles', {}))
    rotation = shape_info.get('rotation', 0)
    clip_path = shape_info.get('clipPath', '')
    bg_color = parse_color(styles.get('backgroundColor'))
    text_content = shape_info.get('text', element.get('text', '')).strip()

    shape_type = MSO_SHAPE.RECTANGLE # Default
    
    # Enhanced clip-path mapping to PowerPoint shapes
    if 'polygon' in clip_path:
        # Normalize the clip-path string for easier pattern matching
        normalized_clip = clip_path.lower().replace(' ', '').replace('polygon(', '').replace(')', '')
        
        # Triangle patterns - Downward facing triangle (like triangle-arrow-4)
        # Pattern: polygon(50% 100%, 0 0, 100% 0) - peak at bottom, base at top
        if ('50%100%' in normalized_clip and '00' in normalized_clip and '100%0' in normalized_clip) or \
           ('50%100%' in normalized_clip and '0%0%' in normalized_clip and '100%0%' in normalized_clip):
            shape_type = MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE
            rotation = 180  # Rotate to point downward
        
        # Triangle patterns - Right facing triangle (like line-arrow)
        # Pattern: polygon(100% 50%, 32% 1%, 32% 99%) - peak at right, base vertical on left
        elif ('100%50%' in normalized_clip and '32%1%' in normalized_clip and '32%99%' in normalized_clip) or \
             ('100%50%' in normalized_clip and any(x in normalized_clip for x in ['30%0%', '30%100%', '35%0%', '35%100%'])):
            shape_type = MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE
            rotation = 90  # Rotate to point right
        
        # Enhanced Chevron patterns
        # Original chevron: polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)
        elif ('50%0%' in normalized_clip and '100%50%' in normalized_clip and '50%100%' in normalized_clip and '0%50%' in normalized_clip):
            shape_type = MSO_AUTO_SHAPE_TYPE.CHEVRON
        
        # Complex chevron: polygon(75% 0%, 100% 50%, 75% 100%, 21% 100%, 54% 50%, 19% 0)
        elif ('75%0%' in normalized_clip and '100%50%' in normalized_clip and '75%100%' in normalized_clip and 
              '21%100%' in normalized_clip and '54%50%' in normalized_clip and '19%0' in normalized_clip):
            shape_type = MSO_AUTO_SHAPE_TYPE.CHEVRON
        
        # Right arrow patterns
        # Pattern: polygon(100% 50%, 75% 90%, ...) - arrow pointing right
        elif '100%50%' in normalized_clip and '75%90%' in normalized_clip:
            shape_type = MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW
        
        # Left arrow patterns  
        # Pattern: polygon(0% 50%, 25% 90%, ...) - arrow pointing left
        elif '0%50%' in normalized_clip and '25%90%' in normalized_clip:
            shape_type = MSO_AUTO_SHAPE_TYPE.LEFT_ARROW
        
        # Up arrow patterns
        # Pattern: polygon(50% 0%, 90% 25%, ...) - arrow pointing up
        elif '50%0%' in normalized_clip and '90%25%' in normalized_clip:
            shape_type = MSO_AUTO_SHAPE_TYPE.UP_ARROW
        
        # Down arrow patterns
        # Pattern: polygon(50% 100%, 90% 75%, ...) - arrow pointing down
        elif '50%100%' in normalized_clip and '90%75%' in normalized_clip:
            shape_type = MSO_AUTO_SHAPE_TYPE.DOWN_ARROW
        
        # Additional triangle orientations
        # Left-facing triangle: polygon(0% 50%, 100% 0%, 100% 100%)
        elif ('0%50%' in normalized_clip and '100%0%' in normalized_clip and '100%100%' in normalized_clip):
            shape_type = MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE
            rotation = 270  # Rotate to point left
        
        # Up-facing triangle: polygon(50% 0%, 0% 100%, 100% 100%)
        elif ('50%0%' in normalized_clip and '0%100%' in normalized_clip and '100%100%' in normalized_clip):
            shape_type = MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE
            rotation = 0  # Default upward orientation
        
        # Hexagon pattern: polygon(30% 0%, 70% 0%, 100% 50%, 70% 100%, 30% 100%, 0% 50%)
        elif ('30%0%' in normalized_clip and '70%0%' in normalized_clip and '100%50%' in normalized_clip and 
              '70%100%' in normalized_clip and '30%100%' in normalized_clip and '0%50%' in normalized_clip):
            shape_type = MSO_AUTO_SHAPE_TYPE.HEXAGON
        
        # Pentagon pattern: polygon(50% 0%, 100% 38%, 82% 100%, 18% 100%, 0% 38%)
        elif ('50%0%' in normalized_clip and '100%38%' in normalized_clip and '82%100%' in normalized_clip and 
              '18%100%' in normalized_clip and '0%38%' in normalized_clip):
            shape_type = MSO_AUTO_SHAPE_TYPE.PENTAGON
        
        # Octagon pattern: polygon(30% 0%, 70% 0%, 100% 30%, 100% 70%, 70% 100%, 30% 100%, 0% 70%, 0% 30%)
        elif ('30%0%' in normalized_clip and '70%0%' in normalized_clip and '100%30%' in normalized_clip and 
              '100%70%' in normalized_clip and '70%100%' in normalized_clip and '30%100%' in normalized_clip and 
              '0%70%' in normalized_clip and '0%30%' in normalized_clip):
            shape_type = MSO_AUTO_SHAPE_TYPE.OCTAGON
        
        # Star patterns: polygon(50% 0%, 61% 35%, 98% 35%, 68% 57%, 79% 91%, 50% 70%, 21% 91%, 32% 57%, 2% 35%, 39% 35%)
        elif ('50%0%' in normalized_clip and '61%35%' in normalized_clip and '98%35%' in normalized_clip and 
              '68%57%' in normalized_clip and '79%91%' in normalized_clip and '50%70%' in normalized_clip):
            shape_type = MSO_AUTO_SHAPE_TYPE.STAR_5_POINTED
        
        # Diamond/Rhombus pattern: polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)
        elif ('50%0%' in normalized_clip and '100%50%' in normalized_clip and '50%100%' in normalized_clip and 
              '0%50%' in normalized_clip and normalized_clip.count('%') == 8):  # Exactly 4 points
            shape_type = MSO_AUTO_SHAPE_TYPE.DIAMOND
        
        # Parallelogram pattern: polygon(25% 0%, 100% 0%, 75% 100%, 0% 100%)
        elif ('25%0%' in normalized_clip and '100%0%' in normalized_clip and '75%100%' in normalized_clip and 
              '0%100%' in normalized_clip):
            shape_type = MSO_AUTO_SHAPE_TYPE.PARALLELOGRAM
        
        # Trapezoid pattern: polygon(20% 0%, 80% 0%, 100% 100%, 0% 100%)
        elif ('20%0%' in normalized_clip and '80%0%' in normalized_clip and '100%100%' in normalized_clip and 
              '0%100%' in normalized_clip):
            shape_type = MSO_AUTO_SHAPE_TYPE.TRAPEZOID

    try:
        shape = slide.shapes.add_shape(
            shape_type,
            pixels_to_emu(x), pixels_to_emu(y),
            pixels_to_emu(width), pixels_to_emu(height)
        )

        if bg_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = bg_color
        else:
            shape.fill.background()
        
        shape.line.fill.background()
        shape.shadow.inherit = False
        
        # Apply rotation if needed
        if rotation != 0:
            shape.rotation = rotation

        # Add text to the shape if present
        if text_content:
            text_frame = shape.text_frame
            text_frame.clear()
            text_frame.word_wrap = True
            
            # Set text alignment based on styles
            text_align = styles.get('textAlign', 'left')
            justify_content = styles.get('justifyContent', 'left')
            align_items = styles.get('alignItems', 'center')
            
            # Determine vertical alignment
            if align_items == 'center':
                text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            elif align_items == 'flex-start' or align_items == 'start':
                text_frame.vertical_anchor = MSO_ANCHOR.TOP
            elif align_items == 'flex-end' or align_items == 'end':
                text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
            else:
                text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

            # Apply padding from styles
            padding_left = safe_float(styles.get('paddingLeft', '0px'))
            padding_right = safe_float(styles.get('paddingRight', '0px'))
            padding_top = safe_float(styles.get('paddingTop', '0px'))
            padding_bottom = safe_float(styles.get('paddingBottom', '0px'))
            
            text_frame.margin_left = pixels_to_emu(padding_left)
            text_frame.margin_right = pixels_to_emu(padding_right)
            text_frame.margin_top = pixels_to_emu(padding_top)
            text_frame.margin_bottom = pixels_to_emu(padding_bottom)

            p = text_frame.paragraphs[0]
            run = p.add_run()
            run.text = text_content

            # Apply text formatting from styles
            font = run.font
            font_size_px = safe_float(styles.get('fontSize', '12').replace('px', ''))
            font.size = Pt(max(6, get_font_size_pt(font_size_px)))
            
            font_family = styles.get('fontFamily', 'Arial')
            if font_family:
                # Clean font family string and handle special fonts
                font_family = font_family.split(',')[0].strip('"\'')
                font.name = font_family
            
            font.bold = styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
            font.italic = styles.get('fontStyle', 'normal') == 'italic'
            
            # Apply text color
            text_color = parse_color(styles.get('color', 'black'))
            if text_color:
                font.color.rgb = text_color

            # Apply text alignment
            text_align = styles.get('textAlign', 'left')
            if text_align == 'center' or justify_content == 'center':
                p.alignment = PP_ALIGN.CENTER
            elif text_align == 'right' or justify_content == 'flex-end':
                p.alignment = PP_ALIGN.RIGHT
            else:
                p.alignment = PP_ALIGN.LEFT

    except Exception as e:
        print(f"Failed to add shape element: {e}")

def create_pptx_from_json(json_path, output_path=None):
    """Enhanced PowerPoint generation with precise positioning"""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            slides_data = json.load(f)
    except Exception as e:
        print(f"Error reading JSON file: {e}")
        return
   
    if not slides_data:
        print("No slides found in JSON")
        return
   
    # Get slide dimensions from first slide
    first_slide = slides_data[0]
    slide_width = safe_int(first_slide.get('slideWidth', 1039))
    slide_height = safe_int(first_slide.get('slideHeight', 585))
   
    # Create presentation with precise dimensions
    prs = Presentation()
    prs.slide_width = pixels_to_emu(slide_width)
    prs.slide_height = pixels_to_emu(slide_height)
   
    for slide_data in slides_data:
        slide_layout = prs.slide_layouts[6] # Blank layout
        slide = prs.slides.add_slide(slide_layout)
       
        # Set slide background
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = parse_color('#ffffff') or RGBColor(255, 255, 255)
       
        # Add slide background styling
        slide_styles = slide_data.get('slideStyles', {})
        if slide_styles:
            add_bg_shape(slide, slide_styles, 0, 0, slide_width, slide_height)
       
        elements = slide_data.get('elements', [])
       
        # Sort elements primarily by z-index, then by y, x for consistent layering
        elements_sorted = sorted(elements, key=lambda e: (e.get('zIndex', 0), e.get('y', 0), e.get('x', 0)))
       
        # Track processed elements to prevent duplicates
        processed_element_ids = set()
        processed_text_content = set()
        
        # Store table information for overlay positioning
        table_positions = {}
       
        # Process each element
        for element in elements_sorted:
            element_id = f"{element.get('type', 'unknown')}-{element.get('x', 0)}-{element.get('y', 0)}-{element.get('className', '')}"
            if element_id in processed_element_ids:
                continue
            processed_element_ids.add(element_id)
            
            element_type = element.get('type')
            element_class = element.get('className', '')
            parent_has_shadow = False

            # Handle complex shapes first (includes arrows) - prevent text duplication
            if element.get('shapeInfo'):
                shape_text = element.get('shapeInfo', {}).get('text', element.get('text', '')).strip()
                shape_class = element.get('className', '')
                
                # Create a unique identifier for this text content + class + position
                text_class_id = f"{shape_text}-{shape_class}-{element.get('x', 0)}-{element.get('y', 0)}"
                
                # Skip if this exact text with same class at similar position was already processed
                if shape_text and text_class_id in processed_text_content:
                    continue
                
                # Generic container/child duplicate detection
                if 'container' in shape_class:
                    # Check if there's a corresponding child element with the same text at similar position
                    has_child_element = any(
                        other_elem.get('shapeInfo', {}).get('text', '').strip() == shape_text and
                        other_elem.get('className', '') != shape_class and
                        'container' not in other_elem.get('className', '') and
                        abs(other_elem.get('x', 0) - element.get('x', 0)) < 50 and
                        abs(other_elem.get('y', 0) - element.get('y', 0)) < 50
                        for other_elem in elements_sorted
                        if other_elem.get('shapeInfo')
                    )
                    
                    if has_child_element:
                        continue
                
                # Add to processed text content
                if shape_text:
                    processed_text_content.add(text_class_id)
                
                add_shape_element(slide, element, slide_width, slide_height)
                continue

            if element.get('footerElements'):
                # Process individual footer elements with space-between positioning
                footer_elements = element.get('footerElements', [])
                for footer_elem in footer_elements:
                    if footer_elem.get('type') == 'img':
                        add_image_element(slide, footer_elem, slide_width, slide_height, parent_has_shadow)
                    elif footer_elem.get('type') in ['span', 'div']:
                        add_text_element(slide, footer_elem, slide_width, slide_height, parent_has_shadow)
                continue

            if element_type == 'overlay':
                # Pass table positions to overlay processing
                element['_table_positions'] = table_positions
                add_overlay_element(slide, element, slide_width, slide_height)
            elif element.get('inlineGroup'):
                add_inline_group_element(slide, element, slide_width, slide_height, parent_has_shadow)
            elif element_type in ['ul', 'ol']:
                add_list_element(slide, element, slide_width, slide_height, parent_has_shadow)
            elif element_type == 'table':
                # Store table position for overlay calculations
                table_key = f"table_{element.get('x', 0)}_{element.get('y', 0)}"
                table_positions[table_key] = {
                    'x': element.get('x', 0),
                    'y': element.get('y', 0),
                    'table_info': element.get('tableInfo', {})
                }
                
                # Add the table
                add_table_element(slide, element, slide_width, slide_height, parent_has_shadow)
            elif element_type == 'img':
                add_image_element(slide, element, slide_width, slide_height, parent_has_shadow)
            elif element_type == 'a':
                add_link_element(slide, element, slide_width, slide_height, parent_has_shadow)
            elif element_type == 'canvas':
                add_chart_element(slide, element, slide_width, slide_height)
            elif element_type == 'span':
                add_text_element(slide, element, slide_width, slide_height, parent_has_shadow)
            elif element_type == 'pseudo':
                add_pseudo_element(slide, element, slide_width, slide_height)
            elif element_type == 'div' and 'chart' in element.get('className', '') and element.get('chartConfig'):
                add_chart_element(slide, element, slide_width, slide_height)
            elif element_type in ['div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']:  # Removed ul, ol, li
                # Prevent duplication for div elements
                if (element.get('text', '').strip() or
                    has_any_border(element.get('styles', {})) or
                    parse_color(element.get('styles', {}).get('backgroundColor')) or
                    element.get('styles', {}).get('boxShadow', 'none') != 'none'):
                    
                    # Skip if this is a div that only contains processed inline elements
                    if (element.get('text', '').strip() and not element.get('inlineGroup')):
                        add_text_element(slide, element, slide_width, slide_height, parent_has_shadow)
                    elif not element.get('text', '').strip():
                        # Check if it's a structural div that should be rendered
                        if not element.get('footerElements'):
                            x = element.get('x', 0)
                            y = element.get('y', 0)
                            width = max(1, element.get('width', 100))
                            height = max(1, element.get('height', 100))
                            add_bg_shape(slide, element.get('styles', {}), x, y, width, height)

    if output_path is None:
        base_name = os.path.splitext(os.path.basename(json_path))[0]
        output_path = f"{base_name}_output.pptx"
    try:
        prs.save(output_path)
        print(f"Presentation saved as '{output_path}' with {len(slides_data)} slide(s)")
        print(f"Slide dimensions: {slide_width}x{slide_height} pixels")
    except Exception as e:
        print(f"Error saving presentation: {e}")

if __name__ == "__main__":
    create_pptx_from_json('slides_data.json', 'output.pptx')