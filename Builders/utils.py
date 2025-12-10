import re
import os
import base64
import math
import requests
from io import BytesIO
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image, ImageDraw
from pptx.oxml.xmlchemy import OxmlElement

FONT_SCALE_FACTOR = 0.733275

# HTML tags that represent text content in table cells
TEXT_HTML_TAGS = ['p', 'span', 'a', 'em', 'strong', 'b', 'i', 'u', 'br', 'ol', 'ul', 'li']

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
    return px * FONT_SCALE_FACTOR

def get_font_size_pt(font_size_px):
    """
    Convert a font size in pixels to a point size suitable for pptx.
    Applies a global FONT_SCALE_FACTOR and clamps to a sensible minimum to
    avoid extremely small fonts.
    """
    if font_size_px <= 0:
        return 12
    # Apply scaling factor to reduce font size slightly for better HTML matching
    scaled_size = max(6, px_to_pt(font_size_px)  )
    return scaled_size

def get_font_pt_to_px(pt):
    """Convert pptx font size (pt) back to px"""
    return pt / FONT_SCALE_FACTOR


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

def is_bold(font_weight):
    """Check if font weight indicates bold text"""
    if font_weight is None:
        return False
    if isinstance(font_weight, bool):
        return font_weight
    if isinstance(font_weight, (int, float)):
        return font_weight >= 600
    weight_str = str(font_weight).lower().strip()
    return weight_str in ['bold', '600', '700', '800', '900']

def get_para_alignment(text_align):
    """Convert CSS text-align to PowerPoint paragraph alignment."""
    if text_align == 'center':
        return PP_ALIGN.CENTER
    elif text_align == 'right':
        return PP_ALIGN.RIGHT
    elif text_align == 'justify':
        return PP_ALIGN.JUSTIFY
    return PP_ALIGN.LEFT

def get_vertical_alignment(align_items):
    """Convert CSS align-items to PowerPoint vertical anchor."""
    if align_items == 'center':
        return MSO_ANCHOR.MIDDLE
    elif align_items in ['flex-start', 'start', 'top']:
        return MSO_ANCHOR.TOP
    elif align_items in ['flex-end', 'end', 'bottom']:
        return MSO_ANCHOR.BOTTOM
    return MSO_ANCHOR.MIDDLE

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

def make_rounded_image(input_source, radius):
    """Create a rounded image from input source (file path or BytesIO).
    
    Args:
        input_source: File path string or BytesIO object
        radius: Corner radius in pixels
    
    Returns:
        BytesIO: Rounded image as BytesIO stream
    """
    if isinstance(input_source, BytesIO):
        input_source.seek(0)
    im = Image.open(input_source).convert("RGBA")
    mask = Image.new("L", im.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle((0, 0) + im.size, radius=radius, fill=255)
    im.putalpha(mask)
    output = BytesIO()
    im.save(output, "PNG")
    output.seek(0)
    return output        

def parse_linear_gradient(gradient_str, total_width, reference_id=0):
    """Parse linear-gradient color stops (hex, rgb, rgba) into solid segments.

    Supports patterns like:
      linear-gradient(to right, rgb(37,29,95) 0%, rgb(37,29,95) 12%, rgb(217,217,217) 12%, rgb(217,217,217) 100%)
      linear-gradient(to right, #251d5f 0%, #251d5f 12%, #D9D9D9 12%, #D9D9D9 100%)
      
    Args:
        gradient_str (str): The CSS linear-gradient string to parse.
        total_width (float): Total width in pixels to compute segment widths.
        reference_id (int|str): Identifier for diagnostic context.
        
    Returns:
        list: List of (x, width, RGBColor) tuples representing gradient segments.
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
        print(f"[{reference_id}] Gradient parse failed: {e}")
        return []

def fetch_image(image_url, reference_id=0):
    """Fetch image bytes from URL.
    
    Args:
        image_url (str): URL or data URL of the image.
        reference_id (int|str): Identifier for diagnostic context.
    
    Returns:
        BytesIO: Image bytes or None on failure.
    """
    if not image_url:
        return None
    
    try:
        if image_url.startswith('data:'):
            _, data = image_url.split(',', 1)
            img_data = base64.b64decode(data)
            return BytesIO(img_data)
        elif image_url.startswith('http'):
            response = requests.get(image_url, timeout=10)
            if response.status_code == 200:
                return BytesIO(response.content)
        elif os.path.exists(image_url):
            with open(image_url, 'rb') as f:
                return BytesIO(f.read())
    except Exception as e:
        print(f"[{reference_id}] Failed to fetch image: {image_url} - {e}")
    return None
