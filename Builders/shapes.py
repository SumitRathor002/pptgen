import re
import os
import base64
import math
import requests
from io import BytesIO
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE, MSO_COLOR_TYPE
import requests
import os
import re
import math
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import os

from Builders.utils import pixels_to_emu, parse_color, safe_float, get_font_size_pt, is_bold, get_para_alignment, get_vertical_alignment, parse_border_radius, is_uniform_border, has_any_border, get_border_info, parse_linear_gradient, make_rounded_image, px_to_pt
from Builders.logger import logger

def create_precise_border_shapes(slide, x, y, width, height, border_info, reference_id=0):
    """Create precise rectangle shapes to represent individual borders.

    This function draws separate rectangle shapes for each side that requires a
    border (top, right, bottom, left). It uses pixel coordinates converted to
    EMU using pixels_to_emu and applies color and dash styles where provided.

    Args:
        slide: pptx Slide object where shapes will be added.
        x (float): X coordinate (pixels) of the target bounding box.
        y (float): Y coordinate (pixels) of the target bounding box.
        width (float): Width (pixels) of the target bounding box.
        height (float): Height (pixels) of the target bounding box.
        border_info (dict): Mapping of sides to border metadata. Each entry must
            contain keys: 'has_border', 'width', 'color', 'style'.
        reference_id (int|str): Identifier for diagnostic context.

    Returns:
        list: List of pptx shape objects that were created for borders.
    """
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
            logger.error(f"Error creating {side} border: {e}", reference_id=reference_id)
    return shapes_created


def add_bg_shape(slide, styles, x, y, width, height, reference_id=0):
    """Add a background shape to a slide with support for radius, shadow and borders.

    This function creates the primary background shape for a visual element and
    handles:
      - background color or transparent fill
      - rounded corners (borderRadius)
      - uniform borders and non-uniform per-side borders
      - box shadow application

    Args:
        slide: pptx Slide object where shapes will be added.
        styles (dict): CSS-like style dictionary used to derive visuals.
        x (float): X coordinate (pixels) for the shape.
        y (float): Y coordinate (pixels) for the shape.
        width (float): Width (pixels) for the shape.
        height (float): Height (pixels) for the shape.
        reference_id (int|str): Identifier for diagnostic context.

    Returns:
        list: List of pptx shape objects created for the background and any border shapes.
    """
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
            logger.error(f"Error adding bg shape: {e}", reference_id=reference_id)
    # Always handle non-uniform borders for all sides
    if has_any_border_sides and not has_uniform_border:
        border_info = get_border_info(styles)
        border_shapes = create_precise_border_shapes(
            slide, x, y, width, height, border_info,
            reference_id=reference_id
        )
        shapes_created.extend(border_shapes)
    return shapes_created


def apply_shadow(shape, box_shadow_str):
    """Apply a shadow to a pptx shape from a CSS-like box-shadow string.

    The supported box-shadow format is a simplified space-separated list:
      offsetX offsetY blurRadius [spread] [color]
    Units are expected to be pixels or numeric strings convertible by safe_float.
    The function maps parsed values to pptx shadow properties and sets color/transparency.

    Args:
        shape: pptx shape object whose shadow will be modified.
        box_shadow_str (str): CSS-like box-shadow string or 'none'.

    Returns:
        None

    Notes:
        - If parsing fails or color cannot be resolved, no shadow is applied.
        - Alpha in rgba(...) is supported; expected range is 0-255 in current parser.
    """
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


def add_image_element(slide, element, slide_width, slide_height, parent_has_shadow=False, img_dict={}, reference_id=0):
    """Add an image element to the slide with support for radius, border and shadow.

    This function:
      - Loads image data from data URLs, HTTP, or local filesystem.
      - Ensures image is positioned within slide bounds.
      - Optionally rounds corners by creating a rounded image buffer.
      - Adds the picture and, if required, adds a transparent shape on top to
        represent border and/or shadow (and then reorders shapes to sit behind).

    Args:
        slide: pptx Slide object where the image will be added.
        element (dict): Element metadata containing mediaInfo, styles, x, y, width, height.
        slide_width (int): Width of the slide in pixels.
        slide_height (int): Height of the slide in pixels.
        parent_has_shadow (bool): If parent already applies shadow, child shadow is skipped.
        img_dict (dict): Optional cache mapping image src to BytesIO or path.
        reference_id (int|str): Identifier for diagnostic context.

    Returns:
        None
    """
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

    temp_path = img_dict[img_src] if img_dict.get(img_src) else None
    if not temp_path:
        if img_src.startswith('data:'):
            _, data = img_src.split(',', 1)
            img_data = base64.b64decode(data)
            temp_path = BytesIO(img_data)
            
        elif img_src.startswith('http'):
            response = requests.get(img_src, timeout=10)
            if response.status_code == 200:
                temp_path = BytesIO(response.content)

        elif os.path.exists(img_src):
            with open(img_src, 'rb') as f:
                temp_path = BytesIO(f.read())
   
    if not temp_path:
        logger.error(f"Image not found: {img_src}", reference_id=reference_id)
        return
       
    image_to_add = temp_path
    if has_radius:
        scale_x = natural_width / width if width > 0 else 1
        radius_natural = int(radius_display * scale_x)
        temp_rounded = make_rounded_image(temp_path, radius_natural)
        image_to_add = temp_rounded
   
    # Add image with precise positioning
    picture = slide.shapes.add_picture(
        image_to_add,
        pixels_to_emu(x), pixels_to_emu(y),
        pixels_to_emu(width), pixels_to_emu(height)
    )
    picture.shadow.inherit = False
   
    # Handle borders and shadows
    if not any([has_border, has_radius, has_shadow]):
        return

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


def add_pseudo_element(slide, element, slide_width, slide_height, reference_id=0):
    """Render CSS pseudo-elements (::before / ::after) as PPT shapes or textboxes.

    This function supports:
      - Rendering backgrounds including linear gradients (split into segments).
      - Rendering simple text content for list-item bullets (::before) and other pseudo text.
      - Ensuring positioning remains within slide bounds and applying formatting.

    Args:
        slide: pptx Slide object where pseudo element visuals will be added.
        element (dict): Element metadata including coordinates, styles, pseudoType and text.
        slide_width (int): Width of the slide in pixels.
        slide_height (int): Height of the slide in pixels.
        reference_id (int|str): Identifier for diagnostic context.

    Returns:
        None
    """
    x = safe_float(element.get('x', 0))
    y = safe_float(element.get('y', 0))
    width = max(1, safe_float(element.get('width', 0)))
    height = max(1, safe_float(element.get('height', 0)))
    
    styles = element.get('styles', {})
    text = element.get('text', '')
    pseudo_type = element.get('pseudoType', '')
    parent_tag = element.get('parentTagName', '')
    x = max(0, min(x, slide_width - width))
    y = max(0, min(y, slide_height - height))
    
    background_full = styles.get('background', '')
    has_gradient = isinstance(background_full, str) and 'linear-gradient' in background_full

    if has_gradient:
        segments = parse_linear_gradient(background_full, width, reference_id)
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
                        logger.error(f"[{reference_id}] Error creating gradient segment: {e}", reference_id=reference_id)
        else:
            # Fallback to solid color
            add_bg_shape(slide, styles, x, y, width, height, reference_id=reference_id)
    else:
        # Standard pseudo element rendering - ensure proper background color handling
        bg_color = parse_color(styles.get('backgroundColor'))
        if bg_color or has_gradient:
            add_bg_shape(slide, styles, x, y, width, height, reference_id=reference_id)

    if (
        pseudo_type == '::before'
        and parent_tag == 'li'
        and text
        and 'parentText' in element
        and element['parentText'].strip()
    ):
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
        font_bullet.bold = is_bold(styles.get('fontWeight'))
        font_bullet.italic = styles.get('fontStyle') == 'italic'
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
        font_text.bold = is_bold(parent_styles.get('fontWeight'))
        font_text.italic = parent_styles.get('fontStyle', 'normal') == 'italic'
        color2 = parse_color(parent_styles.get('color', styles.get('color', 'black')))
        if color2:
            font_text.color.rgb = color2
        text_align = parent_styles.get('textAlign', styles.get('textAlign'))
        p.alignment = get_para_alignment(text_align)
        textbox.fill.background()
        textbox.line.fill.background()
        textbox.shadow.inherit = False
        return

    # Add text content for pseudo elements (usually none for decorative elements)
    if not text or text in ['""', "''", 'none']:
        return

    textbox = slide.shapes.add_textbox(
        pixels_to_emu(x), pixels_to_emu(y),
        pixels_to_emu(width), pixels_to_emu(height)
    )
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    # Precise margin handling
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
    font.bold = is_bold(styles.get('fontWeight'))
    font.italic = styles.get('fontStyle') == 'italic'
    
    color = parse_color(styles.get('color', 'black'))
    if color:
        font.color.rgb = color
    
    text_align = styles.get('textAlign', 'left')
    p.alignment = get_para_alignment(text_align)
    
    # Make textbox transparent
    textbox.fill.background()
    textbox.line.fill.background()
    textbox.shadow.inherit = False


def get_clip_path_shape(clip_path, rotation):
    """Map a CSS clip-path polygon to an equivalent PPT auto-shape type and rotation.

    This helper examines normalized polygon coordinates (as a string) and attempts
    to select a corresponding MSO_AUTO_SHAPE_TYPE. For some asymmetric shapes,
    a rotation angle is returned to orient the PPT shape similarly.

    Args:
        clip_path (str): The CSS clip-path string (expects polygon(...) style).
        rotation (int|float): Incoming rotation value that may be adjusted.

    Returns:
        tuple: (shape_type, rotation) where shape_type is an MSO_AUTO_SHAPE_TYPE and
               rotation is an integer/float representing degrees to apply.

    Notes:
        - This is a heuristic mapping and may not cover arbitrary complex polygons.
        - Returns default shape_type and rotation unchanged if no pattern matches.
    """
    shape_type = MSO_SHAPE.RECTANGLE  # Default
    
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
    
    return shape_type, rotation


def add_shape_element(slide, element):
    """Add a generic shape element to the slide with text and clip-path mapping.

    This function supports:
      - Mapping CSS clip-path polygons to PPT auto-shapes (via get_clip_path_shape).
      - Filling shapes with background color or leaving them transparent.
      - Applying rotation, padding, text formatting and alignment from styles.

    Args:
        slide: pptx Slide object where the shape will be added.
        element (dict): Element metadata containing 'shapeInfo' or inline styles/text.
        slide_width: Width of the slide (optional, for bounds checking).
        slide_height: Height of the slide (optional, for bounds checking).

    Returns:
        pptx shape object: The created shape. If shapeInfo is missing, returns None.
    """
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

    shape_type = MSO_SHAPE.RECTANGLE  # Default
    
    # Enhanced clip-path mapping to PowerPoint shapes
    if 'polygon' in clip_path:
        shape_type, rotation = get_clip_path_shape(clip_path, rotation)

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

    if not text_content:
        return shape
    
    # Add text to the shape
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
        # Clean font family string (remove quotes and extra info)
        font_family = font_family.split(',')[0].strip('"\'')
        font.name = font_family
    
    font.bold = styles.get('fontWeight', '400') in ['bold', '600', '700', '800', '900']
    font.italic = styles.get('fontStyle', 'normal') == 'italic'
    
    # Apply text color
    text_color = parse_color(styles.get('color', 'black'))
    if text_color:
        font.color.rgb = text_color

    # Apply text alignment
    if text_align == 'center' or justify_content == 'center':
        p.alignment = PP_ALIGN.CENTER
    elif text_align == 'right' or justify_content == 'flex-end':
        p.alignment = PP_ALIGN.RIGHT
    else:
        p.alignment = PP_ALIGN.LEFT
    
    return shape
