import os
import json
from pptx import Presentation
from pptx.dml.color import RGBColor
import concurrent.futures

from Builders.utils import fetch_image, safe_int, pixels_to_emu, parse_color, has_any_border
from Builders.texts import add_text_element, add_list_element, add_link_element, add_inline_group_element, add_overlay_element
from Builders.shapes import add_bg_shape, add_image_element, add_shape_element, add_pseudo_element
from Builders.tables import add_table_element
from Builders.charts import add_chart_element
from Builders.logger import logger

def parse_images(slides_data):
    """
    Extract image URLs from slide data.

    Iterates slides_data and collects image source URLs found in elements and
    footerElements. Returns a list of image URLs (may contain duplicates).
    
    Parameters:
        slides_data (list): List of slide data dictionaries.

    Returns:
        list: A list of image URL strings to be fetched.
    """
    images_to_fetch = set()
    for slide_data in slides_data:
        # For hierarchical structure, use 'children' key
        elements = slide_data.get('elements', slide_data.get('children', []))
        elements_sorted = sorted(elements, key=lambda e: (e.get('zIndex', 0), e.get('y', 0), e.get('x', 0)))
        # Track processed elements to prevent duplicates
        processed_element_ids = set()
        # Process each element
        for element in elements_sorted:
            element_id = f"{element.get('type', 'unknown')}-{element.get('x', 0)}-{element.get('y', 0)}-{element.get('className', '')}"
            if element_id in processed_element_ids:
                continue
            processed_element_ids.add(element_id)
            element_type = element.get('type')
            if element.get('footerElements'):
                # Process individual footer elements with space-between positioning
                footer_elements = element.get('footerElements', [])
                for footer_elem in footer_elements:
                    if footer_elem.get('type') == 'img':
                        images_to_fetch.add(
                            footer_elem.get('mediaInfo', {}).get('src', '')
                        )

            elif element_type == 'img':
                # Handle both direct src and mediaInfo structure
                src = element.get('mediaInfo', {}).get('src', '') or element.get('src', '')
                images_to_fetch.add(src)

            elif element_type == 'table':
                for row in element.get('tableInfo', {}).get('rows', []):
                    for cell in row.get('cells', []):
                        cell_content = cell.get('htmlContent', [])
                        for content in cell_content:
                            if content.get('type') == 'img':
                                images_to_fetch.add(
                                    content.get('mediaInfo', {}).get('src', '')
                                )
            
            # Also check children recursively for hierarchical structure
            _collect_images_recursive(element, images_to_fetch)

    return list(images_to_fetch)

def prefetch_images(images_to_fetch, reference_id, img_dict=None):
    """
    Concurrently fetch image bytes for provided URLs.

    Uses a ThreadPoolExecutor to call fetch_image for each URL and returns a
    mapping of URL -> image bytes (or None on failure).

    Parameters:
        images_to_fetch (list): List of image URL strings.
        reference_id (str): Identifier passed to fetch_image for logging/tracking.
        img_dict (dict): Optional existing dict to add images to.

    Returns:
        dict: Mapping from image URL to fetched bytes (or None).
    """
    if not img_dict:
        img_dict = {}

    if not images_to_fetch:
        return img_dict

    with concurrent.futures.ThreadPoolExecutor(max_workers=min(10, len(images_to_fetch))) as executor:
        futures = {
            executor.submit(
                fetch_image, 
                image_url=url, 
                reference_id=reference_id
            ): url for url in images_to_fetch if url not in img_dict
        }
        for future in concurrent.futures.as_completed(futures):
            url = futures[future]
            try:
                image_bytes = future.result()
                img_dict[url] = image_bytes
            except Exception as e:
                logger.error(f"Error fetching {url}: {e}", reference_id=reference_id)
                img_dict[url] = None
    
    return img_dict

def _collect_images_recursive(element, images_set):
    """Recursively collect image URLs from hierarchical element structure."""
    if not element:
        return
    
    if element.get('type') == 'img':
        src = element.get('mediaInfo', {}).get('src', '') or element.get('src', '')
        if src:
            images_set.add(src)
    
    for child in element.get('children', []):
        _collect_images_recursive(child, images_set)


def _flatten_elements_recursive(elements, flattened_list=None):
    """Recursively flatten hierarchical elements into a single list for rendering.
    
    This function traverses nested 'children' arrays and collects all elements
    that need to be rendered, preserving their absolute positions.
    
    Parameters:
        elements (list): List of element dictionaries, potentially with nested children.
        flattened_list (list): Accumulator list for flattened elements.
    
    Returns:
        list: Flat list of all elements including nested children.
    """
    if flattened_list is None:
        flattened_list = []
    
    for element in elements:
        if not element:
            continue
        
        # Add the current element to the flattened list
        flattened_list.append(element)
        
        # Recursively process children
        children = element.get('children', [])
        if children:
            _flatten_elements_recursive(children, flattened_list)
    
    return flattened_list


def add_div_and_heading_elements(element, slide, slide_width, slide_height, parent_has_shadow, reference_id):
    """
    Render div or heading style elements on a slide.

    Decides whether to add text, background shapes or skip rendering depending
    on element content and styles. Prevents adding duplicate structural divs
    that contain only inline or processed elements.

    Parameters:
        element (dict): Element dictionary describing the div/heading.
        slide (pptx.slide.Slide): Slide to which content will be added.
        slide_width (int): Slide width in pixels.
        slide_height (int): Slide height in pixels.
        parent_has_shadow (bool): Whether parent element has shadow styling.
        reference_id (str): Reference id used when creating background shapes.

    Returns:
        None
    """
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
                add_bg_shape(slide, element.get('styles', {}), x, y, width, height, reference_id=reference_id)


def add_slide(prs, slide_data, slide_width, slide_height, reference_id, img_dict={}, slide=None, add_slide_bg=True):
    """
    Add a single slide to the Presentation based on slide_data.

    Creates a blank slide, applies slide background and iterates through
    elements in z-order delegating rendering to specific builder functions.
    Tracks processed elements to avoid duplication and collects table positions
    for overlay alignment.

    Parameters:
        prs (pptx.presentation.Presentation): Presentation instance to append slide.
        slide_data (dict): Slide data describing elements and styles.
        slide_width (int): Slide width in pixels.
        slide_height (int): Slide height in pixels.
        reference_id (str): Reference id used for image/background references.
        img_dict (dict): Optional map of image URL -> bytes for pre-fetched images.
        slide (pptx.slide.Slide): Optional existing slide to use instead of creating new.
        add_slide_bg (bool): Whether to add slide background styling.

    Returns:
        None
    """
    if slide is None:
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)
    
    # Set slide background
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = parse_color('#ffffff') or RGBColor(255, 255, 255)
    
    # Add slide background styling
    slide_styles = slide_data.get('slideStyles', {})
    if slide_styles and add_slide_bg:
        add_bg_shape(slide, slide_styles, 0, 0, slide_width, slide_height, reference_id=reference_id)
    
    # Support both flat 'elements' and hierarchical 'children' structure
    top_level_elements = slide_data.get('elements', slide_data.get('children', []))
    
    # Flatten all nested children into a single list for rendering
    all_elements = _flatten_elements_recursive(top_level_elements)
    
    elements_sorted = sorted(all_elements, key=lambda e: (e.get('zIndex', 0), e.get('y', 0), e.get('x', 0)))
    
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
            
            add_shape_element(slide, element)
            continue

        if element.get('footerElements'):
            # Process individual footer elements with space-between positioning
            footer_elements = element.get('footerElements', [])
            for footer_elem in footer_elements:
                if footer_elem.get('type') == 'img':
                    add_image_element(slide, footer_elem, slide_width, slide_height, parent_has_shadow, img_dict=img_dict, reference_id=reference_id)
                elif footer_elem.get('type') in ['span', 'div']:
                    add_text_element(slide, footer_elem, slide_width, slide_height, parent_has_shadow)
            continue

        if element_type == 'overlay':
            # Pass table positions to overlay processing
            element['_table_positions'] = table_positions
            add_overlay_element(slide, element, slide_width, slide_height, reference_id=reference_id)
        elif element.get('inlineGroup'):
            add_inline_group_element(slide, element, slide_width, slide_height, reference_id=reference_id)
        elif element_type in ['ul', 'ol']:
            add_list_element(slide, element, parent_has_shadow=parent_has_shadow, reference_id=reference_id)
        elif element_type == 'table':
            # Store table position for overlay calculations
            table_key = f"table_{element.get('x', 0)}_{element.get('y', 0)}"
            table_positions[table_key] = {
                'x': element.get('x', 0),
                'y': element.get('y', 0),
                'table_info': element.get('tableInfo', {})
            }
            
            # Add the table
            add_table_element(slide, element, slide_width, slide_height, reference_id, parent_has_shadow, img_dict=img_dict)
        elif element_type == 'img':
            add_image_element(slide, element, slide_width, slide_height, parent_has_shadow, img_dict=img_dict, reference_id=reference_id)
        elif element_type == 'a':
            add_link_element(slide, element, slide_width, slide_height)
        elif element_type == 'canvas':
            add_chart_element(slide, element, slide_width, slide_height, reference_id=reference_id)
        elif element_type == 'span':
            add_text_element(slide, element, slide_width, slide_height, parent_has_shadow)
        elif element_type == 'pseudo':
            add_pseudo_element(slide, element, slide_width, slide_height, reference_id=reference_id)
        elif element_type == 'div' and 'chart' in element.get('className', '') and element.get('chartConfig'):
            add_chart_element(slide, element, slide_width, slide_height, reference_id=reference_id)
        elif element_type in ['div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            add_div_and_heading_elements(element, slide, slide_width, slide_height, parent_has_shadow, reference_id)


def create_pptx_from_json(slides_data, reference_id=0):
    """Enhanced PowerPoint generation with hierarchical data processing and image prefetching.
    
    Parameters:
        slides_data (list): List of slide data dictionaries.
        reference_id (int|str): Identifier for logging/tracking.
    
    Returns:
        pptx.presentation.Presentation: The constructed Presentation object or None on failure.
    """
    if not slides_data:
        logger.warning("No slides found in data", reference_id=reference_id)
        return
    
    # Prefetch images for better performance
    images_to_fetch = parse_images(slides_data)
    img_dict = prefetch_images(images_to_fetch, reference_id)
   
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
            add_bg_shape(slide, slide_styles, 0, 0, slide_width, slide_height, reference_id=reference_id)
       
        # Get root children from hierarchical structure
        children = slide_data.get('children', [])
       
        # Sort elements primarily by z-index, then by y, x for consistent layering
        elements_sorted = sorted(children, key=lambda e: (e.get('zIndex', 0), e.get('y', 0), e.get('x', 0)))
       
        # Track processed elements to prevent duplicates
        processed_element_ids = set()
        processed_text_content = set()
        
        # Store table information for overlay positioning
        table_positions = {}
       
        # Process each element hierarchically
        for element in elements_sorted:
            element_id = f"{element.get('type', 'unknown')}-{element.get('x', 0)}-{element.get('y', 0)}-{element.get('className', '')}"
            if element_id in processed_element_ids:
                continue
            processed_element_ids.add(element_id)
            
            try:
                from pptx_script import process_element_recursive
                process_element_recursive(slide, element, slide_width, slide_height, processed_text_content, table_positions, img_dict=img_dict, reference_id=reference_id)
            except Exception as e:
                logger.error(f"Error processing element {element_id}: {e}", reference_id=reference_id)

    return prs
