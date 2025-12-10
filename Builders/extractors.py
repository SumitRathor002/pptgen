"""
Helper functions for extracting structured data from hierarchical element structures.
"""

def extract_table_info_from_hierarchy(table_element):
    """Extract table info from hierarchical structure"""
    children = table_element.get('children', [])
    
    # Find tbody, thead, tfoot or direct tr elements
    all_rows = []
    for child in children:
        child_type = child.get('type')
        if child_type in ['tbody', 'thead', 'tfoot']:
            # These contain tr children
            for tr in child.get('children', []):
                if tr.get('type') == 'tr':
                    all_rows.append(tr)
        elif child_type == 'tr':
            all_rows.append(child)
    
    # Calculate row and column count
    row_count = len(all_rows)
    max_cols = 0
    
    for tr in all_rows:
        col_count = 0
        for cell in tr.get('children', []):
            if cell.get('type') in ['td', 'th']:
                cell_info = cell.get('cellInfo', {})
                col_count += cell_info.get('colSpan', 1)
        max_cols = max(max_cols, col_count)
    
    # Build rows structure
    rows = []
    for idx, tr in enumerate(all_rows):
        row_data = {
            'index': idx,
            'rect': {
                'x': tr.get('x', 0),
                'y': tr.get('y', 0),
                'width': tr.get('width', 0),
                'height': tr.get('height', 0)
            },
            'styles': tr.get('styles', {}),
            'cells': []
        }
        
        for cell in tr.get('children', []):
            if cell.get('type') in ['td', 'th']:
                cell_info = cell.get('cellInfo', {})
                
                # Extract inline group if cell has inline children
                inline_group = None
                children = cell.get('children', [])
                has_inline = any(c.get('type') in ['span', 'strong', 'em', 'b', 'i', 'u', 'br'] for c in children)
                if has_inline:
                    inline_group = extract_inline_group_from_hierarchy(cell)
                else:
                    inline_group = None

                from Builders.tables import flatten_cell_content
                # Flatten nested content (e.g., td > div > [p, p, ul]) to extract actual elements
                html_content = flatten_cell_content(children)

                cell_data = {
                    'rect': {
                        'x': cell.get('x', 0),
                        'y': cell.get('y', 0),
                        'width': cell.get('width', 0),
                        'height': cell.get('height', 0)
                    },
                    'colSpan': cell_info.get('colSpan', 1),
                    'rowSpan': cell_info.get('rowSpan', 1),
                    'styles': cell.get('styles', {}),
                    'text': cell.get('text', ''),
                    'htmlContent': html_content
                }
                
                if inline_group:
                    cell_data['inlineGroup'] = inline_group
                row_data['cells'].append(cell_data)
        rows.append(row_data)
    
    return {
        'rowCount': row_count,
        'columnCount': max_cols,
        'rect': {
            'x': table_element.get('x', 0),
            'y': table_element.get('y', 0),
            'width': table_element.get('width', 0),
            'height': table_element.get('height', 0)
        },
        'styles': table_element.get('styles', {}),
        'rows': rows
    }


def extract_list_info_from_hierarchy(list_element):
    """Extract list info from hierarchical structure"""
    children = list_element.get('children', [])
    list_type = list_element.get('type')
    
    items = []
    for child in children:
        if child.get('type') == 'li':
            item_data = {
                'rect': {
                    'x': child.get('x', 0),
                    'y': child.get('y', 0),
                    'width': child.get('width', 0),
                    'height': child.get('height', 0)
                },
                'styles': child.get('styles', {}),
                'text': child.get('text', '')
            }
            
            # Check for inline children
            li_children = child.get('children', [])
            has_inline = any(c.get('type') in ['span', 'strong', 'em', 'b', 'i', 'u', 'br'] for c in li_children)
            if has_inline:
                item_data['inlineGroup'] = extract_inline_group_from_hierarchy(child)
            
            # Check for nested lists
            for nested in li_children:
                if nested.get('type') in ['ul', 'ol']:
                    item_data['nestedList'] = extract_list_info_from_hierarchy(nested)
            
            items.append(item_data)
    
    return {
        'type': list_type,
        'rect': {
            'x': list_element.get('x', 0),
            'y': list_element.get('y', 0),
            'width': list_element.get('width', 0),
            'height': list_element.get('height', 0)
        },
        'styles': list_element.get('styles', {}),
        'items': items,
        'listStyles': list_element.get('styles', {})
    }


def extract_inline_group_from_hierarchy(parent_element):
    """Extract inline group from hierarchical structure"""
    inline_elements = []
    
    # Add direct text if present
    if parent_element.get('text'):
        inline_elements.append({
            'type': 'text',
            'text': parent_element.get('text'),
            'styles': parent_element.get('styles', {})
        })
    
    # Process children
    for child in parent_element.get('children', []):
        child_type = child.get('type')
        
        if child_type == 'br':
            inline_elements.append({'type': 'br'})
        elif child_type in ['span', 'strong', 'em', 'b', 'i', 'u', 'a']:
            inline_elements.append({
                'type': child_type,
                'text': child.get('text', ''),
                'styles': child.get('styles', {})
            })
            # Recursively add nested inline children
            for nested in child.get('children', []):
                if nested.get('type') in ['span', 'strong', 'em', 'b', 'i', 'u', 'br']:
                    if nested.get('type') == 'br':
                        inline_elements.append({'type': 'br'})
                    else:
                        inline_elements.append({
                            'type': nested.get('type'),
                            'text': nested.get('text', ''),
                            'styles': nested.get('styles', {})
                        })
    
    return {
        'groupRect': {
            'x': parent_element.get('x', 0),
            'y': parent_element.get('y', 0),
            'width': parent_element.get('width', 0),
            'height': parent_element.get('height', 0)
        },
        'styles': parent_element.get('styles', {}),
        'inlineElements': inline_elements
    }
