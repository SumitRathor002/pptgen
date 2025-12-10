import re
import json
from pptx.util import Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_MARK, XL_TICK_LABEL_POSITION

from Builders.utils import parse_color, is_bold, pixels_to_emu, FONT_SCALE_FACTOR, SubElement
from Builders.logger import logger


def parse_data_label_formatter(formatter_str, labels, values, index):
    """
    Parses a Chart.js data_label formatter string and returns the formatted label.
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


def apply_data_label_positioning(data_labels, anchor, align, chart_type_str):
    """Apply Chart.js style positioning to PowerPoint data labels"""
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


def find_chart_type(chart_type_str, datasets, index_axis):
    """
    Map Chart.js-like chart type and dataset hints to pptx XL_CHART_TYPE.

    Uses dataset hints (e.g. presence of fills or pointRadius) and indexAxis to
    return a suitable XL_CHART_TYPE for rendering. Returns None if no mapping
    could be determined.
    """
    
    # Map chart types with better handling
    if chart_type_str == 'pie':
        return XL_CHART_TYPE.PIE
    elif chart_type_str == 'doughnut':
        # Doughnut charts in PowerPoint are rendered as pie charts with a hole size
        return XL_CHART_TYPE.DOUGHNUT
    elif chart_type_str == 'line':
        has_fill = any(dataset.get('fill') is not None and dataset.get('fill') != False for dataset in datasets)
        if has_fill:
            return XL_CHART_TYPE.AREA_STACKED if any(dataset.get('fill') == '-1' for dataset in datasets) else XL_CHART_TYPE.AREA
        else:
            point_radius = datasets[0].get('pointRadius', 0) if datasets else 0
            return XL_CHART_TYPE.LINE_MARKERS if point_radius > 0 else XL_CHART_TYPE.LINE
    elif chart_type_str == 'bar':
        return XL_CHART_TYPE.BAR_CLUSTERED if index_axis == 'y' else XL_CHART_TYPE.COLUMN_CLUSTERED
    
    return None


def build_chart_data(data, options, chart_type):
    """
    Build a pptx CategoryChartData object from chart config data.

    Handles multi-line labels and conditional series labeling depending on legend
    display options to avoid extraneous textual output on charts with a single
    dataset.
    """
    
    chart_data = CategoryChartData()
    labels = data.get('labels', [])
    
    # Enhanced label handling for multi-line labels
    processed_labels = [str(label).replace('\\n', '\n') for label in labels]
    # Reverse category order for horizontal bar charts to match HTML rendering
    is_horizontal_bar = chart_type == XL_CHART_TYPE.BAR_CLUSTERED
    if is_horizontal_bar:
        processed_labels = list(reversed(processed_labels))

    # For doughnut/pie charts without labels, use empty strings for categories
    if chart_type in [XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT] and not processed_labels:
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

    return chart_data


def handle_data_labels(chart, data_label_opts, labels, chart_type_str):
    """
    Configure data label visibility, font and content on a pptx Chart.

    Supports custom formatter strings by applying formatted text to individual
    point labels; otherwise uses pptx data label flags. Applies FONT_SCALE_FACTOR
    to sizes for visually consistent output.
    """
    has_data_labels = data_label_opts.get('display', False)
    
    if not has_data_labels:
        plot = chart.plots[0]
        plot.has_data_labels = False
        return 
    
    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    
    # Get font settings from data labels options - apply scaling to exact values from JSON
    font_opts = data_label_opts.get('font', {})
    label_color_str = data_label_opts.get('color', '#333333')
    label_color = parse_color(label_color_str)
    original_font_size = font_opts.get('size', 8)  # Use exact size from JSON
    # Apply font scaling factor to data labels font size
    scaled_font_size = max(6, int(original_font_size * FONT_SCALE_FACTOR))
    font_weight = is_bold(font_opts.get('weight'))
    font_family = 'Arial'

    # Apply font settings to data labels using scaled values
    data_labels.font.bold = font_weight
    data_labels.font.size = Pt(scaled_font_size)  # Use scaled font size
    data_labels.font.name = font_family
    if label_color:
        data_labels.font.color.rgb = label_color
    
    apply_data_label_positioning(
        data_labels,
        data_label_opts.get('anchor', 'center'),
        data_label_opts.get('align', 'center'),
        chart_type_str
    )

    formatter_str = data_label_opts.get('formatter')
    if not formatter_str:
        # For non-custom formatters, still apply the exact scaled font settings
        data_labels.show_category_name = False
        data_labels.show_value = True
        data_labels.show_percentage = False
        return 
    
    data_labels.show_category_name = False
    data_labels.show_value = False
    data_labels.show_percentage = False
    
    # Apply formatting to individual points with consistent scaled font settings
    for series in chart.series:
        for i, point in enumerate(series.points):
            point.has_data_label = True
            data_label = point.data_label
            values = series.values
            formatted_text = parse_data_label_formatter(formatter_str, labels, values, i)
            data_label.text_frame.text = formatted_text
            
            if not data_label.text_frame.paragraphs:
                return 
            
            for paragraph in data_label.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(scaled_font_size)  # Use scaled size
                    run.font.bold = font_weight
                    run.font.name = font_family
                    if label_color:
                        run.font.color.rgb = label_color


def configure_legend(chart, legend_opts, has_data: bool):
    """
    Configure chart legend visibility, position and font styling.

    Parameters:
      chart: pptx Chart object.
      legend_opts (dict): Legend options from chart config.
      has_data (bool): Default legend presence when display isn't specified.
    """
    legend_display = legend_opts.get('display')
    if legend_display is False:
        chart.has_legend = False
        # Also ensure series names don't show up anywhere else
        # Hide series names from chart title area
        if hasattr(chart, 'chart_title'):
            chart.chart_title.has_text_frame = False

                
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
        legend_labels = legend_opts.get('labels', {})
        legend_font = legend_labels.get('font', {})
        legend_color_str = legend_labels.get('color')
        legend_color = parse_color(legend_color_str)
        scaled_size = int(10 * FONT_SCALE_FACTOR)
        if legend_font.get('size'):
            # Apply font scaling factor to legend font size
            original_size = int(legend_font['size'])
            scaled_size = max(6, int(original_size * FONT_SCALE_FACTOR))
        
        chart.legend.font.size = Pt(scaled_size)
        chart.legend.font.name = 'Arial'
        
        if legend_color:
            chart.legend.font.color.rgb = legend_color

    else:
        # Default behavior when display is not specified
        chart.has_legend = has_data 


def handle_grid_lines(value_axis, category_axis, value_scale, category_scale):
    """
    Apply gridline visibility and color settings to chart axes.

    Skips for pie charts. Reads grid display and color from scale configuration
    and applies them to the corresponding pptx axis gridlines.
    """
    y_grid = value_scale.get('grid', {})
    x_grid = category_scale.get('grid', {})
    
    # Value axis gridlines
    value_axis.has_major_gridlines = bool(y_grid.get('display'))

    y_grid_color = parse_color(y_grid.get('color'))
    if value_axis.has_major_gridlines and y_grid_color: 
        value_axis.major_gridlines.format.line.color.rgb = y_grid_color

    # Category axis gridlines
    category_axis.has_major_gridlines = bool(x_grid.get('display'))
    x_grid_color = parse_color(x_grid.get('color'))
    if category_axis.has_major_gridlines and x_grid_color: 
        category_axis.major_gridlines.format.line.color.rgb = x_grid_color
                

def handle_value_axis(value_axis, value_scale, chart_type_str, all_negative):
    """
    Configure the numeric/value axis: min/max, step size, ticks and visibility.

    Applies FONT_SCALE_FACTOR to tick label font sizes and handles removal of
    tick marks for most chart types for clarity.
    """
    # Handle_value_axis
    val_min = value_scale.get('min')
    val_max = value_scale.get('max')
    val_max = value_scale.get('suggestedMax') or val_max
    
    if val_min is not None:
        value_axis.minimum_scale = val_min
        
    # Handle special case for all-negative bar charts
    if not all_negative and chart_type_str != 'bar' and val_max is not None:
        value_axis.maximum_scale = float(val_max)
        
    # Handle step size
    ticks = value_scale.get('ticks', {})
    step_size = ticks.get('stepSize')
    if step_size:
        value_axis.major_unit = float(step_size)

    # Remove tick marks for bar charts
    if chart_type_str not in ['pie']:
        value_axis.major_tick_mark = XL_TICK_MARK.NONE
        value_axis.minor_tick_mark = XL_TICK_MARK.NONE
        
    if not value_axis.tick_labels:
        return 
    
    tick_font = ticks.get('font', {})
    tick_color = parse_color(ticks.get('color', '#888888'))
    if tick_color:
        value_axis.tick_labels.font.color.rgb = tick_color
        
    if tick_font.get('size'):
        # Apply font scaling factor to value axis tick labels
        original_tick_size = tick_font['size']
        scaled_tick_size = max(6, int(original_tick_size * FONT_SCALE_FACTOR))
        value_axis.tick_labels.font.size = Pt(scaled_tick_size)
        value_axis.tick_labels.font.name = 'Arial'
            
    # Handle axis visibility
    if value_scale.get('display', True) == False:
        value_axis.visible = False
    
    # Hide axis line if border is not displayed
    if value_scale.get('border', {}).get('display') is False:
        value_axis.format.line.fill.background()

    
def handle_category_axis(category_axis, category_scale, chart_type_str, all_negative, reference_id=None):
    """
    Configure category axis (x or y) visual settings such as visibility, ticks,
    font sizing and rotation. Also handles special positioning for horizontal
    bar charts and charts with all-negative values.
    """
    # Enhanced Category axis configuration with font scaling
    
    cat_ticks = category_scale.get('ticks', {})
    category_axis.visible = bool(category_scale.get('display', True))
    # Category axis visibility should only depend on scale display setting, not data labels
    if not category_axis.visible: 
        return 
        
    # Handle axis line display
    border_config = category_scale.get('border', {})
    border_color = parse_color(border_config.get('color', '#666666'))
    if border_config.get('display') is False:
        category_axis.format.line.fill.background()
    elif border_config.get('display') is True and border_color:
        # Ensure axis line is visible and apply color if specified
        category_axis.format.line.color.rgb = border_color
            
    # Configure category axis labels when they should be visible
    if chart_type_str not in ['pie']:
        category_axis.major_tick_mark = XL_TICK_MARK.NONE
        category_axis.minor_tick_mark = XL_TICK_MARK.NONE
        
    
    if category_axis.tick_labels:
        tick_font = cat_ticks.get('font', {})
        scaled_cat_size = int(8 * FONT_SCALE_FACTOR)
        if tick_font.get('size'):
            # Apply font scaling factor to category axis tick labels
            original_cat_size = tick_font['size']
            scaled_cat_size = max(6, int(original_cat_size * FONT_SCALE_FACTOR))
    
        category_axis.tick_labels.font.size = Pt(scaled_cat_size)
        category_axis.tick_labels.font.name = 'Arial'
        
        tick_color = parse_color(cat_ticks.get('color', '#666666'))
        if tick_color:
            category_axis.tick_labels.font.color.rgb = tick_color
        
        # Handle label rotation properly
        max_rotation = cat_ticks.get('maxRotation', 0)
        min_rotation = cat_ticks.get('minRotation', 0)
        if max_rotation == 0 and min_rotation == 0:
            category_axis.tick_labels.orientation = 0  # Horizontal
        elif max_rotation > 0:
            category_axis.tick_labels.orientation = max_rotation
        
        # Set x-axis label position to 'low' and distance from axis for bar charts
        if chart_type_str != 'bar':
            return
            
        try:
            # Set distance from axis for top positioning
            # category_axis.tick_labels.offset = 500
            axis_label_position = XL_TICK_LABEL_POSITION.LOW
            if all_negative:
                # axis_element = category_axis._element
                # For charts with negative values, position x-axis at top
                axis_label_position = XL_TICK_LABEL_POSITION.HIGH
                # Set axis crossing to automatic high for negative data
                # crosses = axis_element.find(qn('c:crosses'))
                # if crosses is not None:
                #     crosses.set('val', 'autoZero')
                # else:
                #     crosses_elem = SubElement(axis_element, 'c:crosses', val='autoZero')            
            category_axis.tick_label_position = axis_label_position
        except Exception as tick_pos_error:
            logger.error(f"Error setting tick label position: {tick_pos_error}", reference_id=reference_id)


def handle_colors(chart, chart_type, datasets):
    """
    Apply colors from dataset configuration to chart series and data points.

    Handles line/area fills, line colors, and per-point fills for bar/column/pie
    charts. When necessary it will inject XML elements (e.g., invertIfNegative)
    to control pptx rendering behavior.
    """
    
    for i, series in enumerate(chart.series):
        if i >= len(datasets):
            continue
        
        dataset = datasets[i]
        # Line/Area chart colors
        if chart_type in [XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS, XL_CHART_TYPE.AREA, XL_CHART_TYPE.AREA_STACKED]:
            border_color = parse_color(dataset.get('borderColor'))
            if border_color:
                series.format.line.color.rgb = border_color
            
            # Enhanced area fill handling based on dataset configuration
            if chart_type in [XL_CHART_TYPE.AREA, XL_CHART_TYPE.AREA_STACKED]:
                bg_color_str = dataset.get('backgroundColor', '')
                
                # Check if backgroundColor is transparent or not set for area fill
                bg_color = parse_color(bg_color_str)
                if bg_color:
                    series.format.fill.solid()
                    series.format.fill.fore_color.rgb = bg_color
                else:
                    # If backgroundColor is transparent or not set, don't fill
                    series.format.fill.background()

        # Bar/Column/Pie chart colors
        elif chart_type in [XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.BAR_CLUSTERED, XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT]:
            background_colors = dataset.get('backgroundColor', [])

            # Reverse colors for horizontal bars to match reversed data
            if chart_type == XL_CHART_TYPE.BAR_CLUSTERED:
                background_colors = list(reversed(background_colors))

            for j, point in enumerate(series.points):
                # Disable "invert if negative" for bar/column charts to prevent automatic color inversion
                if chart_type in [XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.BAR_CLUSTERED, XL_CHART_TYPE.DOUGHNUT]:
                    SubElement(point.format.element, 'c:invertIfNegative', val='0')

                if j >= len(background_colors):
                    continue
                
                color = parse_color(background_colors[j])
                if not color:
                    continue
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = color


def add_chart_element(slide, element, slide_width, slide_height, reference_id=0):
    """Enhanced chart rendering with accurate configuration extraction using modular helper functions.
    
    Args:
        slide: pptx Slide object where the chart will be added.
        element (dict): Element metadata containing chartInfo or chartConfig.
        slide_width (int): Width of the slide in pixels.
        slide_height (int): Height of the slide in pixels.
        reference_id (int|str): Identifier for diagnostic context.
    """
    chart_config = None
    # Use element dimensions directly
    x = max(0, min(element.get('x', 0), slide_width - 10))
    y = max(0, min(element.get('y', 0), slide_height - 10))
    width = max(10, min(element.get('width', 400), slide_width - x))
    height = max(10, min(element.get('height', 300), slide_height - y))
    
    if element.get('chartInfo'):
        chart_info = element['chartInfo']
        chart_data = chart_info.get('chartData')
        canvas_width = chart_info.get('width', width)
        canvas_height = chart_info.get('height', height)
        
        # Use canvas dimensions if they're more accurate
        if canvas_width > 0 and canvas_height > 0:
            width = min(canvas_width, slide_width - x)
            height = min(canvas_height, slide_height - y)
        
        if isinstance(chart_data, dict):
            # Data is already a parsed dictionary
            chart_config = chart_data
        
        if isinstance(chart_data, str):
            # Data is a string, needs parsing
            try:
                chart_config = json.loads(chart_data)
            except json.JSONDecodeError as e:
                logger.error(f"Error parsing chart data string: {e}", reference_id=reference_id)
        
    # Fallback for older format
    if not chart_config:
        chart_config = element.get('chartConfig')

    if not chart_config:
        return

    chart_type_str = chart_config.get('type')
    options = chart_config.get('options', {})
    datasets = chart_config.get('data', {}).get('datasets', [])
    
    # Check if all data values are negative for bar charts
    all_negative = False
    if chart_type_str == 'bar':    
        all_values = [val for dataset in datasets for val in dataset.get('data', []) if isinstance(val, (int, float))]
        all_negative = bool(all_values and all(val < 0 for val in all_values))
    
    
    index_axis = options.get('indexAxis', 'x')
    chart_type = find_chart_type(chart_type_str, datasets=datasets, index_axis=index_axis)
    
    if chart_type is None:
        return
    
    data = chart_config.get('data', {})
    labels = data.get('labels', [])
    chart_data = build_chart_data(data, options, chart_type)

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

    # Enhanced Legend handling with accurate font configuration
    plugins = options.get('plugins', {})
    legend_opts = plugins.get('legend', {})
    
    # properly handle explicit display settings
    configure_legend(chart, legend_opts, has_data = bool(len(data.get('datasets', [])) > 1))
    
    # Enhanced Data Labels with proper font settings from options and scaling
    data_label_opts = plugins.get('datalabels', {})
    handle_data_labels(chart, data_label_opts, labels, chart_type_str)
    
    # Disable chart title to prevent series labels from appearing there
    chart.has_title = False

    if chart_type_str not in ['pie', 'doughnut']:
        scales = options.get('scales', {})
        is_horizontal_bar = chart_type == XL_CHART_TYPE.BAR_CLUSTERED
        
        # Get axis configurations
        category_scale = scales.get('y' if is_horizontal_bar else 'x', {})
        value_scale = scales.get('x' if is_horizontal_bar else 'y', {})
        category_axis = chart.category_axis
        value_axis = chart.value_axis

        handle_value_axis(value_axis, value_scale, chart_type_str, all_negative)
        handle_category_axis(category_axis, category_scale, chart_type_str, all_negative, reference_id=reference_id)
        handle_grid_lines(value_axis, category_axis, value_scale, category_scale)

    if chart_type in [XL_CHART_TYPE.BAR_CLUSTERED, XL_CHART_TYPE.COLUMN_CLUSTERED]:
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
        # overlap in PowerPoint: negative = gap, 0 = touching, positive = overlapping
        # We'll set to 0 for now as Chart.js barPercentage is more about individual bar width
        plot.overlap = 0

    datasets = data.get('datasets', [])
    handle_colors(chart, chart_type, datasets)
