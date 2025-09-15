const puppeteer = require('puppeteer');
const fs = require('fs').promises;

async function extractSlideData(htmlFilePath, outputPath) {
    const browser = await puppeteer.launch({
        headless: true,
        devtools: false,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-web-security']
    });
    const page = await browser.newPage();
    try {
        const htmlContent = await fs.readFile(htmlFilePath, 'utf-8');
        await page.setViewport({ width: 1920, height: 1080 });
        await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
        await page.waitForFunction(() => {
            const images = Array.from(document.querySelectorAll('img'));
            const fonts = document.fonts ? document.fonts.ready : Promise.resolve();
            return Promise.all([fonts, images.every(img => img.complete)]);
        }, { timeout: 15000 }).catch(() => console.log('Some resources may not have loaded'));

        const documentInfo = await page.evaluate(() => {
            const body = document.body;
            const html = document.documentElement;
            const actualWidth = Math.max(
                body.scrollWidth, body.offsetWidth,
                html.clientWidth, html.scrollWidth, html.offsetWidth
            );
            const actualHeight = Math.max(
                body.scrollHeight, body.offsetHeight,
                html.clientHeight, html.scrollHeight, html.offsetHeight
            );
            const viewportWidth = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
            const viewportHeight = Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0);
            let slideElements = Array.from(document.querySelectorAll('.slide'));
            if (slideElements.length === 0) {
                slideElements = Array.from(document.body.children).filter(el => el.tagName === 'DIV' || el.tagName === 'SECTION');
            }
            if (slideElements.length === 0) {
                slideElements = [document.body];
            }
            const slidesInfo = slideElements.map((slide, index) => {
                const rect = slide.getBoundingClientRect();
                return {
                    index: index + 1,
                    rect: {
                        x: rect.left,
                        y: rect.top,
                        width: rect.width,
                        height: rect.height
                    }
                };
            });
            return {
                actualWidth,
                actualHeight,
                viewportWidth,
                viewportHeight,
                slidesCount: slideElements.length,
                slidesInfo
            };
        });

        console.log('Document dimensions:', documentInfo);
        console.log(`Found ${documentInfo.slidesCount} slides`);
        const targetWidth = Math.max(documentInfo.actualWidth, 1920);
        const targetHeight = Math.max(documentInfo.actualHeight, 1080);
        await page.setViewport({ width: targetWidth, height: targetHeight });

        const allSlidesData = await page.evaluate(async (docInfo) => {
            const slides = [];
            let slideElements = Array.from(document.querySelectorAll('.slide'));
            if (slideElements.length === 0) {
                slideElements = Array.from(document.body.children).filter(el => el.tagName === 'DIV' || el.tagName === 'SECTION');
            }
            if (slideElements.length === 0) {
                slideElements = [document.body];
            }
            const IMPORTANT_ELEMENTS = [
                'div', 'span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
                'strong', 'b', 'em', 'i', 'u', 'strike', 'del', 'ins', 'mark', 'small', 'sub', 'sup',
                'ul', 'ol', 'li', 'dl', 'dt', 'dd',
                'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'caption', 'colgroup', 'col',
                'img', 'svg', 'video', 'audio', 'iframe',
                'a', 'button', 'input', 'textarea', 'select', 'option', 'label', 'fieldset', 'legend',
                'blockquote', 'pre', 'code', 'kbd', 'samp', 'var',
                'article', 'section', 'aside', 'nav', 'header', 'footer', 'main',
                'figure', 'figcaption', 'details', 'summary', 'dialog',
                'hr', 'br', 'wbr',
                'abbr', 'address', 'bdi', 'bdo', 'cite', 'dfn', 'q', 'ruby', 'rt', 'rp', 's', 'time',
                'canvas'
            ];

            function safeFloat(value) {
                try {
                    return parseFloat(value.replace(/[^0-9.-]/g, '')) || 0;
                } catch {
                    return 0;
                }
            }

            function getElementId(element) {
                if (!element || !element.tagName) return 'unknown';
                
                const rect = element.getBoundingClientRect();
                const text = element.textContent ? element.textContent.trim().substring(0, 50) : '';
                const src = element.src || element.href || '';
                const tagName = element.tagName.toLowerCase();
                const className = element.className || '';
                const id = element.id || '';
                
                return `${tagName}-${id}-${className}-${rect.left.toFixed(1)}-${rect.top.toFixed(1)}-${rect.width.toFixed(1)}-${rect.height.toFixed(1)}-${text}-${src}`;
            }

            function getResolvedBackgroundColor(elementOrStyle, isPseudo = false) {
                let bg;
                if (isPseudo) {
                    bg = elementOrStyle.backgroundColor;
                } else {
                    const element = elementOrStyle;
                    const style = window.getComputedStyle(element);
                    bg = style.backgroundColor;
                    const hasExplicitBackground = element.style.backgroundColor ||
                        element.getAttribute('style')?.includes('background') ||
                        Array.from(element.classList).some(cls => {
                            try {
                                const rules = Array.from(document.styleSheets).flatMap(sheet =>
                                    Array.from(sheet.cssRules || []));
                                return rules.some(rule =>
                                    rule.selectorText?.includes(`.${cls}`) &&
                                    rule.style?.backgroundColor
                                );
                            } catch {
                                return false;
                            }
                        });
                    if (bg === 'rgb(255, 255, 255)' && !hasExplicitBackground) {
                        let current = element.parentElement;
                        let foundExplicitParentBg = false;
                        while (current && current !== document.body) {
                            const parentStyle = window.getComputedStyle(current);
                            const parentBg = parentStyle.backgroundColor;
                            if (parentBg !== 'rgba(0, 0, 0, 0)' && parentBg !== 'transparent') {
                                foundExplicitParentBg = true;
                                break;
                            }
                            current = current.parentElement;
                        }
                        if (!foundExplicitParentBg) {
                            return 'rgba(255, 255, 255, 0)';
                        }
                    }
                }
                return bg || 'rgba(0, 0, 0, 0)';
            }

            function getTextContent(element, processedElements) {
                const tagName = element.tagName.toLowerCase();
                const styles = window.getComputedStyle(element);
                if (styles.display === 'inline' || styles.display === 'inline-block') {
                    return element.textContent.trim();
                }
                const directText = Array.from(element.childNodes)
                    .filter(node => node.nodeType === Node.TEXT_NODE)
                    .map(node => node.textContent.trim())
                    .join(' ')
                    .trim();
                return directText;
            }

            function hasTextContentChildren(element) {
                const textContentTags = ['div', 'span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'td', 'th', 'ul', 'ol', 'table'];
                return Array.from(element.children).some(child => {
                    const childTag = child.tagName.toLowerCase();
                    return textContentTags.includes(childTag) && child.textContent.trim().length > 0;
                });
            }

            function shouldExtractTextContent(element, processedTextElements) {
                const elementId = getElementId(element);
                if (processedTextElements.has(elementId)) {
                    return false;
                }
                
                if (hasTextContentChildren(element)) {
                    return false;
                }
                
                const tagName = element.tagName.toLowerCase();
                if (['ul', 'ol', 'table', 'tbody', 'thead', 'tfoot'].includes(tagName)) {
                    return false;
                }
                
                // Check if any child element has shapeInfo - if so, don't extract text from parent
                const hasShapeChild = Array.from(element.children).some(child => {
                    const styles = window.getComputedStyle(child);
                    return styles.clipPath && styles.clipPath.startsWith('polygon');
                });
                
                if (hasShapeChild) {
                    return false;
                }
                
                // Additional check: if this element's text content is already captured by a shape element
                const elementText = element.textContent ? element.textContent.trim() : '';
                if (elementText) {
                    const rect = element.getBoundingClientRect();
                    // Check if there's a shape element with the same text at similar position
                    const hasShapeWithSameText = Array.from(document.querySelectorAll('*')).some(other => {
                        if (other === element) return false;
                        const otherStyles = window.getComputedStyle(other);
                        const otherText = other.textContent ? other.textContent.trim() : '';
                        const otherRect = other.getBoundingClientRect();
                        
                        return (otherStyles.clipPath && otherStyles.clipPath.startsWith('polygon')) &&
                               otherText === elementText &&
                               Math.abs(otherRect.left - rect.left) < 50 &&
                               Math.abs(otherRect.top - rect.top) < 50;
                    });
                    
                    if (hasShapeWithSameText) {
                        return false;
                    }
                }
                
                return true;
            }

            function detectRedHighlightWrapper(container, slideContainer) {
                // Look for red highlight wrapper child elements
                const redHighlightWrappers = Array.from(container.children).filter(child => {
                    const childClass = child.className || '';
                    const childStyles = window.getComputedStyle(child);
                    
                    // Check for red highlight wrapper pattern
                    return (
                        childClass.includes('red-highlight-wrapper') ||
                        childClass.includes('table-row-highlight-wrapper') ||
                        (childStyles.position === 'absolute' &&
                         childStyles.borderColor.includes('255, 0, 0') && // Red border
                         (childStyles.borderStyle === 'dashed' || childStyles.borderStyle === 'dotted') &&
                         childStyles.backgroundColor === 'rgba(0, 0, 0, 0)') // Transparent background
                    );
                });
                
                if (redHighlightWrappers.length === 0) {
                    return null;
                }
                
                const wrapper = redHighlightWrappers[0];
                const wrapperStyles = window.getComputedStyle(wrapper);
                const containerRect = container.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                
                // Get the actual wrapper rectangle (it should already be positioned correctly)
                const wrapperRect = wrapper.getBoundingClientRect();
                
                // Handle table row highlight wrappers with fixed width
                if (wrapper.className.includes('table-row-highlight-wrapper')) {
                    // For table row highlights, find the parent row and table
                    const parentCell = wrapper.closest('td');
                    const parentRow = wrapper.closest('tr');
                    const parentTable = wrapper.closest('table');
                    
                    if (parentRow && parentTable) {
                        const rowRect = parentRow.getBoundingClientRect();
                        const tableRect = parentTable.getBoundingClientRect();
                        
                        // Use table-based positioning for row highlights - extend across full table width
                        const extendedRect = {
                            x: Math.round(tableRect.left - slideRect.left - 12), // Left extension beyond table
                            y: Math.round(rowRect.top - slideRect.top - 3), // Top extension above row
                            width: Math.round(tableRect.width + 24), // Full table width + extensions
                            height: Math.round(rowRect.height + 6) // Row height + extensions
                        };
                        
                        console.log(`Table row highlight - Table: x=${tableRect.left}, y=${tableRect.top}, w=${tableRect.width}, Row: y=${rowRect.top}, h=${rowRect.height}`);
                        console.log(`Extended rect: x=${extendedRect.x}, y=${extendedRect.y}, w=${extendedRect.width}, h=${extendedRect.height}`);
                        
                        return {
                            hasRedHighlight: true,
                            extendedRect: extendedRect,
                            borderColor: wrapperStyles.borderColor,
                            borderWidth: wrapperStyles.borderWidth,
                            borderStyle: wrapperStyles.borderStyle,
                            wrapperClass: wrapper.className
                        };
                    }
                }
                
                // Use the wrapper's actual rendered position and size relative to slide
                const extendedRect = {
                    x: Math.round(wrapperRect.left - slideRect.left),
                    y: Math.round(wrapperRect.top - slideRect.top),
                    width: Math.round(wrapperRect.width),
                    height: Math.round(wrapperRect.height)
                };
                
                // Validate that the rectangle makes sense
                if (extendedRect.width <= 0 || extendedRect.height <= 0) {
                    return null;
                }
                
                return {
                    hasRedHighlight: true,
                    extendedRect: extendedRect,
                    borderColor: wrapperStyles.borderColor,
                    borderWidth: wrapperStyles.borderWidth,
                    borderStyle: wrapperStyles.borderStyle,
                    wrapperClass: wrapper.className
                };
            }

            function shouldExtractInlineGroup(element, processedTextElements) {
                const elementId = getElementId(element);
                if (processedTextElements.has(elementId)) {
                    return false;
                }

                const tagName = element.tagName.toLowerCase();
                if (!['p', 'div', 'li', 'td', 'th'].includes(tagName)) return false;
                
                if (element.children.length > 10) {
                    return false;
                }
                
                const hasInlineFormatting = Array.from(element.children).some(child => {
                    const childTag = child.tagName.toLowerCase();
                    return ['strong', 'b', 'em', 'i', 'u', 'mark', 'span'].includes(childTag);
                });
                
                // Check for red highlight wrapper
                const hasRedHighlightWrapper = Array.from(element.children).some(child => {
                    const childClass = child.className || '';
                    const childStyles = window.getComputedStyle(child);
                    return (
                        childClass.includes('red-highlight-wrapper') ||
                        childClass.includes('table-row-highlight-wrapper') ||
                        (childStyles.position === 'absolute' &&
                         childStyles.borderColor.includes('255, 0, 0') &&
                         (childStyles.borderStyle === 'dashed' || childStyles.borderStyle === 'dotted'))
                    );
                });
                
                if (!hasInlineFormatting && !hasRedHighlightWrapper) return false;
                
                const structuredChildren = Array.from(element.children).filter(child => {
                    const childTag = child.tagName.toLowerCase();
                    const childClass = child.className || '';
                    return ['div', 'ul', 'ol', 'table', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(childTag) &&
                           !childClass.includes('red-highlight-wrapper');
                });
               
                if (structuredChildren.length > 0) {
                    return false;
                }
                
                const nonInlineChildren = Array.from(element.children).filter(child => {
                    const childTag = child.tagName.toLowerCase();
                    const childClass = child.className || '';
                    return !['strong', 'b', 'em', 'i', 'u', 'mark', 'span', 'br', 'div'].includes(childTag) ||
                           (childTag === 'div' && !childClass.includes('red-highlight-wrapper'));
                });
                
                if (nonInlineChildren.length > 0) {
                    return false;
                }
                
                return true;
            }

            function getInlineGroup(container, slideContainer, processedTextElements) {
                if (!shouldExtractInlineGroup(container, processedTextElements)) {
                    return null;
                }
                
                const hasInlineFormatting = Array.from(container.children).some(child => {
                    const childTag = child.tagName.toLowerCase();
                    return ['strong', 'b', 'em', 'i', 'u', 'mark', 'span'].includes(childTag);
                });
                
                // Check for red highlight wrapper
                const redHighlightInfo = detectRedHighlightWrapper(container, slideContainer);
                
                if (!hasInlineFormatting && !redHighlightInfo) return null;
                
                const rect = container.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                const inlineElements = [];
                let fullText = '';
                const walkNodes = (node) => {
                    if (node.nodeType === Node.TEXT_NODE) {
                        let text = node.textContent.replace(/\s+/g, ' ');
                        if (text.trim() === '') return;
                        inlineElements.push({
                            type: 'text',
                            text: text,
                            styles: extractComprehensiveStyles(container)
                        });
                        fullText += text;
                    } else if (node.nodeType === Node.ELEMENT_NODE) {
                        const childTag = node.tagName.toLowerCase();
                        
                        // Skip red highlight wrapper elements - they're handled separately
                        if (node.className && (node.className.includes('red-highlight-wrapper') || node.className.includes('table-row-highlight-wrapper'))) {
                            return;
                        }
                        
                        if (['strong', 'b', 'em', 'i', 'u', 'mark', 'span'].includes(childTag)) {
                            let text = node.textContent;
                            const lines = text.split('\n');
                            const normalizedLines = [lines[0], ...lines.slice(1).map(l => l.replace(/^[ \t]+/, ''))];
                            text = normalizedLines.join('\n');
                            if (text.trim() === '') return;
                            inlineElements.push({
                                type: childTag,
                                text: text,
                                styles: extractComprehensiveStyles(node)
                            });
                            fullText += text;
                        } else if (childTag === 'br') {
                            inlineElements.push({
                                type: 'br',
                                text: '\n',
                                styles: extractComprehensiveStyles(container)
                            });
                            fullText += '\n';
                        } else {
                            for (const child of node.childNodes) {
                                walkNodes(child);
                            }
                        }
                    }
                };
                for (const node of container.childNodes) {
                    walkNodes(node);
                }
                const filteredInline = inlineElements.filter(el => el.text.trim() !== '' || el.type === 'br');
                if (filteredInline.length > 0 || redHighlightInfo) {
                    if (filteredInline[0] && filteredInline[0].type === 'text') {
                        filteredInline[0].text = filteredInline[0].text.replace(/^\s+/, '');
                    }
                    const lastIndex = filteredInline.length - 1;
                    if (filteredInline[lastIndex] && filteredInline[lastIndex].type === 'text') {
                        filteredInline[lastIndex].text = filteredInline[lastIndex].text.replace(/\s+$/, '');
                    }
                    fullText = filteredInline.map(el => el.text).join('');
                    const result = {
                        text: fullText,
                        inlineElements: filteredInline,
                        groupRect: {
                            x: Math.round(rect.left - slideRect.left),
                            y: Math.round(rect.top - slideRect.top),
                            width: Math.round(rect.width),
                            height: Math.round(rect.height)
                        },
                        styles: extractComprehensiveStyles(container)
                    };
                    
                    // Add red highlight information if present
                    if (redHighlightInfo) {
                        result.hasRedHighlight = true;
                        result.redHighlightStyles = redHighlightInfo;
                    }
                    
                    return result;
                }
                return null;
            }

            function getListInfo(element, slideContainer) {
                const tagName = element.tagName.toLowerCase();
                const listInfo = {};
                if (['ul', 'ol'].includes(tagName)) {
                    const items = Array.from(element.querySelectorAll(':scope > li'));
                    const styles = window.getComputedStyle(element);
                    const rect = element.getBoundingClientRect();
                    const slideRect = slideContainer.getBoundingClientRect();
                    listInfo.type = tagName;
                    listInfo.itemCount = items.length;
                    listInfo.rect = {
                        x: Math.round(rect.left - slideRect.left),
                        y: Math.round(rect.top - slideRect.top),
                        width: Math.round(rect.width),
                        height: Math.round(rect.height)
                    };
                    listInfo.listStyles = {
                        listStyleType: styles.listStyleType,
                        listStylePosition: styles.listStylePosition,
                        paddingLeft: styles.paddingLeft,
                        marginTop: styles.marginTop,
                        marginBottom: styles.marginBottom
                    };
                    listInfo.items = items.map((item, index) => {
                        const itemRect = item.getBoundingClientRect();
                        const itemStyles = extractComprehensiveStyles(item);
                        const text = item.textContent.trim();
                        const inlineGroup = getInlineGroup(item, slideContainer, new Set());
                        const nestedListElement = item.querySelector(':scope > ul, :scope > ol');
                        const nestedList = nestedListElement ? getListInfo(nestedListElement, slideContainer) : null;
                        const beforePseudo = extractPseudo(item, slideContainer, '::before');
                        return {
                            index,
                            text: text,
                            styles: itemStyles,
                            rect: {
                                x: Math.round(itemRect.left - slideRect.left),
                                y: Math.round(itemRect.top - slideRect.top),
                                width: Math.round(itemRect.width),
                                height: Math.round(itemRect.height)
                            },
                            inlineGroup: inlineGroup,
                            nestedList: nestedList,
                            hasNestedList: !!nestedListElement,
                            bulletInfo: beforePseudo ? {
                                content: beforePseudo.text,
                                position: beforePseudo,
                                styles: beforePseudo.styles
                            } : null
                        };
                    });
                    if (tagName === 'ol') {
                        listInfo.start = element.start || 1;
                        listInfo.reversed = element.reversed || false;
                    }
                }
                return listInfo;
            }

            function getTableInfo(element, slideContainer) {
                const tagName = element.tagName.toLowerCase();
                const tableInfo = {};
                if (tagName === 'table') {
                    const rows = Array.from(element.querySelectorAll('tr'));
                    const rect = element.getBoundingClientRect();
                    const slideRect = slideContainer.getBoundingClientRect();
                    tableInfo.type = 'table';
                    tableInfo.rect = {
                        x: Math.round((rect.left - slideRect.left) * 10) / 10,
                        y: Math.round((rect.top - slideRect.top) * 10) / 10,
                        width: Math.round(rect.width * 10) / 10,
                        height: Math.round(rect.height * 10) / 10
                    };
                    tableInfo.rowCount = rows.length;
                    tableInfo.styles = extractComprehensiveStyles(element);
                    let maxCols = 0;
                    rows.forEach(row => {
                        const cells = Array.from(row.querySelectorAll('td, th'));
                        let colCount = 0;
                        cells.forEach(cell => colCount += cell.colSpan || 1);
                        maxCols = Math.max(maxCols, colCount);
                    });
                    tableInfo.columnCount = maxCols;
                    tableInfo.rows = rows.map((row, rowIndex) => {
                        const cells = Array.from(row.querySelectorAll('td, th'));
                        const rowRect = row.getBoundingClientRect();
                        const rowStyles = extractComprehensiveStyles(row);
                        let actualRowHeight = Math.round(rowRect.height * 10) / 10;
                        const inlineStyle = row.getAttribute('style');
                        if (inlineStyle && inlineStyle.includes('height:')) {
                            const heightMatch = inlineStyle.match(/height:\s*(\d+(?:\.\d+)?)px/);
                            if (heightMatch) {
                                const specifiedHeight = parseFloat(heightMatch[1]);
                                actualRowHeight = Math.min(specifiedHeight, actualRowHeight);
                            }
                        }
                        actualRowHeight = Math.max(12, actualRowHeight);
                        return {
                            index: rowIndex,
                            rect: {
                                x: Math.round((rowRect.left - slideRect.left) * 10) / 10,
                                y: Math.round((rowRect.top - slideRect.top) * 10) / 10,
                                width: Math.round(rowRect.width * 10) / 10,
                                height: actualRowHeight
                            },
                            styles: rowStyles,
                            cells: cells.map((cell, cellIndex) => {
                                const cellRect = cell.getBoundingClientRect();
                                const inlineGroup = getInlineGroup(cell, slideContainer, new Set());
                                const text = inlineGroup ? '' : getTextContent(cell, new Set());
                                return {
                                    type: cell.tagName.toLowerCase(),
                                    text: text,
                                    rect: {
                                        x: Math.round((cellRect.left - slideRect.left) * 10) / 10,
                                        y: Math.round((cellRect.top - slideRect.top) * 10) / 10,
                                        width: Math.round(cellRect.width * 10) / 10,
                                        height: Math.round(cellRect.height * 10) / 10
                                    },
                                    styles: extractComprehensiveStyles(cell),
                                    colSpan: cell.colSpan || 1,
                                    rowSpan: cell.rowSpan || 1,
                                    cellIndex,
                                    inlineGroup
                                };
                            })
                        };
                    });
                }
                return tableInfo;
            }

            function getAccurateImageDimensions(img, container, slideContainer) {
                const imgRect = img.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                return {
                    x: Math.round(imgRect.left - slideRect.left),
                    y: Math.round(imgRect.top - slideRect.top),
                    width: Math.round(imgRect.width),
                    height: Math.round(imgRect.height),
                    naturalWidth: img.naturalWidth || 0,
                    naturalHeight: img.naturalHeight || 0
                };
            }

            function getAccuratePosition(element, slideContainer) {
                const rect = element.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                const styles = window.getComputedStyle(element);
                let x = rect.left - slideRect.left;
                let y = rect.top - slideRect.top;
                let width = rect.width;
                let height = rect.height;
                
                const devicePixelRatio = window.devicePixelRatio || 1;
                x /= devicePixelRatio;
                y /= devicePixelRatio;
                width /= devicePixelRatio;
                height /= devicePixelRatio;
                
                if (styles.display === 'inline' || styles.display === 'inline-block') {
                    try {
                        const range = document.createRange();
                        range.selectNodeContents(element);
                        const rangeRect = range.getBoundingClientRect();
                        if (rangeRect.width > 0 && rangeRect.height > 0) {
                            x = (rangeRect.left - slideRect.left) / devicePixelRatio;
                            y = (rangeRect.top - slideRect.top) / devicePixelRatio;
                            width = rangeRect.width / devicePixelRatio;
                            height = rangeRect.height / devicePixelRatio;
                        }
                    } catch (e) {}
                }
                
                const parent = element.parentElement;
                const parentStyles = parent ? window.getComputedStyle(parent) : null;
                if (parentStyles && (parentStyles.display === 'flex' || parentStyles.display === 'inline-flex')) {
                    const computedRect = element.getBoundingClientRect();
                    x = (computedRect.left - slideRect.left) / devicePixelRatio;
                    y = (computedRect.top - slideRect.top) / devicePixelRatio;
                    width = computedRect.width / devicePixelRatio;
                    height = computedRect.height / devicePixelRatio;
                }
                
                const transform = styles.transform;
                if (transform && transform !== 'none') {
                    try {
                        const matrix = new DOMMatrix(transform);
                        if (matrix.a !== 1 || matrix.d !== 1) {
                            width *= Math.abs(matrix.a);
                            height *= Math.abs(matrix.d);
                        }
                        x += matrix.e / devicePixelRatio;
                        y += matrix.f / devicePixelRatio;
                    } catch (e) {}
                }
                return {
                    x: Math.round(x * 10) / 10,
                    y: Math.round(y * 10) / 10,
                    width: Math.max(1, Math.round(width * 10) / 10),
                    height: Math.max(1, Math.round(height * 10) / 10)
                };
            }

            function extractComprehensiveStyles(elementOrStyle) {
                let styles;
                let isPseudo = false;
                if (elementOrStyle instanceof Element) {
                    styles = window.getComputedStyle(elementOrStyle);
                } else {
                    styles = elementOrStyle;
                    isPseudo = true;
                }
                const customProperties = {};
                for (const prop of styles) {
                    if (prop.startsWith('--')) customProperties[prop] = styles.getPropertyValue(prop);
                }
                const borderProps = {};
                ['Top', 'Right', 'Bottom', 'Left'].forEach(side => {
                    borderProps[`border${side}Width`] = styles[`border${side}Width`];
                    borderProps[`border${side}Style`] = styles[`border${side}Style`];
                    borderProps[`border${side}Color`] = styles[`border${side}Color`];
                });
                return {
                    fontSize: styles.fontSize,
                    fontFamily: styles.fontFamily,
                    fontWeight: styles.fontWeight,
                    fontStyle: styles.fontStyle,
                    lineHeight: styles.lineHeight,
                    textAlign: styles.textAlign,
                    textDecoration: styles.textDecoration,
                    color: styles.color,
                    background: styles.background,
                    backgroundColor: getResolvedBackgroundColor(elementOrStyle, isPseudo),
                    width: styles.width,
                    height: styles.height,
                    padding: styles.padding,
                    paddingTop: styles.paddingTop,
                    paddingRight: styles.paddingRight,
                    paddingBottom: styles.paddingBottom,
                    paddingLeft: styles.paddingLeft,
                    margin: styles.margin,
                    marginTop: styles.marginTop,
                    marginRight: styles.marginRight,
                    marginBottom: styles.marginBottom,
                    marginLeft: styles.marginLeft,
                    border: styles.border,
                    borderWidth: styles.borderWidth,
                    borderStyle: styles.borderStyle,
                    borderColor: styles.borderColor,
                    ...borderProps,
                    borderRadius: styles.borderRadius,
                    position: styles.position,
                    display: styles.display,
                    visibility: styles.visibility,
                    zIndex: styles.zIndex,
                    boxShadow: styles.boxShadow,
                    listStyleType: styles.listStyleType,
                    listStylePosition: styles.listStylePosition,
                    listStyleImage: styles.listStyleImage,
                    overflow: styles.overflow,
                    overflowX: styles.overflowX,
                    overflowY: styles.overflowY,
                    flex: styles.flex,
                    flexDirection: styles.flexDirection,
                    justifyContent: styles.justifyContent,
                    alignItems: styles.alignItems,
                    gap: styles.gap,
                    left: styles.left,
                    top: styles.top,
                    right: styles.right,
                    bottom: styles.bottom,
                    content: styles.content,
                    objectFit: styles.objectFit,
                    objectPosition: styles.objectPosition,
                    maxWidth: styles.maxWidth,
                    maxHeight: styles.maxHeight,
                    minWidth: styles.minWidth,
                    minHeight: styles.minHeight,
                    transform: styles.transform,
                    transformOrigin: styles.transformOrigin,
                    clipPath: styles.clipPath,
                    flexShrink: styles.flexShrink,
                    pointerEvents: styles.pointerEvents,
                    customProperties
                };
            }

            function getPseudoPosition(element, slideContainer, pseudoStyles, pseudo) {
                const parentRect = element.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                let width = safeFloat(pseudoStyles.width);
                let height = safeFloat(pseudoStyles.height);
                const leftRaw = pseudoStyles.left;
                const rightRaw = pseudoStyles.right;
                const topRaw = pseudoStyles.top;
                const bottomRaw = pseudoStyles.bottom;
                const left = leftRaw === 'auto' ? null : safeFloat(leftRaw);
                const right = rightRaw === 'auto' ? null : safeFloat(rightRaw);
                const top = topRaw === 'auto' ? null : safeFloat(topRaw);
                const bottom = bottomRaw === 'auto' ? null : safeFloat(bottomRaw);
                if (width === 0 && left !== null && right !== null) {
                    width = parentRect.width - left - right;
                } else if (width === 0 && right !== null) {
                    const elementStyles = window.getComputedStyle(element);
                    const parentWidth = parentRect.width;
                    width = parentWidth - right;
                } else if (width === 0) {
                    if (pseudo === '::before' && element.tagName.toLowerCase() === 'li') {
                        width = safeFloat(pseudoStyles.fontSize) || 14;
                    } else {
                        width = parentRect.width;
                    }
                }
                if (height === 0 && top !== null && bottom !== null) {
                    height = parentRect.height - top - bottom;
                } else if (height === 0) {
                    if (pseudo === '::after') {
                        height = 2;
                    } else if (pseudo === '::before' && element.tagName.toLowerCase() === 'li') {
                        height = safeFloat(pseudoStyles.fontSize) || 14;
                    } else {
                        height = 2;
                    }
                }
                let x;
                if (left !== null) {
                    x = left;
                } else if (right !== null) {
                    x = parentRect.width - width - right;
                } else {
                    x = 0;
                }
                let y;
                if (top !== null) {
                    y = top;
                } else if (bottom !== null) {
                    y = parentRect.height - height - bottom;
                } else {
                    if (pseudo === '::after') {
                        y = parentRect.height - height;
                    } else if (pseudo === '::before') {
                        y = parentRect.height - height;
                    } else {
                        y = 0;
                    }
                }
                x += parentRect.left - slideRect.left;
                y += parentRect.top - slideRect.top;

                return {
                    x: Math.round(x * 10) / 10,
                    y: Math.round(y * 10) / 10,
                    width: Math.round(width * 10) / 10,
                    height: Math.round(height * 10) / 10
                };
            }

            function extractPseudo(element, slideContainer, pseudo) {
                const pseudoStyles = window.getComputedStyle(element, pseudo);
                if (pseudoStyles.display === 'none' || pseudoStyles.visibility === 'hidden' || pseudoStyles.opacity === '0') return null;
                const content = pseudoStyles.content.replace(/['"]/g, '').trim();
                const hasContent = content !== '' && content !== 'none';
                const hasBackground = pseudoStyles.backgroundColor !== 'rgba(0, 0, 0, 0)' && pseudoStyles.backgroundColor !== 'transparent';
                const hasGradient = pseudoStyles.background && pseudoStyles.background.includes('gradient');
                const hasBorder = ['borderTopWidth', 'borderRightWidth', 'borderBottomWidth', 'borderLeftWidth']
                    .some(prop => safeFloat(pseudoStyles[prop]) > 0);
                const hasBoxShadow = pseudoStyles.boxShadow !== 'none';
                if (!hasContent && !hasBackground && !hasGradient && !hasBorder && !hasBoxShadow) return null;
                const position = getPseudoPosition(element, slideContainer, pseudoStyles, pseudo);
                if (!position || (position.width <= 0 && position.height <= 0)) return null;
                const styles = extractComprehensiveStyles(pseudoStyles);
                
                const parentZ = parseInt(window.getComputedStyle(element).zIndex) || 0;
                let assignedZ = parentZ - 2;
                if (pseudo === '::after') assignedZ = parentZ - 1;
                const elementData = {
                    type: 'pseudo',
                    pseudoType: pseudo,
                    x: position.x,
                    y: position.y,
                    width: position.width,
                    height: position.height,
                    styles,
                    zIndex: assignedZ,
                    parentClassName: element.className || '',
                    parentTagName: element.tagName.toLowerCase()
                };
                if (hasContent) {
                    elementData.text = content;
                }
                return elementData;
            }

            function extractShapeInfo(element, slideContainer) {
                const styles = window.getComputedStyle(element);
                const isShapeByClipPath = styles.clipPath && styles.clipPath.startsWith('polygon');
                const isShapeByTransform = styles.transform !== 'none' && styles.transform.includes('matrix');

                if (!isShapeByClipPath && !isShapeByTransform) return null;

                const rect = element.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                
                let rotation = 0;
                const transform = styles.transform;
                if (transform && transform !== 'none') {
                    try {
                        const matrix = new DOMMatrix(transform);
                        rotation = Math.round(Math.atan2(matrix.b, matrix.a) * (180 / Math.PI));
                    } catch (e) {
                        const match = transform.match(/rotate\(([-\d.]+)deg\)/);
                        if (match) {
                            rotation = parseFloat(match[1]);
                        }
                    }
                }

                const width = element.offsetWidth;
                const height = element.offsetHeight;

                const centerX = rect.left - slideRect.left + rect.width / 2;
                const centerY = rect.top - slideRect.top + rect.height / 2;
                const x = centerX - width / 2;
                const y = centerY - height / 2;

                const elementText = element.textContent ? element.textContent.trim() : '';

                return {
                    type: 'shape',
                    rect: {
                        x: Math.round(x),
                        y: Math.round(y),
                        width: Math.round(width),
                        height: Math.round(height)
                    },
                    styles: extractComprehensiveStyles(element),
                    className: element.className || '',
                    rotation: rotation,
                    clipPath: styles.clipPath,
                    text: elementText,
                    elementId: getElementId(element)
                };
            }

            function extractOverlayInfo(element, slideContainer) {
                const styles = window.getComputedStyle(element);
                const elementClass = element.className || '';
                const rect = element.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                
                // Special handling for red highlight wrappers
                if (elementClass.includes('red-highlight-wrapper') || elementClass.includes('table-row-highlight-wrapper')) {
                    // Find the parent list item or table row
                    const parentLi = element.closest('li');
                    const parentRow = element.closest('tr');
                    
                    if (parentLi) {
                        const parentRect = parentLi.getBoundingClientRect();
                        
                        // Calculate the extended wrapper position based on CSS positioning
                        const top = safeFloat(styles.top) || -4;
                        const left = safeFloat(styles.left) || -30;
                        const right = safeFloat(styles.right) || -20;
                        const bottom = safeFloat(styles.bottom) || -4;
                        
                        // Calculate wrapper dimensions relative to parent li
                        const wrapperRect = {
                            x: Math.round(parentRect.left - slideRect.left + left),
                            y: Math.round(parentRect.top - slideRect.top + top),
                            width: Math.round(parentRect.width - left - right),
                            height: Math.round(parentRect.height - top - bottom)
                        };
                        
                        return {
                            type: 'overlay',
                            rect: wrapperRect,
                            styles: extractComprehensiveStyles(element),
                            className: element.className || '',
                            zIndex: parseInt(styles.zIndex) || 0,
                            pointerEvents: styles.pointerEvents
                        };
                    } else if (parentRow) {
                        // Handle table row highlight wrapper
                        const parentTable = element.closest('table');
                        if (parentTable) {
                            const rowRect = parentRow.getBoundingClientRect();
                            const tableRect = parentTable.getBoundingClientRect();
                            
                            // Calculate wrapper dimensions based on table and row
                            const wrapperRect = {
                                x: Math.round(tableRect.left - slideRect.left - 12),
                                y: Math.round(rowRect.top - slideRect.top - 3),
                                width: Math.round(tableRect.width + 24),
                                height: Math.round(rowRect.height + 6)
                            };
                            
                            console.log(`Overlay table row highlight - Row: ${rowRect.top}, Table: ${tableRect.left}, Extended: x=${wrapperRect.x}, y=${wrapperRect.y}, w=${wrapperRect.width}, h=${wrapperRect.height}`);
                            
                            return {
                                type: 'overlay',
                                rect: wrapperRect,
                                styles: extractComprehensiveStyles(element),
                                className: element.className || '',
                                zIndex: parseInt(styles.zIndex) || 0,
                                pointerEvents: styles.pointerEvents
                            };
                        }
                    }
                }
                
                // Standard overlay handling
                if (styles.borderStyle === 'none' && styles.backgroundColor === 'rgba(0, 0, 0, 0)' && !styles.boxShadow) return null;
                
                return {
                    type: 'overlay',
                    rect: {
                        x: Math.round((rect.left - slideRect.left) * 10) / 10,
                        y: Math.round((rect.top - slideRect.top) * 10) / 10,
                        width: Math.round(rect.width * 10) / 10,
                        height: Math.round(rect.height * 10) / 10
                    },
                    styles: extractComprehensiveStyles(element),
                    className: element.className || '',
                    zIndex: parseInt(styles.zIndex) || 0,
                    pointerEvents: styles.pointerEvents
                };
            }

            function isRedundantInlineElement(element) {
                const tag = element.tagName.toLowerCase();
                if (!['span', 'strong', 'b', 'em', 'i', 'u', 'mark'].includes(tag)) return false;
                const parent = element.parentElement;
                if (!parent) return false;
                const parentTag = parent.tagName.toLowerCase();
                
                if (['div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'td', 'th'].includes(parentTag)) {
                    const elementStyles = window.getComputedStyle(element);
                    const parentStyles = window.getComputedStyle(parent);
                    
                    const hasDifferentStyling = (
                        elementStyles.fontWeight !== parentStyles.fontWeight ||
                        elementStyles.fontStyle !== parentStyles.fontStyle ||
                        elementStyles.color !== parentStyles.color ||
                        elementStyles.backgroundColor !== parentStyles.backgroundColor ||
                        elementStyles.textDecoration !== parentStyles.textDecoration
                    );
                    
                    return !hasDifferentStyling;
                }
                return false;
            }

            function shouldProcessElement(element, slideContainer) {
                if (element === slideContainer) return false;
                if (!slideContainer.contains(element)) return false;
                const tagName = element.tagName.toLowerCase();
                if (!IMPORTANT_ELEMENTS.includes(tagName)) return false;
                const styles = window.getComputedStyle(element);
                if (styles.display === 'none' || styles.visibility === 'hidden' || styles.opacity === '0') return false;
                
                const hasBorder = ['borderTopWidth', 'borderRightWidth', 'borderBottomWidth', 'borderLeftWidth']
                    .some(prop => parseFloat(styles[prop]) > 0 && styles[prop.replace('Width', 'Style')] !== 'none');
                    
                const hasBackground = styles.backgroundColor !== 'rgba(0, 0, 0, 0)' && 
                                    styles.backgroundColor !== 'transparent' &&
                                    styles.backgroundColor !== 'initial';
                                    
                const hasShadow = styles.boxShadow !== 'none';
                const hasGradient = styles.background && styles.background.includes('gradient');
                
                if (hasBorder || hasBackground || hasShadow || hasGradient) {
                    return true;
                }
                
                if (['img', 'canvas', 'svg'].includes(tagName)) {
                    return true;
                }
                
                if (styles.clipPath.includes('polygon')) {
                    return true;
                }
                
                if (['strong', 'b', 'i', 'em', 'u', 'mark', 'li', 'td', 'th', 'span'].includes(tagName)) {
                    return true;
                }
                
                if (tagName === 'div') {
                    const parent = element.parentElement;
                    const parentStyles = parent ? window.getComputedStyle(parent) : null;
                    
                    if (parentStyles && (parentStyles.display === 'flex' || parentStyles.display === 'inline-flex')) {
                        return true;
                    }
                    
                    if (styles.display === 'flex' || styles.display === 'inline-flex') {
                        return true;
                    }
                    
                    if (Array.from(element.classList).some(cls => 
                        cls.includes('chart') || cls.includes('chartjs-chart')) || styles.clipPath.includes('polygon')) {
                        return true;
                    }
                    
                    if (styles.borderStyle !== 'none' && parseFloat(styles.borderWidth) > 0) {
                        return true;
                    }
                }
                
                const text = element.textContent ? element.textContent.trim() : '';
                if (!text && !['img', 'svg', 'video', 'audio', 'iframe', 'hr', 'br', 'canvas'].includes(tagName)) {
                    const rect = element.getBoundingClientRect();
                    if (rect.width === 0 && rect.height === 0) {
                        return false;
                    }
                    
                    if (tagName === 'div' && !hasBorder && !hasBackground && !hasShadow) {
                        return false;
                    }
                }
                
                return true;
            }

            function isOverlayElement(element) {
                const styles = window.getComputedStyle(element);
                const position = styles.position;
                const hasBorder = (
                    (parseFloat(styles.borderTopWidth) > 0 && styles.borderTopStyle !== 'none') ||
                    (parseFloat(styles.borderRightWidth) > 0 && styles.borderRightStyle !== 'none') ||
                    (parseFloat(styles.borderBottomWidth) > 0 && styles.borderBottomStyle !== 'none') ||
                    (parseFloat(styles.borderLeftWidth) > 0 && styles.borderLeftStyle !== 'none')
                );
                const hasZ = parseInt(styles.zIndex) > 0;
                const tag = element.tagName.toLowerCase();
                const elementClass = element.className || '';
                const isContentTag = [
                    'li', 'p', 'span', 'strong', 'b', 'em', 'i', 'u', 'mark'
                ].includes(tag);
                
                // Check for red highlight wrapper pattern
                const isRedHighlightWrapper = (
                    elementClass.includes('red-highlight-wrapper') ||
                    elementClass.includes('table-row-highlight-wrapper') ||
                    elementClass.includes('table-row-highlight-wrapper') ||
                    (position === 'absolute' &&
                     styles.borderColor.includes('255, 0, 0') && // Red border
                     (styles.borderStyle === 'dashed' || styles.borderStyle === 'dotted') &&
                     styles.backgroundColor === 'rgba(0, 0, 0, 0)') // Transparent background
                );
                
                return (
                    ((position === 'absolute' || position === 'fixed' || position === 'relative') &&
                     (hasBorder || styles.pointerEvents === 'none') &&
                     !isContentTag) ||
                    isRedHighlightWrapper
                );
            }

            function getAllDescendants(element) {
                let descendants = [];
                function traverse(el) {
                    if (!el || !el.children) return;
                    for (let child of el.children) {
                        if (child && child.tagName) {
                            descendants.push(child);
                            traverse(child);
                        }
                    }
                }
                traverse(element);
                return descendants;
            }

            function extractFooterInfo(element, slideContainer) {
                const styles = window.getComputedStyle(element);
                const isFooter = (
                    element.className.includes('footer') ||
                    element.tagName.toLowerCase() === 'footer' ||
                    (styles.display === 'flex' && 
                     (styles.justifyContent === 'center' || styles.justifyContent === 'space-between') &&
                     styles.alignItems === 'center')
                );
                
                if (!isFooter) return null;
                
                const footerRect = element.getBoundingClientRect();
                const slideRect = slideContainer.getBoundingClientRect();
                const footerElements = [];
                
                // Get all child elements (spans, images, etc.)
                const children = Array.from(element.children);
                
                if (children.length === 0) {
                    // If no children, check for direct text content
                    const text = element.textContent ? element.textContent.trim() : '';
                    if (text) {
                        footerElements.push({
                            type: 'text',
                            text: text,
                            x: Math.round((footerRect.left - slideRect.left) * 10) / 10,
                            y: Math.round((footerRect.top - slideRect.top) * 10) / 10,
                            width: Math.round(footerRect.width * 10) / 10,
                            height: Math.round(footerRect.height * 10) / 10,
                            styles: extractComprehensiveStyles(element),
                            className: element.className || '',
                            originalIndex: 0
                        });
                    }
                } else {
                    // Process all child elements
                    children.forEach((child, index) => {
                        const childRect = child.getBoundingClientRect();
                        const childStyles = extractComprehensiveStyles(child);
                        const tagName = child.tagName.toLowerCase();
                        
                        const elementData = {
                            type: tagName,
                            x: Math.round((childRect.left - slideRect.left) * 10) / 10,
                            y: Math.round((childRect.top - slideRect.top) * 10) / 10,
                            width: Math.round(childRect.width * 10) / 10,
                            height: Math.round(childRect.height * 10) / 10,
                            styles: childStyles,
                            className: child.className || '',
                            originalIndex: index
                        };
                        
                        // Handle different element types
                        if (tagName === 'img') {
                            elementData.mediaInfo = {
                                src: child.src || '',
                                alt: child.alt || '',
                                naturalWidth: child.naturalWidth || 0,
                                naturalHeight: child.naturalHeight || 0,
                                currentWidth: Math.round(childRect.width * 10) / 10,
                                currentHeight: Math.round(childRect.height * 10) / 10
                            };
                        } else {
                            // For text elements (span, div, etc.)
                            const text = child.textContent ? child.textContent.trim() : '';
                            if (text) {
                                elementData.text = text;
                            }
                        }
                        
                        footerElements.push(elementData);
                    });
                }
                
                return {
                    type: 'footer',
                    footerElements: footerElements,
                    containerRect: {
                        x: Math.round((footerRect.left - slideRect.left) * 10) / 10,
                        y: Math.round((footerRect.top - slideRect.top) * 10) / 10,
                        width: Math.round(footerRect.width * 10) / 10,
                        height: Math.round(footerRect.height * 10) / 10
                    },
                    containerStyles: extractComprehensiveStyles(element),
                    justifyContent: styles.justifyContent,
                    alignItems: styles.alignItems,
                    flexDirection: styles.flexDirection || 'row'
                };
            }

            for (let slideIndex = 0; slideIndex < slideElements.length; slideIndex++) {
                const slideElement = slideElements[slideIndex];
                const slideRect = slideElement.getBoundingClientRect();
                const slide = {
                    slideId: slideIndex + 1,
                    elements: [],
                    slideWidth: parseFloat(slideRect.width.toFixed(2)),
                    slideHeight: parseFloat(slideRect.height.toFixed(2)),
                    slidePosition: {
                        x: parseFloat(slideRect.left.toFixed(2)),
                        y: parseFloat(slideRect.top.toFixed(2))
                    },
                    slideStyles: extractComprehensiveStyles(slideElement)
                };
                const processedElements = new Set();
                const processedTextElements = new Set();
                const allElements = Array.from(slideElement.querySelectorAll('*'));
                const elementsToProcess = allElements.filter(element => {
                    if (isRedundantInlineElement(element)) return false;
                    return shouldProcessElement(element, slideElement);
                });
                const overlayElements = [];
                
                // First, scan for table row highlights with improved accuracy and original CSS extraction
                const foundHighlights = new Set();
                slideElement.querySelectorAll('.table-row-highlight-wrapper').forEach(wrapper => {
                    const parentCell = wrapper.closest('td');
                    const parentRow = wrapper.closest('tr');
                    const parentTable = wrapper.closest('table');
                    
                    if (parentRow && parentTable && parentCell) {
                        // Get the text from the first cell to identify the row
                        const firstCell = parentRow.querySelector('td');
                        const rowText = firstCell ? firstCell.textContent.trim() : '';
                        
                        // Only process if this is one of the target highlight rows and we haven't seen it
                        const targetTexts = [
                            'ネット有利子負債(除く現預金・短期性有価証券)',
                            'ROE',
                            '株主資本比率',
                            '一人当たりの売上高'
                        ];
                        
                        if (!targetTexts.includes(rowText) || foundHighlights.has(rowText)) {
                            return; // Skip this highlight
                        }
                        
                        foundHighlights.add(rowText); // Mark as found
                        
                        const wrapperStyles = window.getComputedStyle(wrapper);
                        const tableRect = parentTable.getBoundingClientRect();
                        const rowRect = parentRow.getBoundingClientRect();
                        const slideRect = slideElement.getBoundingClientRect();
                        
                        // Get row index within the table
                        const rowIndex = Array.from(parentTable.rows).indexOf(parentRow);
                        
                        // Extract styles exactly like .red-highlight-wrapper - use getComputedStyle and inline styles
                        const computedStyles = window.getComputedStyle(wrapper);
                        
                        // Get inline styles directly from element's style attribute
                        const inlineStyle = wrapper.getAttribute('style') || '';
                        
                        // Create styles object with the exact same structure as .red-highlight-wrapper
                        const extractedStyles = {
                            position: computedStyles.position,
                            top: computedStyles.top,
                            left: computedStyles.left,
                            right: computedStyles.right,
                            bottom: computedStyles.bottom,
                            width: computedStyles.width,
                            height: computedStyles.height,
                            border: computedStyles.border,
                            borderWidth: computedStyles.borderWidth,
                            borderStyle: computedStyles.borderStyle,
                            borderColor: computedStyles.borderColor,
                            backgroundColor: computedStyles.backgroundColor,
                            zIndex: computedStyles.zIndex,
                            pointerEvents: computedStyles.pointerEvents,
                            boxSizing: computedStyles.boxSizing
                        };
                        
                        // Calculate relative position within the slide (like .red-highlight-wrapper)
                        const parentCellRect = parentCell.getBoundingClientRect();
                        
                        // Use wrapper's actual computed position relative to the slide
                        const wrapperRect = wrapper.getBoundingClientRect();
                        const relativeX = wrapperRect.left - slideRect.left;
                        const relativeY = wrapperRect.top - slideRect.top;
                        const relativeWidth = wrapperRect.width;
                        const relativeHeight = wrapperRect.height;
                        
                        // Determine color based on the specific row
                        let borderColor = extractedStyles.borderColor;
                        if (rowText === '一人当たりの売上高') {
                            // Override for blue highlight (like in HTML)
                            borderColor = 'rgb(0, 102, 255)';
                            extractedStyles.border = extractedStyles.border.replace(/rgb\([^)]+\)/, borderColor);
                        }
                        
                        console.log(`Found table row highlight: "${rowText}" at rowIndex=${rowIndex}`);
                        console.log(`Wrapper position: x=${relativeX}, y=${relativeY}, w=${relativeWidth}, h=${relativeHeight}`);
                        console.log(`Computed styles:`, extractedStyles);
                        
                        const overlayInfo = {
                            type: 'overlay',
                            className: 'table-row-highlight-wrapper',
                            x: relativeX,
                            y: relativeY,
                            width: relativeWidth,
                            height: relativeHeight,
                            targetRowIndex: rowIndex,
                            targetRowText: rowText,
                            tableY: tableRect.top - slideRect.top,
                            tableX: tableRect.left - slideRect.left,
                            tableWidth: tableRect.width,
                            tableHeight: tableRect.height,
                            rowY: rowRect.top - slideRect.top,
                            rowHeight: rowRect.height,
                            parentCellX: parentCellRect.left - slideRect.left,
                            parentCellY: parentCellRect.top - slideRect.top,
                            parentCellWidth: parentCellRect.width,
                            parentCellHeight: parentCellRect.height,
                            zIndex: parseInt(extractedStyles.zIndex) || 10,
                            styles: extractedStyles, // Use the complete computed styles
                            rect: {
                                x: relativeX,
                                y: relativeY,
                                width: relativeWidth,
                                height: relativeHeight
                            }
                        };
                        
                        overlayElements.push(overlayInfo);
                        
                        // Mark this wrapper as processed to avoid duplicate processing
                        processedElements.add(getElementId(wrapper));
                        processedElements.add(getElementId(parentCell));
                    }
                });
                
                console.log(`Found ${foundHighlights.size} unique table row highlights (expected 4)`);
                
                // Then process other overlay elements
                
                elementsToProcess.forEach(element => {
                    if (isOverlayElement(element)) {
                        const overlayInfo = extractOverlayInfo(element, slideElement);
                        if (overlayInfo) {
                            console.log(`Found overlay element: ${element.className}, type: ${overlayInfo.type}`);
                            overlayElements.push(overlayInfo);
                        }
                        processedElements.add(getElementId(element));
                        for (const child of element.children) processedElements.add(getElementId(child));
                    }
                });
                
                elementsToProcess.sort((a, b) => {
                    const rectA = a.getBoundingClientRect();
                    const rectB = b.getBoundingClientRect();
                    const topDiff = rectA.top - rectB.top;
                    return Math.abs(topDiff) < 5 ? rectA.left - rectB.left : topDiff;
                });
                
                elementsToProcess.forEach(element => {
                    if (processedElements.has(getElementId(element))) return;
                    const elementId = getElementId(element);
                    const tagName = element.tagName.toLowerCase();
                    
                    // Check for footer elements first
                    const footerInfo = extractFooterInfo(element, slideElement);
                    if (footerInfo) {
                        slide.elements.push(footerInfo);
                        processedElements.add(elementId);
                        // Mark all footer children as processed
                        Array.from(element.children).forEach(child => {
                            processedElements.add(getElementId(child));
                            processedTextElements.add(getElementId(child));
                        });
                        return;
                    }

                    const position = getAccuratePosition(element, slideElement);
                    const styles = extractComprehensiveStyles(element);
                    
                    const shapeInfo = extractShapeInfo(element, slideElement);

                    const shouldExtractInline = shouldExtractInlineGroup(element, processedTextElements);
                    const inlineGroup = shouldExtractInline ? getInlineGroup(element, slideElement, processedTextElements) : null;
                    
                    if (processedElements.has(elementId)) {
                        processedElements.add(elementId);
                        return;
                    }
                    processedElements.add(elementId);
                    
                    const elementData = {
                        type: tagName,
                        x: position.x,
                        y: position.y,
                        width: position.width,
                        height: position.height,
                        styles,
                        className: element.className || '',
                        id: element.id || '',
                        zIndex: parseInt(styles.zIndex) || 0,
                        inlineGroup: inlineGroup,
                        shapeInfo: shapeInfo
                    };
                    
                    if (inlineGroup) {
                        processedTextElements.add(elementId);
                        getAllDescendants(element).forEach(desc => {
                            processedTextElements.add(getElementId(desc));
                        });
                    }

                    if (shapeInfo) {
                        processedElements.add(elementId);
                        const shapeText = shapeInfo.text;
                        if (shapeText) {
                            const shapeRect = element.getBoundingClientRect();
                            Array.from(slideElement.querySelectorAll('*')).forEach(other => {
                                if (other === element) return;
                                const otherText = other.textContent ? other.textContent.trim() : '';
                                const otherRect = other.getBoundingClientRect();
                                
                                if (otherText === shapeText &&
                                    Math.abs(otherRect.left - shapeRect.left) < 100 &&
                                    Math.abs(otherRect.top - shapeRect.top) < 100) {
                                    processedTextElements.add(getElementId(other));
                                    processedElements.add(getElementId(other));
                                }
                            });
                        }
                        
                        getAllDescendants(element).forEach(desc => {
                            processedElements.add(getElementId(desc));
                            processedTextElements.add(getElementId(desc));
                        });
                    }

                    if (tagName === 'img') {
                        const imgDimensions = getAccurateImageDimensions(element, element.parentElement, slideElement);
                        elementData.mediaInfo = {
                            src: element.src || '',
                            alt: element.alt || '',
                            naturalWidth: element.naturalWidth || 0,
                            naturalHeight: element.naturalHeight || 0,
                            currentWidth: imgDimensions.width,
                            currentHeight: imgDimensions.height
                        };
                        processedTextElements.add(elementId);
                    }

                    if (tagName === 'canvas' && (element.classList.contains('chartjs-chart') || element.closest('.chart'))) {
                        const canvasRect = element.getBoundingClientRect();
                        const slideRect = slideElement.getBoundingClientRect();
                        let chartConfig = null;
                        try {
                            const chartData = element.getAttribute('data-chart');
                            if (chartData) {
                                chartConfig = JSON.parse(chartData);
                            }
                        } catch(e) {
                            console.log('Could not parse chart data for element', element.id);
                            chartConfig = element.getAttribute('data-chart');
                        }
                        
                        elementData.chartInfo = {
                            chartId: element.id || '',
                            chartClass: element.className || '',
                            chartData: chartConfig,
                            width: Math.round(canvasRect.width),
                            height: Math.round(canvasRect.height),
                            attributeWidth: element.width || element.getAttribute('width') || 0,
                            attributeHeight: element.height || element.getAttribute('height') || 0,
                            canvasRect: {
                                x: Math.round(canvasRect.left - slideRect.left),
                                y: Math.round(canvasRect.top - slideRect.top),
                                width: Math.round(canvasRect.width),
                                height: Math.round(canvasRect.height)
                            }
                        };
                        processedTextElements.add(elementId);
                    }

                    if (tagName === 'svg') {
                        elementData.svgInfo = {
                            svgContent: element.outerHTML,
                            viewBox: element.getAttribute('viewBox') || '',
                            width: element.getAttribute('width') || position.width,
                            height: element.getAttribute('height') || position.height
                        };
                        processedTextElements.add(elementId);
                    }
                    
                    if (!inlineGroup && !shapeInfo && ['div', 'span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(tagName)) {
                        if (shouldExtractTextContent(element, processedTextElements)) {
                            const text = getTextContent(element, processedElements);
                            if (text) {
                                const elementRect = element.getBoundingClientRect();
                                const isDuplicateOfShape = slide.elements.some(existingElement => {
                                    if (existingElement.shapeInfo && existingElement.shapeInfo.text === text) {
                                        const shapeX = existingElement.x;
                                        const shapeY = existingElement.y;
                                        const elementX = elementRect.left - slideRect.left;
                                        const elementY = elementRect.top - slideRect.top;
                                        return Math.abs(shapeX - elementX) < 100 && Math.abs(shapeY - elementY) < 100;
                                    }
                                    return false;
                                });
                                
                                if (!isDuplicateOfShape) {
                                    elementData.text = text;
                                    processedTextElements.add(elementId);
                                    getAllDescendants(element).forEach(desc => {
                                        processedTextElements.add(getElementId(desc));
                                    });
                                }
                            }
                        }
                    }

                    if (['ul', 'ol'].includes(tagName)) {
                        elementData.listInfo = getListInfo(element, slideElement);
                        processedTextElements.add(elementId);
                        getAllDescendants(element).forEach(desc => {
                            processedTextElements.add(getElementId(desc));
                        });
                    }
                    if (['table'].includes(tagName)) {
                        elementData.tableInfo = getTableInfo(element, slideElement);
                        processedTextElements.add(elementId);
                        getAllDescendants(element).forEach(desc => {
                            processedTextElements.add(getElementId(desc));
                        });
                    }

                    slide.elements.push(elementData);

                    const before = extractPseudo(element, slideElement, '::before');
                    if (before) slide.elements.push(before);
                    const after = extractPseudo(element, slideElement, '::after');
                    if (after) slide.elements.push(after);

                    if (['ul', 'ol', 'table'].includes(tagName)) {
                        getAllDescendants(element).forEach(desc => {
                            processedElements.add(getElementId(desc));
                        });
                    }
                });

                const missedElements = Array.from(slideElement.querySelectorAll('img, canvas, svg')).filter(element => 
                    !processedElements.has(getElementId(element))
                );
                
                missedElements.forEach(element => {
                    const elementId = getElementId(element);
                    const tagName = element.tagName.toLowerCase();
                    const position = getAccuratePosition(element, slideElement);
                    const styles = extractComprehensiveStyles(element);
                    
                    const elementData = {
                        type: tagName,
                        x: position.x,
                        y: position.y,
                        width: position.width,
                        height: position.height,
                        styles,
                        className: element.className || '',
                        id: element.id || '',
                        zIndex: parseInt(styles.zIndex) || 0
                    };

                    if (tagName === 'img') {
                        const imgDimensions = getAccurateImageDimensions(element, element.parentElement, slideContainer);
                        elementData.mediaInfo = {
                            src: element.src || '',
                            alt: element.alt || '',
                            naturalWidth: element.naturalWidth || 0,
                            naturalHeight: element.naturalHeight || 0,
                            currentWidth: imgDimensions.width,
                            currentHeight: imgDimensions.height
                        };
                    }

                    if (tagName === 'canvas' && (element.classList.contains('chartjs-chart') || element.closest('.chart'))) {
                        const canvasRect = element.getBoundingClientRect();
                        const slideRect = slideElement.getBoundingClientRect();
                        let chartConfig = null;
                        try {
                            const chartData = element.getAttribute('data-chart');
                            if (chartData) {
                                chartConfig = JSON.parse(chartData);
                            }
                        } catch(e) {
                            console.log('Could not parse chart data for element', element.id);
                            chartConfig = element.getAttribute('data-chart');
                        }
                        
                        elementData.chartInfo = {
                            chartId: element.id || '',
                            chartClass: element.className || '',
                            chartData: chartConfig,
                            width: Math.round(canvasRect.width),
                            height: Math.round(canvasRect.height),
                            attributeWidth: element.width || element.getAttribute('width') || 0,
                            attributeHeight: element.height || element.getAttribute('height') || 0,
                            canvasRect: {
                                x: Math.round(canvasRect.left - slideRect.left),
                                y: Math.round(canvasRect.top - slideRect.top),
                                width: Math.round(canvasRect.width),
                                height: Math.round(canvasRect.height)
                            }
                        };
                    }

                    if (tagName === 'svg') {
                        elementData.svgInfo = {
                            svgContent: element.outerHTML,
                            viewBox: element.getAttribute('viewBox') || '',
                            width: element.getAttribute('width') || position.width,
                            height: element.getAttribute('height') || position.height
                        };
                    }

                    slide.elements.push(elementData);
                    processedElements.add(elementId);
                });

                slide.elements.push(...overlayElements);
                slide.elements.sort((a, b) => a.zIndex - b.zIndex || a.y - b.y || a.x - b.x);
                slides.push(slide);
            }
            return slides;
        }, documentInfo);

        await fs.writeFile(outputPath, JSON.stringify(allSlidesData, null, 2), 'utf-8');
        console.log(`Successfully extracted ${allSlidesData.length} slides to ${outputPath}`);
    } catch (err) {
        console.error('Error processing slides:', err);
    } finally {
        await browser.close();
    }
}

const htmlFilePath = 'input.html';
const outputPath = 'slides_data.json';
extractSlideData(htmlFilePath, outputPath).catch(err => {
    console.error('Error:', err);
});