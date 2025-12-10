const { Cluster } = require('puppeteer-cluster');
const fs = require('fs');
const path = require('path');

const clusterHolder = { cluster: null };

async function initCluster() {
    clusterHolder.cluster = await Cluster.launch({
        concurrency: Cluster.CONCURRENCY_PAGE,
        maxConcurrency: 15,
        timeout: 30000,
        monitor: true,
        puppeteerOptions: {
            headless: 'new',
            args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-web-security'],
        },
    });
}

async function extractSlideData(htmlContent) {
    return await clusterHolder.cluster.execute(htmlContent, async ({ page, data }) => {
        try {
            await page.setViewport({ width: 1920, height: 1080 });
            await page.setContent(data, { waitUntil: 'load', timeout: 15000 });

            // Wait for all resources (images, fonts, etc.) to load
            await page.waitForFunction(() => {
                const images = Array.from(document.querySelectorAll('img'));
                const fonts = document.fonts ? document.fonts.ready : Promise.resolve();
                return Promise.all([fonts, images.every(img => img.complete)]);
            }, { timeout: 15000 }).catch(() => console.log('Some resources may not have loaded'));

            let documentInfo = await page.evaluate(() => {
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

                function safeFloat(value) {
                    try {
                        return parseFloat(value.replace(/[^0-9.-]/g, '')) || 0;
                    } catch {
                        return 0;
                    }
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
                        verticalAlign: styles.verticalAlign,
                        customProperties
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

                function extractDirectTextContent(element) {
                    return Array.from(element.childNodes)
                        .filter(node => node.nodeType === Node.TEXT_NODE)
                        .map(node => node.textContent.trim())
                        .join(' ')
                        .trim();
                }

                function getTableMetadata(tableElement) {
                    const rows = Array.from(tableElement.querySelectorAll('tr'));
                    let maxCols = 0;
                    
                    rows.forEach(row => {
                        const cells = Array.from(row.querySelectorAll('td, th'));
                        let colCount = 0;
                        cells.forEach(cell => colCount += cell.colSpan || 1);
                        maxCols = Math.max(maxCols, colCount);
                    });
                    
                    return {
                        rowCount: rows.length,
                        columnCount: maxCols,
                        rows: rows.map((row, rowIndex) => {
                            const cells = Array.from(row.querySelectorAll('td, th'));
                            return {
                                index: rowIndex,
                                cellCount: cells.length
                            };
                        })
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
                    
                    // Always include zIndex from pseudoStyles if present
                    let zIndex = parseInt(pseudoStyles.zIndex);
                    if (isNaN(zIndex)) {
                        // fallback to parent zIndex logic
                        const parentZ = parseInt(window.getComputedStyle(element).zIndex) || 0;
                        zIndex = pseudo === '::after' ? parentZ - 1 : parentZ - 2;
                    }

                    const elementData = {
                        type: 'pseudo',
                        pseudoType: pseudo,
                        x: position.x,
                        y: position.y,
                        width: position.width,
                        height: position.height,
                        styles,
                        zIndex: zIndex,
                        parentClassName: element.className || '',
                        parentTagName: element.tagName.toLowerCase(),
                        children: []
                    };
                    if (hasContent) {
                        elementData.text = content;
                    }
                    return elementData;
                }

                function shouldProcessElement(element) {
                    const styles = window.getComputedStyle(element);
                    if (styles.display === 'none' || styles.visibility === 'hidden' || styles.opacity === '0') return false;
                    return true;
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
                            if (match) rotation = parseFloat(match[1]);
                        }
                    }

                    const width = element.offsetWidth;
                    const height = element.offsetHeight;
                    const centerX = rect.left - slideRect.left + rect.width / 2;
                    const centerY = rect.top - slideRect.top + rect.height / 2;
                    const x = centerX - width / 2;
                    const y = centerY - height / 2;

                    return {
                        isShape: true,
                        type: 'shape',
                        rotation: rotation,
                        clipPath: styles.clipPath,
                        transform: styles.transform,
                        rect: {
                            x: Math.round(x),
                            y: Math.round(y),
                            width: Math.round(width),
                            height: Math.round(height)
                        }
                    };
                }

                function processElementHierarchy(element, slideContainer, depth = 0) {
                    if (!shouldProcessElement(element)) return null;

                    const tagName = element.tagName.toLowerCase();
                    const position = getAccuratePosition(element, slideContainer);
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
                        zIndex: parseInt(styles.zIndex) || 0,
                        children: []
                    };

                    // Check if this element is a shape
                    const shapeInfo = extractShapeInfo(element, slideContainer);
                    if (shapeInfo) {
                        elementData.shapeInfo = shapeInfo;
                    }

                    // Extract direct text content (not from children)
                    const directText = extractDirectTextContent(element);
                    if (directText) {
                        elementData.text = directText;
                    }

                    // Handle special element types
                    if (tagName === 'img') {
                        elementData.mediaInfo = {
                            src: element.src || '',
                            alt: element.alt || '',
                            naturalWidth: element.naturalWidth || 0,
                            naturalHeight: element.naturalHeight || 0,
                            currentWidth: position.width,
                            currentHeight: position.height
                        };
                    }

                    if (tagName === 'a') {
                        elementData.linkInfo = {
                            href: element.href || '',
                            target: element.target || '',
                            text: element.textContent.trim()
                        };
                    }

                    if (tagName === 'canvas') {
                        if (element.classList.contains('chartjs-chart') || element.closest('.chart')) {
                            let chartConfig = null;
                            try {
                                const chartData = element.getAttribute('data-chart');
                                if (chartData) {
                                    chartConfig = JSON.parse(chartData);
                                }
                            } catch(e) {
                                chartConfig = element.getAttribute('data-chart');
                            }
                            
                            elementData.chartInfo = {
                                chartId: element.id || '',
                                chartClass: element.className || '',
                                chartData: chartConfig,
                                width: Math.round(position.width),
                                height: Math.round(position.height),
                                attributeWidth: element.width || element.getAttribute('width') || 0,
                                attributeHeight: element.height || element.getAttribute('height') || 0
                            };
                        }
                    }

                    if (tagName === 'svg') {
                        elementData.svgInfo = {
                            svgContent: element.outerHTML,
                            viewBox: element.getAttribute('viewBox') || '',
                            width: element.getAttribute('width') || position.width,
                            height: element.getAttribute('height') || position.height
                        };
                    }

                    // Add table metadata
                    if (tagName === 'table') {
                        elementData.tableInfo = getTableMetadata(element);
                    }

                    // Add cell metadata for td/th elements
                    if (tagName === 'td' || tagName === 'th') {
                        const row = element.parentElement;
                        const table = row ? row.parentElement : null;
                        let rowIndex = -1;
                        
                        if (row && table) {
                            const allRows = Array.from(table.querySelectorAll('tr'));
                            rowIndex = allRows.indexOf(row);
                        }
                        
                        elementData.cellInfo = {
                            colSpan: element.colSpan || 1,
                            rowSpan: element.rowSpan || 1,
                            cellIndex: element.cellIndex >= 0 ? element.cellIndex : -1,
                            rowIndex: rowIndex
                        };
                    }

                    // Extract pseudo-elements
                    const before = extractPseudo(element, slideContainer, '::before');
                    if (before) {
                        elementData.pseudoBefore = before;
                    }
                    const after = extractPseudo(element, slideContainer, '::after');
                    if (after) {
                        elementData.pseudoAfter = after;
                    }

                    // Process children recursively
                    Array.from(element.children).forEach(child => {
                        const childData = processElementHierarchy(child, slideContainer, depth + 1);
                        if (childData) {
                            elementData.children.push(childData);
                        }
                    });

                    return elementData;
                }

                for (let slideIndex = 0; slideIndex < slideElements.length; slideIndex++) {
                    const slideElement = slideElements[slideIndex];
                    const slideRect = slideElement.getBoundingClientRect();
                    const slide = {
                        slideId: slideIndex + 1,
                        slideWidth: parseFloat(slideRect.width.toFixed(2)),
                        slideHeight: parseFloat(slideRect.height.toFixed(2)),
                        slidePosition: {
                            x: parseFloat(slideRect.left.toFixed(2)),
                            y: parseFloat(slideRect.top.toFixed(2))
                        },
                        slideStyles: extractComprehensiveStyles(slideElement),
                        children: []
                    };

                    Array.from(slideElement.children).forEach(child => {
                        const elementData = processElementHierarchy(child, slideElement, 0);
                        if (elementData) {
                            slide.children.push(elementData);
                        }
                    });

                    slides.push(slide);
                }
                return slides;
            }, documentInfo);
            
            return allSlidesData;
        } catch (err) {
            console.error('Error processing slides:', err);
        } finally {
            if (page && !page.isClosed()) await page.close();
        }
    });
}

async function htmlToImage(htmlFilePath) {
  return await clusterHolder.cluster.execute(htmlFilePath, async ({ page, data }) => {
    await page.setViewport({ width: 1920, height: 1080 });
    await page.goto(`file://${path.resolve(htmlFilePath)}`);
    await page.waitForSelector('#ready', { timeout: 15000 });
    return await page.screenshot({ fullPage: true });
  });
}

async function main() {
    const absInputPath = path.resolve('dolbix.html');
    const outputJsonPath = path.resolve('slides_data.json');

    const htmlContent = fs.readFileSync(absInputPath, 'utf-8');
    
    await initCluster();
    const slidesData = await extractSlideData(htmlContent);
    if (!slidesData) {
        console.error('No slide data extracted. Aborting write.');
        await clusterHolder.cluster.idle();
        await clusterHolder.cluster.close();
        process.exit(2);
    }
    fs.writeFileSync(outputJsonPath, JSON.stringify(slidesData, null, 2), 'utf-8');
    await clusterHolder.cluster.idle();
    await clusterHolder.cluster.close();
    console.log(`Extracted data written to ${outputJsonPath}`);
}

if (require.main === module) {
    main().catch(err => {
        console.error('Fatal error:', err);
        process.exit(1);
    });
}

module.exports = { initCluster, extractSlideData, htmlToImage };