const fs = require('fs');
const path = require('path');
function loadModule(name) {
    try {
        return require(name);
    } catch (e) {
        try {
            return require(path.join(__dirname, 'node_modules', name));
        } catch (e2) {
            return require(`C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\node_modules\\${name}`);
        }
    }
}
const pptxgen = loadModule('pptxgenjs');
const AdmZip = loadModule('adm-zip');

const BASE_DIR = __dirname;
const REF_DIR = fs.existsSync(path.join(BASE_DIR, '..', 'BombayDC_Decks_With_BG_Images'))
    ? path.join(BASE_DIR, '..', 'BombayDC_Decks_With_BG_Images')
    : (fs.existsSync("C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images")
        ? "C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images"
        : path.join(BASE_DIR, 'source_decks'));

const OUT_DIR_1 = path.join(BASE_DIR, 'SALES DECKS', 'new MD');
const OUT_DIR_2 = path.join(BASE_DIR, 'new MD');
const MEDIA_EXTRACT_DIR = path.join(BASE_DIR, 'extracted_all_media');

[OUT_DIR_1, OUT_DIR_2, MEDIA_EXTRACT_DIR].forEach(dir => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

const fileMap = {
    "01_BombayDC_BFSI_POV_Deck.pptx": "BombayDC_BFSI_POV_Deck (1)_BDC_Styled.pptx",
    "02_BFSI_Platform_Pitch_V4.pptx": "V4_BDC_Styled.pptx",
    "03_BFSI_CXO_Intro_V5.pptx": "V5_BDC_Styled.pptx",
    "04_Consumer_Platforms_Deck.pptx": "Consumer Platforms deck_BDC_Styled.pptx",
    "05_Corporate_Brand_Platforms_Deck.pptx": "BombayDC_Corporate_Brand_Platforms_Deck (1)_BDC_Styled.pptx",
    "06_Corporate_Digital_Platforms_Deck_v2.pptx": "Corporate_Digital_Platforms_Deck_v2_BDC_Styled.pptx",
    "07_Ecommerce_POV_Deck_Rebuilt.pptx": "BombayDC_Ecommerce_POV_Deck_Rebuilt (1)_BDC_Styled.pptx",
    "08_Ecommerce_V5.pptx": "ecomm-V5_BDC_Styled.pptx",
    "09_B2B_POV_Deck.pptx": "BombayDC_B2B_POV_Deck-2 (1)_BDC_Styled.pptx",
    "10_Enterprise_Platforms_POV.pptx": "BombayDC_Enterprise_Platforms_POV_BDC_Styled.pptx",
    "11_Intro_Deck.pptx": "Intro deck_BDC_Styled.pptx"
};

const bgMappingsPath = fs.existsSync(path.join(BASE_DIR, 'deck_bg_mappings.json'))
    ? path.join(BASE_DIR, 'deck_bg_mappings.json')
    : "C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\deck_bg_mappings.json";
const bgMappings = fs.existsSync(bgMappingsPath) ? JSON.parse(fs.readFileSync(bgMappingsPath, 'utf8')) : {};

const coverBgMapPath = fs.existsSync(path.join(BASE_DIR, 'cover_bg_map.json'))
    ? path.join(BASE_DIR, 'cover_bg_map.json')
    : "C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\cover_bg_map.json";
const coverBgMap = fs.existsSync(coverBgMapPath) ? JSON.parse(fs.readFileSync(coverBgMapPath, 'utf8')) : {};

// LOAD BDC DECK (COPY) EXACT CLOSING SLIDE
const refDeckPath = fs.existsSync(path.join(BASE_DIR, 'BDC Deck (Copy).pptx'))
    ? path.join(BASE_DIR, 'BDC Deck (Copy).pptx')
    : "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\BDC Deck (Copy).pptx";
const refDeckZip = fs.existsSync(refDeckPath) ? new AdmZip(refDeckPath) : null;
const refSlideEntries = refDeckZip ? refDeckZip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml')) : [];
if (refSlideEntries.length > 0) {
    refSlideEntries.sort((a, b) => parseInt(a.entryName.match(/slide(\d+)\.xml/)[1]) - parseInt(b.entryName.match(/slide(\d+)\.xml/)[1]));
}
const closingSlideEntry = refSlideEntries.length > 0 ? refSlideEntries[refSlideEntries.length - 1] : null;
const closingSlideXml = (refDeckZip && closingSlideEntry) ? refDeckZip.readAsText(closingSlideEntry) : "";

const closingSlideRelPath = closingSlideEntry ? `ppt/slides/_rels/${path.basename(closingSlideEntry.entryName)}.rels` : "";
const closingSlideRelEntry = refDeckZip ? refDeckZip.getEntries().find(e => e.entryName === closingSlideRelPath) : null;

const closingMediaDir = path.join(MEDIA_EXTRACT_DIR, "BDC_Deck_Copy_Closing");
if (!fs.existsSync(closingMediaDir)) fs.mkdirSync(closingMediaDir, { recursive: true });
if (refDeckZip) {
    refDeckZip.getEntries().filter(e => e.entryName.startsWith('ppt/media/')).forEach(m => {
        fs.writeFileSync(path.join(closingMediaDir, path.basename(m.entryName)), m.getData());
    });
}

const closingRelMap = {};
if (closingSlideRelEntry && refDeckZip) {
    const relXml = refDeckZip.readAsText(closingSlideRelEntry);
    const relMatches = [...relXml.matchAll(/<Relationship[\s\S]*?\/>/g)];
    relMatches.forEach(rel => {
        const idMatch = rel[0].match(/Id="([^"]+)"/);
        const targetMatch = rel[0].match(/Target="\.\.\/media\/([^"]+)"/);
        if (idMatch && targetMatch) {
            closingRelMap[idMatch[1]] = path.join(closingMediaDir, targetMatch[1]);
        }
    });
}

function decodeXmlEntities(str) {
    if (!str) return "";
    return str
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, "&")
        .replace(/&quot;/g, '"')
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .trim();
}

// AUTO-DETECT slide background colour from source PPTX XML
// Returns the hex colour string (e.g. "034E48") or null if not found
function detectSlideBg(slideXml) {
    const bgMatch = slideXml.match(/<p:bg>[\s\S]*?<\/p:bg>/);
    if (!bgMatch) return null;
    const solidClr = bgMatch[0].match(/a:srgbClr val="([^"]+)"/i);
    if (solidClr) return solidClr[1].toUpperCase();
    // Theme colour references
    const schemeClr = bgMatch[0].match(/a:schemeClr val="([^"]+)"/i);
    if (schemeClr) {
        const s = schemeClr[1].toLowerCase();
        if (s === 'dk1' || s === 'dk2') return '121212';
        if (s === 'lt1' || s === 'lt2') return 'ECE9E4';
    }
    return null;
}

// ==========================================
// 1. STANDARDIZED SPACING TOKENS (INCHES)
// ==========================================
const SPACING = {
    XS: 0.05,       // Tight cluster (e.g. number + label, pill text offset)
    SM: 0.08,       // Between category/header and immediate body/divider
    MD: 0.12,       // Between list items/bullets within a block
    LG: 0.20,       // Between sibling blocks / grid rows
    XL: 0.30,       // Between major vertical sections
    CONTENT_START_Y: 1.60, // Standard start Y for slide content
    CONTENT_MAX_BOTTOM: 5.08, // Maximum Y bottom boundary
    HEADER_DIVIDER_Y: 1.35, // Fixed header line divider
    CARD_PILL_H: 0.50, // Standard card pill height
};

// Returns true when hex colour luminance < 128 (perceptual threshold)
function isColorDark(hex) {
    if (!hex || hex.length < 6) return false;
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);
    return (r * 299 + g * 587 + b * 114) / 1000 < 128;
}

// Round any font size to the nearest EVEN whole number
function toEvenPt(n) {
    return Math.round(n / 2) * 2;
}

// Word-wrap and line measurement for Inter font
function measureParagraphHeight(text, widthInches, fontSizePt, isBullet, paraSpaceAfterPt = 2.0) {
    if (!text || text.trim().length === 0) return 0;
    const effectiveWidth = Math.max(0.4, isBullet ? widthInches - 0.16 : widthInches);
    // Character width estimate for Inter font: ~0.52 * (fontSize / 72)
    const avgCharWidth = (fontSizePt / 72) * 0.52;
    const charsPerLine = Math.max(1, Math.floor(effectiveWidth / avgCharWidth));
    
    // Compute wrapped line count with word boundary wrapping
    const words = text.trim().split(/\s+/);
    let lineCount = 1;
    let currentLineLen = 0;
    words.forEach(word => {
        if (currentLineLen + word.length + 1 > charsPerLine) {
            lineCount++;
            currentLineLen = word.length;
        } else {
            currentLineLen += word.length + 1;
        }
    });
    
    const lineHeight = (fontSizePt / 72) * 1.22;
    const spaceAfter = paraSpaceAfterPt / 72;
    return (lineCount * lineHeight) + spaceAfter;
}

function calculateTextShapeHeight(spXml, w, customFontSize = null) {
    const txBodyMatch = spXml.match(/<p:txBody>[\s\S]*?<\/p:txBody>/) || spXml.match(/<a:txBody>[\s\S]*?<\/a:txBody>/);
    if (!txBodyMatch) return 0.20;
    
    const pMatches = [...txBodyMatch[0].matchAll(/<a:p>[\s\S]*?<\/a:p>/g)];
    let totalH = 0.02;
    const pCount = pMatches.length;

    pMatches.forEach((p, pIdx) => {
        const pXml = p[0];
        const szMatch = pXml.match(/sz="(\d+)"/);
        let fontSize = customFontSize ? customFontSize : (szMatch ? parseFloat((parseInt(szMatch[1]) / 100).toFixed(1)) : 8.0);
        fontSize = toEvenPt(fontSize);
        
        if (fontSize >= 40) fontSize = 48.0;
        else if (fontSize >= 16) fontSize = 18.0;
        else fontSize = 8.0;
        
        const isBullet = pXml.includes('<a:buChar') || pXml.includes('<a:buAutoNum');
        const rawTxt = [...pXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join('');
        
        if (rawTxt.trim().length === 0) return;
        
        if (fontSize >= 40) {
            totalH += 0.52;
            return;
        }

        const spaceAfter = isBullet ? 2.0 : (pIdx === pCount - 1 ? 0 : 3.0);
        totalH += measureParagraphHeight(rawTxt, w, fontSize, isBullet, spaceAfter);
    });
    
    return Math.max(parseFloat(totalH.toFixed(2)), 0.20);
}

function getTopLevelShapes(xml) {
    const shapes = [];
    const spTreeMatch = xml.match(/<p:spTree>([\s\S]*?)<\/p:spTree>/);
    if (!spTreeMatch) return shapes;
    
    const content = spTreeMatch[1];
    let pos = 0;
    
    while (pos < content.length) {
        const nextSp = content.indexOf('<p:sp>', pos);
        const nextPic = content.indexOf('<p:pic>', pos);
        const nextCxn = content.indexOf('<p:cxnSp>', pos);
        
        const indices = [nextSp, nextPic, nextCxn].filter(i => i !== -1);
        if (indices.length === 0) break;
        
        const start = Math.min(...indices);
        let tag = '<p:sp>';
        let closeTag = '</p:sp>';
        if (start === nextPic) { tag = '<p:pic>'; closeTag = '</p:pic>'; }
        if (start === nextCxn) { tag = '<p:cxnSp>'; closeTag = '</p:cxnSp>'; }
        
        const end = content.indexOf(closeTag, start);
        if (end === -1) break;
        
        const shapeXml = content.substring(start, end + closeTag.length);
        shapes.push({ tag, xml: shapeXml });
        pos = end + closeTag.length;
    }
    return shapes;
}

async function processDeck(refFileName, outFileName) {
    console.log(`\n==================================================`);
    console.log(`Standardizing & QA Processing: ${refFileName} -> ${outFileName}`);
    console.log(`==================================================`);

    const refFilePath = path.join(REF_DIR, refFileName);
    if (!fs.existsSync(refFilePath)) {
        console.warn(`Source ref file not found at ${refFilePath}, skipping.`);
        return;
    }

    const zip = new AdmZip(refFilePath);
    const zipEntries = zip.getEntries();

    const deckMediaDir = path.join(MEDIA_EXTRACT_DIR, refFileName.replace('.pptx', ''));
    if (!fs.existsSync(deckMediaDir)) fs.mkdirSync(deckMediaDir, { recursive: true });

    const mediaEntries = zipEntries.filter(e => e.entryName.startsWith('ppt/media/'));
    mediaEntries.forEach(m => {
        fs.writeFileSync(path.join(deckMediaDir, path.basename(m.entryName)), m.getData());
    });

    const slideEntries = zipEntries.filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'));
    slideEntries.sort((a, b) => {
        const numA = parseInt(a.entryName.match(/slide(\d+)\.xml/)[1]);
        const numB = parseInt(b.entryName.match(/slide(\d+)\.xml/)[1]);
        return numA - numB;
    });

    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    pres.title = refFileName.replace('.pptx', '');

    const deckBgInfo = bgMappings[refFileName];

    for (let sIdx = 0; sIdx < slideEntries.length; sIdx++) {
        const sEntry = slideEntries[sIdx];
        const sNum = sIdx + 1;
        const isClosingSlide = (sIdx === slideEntries.length - 1);
        const darkSlidesList = (deckBgInfo && deckBgInfo.dark_slides) ? deckBgInfo.dark_slides : [];
        // Start with config-based flag; will be overridden by XML detection below
        let isDarkThemeSlide = sIdx === 0 || isClosingSlide || darkSlidesList.includes(sNum);
        
        const xml = isClosingSlide ? closingSlideXml : zip.readAsText(sEntry);

        // AUTO-DETECT background from source XML — overrides the hardcoded dark_slides list
        let detectedBgColor = null;
        if (!isClosingSlide && sIdx > 0) {
            const detected = detectSlideBg(xml);
            if (detected) {
                detectedBgColor = detected;
                isDarkThemeSlide = isColorDark(detected);
            }
        }

        let relMap = {};
        if (isClosingSlide) {
            relMap = closingRelMap;
        } else {
            const relPath = `ppt/slides/_rels/${path.basename(sEntry.entryName)}.rels`;
            const relEntry = zipEntries.find(e => e.entryName === relPath);
            if (relEntry) {
                const relXml = zip.readAsText(relEntry);
                const relMatches = [...relXml.matchAll(/<Relationship[\s\S]*?\/>/g)];
                relMatches.forEach(rel => {
                    const idMatch = rel[0].match(/Id="([^"]+)"/);
                    const targetMatch = rel[0].match(/Target="\.\.\/media\/([^"]+)"/);
                    if (idMatch && targetMatch) {
                        relMap[idMatch[1]] = path.join(deckMediaDir, targetMatch[1]);
                    }
                });
            }
        }

        const pptSlide = pres.addSlide();

        if (isClosingSlide) {
            pptSlide.background = { color: "034E48" };

            const hdrImg = path.join(MEDIA_EXTRACT_DIR, "BDC_Deck_Copy_Closing", "image-10-1.png");
            if (fs.existsSync(hdrImg)) {
                pptSlide.addImage({ path: hdrImg, x: 0.16, y: 0.0, w: 9.69, h: 0.57 });
            } else {
                pptSlide.addText("BOMBAYDC", { x: 0.16, y: 0.28, w: 2.0, h: 0.3, fontSize: 11.5, fontFace: "Inter Medium", color: "ECE9E4", align: "left" });
                pptSlide.addText("bombaydc.com", { x: 7.84, y: 0.28, w: 2.0, h: 0.3, fontSize: 11.5, fontFace: "Inter Medium", color: "ECE9E4", align: "right" });
            }

            pptSlide.addText("LET'S BUILD", { x: 1.99, y: 0.65, w: 7.85, h: 0.55, fontSize: 36.0, fontFace: "Inter Medium", color: "ECE9E4", valign: "top", margin: 0 });
            pptSlide.addText("WHAT'S NEXT.", { x: 1.99, y: 1.15, w: 7.85, h: 0.55, fontSize: 36.0, fontFace: "Inter Medium", color: "ECE9E4", valign: "top", margin: 0 });
            pptSlide.addText("Explore our work, sectors, and point of view at www.bombaydc.com", { x: 1.99, y: 1.70, w: 7.85, h: 0.35, fontSize: 9.0, fontFace: "Inter", color: "B4B4B4", valign: "top", margin: 0 });

            pptSlide.addShape(pres.shapes.LINE, { x: 1.99, y: 2.15, w: 7.85, h: 0, line: { color: "3E8D86", width: 0.5 } });
            pptSlide.addShape(pres.shapes.RECTANGLE, { x: 1.99, y: 2.50, w: 3.8, h: 1.65, fill: { color: "000000" }, line: { width: 0 } });

            const profileImg = path.join(MEDIA_EXTRACT_DIR, "BDC_Deck_Copy_Closing", "image-10-2.png");
            if (fs.existsSync(profileImg)) {
                pptSlide.addImage({ path: profileImg, x: 2.24, y: 2.75, w: 1.15, h: 1.15, rounding: true });
            }

            pptSlide.addText("Siddesh Pednekar", { x: 3.55, y: 2.75, w: 2.1, h: 0.3, fontSize: 11.0, fontFace: "Inter Medium", color: "FFFFFF", valign: "top", margin: 0 });
            pptSlide.addText("Partner & COO", { x: 3.55, y: 3.10, w: 2.1, h: 0.2, fontSize: 9.0, fontFace: "Inter", color: "B4B4B4", valign: "top", margin: 0 });
            pptSlide.addText("sid@bombaydc.com", { x: 3.55, y: 3.33, w: 2.1, h: 0.2, fontSize: 9.0, fontFace: "Inter", color: "B4B4B4", valign: "top", margin: 0 });
            pptSlide.addText("9819981354", { x: 3.55, y: 3.55, w: 2.1, h: 0.2, fontSize: 9.0, fontFace: "Inter", color: "B4B4B4", valign: "top", margin: 0 });

            pptSlide.addText("CONFIDENTIAL AND PROPRIETARY | © BombayDC. This material is intended solely for your internal use and any use of this material without specific permission of BombayDC is strictly prohibited. All rights reserved.", {
                x: 1.99, y: 4.85, w: 7.5, h: 0.45, fontSize: 9.0, fontFace: "Inter", color: "B4B4B4", lineSpacingMultiple: 1.2, margin: 0
            });

            console.log(`  Slide ${sNum}/${slideEntries.length} processed (Master Closing Slide).`);
            continue;
        } else if (sIdx === 0 && coverBgMap[refFileName] && fs.existsSync(coverBgMap[refFileName])) {
            pptSlide.background = { path: coverBgMap[refFileName] };
        } else {
            // Use detected source colour if available; otherwise fall back to normalised light/dark
            const bgColor = detectedBgColor || (isDarkThemeSlide ? "121212" : "ECE9E4");
            pptSlide.background = { color: bgColor };
        }

        const shapes = getTopLevelShapes(xml);

        // PRE-PASS 2: Unified Header Processing (Title + Subline in ONE Single Container)
        const FIXED_DIVIDER_Y = 1.35;
        let headerTitleText = "";
        let headerSublineText = "";
        let hasHeaderTitle = false;

        shapes.forEach(shapeObj => {
            const offMatch = shapeObj.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            if (!offMatch) return;
            const py = parseFloat((parseInt(offMatch[2]) / 914400).toFixed(3));
            const px = parseFloat((parseInt(offMatch[1]) / 914400).toFixed(3));
            const szMatch = shapeObj.xml.match(/sz="(\d+)"/);
            const fontSize = szMatch ? parseFloat((parseInt(szMatch[1]) / 100).toFixed(1)) : 0;

            if (py < 0.85 && px >= 1.8 && fontSize >= 16 && shapeObj.xml.includes('<a:t>')) {
                const rawTxt = [...shapeObj.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').replace(/\s+/g, ' ');
                headerTitleText = decodeXmlEntities(rawTxt).trim().toUpperCase();
                if (headerTitleText.length > 0) hasHeaderTitle = true;
            }

            if (py >= 0.75 && py <= 1.30 && px >= 1.8 && fontSize > 0 && fontSize < 14 && shapeObj.xml.includes('<a:t>')) {
                const rawTxt = [...shapeObj.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').replace(/\s+/g, ' ').trim();
                if (rawTxt.length > 3 && !rawTxt.startsWith('1') && !rawTxt.startsWith('2') && !rawTxt.startsWith('3') && !rawTxt.startsWith('4') && !rawTxt.startsWith('5')) {
                    headerSublineText = decodeXmlEntities(rawTxt);
                }
            }
        });

        const isTitle2Lines = headerTitleText.length > 38;

        if (hasHeaderTitle && !isClosingSlide) {
            const displaySublineAbove = !isTitle2Lines ? headerSublineText : "";

            const headerRuns = [
                {
                    text: headerTitleText,
                    options: {
                        color: isDarkThemeSlide ? "FFFFFF" : "1A1A1A",
                        fontSize: toEvenPt(18.8),
                        bold: false,
                        fontFace: "Inter Medium",
                        breakLine: displaySublineAbove ? true : false,
                        charSpacing: 0.3,
                        paraSpaceAfter: displaySublineAbove ? 4 : 0
                    }
                }
            ];

            if (displaySublineAbove) {
                headerRuns.push({
                    text: displaySublineAbove,
                    options: {
                        color: isDarkThemeSlide ? "B0B0B4" : "5A5A5E",
                        fontSize: 8.0,
                        bold: false,
                        fontFace: "Inter",
                        paraSpaceBefore: 2
                    }
                });
            }

            pptSlide.addText(headerRuns, {
                x: 1.99,
                y: 0.57,
                w: 7.85,
                h: 0.76,
                valign: "top",
                margin: [0, 0, 0, 0]
            });
        }

        let movedSublineShift = 0;
        if (isTitle2Lines && headerSublineText && !isClosingSlide) {
            movedSublineShift = 0.30;
            pptSlide.addText([{
                text: headerSublineText,
                options: {
                    color: isDarkThemeSlide ? "D0D0D4" : "4A4A4E",
                    fontSize: 8.0,
                    bold: false,
                    fontFace: "Inter",
                    paraSpaceAfter: 3
                }
            }], {
                x: 1.99,
                y: 1.60,
                w: 7.85,
                h: 0.26,
                valign: "top",
                margin: [0, 0, 0, 0]
            });
        }

        let coverTitleBottomY = 0;

        // CUSTOM LAYOUT: Consumer Lifecycle (Two-Section Split: Acquisition vs. Retention)
        const isConsumerLifecycleSlide = shapes.some(s => s.xml.includes('DISCOVER')) && shapes.some(s => s.xml.includes('FIRST VALUE')) && shapes.some(s => s.xml.includes('REPEAT'));

        if (!isClosingSlide && sIdx > 0 && isConsumerLifecycleSlide) {
            // Draw slide index number
            pptSlide.addText([{
                text: `${sNum}`,
                options: {
                    color: isDarkThemeSlide ? "6A6A6B" : "9A9A9E",
                    fontSize: 18.0,
                    bold: false,
                    fontFace: "Inter Medium"
                }
            }], {
                x: 0.16,
                y: 0.57,
                w: 1.68,
                h: 0.33,
                valign: "top",
                margin: [0, 0, 0, 0]
            });

            // Draw header divider line
            pptSlide.addShape(pres.shapes.LINE, {
                x: 0.16, y: SPACING.HEADER_DIVIDER_Y, w: 9.68, h: 0,
                line: { color: isDarkThemeSlide ? "555558" : "B0B0B4", width: 0.50 }
            });

            // Two-Section Split: Acquisition Moments vs. Retention & Growth
            const col1X = 1.99;
            const col2X = 5.95;
            const colW = 3.80;
            const startContentY = SPACING.CONTENT_START_Y + movedSublineShift;

            // Column 1 Header: ACQUISITION MOMENTS
            pptSlide.addText([{
                text: "ACQUISITION MOMENTS",
                options: {
                    color: "4DB89A",
                    fontSize: 8.0,
                    bold: true,
                    fontFace: "Inter Bold",
                    charSpacing: 0.5
                }
            }], {
                x: col1X,
                y: startContentY,
                w: colW,
                h: 0.22,
                valign: "top",
                margin: [0, 0, 0, 0]
            });

            // Column 2 Header: RETENTION & GROWTH
            pptSlide.addText([{
                text: "RETENTION & GROWTH",
                options: {
                    color: "4DB89A",
                    fontSize: 8.0,
                    bold: true,
                    fontFace: "Inter Bold",
                    charSpacing: 0.5
                }
            }], {
                x: col2X,
                y: startContentY,
                w: colW,
                h: 0.22,
                valign: "top",
                margin: [0, 0, 0, 0]
            });

            // Divider lines under column headers
            const headerLineY = startContentY + 0.26;
            pptSlide.addShape(pres.shapes.LINE, {
                x: col1X, y: headerLineY, w: colW, h: 0,
                line: { color: isDarkThemeSlide ? "38383C" : "D4D4D8", width: 0.35 }
            });
            pptSlide.addShape(pres.shapes.LINE, {
                x: col2X, y: headerLineY, w: colW, h: 0,
                line: { color: isDarkThemeSlide ? "38383C" : "D4D4D8", width: 0.35 }
            });

            const acqItems = [
                { num: "01", title: "DISCOVER", desc: "The proposition is not clear or relevant enough." },
                { num: "02", title: "JOIN", desc: "Onboarding asks for effort before proving value." },
                { num: "03", title: "FIRST VALUE", desc: "The meaningful outcome arrives too late." }
            ];

            const retItems = [
                { num: "04", title: "REPEAT", desc: "The product does not create a useful habit." },
                { num: "05", title: "STAY", desc: "Relevance falls as needs and context change." },
                { num: "06", title: "GROW", desc: "The next paid or deeper action feels generic." }
            ];

            const itemY = headerLineY + SPACING.SM;
            const itemGap = 0.68;

            for (let i = 0; i < 3; i++) {
                const acq = acqItems[i];
                const ret = retItems[i];
                const curRowY = parseFloat((itemY + (i * itemGap)).toFixed(2));

                // Left Item
                pptSlide.addText([
                    { text: `${acq.num}.  ${acq.title}`, options: { color: isDarkThemeSlide ? "FFFFFF" : "034E48", fontSize: 8.0, bold: true, fontFace: "Inter Bold", breakLine: true, paraSpaceAfter: 2 } },
                    { text: acq.desc, options: { color: isDarkThemeSlide ? "B0B0B4" : "5A5A5E", fontSize: 8.0, bold: false, fontFace: "Inter" } }
                ], {
                    x: col1X,
                    y: curRowY,
                    w: colW,
                    h: 0.52,
                    valign: "top",
                    margin: [0, 0, 0, 0]
                });

                // Right Item
                pptSlide.addText([
                    { text: `${ret.num}.  ${ret.title}`, options: { color: isDarkThemeSlide ? "FFFFFF" : "034E48", fontSize: 8.0, bold: true, fontFace: "Inter Bold", breakLine: true, paraSpaceAfter: 2 } },
                    { text: ret.desc, options: { color: isDarkThemeSlide ? "B0B0B4" : "5A5A5E", fontSize: 8.0, bold: false, fontFace: "Inter" } }
                ], {
                    x: col2X,
                    y: curRowY,
                    w: colW,
                    h: 0.52,
                    valign: "top",
                    margin: [0, 0, 0, 0]
                });

                if (i < 2) {
                    const rowSepY = curRowY + 0.56;
                    pptSlide.addShape(pres.shapes.LINE, {
                        x: col1X, y: rowSepY, w: colW, h: 0,
                        line: { color: isDarkThemeSlide ? "28282C" : "E4E4E8", width: 0.25 }
                    });
                    pptSlide.addShape(pres.shapes.LINE, {
                        x: col2X, y: rowSepY, w: colW, h: 0,
                        line: { color: isDarkThemeSlide ? "28282C" : "E4E4E8", width: 0.25 }
                    });
                }
            }

            // Bottom Callout Box
            const calloutY = 4.40;
            pptSlide.addShape(pres.shapes.RECTANGLE, {
                x: 1.99,
                y: calloutY,
                w: 7.85,
                h: 0.46,
                fill: { color: isDarkThemeSlide ? "142321" : "E8F3F1" },
                line: { color: "034E48", width: 0.5 }
            });

            pptSlide.addText([{
                text: "The opportunity is to find the highest-value leak, redesign the behaviour around it, and measure whether the business outcome moves.",
                options: {
                    color: isDarkThemeSlide ? "4DB89A" : "034E48",
                    fontSize: 8.0,
                    bold: false,
                    fontFace: "Inter Medium"
                }
            }], {
                x: 2.15,
                y: calloutY + 0.08,
                w: 7.50,
                h: 0.30,
                valign: "middle",
                margin: [0, 0, 0, 0]
            });

            continue;
        }

        // PRE-PASS 1: Group and align Card Grids into uniform rows
        const cardPills = [];
        const allCardPills = [];
        const cardRowMap = {};
        const renderedCardBoxes = new Set();
        const rowLayoutMap = {};
        const rowLineYMap = {};

        shapes.forEach(shapeObj => {
            const offMatch = shapeObj.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const extMatch = shapeObj.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            if (!offMatch || !extMatch) return;
            const px = parseFloat((parseInt(offMatch[1]) / 914400).toFixed(3));
            const py = parseFloat((parseInt(offMatch[2]) / 914400).toFixed(3));
            const pw = parseFloat((parseInt(extMatch[1]) / 914400).toFixed(3));
            const ph = parseFloat((parseInt(extMatch[2]) / 914400).toFixed(3));
            
            const spPrMatch = shapeObj.xml.match(/<p:spPr>[\s\S]*?<\/p:spPr>/);
            const fillMatch = spPrMatch ? spPrMatch[0].match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/) : null;
            const fill = fillMatch ? fillMatch[1].toUpperCase() : null;
            const isCardPill = fill && pw < 4.0 && ["034E48", "4DB89A", "1C1C1E", "224B12", "1A3632", "0A3B36", "08322D", "0D524A", "004B44", "024E48", "034D47", "333333", "222222", "1E1E1E", "2A2A2C", "3A3A3C"].includes(fill);
            
            if (isCardPill) {
                cardPills.push({ x: px, y: py, w: pw, h: ph });
                allCardPills.push({ x: px, y: py });
            }
        });

        if (cardPills.length >= 2) {
            const pillRows = [];
            cardPills.sort((a, b) => a.y - b.y);
            cardPills.forEach(p => {
                let row = pillRows.find(r => Math.abs(r.y - p.y) < 0.45);
                if (!row) {
                    row = { y: p.y, pills: [] };
                    pillRows.push(row);
                }
                row.pills.push(p);
            });

            const hasTopIntro = shapes.some(sh => {
                const offM = sh.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                const extM = sh.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
                if (!offM || !extM || !sh.xml.includes('<a:t>')) return false;
                const py = parseFloat((parseInt(offM[2]) / 914400).toFixed(3));
                const pw = parseFloat((parseInt(extM[1]) / 914400).toFixed(3));
                return (py >= 1.50 && py <= 1.95 && pw >= 5.0);
            });

            const baseCardStartY = hasTopIntro ? (SPACING.CONTENT_START_Y + 0.45 + movedSublineShift) : (SPACING.CONTENT_START_Y + SPACING.XS + movedSublineShift);
            const CARD_TEXT_OFFSET = SPACING.CARD_PILL_H + SPACING.SM;

            if (pillRows.length === 1) {
                const r1Y = baseCardStartY;
                pillRows[0].pills.forEach(p => {
                    cardRowMap[`${p.x.toFixed(2)}_${p.y.toFixed(2)}`] = { cardX: p.x, cardY: r1Y, textY: r1Y + CARD_TEXT_OFFSET };
                });
            } else if (pillRows.length >= 2) {
                const r1Y = baseCardStartY;
                let maxR1TextH = 0.40;
                shapes.forEach(sh => {
                    const offM = sh.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                    const extM = sh.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
                    if (!offM || !extM || !sh.xml.includes('<a:t>')) return;
                    const py = parseFloat((parseInt(offM[2]) / 914400).toFixed(3));
                    const pw = parseFloat((parseInt(extM[1]) / 914400).toFixed(3));
                    if (py > pillRows[0].y + 0.10 && py <= (pillRows[1] ? pillRows[1].y + 0.10 : 3.20) && pw < 4.0) {
                        const estH = calculateTextShapeHeight(sh.xml, pw);
                        if (estH > maxR1TextH) maxR1TextH = estH;
                    }
                });

                // Row 2 placed additively based on measured Row 1 text height + standard row gap
                const r2Y = parseFloat(Math.min(3.45, Math.max(3.20, r1Y + SPACING.CARD_PILL_H + maxR1TextH + SPACING.LG)).toFixed(2));
                pillRows[0].pills.forEach(p => {
                    cardRowMap[`${p.x.toFixed(2)}_${p.y.toFixed(2)}`] = { cardX: p.x, cardY: r1Y, textY: r1Y + CARD_TEXT_OFFSET };
                });
                pillRows[1].pills.forEach(p => {
                    cardRowMap[`${p.x.toFixed(2)}_${p.y.toFixed(2)}`] = { cardX: p.x, cardY: r2Y, textY: r2Y + CARD_TEXT_OFFSET };
                });

                let maxR2TextH = 0.40;
                shapes.forEach(sh => {
                    const offM = sh.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                    const extM = sh.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
                    if (!offM || !extM || !sh.xml.includes('<a:t>')) return;
                    const py = parseFloat((parseInt(offM[2]) / 914400).toFixed(3));
                    const pw = parseFloat((parseInt(extM[1]) / 914400).toFixed(3));
                    if (py > pillRows[1].y + 0.10 && pw < 4.0) {
                        const estH = calculateTextShapeHeight(sh.xml, pw);
                        if (estH > maxR2TextH) maxR2TextH = estH;
                    }
                });

                // Footnote placed additively with SPACING.MD gap below tallest Row 2 content
                const bottomNoteY = Math.min(SPACING.CONTENT_MAX_BOTTOM - 0.25, Math.max(4.00, r2Y + CARD_TEXT_OFFSET + maxR2TextH + SPACING.MD));
                shapes.forEach(sh => {
                    const offM = sh.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                    const extM = sh.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
                    if (!offM || !extM || !sh.xml.includes('<a:t>')) return;
                    const py = parseFloat((parseInt(offM[2]) / 914400).toFixed(3));
                    const pw = parseFloat((parseInt(extM[1]) / 914400).toFixed(3));
                    if (py >= 3.20 && pw >= 5.0) {
                        rowLayoutMap[py.toFixed(2)] = { finalY: parseFloat(bottomNoteY.toFixed(2)) };
                    }
                });
            }
        }

        // PRE-PASS 3: Timeline & Row-List Layout Engine (Strictly for Timeline / Row-List layouts)

        const hasCategoryHeader = shapes.some(s => {
            if (!s.xml.includes('<a:t>')) return false;
            const rawTxt = [...s.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').trim();
            return (rawTxt === "CATEGORY" || rawTxt === "DESCRIPTION");
        });

        const lineDividers = shapes.filter(s => (s.tag === '<p:cxnSp>' || s.xml.includes('prst="line"')) && !s.xml.includes('<a:t>'));
        const distinctRowLines = new Set();
        lineDividers.forEach(s => {
            const extM = s.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            const offM = s.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            if (extM && offM) {
                const w = parseFloat((parseInt(extM[1]) / 914400).toFixed(3));
                const y = parseFloat((parseInt(offM[2]) / 914400).toFixed(3));
                if (w >= 4.0 && y > 1.40) {
                    distinctRowLines.add(y.toFixed(1));
                }
            }
        });
        const hasStackedRowDividers = distinctRowLines.size >= 2;

        const isTimelineOrRowListSlide = shapes.some(s => {
            if (!s.xml.includes('<a:t>')) return false;
            const rawTxt = [...s.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').trim();
            return (rawTxt.endsWith("Days") || rawTxt.endsWith("Weeks") || rawTxt.endsWith("Months") || rawTxt.startsWith("Phase") || rawTxt === "CATEGORY" || rawTxt === "DESCRIPTION");
        }) || hasStackedRowDividers;

        if (!isClosingSlide && sIdx > 0 && isTimelineOrRowListSlide && cardPills.length < 2) {
            const listShapes = shapes.filter(s => {
                const offMatch = s.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                const extMatch = s.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
                if (!offMatch || !s.xml.includes('<a:t>')) return false;
                const py = parseFloat((parseInt(offMatch[2]) / 914400).toFixed(3));
                const px = parseFloat((parseInt(offMatch[1]) / 914400).toFixed(3));
                const pw = extMatch ? parseFloat((parseInt(extMatch[1]) / 914400).toFixed(3)) : 3.0;
                if (py < 1.55 || px < 1.0 || pw >= 5.0) return false;
                const rawTxt = [...s.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ');
                const clean = decodeXmlEntities(rawTxt).trim();
                if (clean === "CATEGORY" || clean === "DESCRIPTION") return false;
                const isExplicitBottom = py >= 3.8 && (clean.startsWith("New team structure:") || clean.startsWith("Illustrative use cases") || clean.startsWith("The result:") || clean.startsWith("What enterprises now expect:") || clean.startsWith("85% client repeat") || clean.startsWith("What the platform"));
                if (isExplicitBottom) return false;
                return clean.length > 0 && !clean.includes('CONFIDENTIAL') && !clean.includes('bombaydc.com');
            });

            const clusters = [];
            listShapes.forEach(s => {
                const offMatch = s.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                const py = parseFloat((parseInt(offMatch[2]) / 914400).toFixed(3));
                const px = parseFloat((parseInt(offMatch[1]) / 914400).toFixed(3));
                const extMatch = s.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
                const pw = extMatch ? parseFloat((parseInt(extMatch[1]) / 914400).toFixed(3)) : 3.0;

                let c = clusters.find(cl => Math.abs(cl.origY - py) < 0.28);
                if (!c) {
                    c = { origY: py, shapes: [] };
                    clusters.push(c);
                }
                c.shapes.push({ s, py, px, pw, xml: s.xml });
            });

            if (clusters.length >= 2) {
                clusters.sort((a, b) => a.origY - b.origY);

                const hasTopIntroText = shapes.some(sh => {
                    const offM = sh.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                    const extM = sh.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
                    if (!offM || !extM || !sh.xml.includes('<a:t>')) return false;
                    const py = parseFloat((parseInt(offM[2]) / 914400).toFixed(3));
                    const pw = parseFloat((parseInt(extM[1]) / 914400).toFixed(3));
                    return (py >= 1.50 && py <= 1.90 && pw >= 5.0);
                });

                const bottomShapes = shapes.filter(sh => {
                    const offM = sh.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                    if (!offM || !sh.xml.includes('<a:t>')) return false;
                    const py = parseFloat((parseInt(offM[2]) / 914400).toFixed(3));
                    const rawTxt = [...sh.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ');
                    const clean = decodeXmlEntities(rawTxt).trim();
                    return py >= 3.8 && (clean.startsWith("New team structure:") || clean.startsWith("Illustrative use cases") || clean.startsWith("The result:") || clean.startsWith("What enterprises now expect:") || clean.startsWith("85% client repeat") || clean.startsWith("What the platform"));
                });

                const startY = (hasTopIntroText ? (SPACING.CONTENT_START_Y + 0.48) : (SPACING.CONTENT_START_Y + SPACING.SM)) + movedSublineShift;
                const maxClusterEndY = (bottomShapes.length > 0) ? (3.75 + movedSublineShift) : (SPACING.CONTENT_MAX_BOTTOM - 0.30 + movedSublineShift);
                const totalAvailH = maxClusterEndY - startY;
                const gapBetweenRows = SPACING.XS;
                const uniformRowH = parseFloat(((totalAvailH / clusters.length) - gapBetweenRows).toFixed(2));

                let curY = startY;
                clusters.forEach((c, idx) => {
                    const rowContentY = curY;
                    const rowDividerY = parseFloat((curY + uniformRowH).toFixed(2));

                    c.shapes.forEach(sh => {
                        rowLayoutMap[sh.py.toFixed(2)] = { finalY: rowContentY, lineY: rowDividerY, rowH: uniformRowH };
                    });

                    shapes.forEach(lineSh => {
                        if (lineSh.tag === '<p:cxnSp>' || lineSh.xml.includes('prst="line"')) {
                            const lOff = lineSh.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                            if (lOff) {
                                const lY = parseFloat((parseInt(lOff[2]) / 914400).toFixed(3));
                                if (lY > c.origY - 0.05 && lY <= c.origY + 0.90) {
                                    rowLineYMap[lY.toFixed(2)] = rowDividerY;
                                }
                            }
                        }
                    });

                    curY = parseFloat((rowDividerY + gapBetweenRows).toFixed(2));
                });

                bottomShapes.sort((a, b) => {
                    const offA = a.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                    const offB = b.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                    return parseInt(offA[2]) - parseInt(offB[2]);
                });

                let bY = Math.min(SPACING.CONTENT_MAX_BOTTOM - 0.70, Math.max(3.90, curY + SPACING.SM));
                bottomShapes.forEach(sh => {
                    const offM = sh.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                    const py = parseFloat((parseInt(offM[2]) / 914400).toFixed(3));
                    rowLayoutMap[py.toFixed(2)] = { finalY: bY };
                    bY = parseFloat((bY + 0.35).toFixed(2));
                });
            }
        }

        // Sort shapes vertically by top Y position so text rendering flows downwards
        shapes.sort((a, b) => {
            const offA = a.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const offB = b.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const yA = offA ? parseInt(offA[2]) : 0;
            const yB = offB ? parseInt(offB[2]) : 0;
            return yA - yB;
        });

        // ZERO-OVERLAP BOUNDING BOX TRACKING
        const occupiedBoxes = [];

        function recordOccupiedBox(box) {
            occupiedBoxes.push({
                x: box.x,
                y: box.y,
                w: box.w,
                h: box.h,
                bottom: parseFloat((box.y + box.h).toFixed(3)),
                txt: box.txt || ""
            });
        }

        function resolveVerticalCollision(proposedX, proposedY, proposedW, proposedH, minGap = SPACING.SM) {
            let adjustedY = proposedY;
            for (const prev of occupiedBoxes) {
                const hOverlap = (proposedX < prev.x + prev.w - 0.06) && (proposedX + proposedW > prev.x + 0.06);
                if (hOverlap) {
                    if (adjustedY < prev.bottom + minGap) {
                        adjustedY = parseFloat((prev.bottom + minGap).toFixed(3));
                    }
                }
            }
            return adjustedY;
        }

        shapes.forEach(shapeObj => {
            const spXml = shapeObj.xml;
            const shapeTag = shapeObj.tag;

            const offMatch = spXml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const extMatch = spXml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            
            if (!offMatch || !extMatch) return;

            const origX = parseFloat((parseInt(offMatch[1]) / 914400).toFixed(3));
            const origY = parseFloat((parseInt(offMatch[2]) / 914400).toFixed(3));
            let x = origX;
            let y = origY;
            let w = parseFloat((parseInt(extMatch[1]) / 914400).toFixed(3));
            let h = parseFloat((parseInt(extMatch[2]) / 914400).toFixed(3));

            if (w >= 9.9 && h >= 5.5 && shapeTag !== '<p:pic>') return;

            if (shapeTag !== '<p:pic>' && !spXml.includes('<a:t>')) {
                const _spPrDecor = spXml.match(/<p:spPr>[\s\S]*?<\/p:spPr>/);
                const _fillDecor = _spPrDecor ? _spPrDecor[0].match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/) : null;
                const _fColorDecor = _fillDecor ? _fillDecor[1].toUpperCase() : null;
                if (_fColorDecor && ["034E48","1D1D1F","121212","000000","0F2C28","1C1C1E","0A3B36"].includes(_fColorDecor)
                    && w > 7.0 && h > 0 && h < 0.65) return;
            }

            const rawTxtForCheck = [...spXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').replace(/\s+/g, ' ').trim();
            const cleanTxtForCheck = decodeXmlEntities(rawTxtForCheck);

            // FIX: Cover slide — skip header-bar text shapes AND floating bombaydc.com URL anywhere on cover
            if (sIdx === 0 && spXml.includes('<a:t>') && origY < 0.65) return;
            if (sIdx === 0 && (cleanTxtForCheck === 'bombaydc.com' || cleanTxtForCheck.toLowerCase() === 'bombaydc.com' || cleanTxtForCheck === 'BOMBAYDC' || cleanTxtForCheck.toLowerCase().includes('www.bombaydc'))) return;

            // FIX: Skip CATEGORY / DESCRIPTION table-header labels everywhere (layout engine handles spacing; we don't render them)
            if (!isClosingSlide && (cleanTxtForCheck === "CATEGORY" || cleanTxtForCheck === "DESCRIPTION")) return;

            if (!isClosingSlide && sIdx > 0 && spXml.includes('<a:t>')) {
                const szMatch = spXml.match(/sz="(\d+)"/);
                const fontSize = szMatch ? parseFloat((parseInt(szMatch[1]) / 100).toFixed(1)) : 0;
                
                if (origY < 0.60 && (cleanTxtForCheck === "BOMBAYDC" || cleanTxtForCheck.includes("bombaydc.com") || (origX < 0.5 && fontSize < 12) || (origX > 7.0 && fontSize < 12))) return;

                if (origY < 0.85 && fontSize >= 16 && origX >= 1.8) return;
                if (headerTitleText && cleanTxtForCheck.toUpperCase() === headerTitleText) return;
                if (headerTitleText && cleanTxtForCheck.length > 10 && headerTitleText.includes(cleanTxtForCheck.toUpperCase())) return;

                if (origY >= 0.75 && origY <= 1.30 && w >= 5.0 && fontSize < 14 && fontSize > 0) return;
                if (headerSublineText && cleanTxtForCheck === headerSublineText) return;
                if (headerSublineText && cleanTxtForCheck.length > 8 && (headerSublineText.includes(cleanTxtForCheck) || cleanTxtForCheck.includes(headerSublineText.substring(0, 15)))) return;
            }

            // Match card pill or text shape via cardRowMap (ONLY FOR SHAPES w < 4.0!)
            let isRowMapElement = false;
            let matchedCardInfo = null;

            if (w < 4.0 && !isClosingSlide && cardPills.length >= 2) {
                let matchingPill = null;
                let bestPillY = -1;
                cardPills.forEach(p => {
                    if (Math.abs(p.x - origX) < 0.40 && (origY >= p.y - 0.08)) {
                        if (p.y > bestPillY) {
                            bestPillY = p.y;
                            matchingPill = p;
                        }
                    }
                });

                if (matchingPill) {
                    const pillKey = `${matchingPill.x.toFixed(2)}_${matchingPill.y.toFixed(2)}`;
                    if (cardRowMap[pillKey]) {
                        const rowInfo = cardRowMap[pillKey];
                        x = rowInfo.cardX;
                        if (origY <= matchingPill.y + 0.30) {
                            y = rowInfo.cardY;
                            matchedCardInfo = { cardX: rowInfo.cardX, cardY: rowInfo.cardY, textY: rowInfo.textY, isHeader: true };
                        } else {
                            y = rowInfo.textY;
                            matchedCardInfo = { cardX: rowInfo.cardX, cardY: rowInfo.cardY, textY: rowInfo.textY, isHeader: false };
                        }
                        isRowMapElement = true;
                    }
                }
            }

            if (!isClosingSlide && origY >= 1.35 && !isRowMapElement) {
                const origYKey = origY.toFixed(2);
                if (rowLayoutMap[origYKey] !== undefined) {
                    y = rowLayoutMap[origYKey].finalY;
                } else {
                    y = origY + movedSublineShift;
                }
            }

            const hasTxBody = spXml.includes('<a:t>');
            const spPrMatch = spXml.match(/<p:spPr>[\s\S]*?<\/p:spPr>/);
            const spPrXml = spPrMatch ? spPrMatch[0] : "";
            const fillMatch = spPrXml.match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/);
            let shapeBgFill = fillMatch ? fillMatch[1] : null;

            if (!isClosingSlide && shapeBgFill && ["1A3632", "224B12", "4DB89A", "0A3B36", "08322D", "0D524A", "004B44", "024E48", "034D47"].includes(shapeBgFill.toUpperCase())) {
                shapeBgFill = "034E48";
            }

            const isLine = (shapeTag === '<p:cxnSp>' || spPrXml.includes('prst="line"') || h === 0 || w === 0) && !hasTxBody;

            // Constrain 2-column side-by-side widths when there is a side-by-side column OR image at the same vertical position (TEXT SHAPES ONLY)
            if (!isClosingSlide && !isLine && hasTxBody && x >= 1.8 && x < 4.0 && w > 4.5 && origY > 1.4) {
                const hasRightColOrPic = shapes.some(otherSh => {
                    const oOff = otherSh.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                    const oExt = otherSh.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
                    if (!oOff || !oExt) return false;
                    const ox = parseFloat((parseInt(oOff[1]) / 914400).toFixed(3));
                    const oy = parseFloat((parseInt(oOff[2]) / 914400).toFixed(3));
                    const ow = parseFloat((parseInt(oExt[1]) / 914400).toFixed(3));
                    const isRightCol = (ox >= 5.0 && ow < 4.5 && Math.abs(oy - origY) < 0.20 && otherSh.xml.includes('<a:t>'));
                    const isRightPic = (otherSh.tag === '<p:pic>' && ox >= 5.5);
                    return isRightCol || isRightPic;
                });
                if (hasRightColOrPic) {
                    w = 3.90;
                }
            }

            if (shapeTag === '<p:pic>') {
                const rEmbedMatch = spXml.match(/r:embed="([^"]+)"/);
                if (rEmbedMatch && relMap[rEmbedMatch[1]]) {
                    const imgFile = relMap[rEmbedMatch[1]];
                    if (fs.existsSync(imgFile)) {
                        if (w >= 9.9 && h >= 5.5) return;
                        const isProfilePic = isClosingSlide && (w < 2.0 && y > 2.0);
                        pptSlide.addImage({ path: imgFile, x, y, w, h, rounding: isProfilePic });
                        recordOccupiedBox({ x, y, w, h, txt: "[IMAGE]" });
                    }
                }
                return;
            }

            if (isLine && (w > 0 || h > 0)) {
                const isVertical = (w === 0 || (h > 0 && w < 0.1));
                const lnColorMatch = spPrXml.match(/<a:ln[\s\S]*?<a:srgbClr val="([^"]+)"/);
                let lnColor = lnColorMatch ? lnColorMatch[1] : "B4B4B4";
                
                const wMatch = spPrXml.match(/<a:ln w="(\d+)"/);
                const nativeW = wMatch ? parseFloat((parseInt(wMatch[1]) / 12700).toFixed(2)) : 0.5;
                const finalWidth = isClosingSlide ? nativeW : (isVertical ? 0.30 : 0.50);

                let lineY = y;
                if (!isClosingSlide && !isVertical && origY >= 0.90 && origY <= 1.50) {
                    lineY = SPACING.HEADER_DIVIDER_Y;
                } else if (!isClosingSlide && !isVertical && origY > 1.50 && origY < 1.90 && hasCategoryHeader) {
                    // Skip obsolete CATEGORY header line (1.850 in source) because CATEGORY header was skipped
                    return;
                } else if (!isClosingSlide && !isVertical && rowLineYMap[origY.toFixed(2)] !== undefined) {
                    lineY = rowLineYMap[origY.toFixed(2)];
                } else if (!isClosingSlide && !isVertical && origY >= 2.80 && origY <= 3.00 && w < 4.0 && !isTimelineOrRowListSlide) {
                    lineY = 2.915 + movedSublineShift;
                } else if (!isClosingSlide && !isVertical && origY > 1.50) {
                    let prevBottomInCol = SPACING.HEADER_DIVIDER_Y;
                    for (const prev of occupiedBoxes) {
                        if (x < prev.x + prev.w && x + w > prev.x) {
                            if (prev.bottom > prevBottomInCol) prevBottomInCol = prev.bottom;
                        }
                    }
                    if (prevBottomInCol > SPACING.HEADER_DIVIDER_Y && (lineY - prevBottomInCol > SPACING.XL)) {
                        lineY = parseFloat((prevBottomInCol + SPACING.SM).toFixed(3));
                    } else {
                        lineY = resolveVerticalCollision(x, lineY, w, 0.05, SPACING.XS);
                    }
                }

                if (!isClosingSlide) {
                    const isHeaderLine = (lineY <= 1.36);
                    lnColor = isDarkThemeSlide
                        ? (isHeaderLine ? "555558" : "38383C")
                        : (isHeaderLine ? "B0B0B4" : "D4D4D8");
                }

                pptSlide.addShape(pres.shapes.LINE, {
                    x, y: lineY, 
                    w: isVertical ? 0 : w, 
                    h: isVertical ? h : 0,
                    line: { color: lnColor, width: finalWidth, transparency: isClosingSlide ? 0 : 20 }
                });

                if (!isVertical && lineY > 1.36) {
                    recordOccupiedBox({ x, y: lineY, w, h: 0.05, txt: "[DIVIDER LINE]" });
                }
                return;
            }

            const origHasFill = shapeBgFill && shapeBgFill !== "none";
            const isCardPillShape = (origHasFill && w < 4.0 && ["034E48", "1C1C1E", "333333", "222222", "1E1E1E", "4DB89A", "2A2A2C"].includes(shapeBgFill.toUpperCase())) ||
                                    allCardPills.some(p => Math.abs(p.x - origX) < 0.35 && Math.abs(p.y - origY) < 0.35);

            // Detect single-letter acronym boxes (B, E, A, M etc.) — preserve original height
            const spFirstTxt = [...spXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join('').trim();
            const spFirstSzM = spXml.match(/sz="(\d+)"/);
            const spFirstFs = spFirstSzM ? parseInt(spFirstSzM[1]) / 100 : 0;
            const isAcronymLetter = isCardPillShape && spFirstTxt.length <= 2 && spFirstFs >= 36;

            const CARD_PILL_HEIGHT = SPACING.CARD_PILL_H;
            if (isCardPillShape && !isClosingSlide) {
                // Always use green (034E48) — on dark slides 1C1C1E is invisible against 121212 background
                shapeBgFill = "034E48";
                if (h < CARD_PILL_HEIGHT) h = CARD_PILL_HEIGHT;
            }

            // Draw card pill rectangles — use full original height for acronym letter boxes
            if (shapeBgFill && shapeBgFill !== "none" && isCardPillShape) {
                const boxPosKey = `${Math.round(x * 10)}_${Math.round(y * 10)}`;
                if (!renderedCardBoxes.has(boxPosKey)) {
                    renderedCardBoxes.add(boxPosKey);
                    const pillDrawH = isAcronymLetter ? h : CARD_PILL_HEIGHT;
                    pptSlide.addShape(pres.shapes.RECTANGLE, {
                        x, y, w, h: pillDrawH,
                        fill: { color: shapeBgFill },
                        line: { color: shapeBgFill, width: 0 }
                    });
                    recordOccupiedBox({ x, y, w, h: pillDrawH, txt: "[CARD PILL]" });
                }
            }

            // Draw wide section-header banners (e.g. "Illustrative use cases" with dark fill, w>=4.0)
            if (!isClosingSlide && hasTxBody && origHasFill &&
                ["034E48", "1C1C1E"].includes(shapeBgFill.toUpperCase()) && w >= 4.0 && h <= 0.55) {
                const bannerH = Math.max(h, 0.28);
                pptSlide.addShape(pres.shapes.RECTANGLE, {
                    x, y, w, h: bannerH,
                    fill: { color: "034E48" },
                    line: { color: "034E48", width: 0 }
                });
                recordOccupiedBox({ x, y, w, h: bannerH, txt: "[SECTION BANNER]" });
            }

            if (hasTxBody) {
                const txBodyMatch = spXml.match(/<p:txBody>[\s\S]*?<\/p:txBody>/) || spXml.match(/<a:txBody>[\s\S]*?<\/a:txBody>/);
                if (txBodyMatch) {
                    const pMatches = [...txBodyMatch[0].matchAll(/<a:p>[\s\S]*?<\/a:p>/g)];
                    const textRuns = [];
                    let maxFontSize = 8.55;

                    pMatches.forEach((p, pIdx) => {
                        const pXml = p[0];
                        let isBold = pXml.includes('b="1"') || pXml.includes('b="true"');
                        const isBullet = pXml.includes('<a:buChar') || pXml.includes('<a:buAutoNum');
                        const szMatch = pXml.match(/sz="(\d+)"/);
                        
                        let fontSize = szMatch ? parseFloat((parseInt(szMatch[1]) / 100).toFixed(1)) : 8.55;
                        
                        const isDarkCardBg = ["034E48", "1C1C1E", "333333", "222222", "1E1E1E"].includes((shapeBgFill || "").toUpperCase());
                        const isDarkContext = isDarkThemeSlide || isDarkCardBg;
                        let color = isDarkContext ? "FFFFFF" : "1D1D1F";

                        const clrMatch = pXml.match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/) || pXml.match(/<a:rPr[\s\S]*?<a:srgbClr val="([^"]+)"/);
                        if (clrMatch) {
                            const cVal = clrMatch[1].toUpperCase();
                            if (cVal === "4DB89A" || cVal === "034E48" || cVal === "0A3B36" || cVal === "08322D") {
                                color = isDarkContext ? "4DB89A" : "034E48";
                            }
                        }

                        const rawTxt = [...pXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join('');
                        const cleanTxt = decodeXmlEntities(rawTxt);
                        const isHugeTitle = cleanTxt.includes("LET'S BUILD") || cleanTxt.includes("WHAT'S NEXT");

                        const isColumnCategoryTitle = (!isClosingSlide && origY >= 2.2 && origY <= 3.0 && w <= 3.4 && cleanTxt.length < 40 && !isBullet && !cleanTxt.includes('.'));
                        const isBottomSectionHeader = (!isClosingSlide && (cleanTxt.startsWith("New team structure:") || cleanTxt.startsWith("Illustrative use cases") || cleanTxt.startsWith("The result:") || cleanTxt.startsWith("What enterprises now expect:") || cleanTxt.startsWith("85% client repeat") || cleanTxt.startsWith("What the platform")));
                        const isTimelineBadge = (!isClosingSlide && (cleanTxt.endsWith("Days") || cleanTxt.endsWith("Weeks") || cleanTxt.endsWith("Months") || cleanTxt.startsWith("Phase") || cleanTxt.startsWith("Step")));

                        if (isHugeTitle) {
                            fontSize = 36.0;
                            isBold = false;
                            color = "ECE9E4";
                        } else if (cleanTxt && cleanTxt.trim().length <= 2 && (fontSize >= 28 || fontSize >= 50)) {
                            fontSize = 48.0;
                            isBold = false;
                            color = isDarkContext ? "ECE9E4" : "034E48";
                        } else if (isTimelineBadge) {
                            fontSize = 8.0;
                            isBold = true;
                            color = isDarkContext ? "4DB89A" : "034E48";
                        } else if (sIdx === 0 && fontSize >= 24) {
                            fontSize = 28.0;
                            color = "FFFFFF";
                        } else if (!isClosingSlide) {
                            if (fontSize >= 16 || (origY < 1.35 && origX >= 1.8)) {
                                fontSize = 18.8;
                                color = isDarkContext ? "FFFFFF" : "1A1A1A";
                            } else if (fontSize >= 10.0 || isColumnCategoryTitle || isBottomSectionHeader || (cleanTxt === cleanTxt.toUpperCase() && cleanTxt.length < 50 && !isBullet && !cleanTxt.includes('.')) || cleanTxt.endsWith(':')) {
                                fontSize = 8.0; // 8pt bold inside card boxes and category headers
                                isBold = true;
                                if (isDarkContext) {
                                    color = isDarkCardBg ? "4DB89A" : "FFFFFF";
                                } else {
                                    color = "034E48";
                                }
                            } else {
                                fontSize = 8.0; // Standard body text font size (8pt)
                                if (isDarkContext) {
                                    color = "FFFFFF";
                                } else {
                                    color = "2C2C2E";
                                }
                            }
                        }

                        if (fontSize > maxFontSize) maxFontSize = fontSize;

                        if (cleanTxt) {
                            const finalRunColor = (isCardPillShape || (matchedCardInfo && matchedCardInfo.isHeader)) ? "FFFFFF" : color;
                            textRuns.push({
                                text: (isBullet ? "• " : "") + cleanTxt,
                                options: {
                                    color: finalRunColor,
                                    fontSize: toEvenPt(fontSize),
                                    bold: isHugeTitle ? false : isBold,
                                    fontFace: isHugeTitle ? "Inter" : (isClosingSlide ? (isBold ? "Inter Bold" : "Inter") : (isBold ? "Inter Medium" : "Inter")),
                                    breakLine: pIdx < pMatches.length - 1,
                                    paraSpaceAfter: isBullet ? 2.0 : (pMatches.length > 5 ? 1.5 : 3.0),
                                    paraSpaceBefore: (!isBullet && pIdx > 0 && cleanTxt.length > 5) ? (pMatches.length > 5 ? 1.5 : 2.5) : 0
                                }
                            });
                        }
                    });

                    if (textRuns.length > 0) {
                        const isHeaderOnPill = matchedCardInfo && matchedCardInfo.isHeader;
                        let align = isHeaderOnPill ? "center" : "left";
                        let valign = isHeaderOnPill ? "middle" : "top";

                        let finalY = y;
                        let estH = calculateTextShapeHeight(spXml, w, maxFontSize);
                        // Acronym letters (B/E/A/M): use original shape height, not 0.50in pill height
                        const isAcronymTextBox = maxFontSize >= 36 && cleanTxtForCheck.trim().length <= 2;
                        let finalH = isAcronymTextBox ? h
                                   : isHeaderOnPill ? SPACING.CARD_PILL_H
                                   : parseFloat(Math.max(0.20, estH).toFixed(2));

                        if (sIdx === 0) {
                            if (cleanTxtForCheck.includes("CONFIDENTIAL") || cleanTxtForCheck.includes("PROPRIETARY")) {
                                x = 0.50;
                                finalY = 4.96;
                                w = 3.60;
                                finalH = 0.35;
                            } else if (cleanTxtForCheck.includes("Created By")) {
                                x = 0.50;
                                finalY = 3.90;
                                w = 5.50;
                                finalH = 0.40;
                            } else {
                                const isCoverTitle = (maxFontSize >= 20 || cleanTxtForCheck.length > 25 || origY < 2.3);
                                if (isCoverTitle && coverTitleBottomY === 0) {
                                    finalY = 1.95;
                                    const titleLines = Math.max(1, Math.ceil(cleanTxtForCheck.length / 36));
                                    const realTitleH = parseFloat((titleLines * 0.45).toFixed(2));
                                    finalH = realTitleH;
                                    coverTitleBottomY = finalY + realTitleH;
                                } else {
                                    finalY = coverTitleBottomY > 0 ? parseFloat((coverTitleBottomY + SPACING.SM).toFixed(2)) : 2.50;
                                    finalH = 0.50;
                                    coverTitleBottomY = parseFloat((finalY + finalH).toFixed(2));
                                }
                            }
                        } else if (!isClosingSlide) {
                            if (!isRowMapElement || (matchedCardInfo && !matchedCardInfo.isHeader)) {
                                finalY = resolveVerticalCollision(x, finalY, w, finalH, SPACING.SM);
                            }
                        }

                        // Prevent cuts / overflow by dynamic height clamping and bottom margin floor protection
                        if (!isClosingSlide && finalY + finalH > SPACING.CONTENT_MAX_BOTTOM) {
                            if (finalY + finalH > SPACING.CONTENT_MAX_BOTTOM + 0.15) {
                                console.warn(`    [OVERFLOW WARNING] Slide ${sNum}: Element at x=${x.toFixed(2)} y=${finalY.toFixed(2)} bottom=${(finalY+finalH).toFixed(2)} exceeds boundary ${SPACING.CONTENT_MAX_BOTTOM}`);
                            }
                            finalH = parseFloat((SPACING.CONTENT_MAX_BOTTOM - finalY).toFixed(2));
                            if (finalH < 0.20) finalH = 0.20;
                        }

                        const textOpts = {
                            x, 
                            y: finalY, 
                            w, 
                            h: finalH,
                            align,
                            valign,
                            margin: [0, 0, 0, 0]
                        };

                        if (maxFontSize >= 40) {
                            textOpts.lineSpacingMultiple = 1.0;
                        } else if (maxFontSize >= 16) {
                            textOpts.lineSpacingMultiple = 1.18;
                        } else {
                            textOpts.lineSpacingMultiple = (textRuns.length > 4 || finalH > 1.2) ? 1.22 : 1.28;
                        }

                        pptSlide.addText(textRuns, textOpts);
                        recordOccupiedBox({ x, y: finalY, w, h: finalH, txt: cleanTxtForCheck });
                    }
                }
            }
        });

        console.log(`  Slide ${sNum}/${slideEntries.length} processed.`);
    }

    // OUTPUT ONLY TO new MD — no syncing to any other folders
    const outPath1 = path.join(OUT_DIR_1, outFileName);
    const outPath2 = path.join(OUT_DIR_2, outFileName);
    await pres.writeFile({ fileName: outPath1 });
    if (OUT_DIR_1 !== OUT_DIR_2) {
        fs.copyFileSync(outPath1, outPath2);
    }
    console.log(`  Written to: ${outPath1}`);
}

async function main() {
    for (const refFileName of Object.keys(fileMap)) {
        const outFileName = fileMap[refFileName];
        await processDeck(refFileName, outFileName);
    }
    console.log(`\n🎉 All 11 presentation decks rebuilt. Files are in: SALES DECKS/new MD/ and new MD/`);
}

main().catch(console.error);
