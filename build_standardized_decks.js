const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const AdmZip = require('adm-zip');

const REF_DIR = "C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images";
const OUT_DIR_1 = "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\new MD";
const OUT_DIR_2 = "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\new MD";
const MEDIA_EXTRACT_DIR = "C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\extracted_all_media";

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

const bgMappings = JSON.parse(fs.readFileSync('deck_bg_mappings.json', 'utf8'));
const coverBgMap = fs.existsSync('cover_bg_map.json') ? JSON.parse(fs.readFileSync('cover_bg_map.json', 'utf8')) : {};

// LOAD BDC DECK (COPY) EXACT CLOSING SLIDE
const refDeckZip = new AdmZip("C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\BDC Deck (Copy).pptx");
const refSlideEntries = refDeckZip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'));
refSlideEntries.sort((a, b) => parseInt(a.entryName.match(/slide(\d+)\.xml/)[1]) - parseInt(b.entryName.match(/slide(\d+)\.xml/)[1]));
const closingSlideEntry = refSlideEntries[refSlideEntries.length - 1];
const closingSlideXml = refDeckZip.readAsText(closingSlideEntry);

const closingSlideRelPath = `ppt/slides/_rels/${path.basename(closingSlideEntry.entryName)}.rels`;
const closingSlideRelEntry = refDeckZip.getEntries().find(e => e.entryName === closingSlideRelPath);

const closingMediaDir = path.join(MEDIA_EXTRACT_DIR, "BDC_Deck_Copy_Closing");
if (!fs.existsSync(closingMediaDir)) fs.mkdirSync(closingMediaDir, { recursive: true });
refDeckZip.getEntries().filter(e => e.entryName.startsWith('ppt/media/')).forEach(m => {
    fs.writeFileSync(path.join(closingMediaDir, path.basename(m.entryName)), m.getData());
});

const closingRelMap = {};
if (closingSlideRelEntry) {
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

function calculateTextShapeHeight(spXml, w) {
    const txBodyMatch = spXml.match(/<p:txBody>[\s\S]*?<\/p:txBody>/) || spXml.match(/<a:txBody>[\s\S]*?<\/a:txBody>/);
    if (!txBodyMatch) return 0.35;
    
    const pMatches = [...txBodyMatch[0].matchAll(/<a:p>[\s\S]*?<\/a:p>/g)];
    let totalH = 0.10;
    const pCount = pMatches.length;

    pMatches.forEach(p => {
        const pXml = p[0];
        const szMatch = pXml.match(/sz="(\d+)"/);
        let fontSize = szMatch ? parseFloat((parseInt(szMatch[1]) / 100).toFixed(1)) : 9.0;
        // Normalize font sizes to design system
        if (pCount > 5 && fontSize <= 11) fontSize = 9.0;
        else if (fontSize >= 9.0) fontSize = 9.0;
        
        const isBullet = pXml.includes('<a:buChar') || pXml.includes('<a:buAutoNum');
        const rawTxt = [...pXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join('');
        
        if (rawTxt.trim().length === 0) return;
        
        // More accurate chars-per-line for Inter 9pt at given box width
        const charsPerLine = Math.max(1, Math.floor(w * 14.5));
        const lines = Math.ceil(rawTxt.length / charsPerLine);
        
        const lineHeight = (fontSize / 72) * 1.45; // 1.45x line height for clean readability
        const paraSpace = (isBullet ? 3 : 5) / 72;  // generous para gap for breathing room
        
        totalH += (lines * lineHeight) + paraSpace;
    });
    
    return Math.max(totalH, 0.35);
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
    const zip = new AdmZip(refFilePath);
    const zipEntries = zip.getEntries();

    const deckMediaDir = path.join(MEDIA_EXTRACT_DIR, refFileName.replace('.pptx', ''));
    if (!fs.existsSync(deckMediaDir)) fs.mkdirSync(deckMediaDir, { recursive: true });

    // Extract all media files for this deck
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
        const isDarkThemeSlide = sIdx === 0 || isClosingSlide || darkSlidesList.includes(sNum);
        
        const xml = isClosingSlide ? closingSlideXml : zip.readAsText(sEntry);

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

        let bgPath = null;
        if (deckBgInfo && deckBgInfo.slides[sIdx]) {
            const imgs = deckBgInfo.slides[sIdx].images;
            bgPath = imgs.find(i => i.includes('slide_bg') || i.endsWith('.jpeg') || i.endsWith('.jpg')) || null;
        }

        if (isClosingSlide) {
            // RENDER CLEAN, PERFECT CLOSING SLIDE MASTER LAYOUT
            pptSlide.background = { color: "034E48" };

            // 1. Top Header Image
            const hdrImg = path.join(MEDIA_EXTRACT_DIR, "BDC_Deck_Copy_Closing", "image-10-1.png");
            if (fs.existsSync(hdrImg)) {
                pptSlide.addImage({ path: hdrImg, x: 0.16, y: 0.0, w: 9.69, h: 0.57 });
            } else {
                pptSlide.addText("BOMBAYDC", { x: 0.16, y: 0.28, w: 2.0, h: 0.3, fontSize: 11.5, fontFace: "Inter Medium", color: "ECE9E4", align: "left" });
                pptSlide.addText("bombaydc.com", { x: 7.84, y: 0.28, w: 2.0, h: 0.3, fontSize: 11.5, fontFace: "Inter Medium", color: "ECE9E4", align: "right" });
            }

            // 2. Title & Subtitle (Style 1 Title & Style 3 Body/Subtitle)
            pptSlide.addText("LET'S BUILD", { x: 1.99, y: 0.65, w: 7.85, h: 0.55, fontSize: 36.0, fontFace: "Inter Medium", color: "ECE9E4", valign: "top", margin: 0 });
            pptSlide.addText("WHAT'S NEXT.", { x: 1.99, y: 1.15, w: 7.85, h: 0.55, fontSize: 36.0, fontFace: "Inter Medium", color: "ECE9E4", valign: "top", margin: 0 });
            pptSlide.addText("Explore our work, sectors, and point of view at www.bombaydc.com", { x: 1.99, y: 1.70, w: 7.85, h: 0.35, fontSize: 9.0, fontFace: "Inter", color: "B4B4B4", valign: "top", margin: 0 });

            // 3. Horizontal Line Divider
            pptSlide.addShape(pres.shapes.LINE, { x: 1.99, y: 2.15, w: 7.85, h: 0, line: { color: "3E8D86", width: 0.5 } });

            // 4. Black Contact Card Box
            pptSlide.addShape(pres.shapes.RECTANGLE, { x: 1.99, y: 2.50, w: 3.8, h: 1.65, fill: { color: "000000" }, line: { width: 0 } });

            // 5. Contact Profile Photo
            const profileImg = path.join(MEDIA_EXTRACT_DIR, "BDC_Deck_Copy_Closing", "image-10-2.png");
            if (fs.existsSync(profileImg)) {
                pptSlide.addImage({ path: profileImg, x: 2.24, y: 2.75, w: 1.15, h: 1.15, rounding: true });
            }

            // 6. Contact Details Text (Style 2 Heading & Style 3 Body)
            pptSlide.addText("Siddesh Pednekar", { x: 3.55, y: 2.75, w: 2.1, h: 0.3, fontSize: 11.0, fontFace: "Inter Medium", color: "FFFFFF", valign: "top", margin: 0 });
            pptSlide.addText("Partner & COO", { x: 3.55, y: 3.10, w: 2.1, h: 0.2, fontSize: 9.0, fontFace: "Inter", color: "B4B4B4", valign: "top", margin: 0 });
            pptSlide.addText("sid@bombaydc.com", { x: 3.55, y: 3.33, w: 2.1, h: 0.2, fontSize: 9.0, fontFace: "Inter", color: "B4B4B4", valign: "top", margin: 0 });
            pptSlide.addText("9819981354", { x: 3.55, y: 3.55, w: 2.1, h: 0.2, fontSize: 9.0, fontFace: "Inter", color: "B4B4B4", valign: "top", margin: 0 });

            // 7. Legal Confidentiality Footer (Style 3 Body)
            pptSlide.addText("CONFIDENTIAL AND PROPRIETARY | © BombayDC. This material is intended solely for your internal use and any use of this material without specific permission of BombayDC is strictly prohibited. All rights reserved.", {
                x: 1.99, y: 4.85, w: 7.5, h: 0.45, fontSize: 9.0, fontFace: "Inter", color: "B4B4B4", lineSpacingMultiple: 1.2, margin: 0
            });

            console.log(`  Slide ${sNum}/${slideEntries.length} processed (Master Closing Slide).`);
            continue;
        } else if (sIdx === 0 && coverBgMap[refFileName] && fs.existsSync(coverBgMap[refFileName])) {
            // COVER SLIDE ONLY: Set original cover background image!
            pptSlide.background = { path: coverBgMap[refFileName] };
        } else {
            const bgColor = isDarkThemeSlide ? "121212" : "ECE9E4";
            pptSlide.background = { color: bgColor };
        }

        // 2. Parse top-level shapes from XML
        const shapes = getTopLevelShapes(xml);

        // PRE-PASS 1: Group and align Card Grids into uniform rows
        const cardPills = [];
        const allCardPills = [];
        const cardRowMap = {};
        const renderedCardBoxes = new Set();

        shapes.forEach(shapeObj => {
            const offMatch = shapeObj.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const extMatch = shapeObj.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            if (!offMatch || !extMatch) return;
            const px = parseFloat((parseInt(offMatch[1]) / 914400).toFixed(3));
            const py = parseFloat((parseInt(offMatch[2]) / 914400).toFixed(3));
            const ph = parseFloat((parseInt(extMatch[2]) / 914400).toFixed(3));
            
            const spPrMatch = shapeObj.xml.match(/<p:spPr>[\s\S]*?<\/p:spPr>/);
            const fillMatch = spPrMatch ? spPrMatch[0].match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/) : null;
            const fill = fillMatch ? fillMatch[1].toUpperCase() : null;
            const isCardPill = fill && ["034E48", "4DB89A", "1C1C1E", "224B12", "1A3632", "0A3B36", "08322D", "0D524A", "004B44", "024E48", "034D47", "333333", "222222", "1E1E1E", "2A2A2C", "3A3A3C"].includes(fill);
            
            if (isCardPill) {
                cardPills.push({ x: px, y: py, h: ph });
                allCardPills.push({ x: px, y: py });
            }
        });

        if (cardPills.length >= 2) {
            const row1Pills = cardPills.filter(p => p.y < 2.8);
            const row2Pills = cardPills.filter(p => p.y >= 2.8);

            if (row1Pills.length > 0) {
                const r1Y = 2.05;
                row1Pills.forEach(p => {
                    cardRowMap[`${p.x.toFixed(2)}_${p.y.toFixed(2)}`] = { cardX: p.x, cardY: r1Y, textY: r1Y + 0.62 };
                });
            }
            if (row2Pills.length > 0) {
                const r2Y = 3.50;
                row2Pills.forEach(p => {
                    cardRowMap[`${p.x.toFixed(2)}_${p.y.toFixed(2)}`] = { cardX: p.x, cardY: r2Y, textY: r2Y + 0.62 };
                });
            }
        }

        // PRE-PASS 2: Unified Header Processing (Title + Subline in ONE Single Container)
        // STRICT RULE: Divider line is FIXED at y: 1.35 in, Header Container is bounded between y: 0.57 in and y: 1.33 in.
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

            // Detect title shape (top Y < 0.85, X >= 1.8, fontSize >= 16)
            if (py < 0.85 && px >= 1.8 && fontSize >= 16 && shapeObj.xml.includes('<a:t>')) {
                const rawTxt = [...shapeObj.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').replace(/\s+/g, ' ');
                headerTitleText = decodeXmlEntities(rawTxt).trim().toUpperCase();
                if (headerTitleText.length > 0) hasHeaderTitle = true;
            }

            // Detect subline shape (Y 0.75-1.30, X >= 1.8, fontSize < 14)
            if (py >= 0.75 && py <= 1.30 && px >= 1.8 && fontSize > 0 && fontSize < 14 && shapeObj.xml.includes('<a:t>')) {
                const rawTxt = [...shapeObj.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').replace(/\s+/g, ' ').trim();
                if (rawTxt.length > 3 && !rawTxt.startsWith('1') && !rawTxt.startsWith('2') && !rawTxt.startsWith('3') && !rawTxt.startsWith('4') && !rawTxt.startsWith('5')) {
                    headerSublineText = decodeXmlEntities(rawTxt);
                }
            }
        });

        // Detect if title wraps to 2 lines
        const isTitle2Lines = headerTitleText.length > 38;

        // isDarkThemeSlide already detected at the start of loop

        // Top header picture bar is rendered consistently across all slides via picture shape loop

        // RENDER HEADER CONTAINER ABOVE DIVIDER (y: 0.57 in)
        if (hasHeaderTitle && !isClosingSlide) {
            // USER RULE: If title wraps to 2 lines, do NOT display subtitle above divider.
            // Only display subtitle above divider if title is 1 line.
            const displaySublineAbove = !isTitle2Lines ? headerSublineText : "";

            const headerRuns = [
                {
                    text: headerTitleText,
                    options: {
                        color: isDarkThemeSlide ? "FFFFFF" : "1A1A1A",
                        fontSize: 18.8,
                        bold: false,
                        fontFace: "Inter Medium",
                        breakLine: displaySublineAbove ? true : false,
                        charSpacing: 0.3, // Slight letter-spacing for sharp modern look
                        paraSpaceAfter: displaySublineAbove ? 5 : 0
                    }
                }
            ];

            if (displaySublineAbove) {
                headerRuns.push({
                    text: displaySublineAbove,
                    options: {
                        color: isDarkThemeSlide ? "B4B4B4" : "6A6A6B",
                        fontSize: 9.0, // Style 3: Body Text / Subline
                        bold: false,
                        fontFace: "Inter",
                        paraSpaceBefore: 3
                    }
                });
            }

            pptSlide.addText(headerRuns, {
                x: 1.99,
                y: 0.57,
                w: 7.85,
                h: 0.76, // Fits safely within 0.57 in to 1.33 in band
                valign: "top",
                margin: [0, 0, 0, 0]
            });
        }

        // USER RULE: If title wraps to 2 lines and subline exists, render subline as first line of content below divider (at y: 1.60 in)
        let movedSublineShift = 0;
        if (isTitle2Lines && headerSublineText && !isClosingSlide) {
            movedSublineShift = 0.35; // Pushes subsequent content down from y: 1.60 in to y: 1.95 in
            pptSlide.addText([{
                text: headerSublineText,
                options: {
                    color: isDarkThemeSlide ? "ECE9E4" : "555555",
                    fontSize: 9.0, // Style 3: Body Text / Subline
                    bold: false,
                    fontFace: "Inter",
                    paraSpaceAfter: 4
                }
            }], {
                x: 1.99,
                y: 1.60,
                w: 7.85,
                h: 0.30,
                valign: "top",
                margin: [0, 0, 0, 0]
            });
        }

        // Per-slide layout state (reset cleanly for EVERY slide)
        const columnBottoms = {};
        const origYSnapCache = {};
        let coverTextShapeCount = 0;
        let coverTitleBottomY = 0;
        let maxHeaderBottom = 0;

        // PRE-PASS 3: Column Header Lower Bounds & Acronym/Framework Number Flow Engine
        // ─────────────────────────────────────────────────────────────────────────────
        let gapCompressShift = 0;
        const columnHeaderLowerBounds = {}; // colKey -> post-shift bottom-Y of column header label
        const bigNumberBottoms = {};        // colKey -> post-shift bottom-Y of big 60pt number (1 2 3 4 / B E A M)

        if (!isClosingSlide && sIdx > 0) {
            let minContentOrigY = Infinity;

            shapes.forEach(shapeObj => {
                const off2 = shapeObj.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/); 
                const ext2 = shapeObj.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/); 
                if (!off2 || !ext2 || !shapeObj.xml.includes('<a:t>')) return;

                const p3OrigY = parseFloat((parseInt(off2[2]) / 914400).toFixed(3));
                const p3OrigX = parseFloat((parseInt(off2[1]) / 914400).toFixed(3));
                const p3W     = parseFloat((parseInt(ext2[1]) / 914400).toFixed(3));
                const p3H     = parseFloat((parseInt(ext2[2]) / 914400).toFixed(3));

                if (p3OrigY <= 1.75 || p3OrigX < 1.0) return; // skip header title, subline, index area, and top subline bullets
                const rawTxt3 = [...shapeObj.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ');
                const cleanTxt3 = decodeXmlEntities(rawTxt3).trim();
                if (!cleanTxt3) return;
                if (cleanTxt3.includes('CONFIDENTIAL') || cleanTxt3.includes('bombaydc.com')) return;

                // Track minimum content Y for gap compression
                minContentOrigY = Math.min(minContentOrigY, p3OrigY);

                const szMatch = shapeObj.xml.match(/sz="(\d+)"/);
                const fontSize = szMatch ? parseFloat((parseInt(szMatch[1]) / 100).toFixed(1)) : 0;

                // Detect Big Numbers / Acronym Letters (e.g. 1 2 3 4 / B E A M at fontSize >= 28 or single char)
                if (fontSize >= 28 || (cleanTxt3.length <= 2 && (cleanTxt3.match(/^[0-9A-Z]$/i) || fontSize >= 20))) {
                    const colKey = Math.round(p3OrigX * 2) / 2;
                    const numStartY = Math.max(1.80 + movedSublineShift, p3OrigY + movedSublineShift);
                    const numBottom = numStartY + 0.75; // 60pt font height is ~0.75in
                    bigNumberBottoms[colKey] = parseFloat(numBottom.toFixed(3));
                }

                // Detect column header overlay labels (e.g. INTELLIGENT PRODUCTS, A PARALLEL LIFECYCLE, A NEW TEAM MODEL)
                if (p3OrigY >= 1.70 && p3OrigY <= 3.50 && p3W >= 1.0 && p3W < 4.0) {
                    const normTxt3 = cleanTxt3.replace(/\s+/g, ' ').trim();
                    if (normTxt3 === normTxt3.toUpperCase() && normTxt3.length > 3 && normTxt3.length < 55) {
                        const colKey = Math.round(p3OrigX * 2) / 2;
                        // Target Y for column headers is ~2.03 in. Body text below must snap to 2.45 in!
                        const targetHeaderBottom = 2.45;
                        if (!columnHeaderLowerBounds[colKey] || targetHeaderBottom > columnHeaderLowerBounds[colKey].bottomY) {
                            columnHeaderLowerBounds[colKey] = { bottomY: targetHeaderBottom, origY: p3OrigY };
                        }
                    }
                }
            });

            // Compute gap compression shift
            if (minContentOrigY < Infinity) {
                const rawCompress = minContentOrigY - 1.80;
                if (rawCompress > 0.15) { // Pull up ANY content gap > 0.15 in to close empty voids
                    gapCompressShift = parseFloat((rawCompress - 0.05).toFixed(3));
                }
            }
        }

        // PRE-PASS 4: Dynamic Timeline & Row-List Layout Engine
        // ──────────────────────────────────────────────────────
        // Dynamically calculates standard row height based on the row with the most content lines,
        // reduces empty vertical gaps, and aligns row divider lines & bottom content perfectly.
        const rowLayoutMap = {}; // origYKey -> { finalY, lineY }
        const rowLineYMap = {};  // origLineYKey -> finalLineY

        if (!isClosingSlide && sIdx > 0) {
            const listShapes = shapes.filter(s => {
                const offMatch = s.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                const extMatch = s.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
                if (!offMatch || !s.xml.includes('<a:t>')) return false;
                const py = parseFloat((parseInt(offMatch[2]) / 914400).toFixed(3));
                const px = parseFloat((parseInt(offMatch[1]) / 914400).toFixed(3));
                const pw = extMatch ? parseFloat((parseInt(extMatch[1]) / 914400).toFixed(3)) : 3.0;
                if (py < 1.70 || px < 1.0 || pw >= 5.0) return false; // PRE-PASS 4 row-clustering applies ONLY to multi-column list items (pw < 5.0)
                const rawTxt = [...s.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ');
                const clean = decodeXmlEntities(rawTxt).trim();
                return clean.length > 0 && !clean.includes('CONFIDENTIAL') && !clean.includes('bombaydc.com');
            });

            // Group into horizontal row clusters
            const clusters = [];
            listShapes.forEach(s => {
                const offMatch = s.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                const py = parseFloat((parseInt(offMatch[2]) / 914400).toFixed(3));
                const px = parseFloat((parseInt(offMatch[1]) / 914400).toFixed(3));
                const extMatch = s.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
                const pw = extMatch ? parseFloat((parseInt(extMatch[1]) / 914400).toFixed(3)) : 3.0;
                
                const rawTxt = [...s.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ');
                const clean = decodeXmlEntities(rawTxt).trim();
                if (clean.includes("team structure") || clean.includes("use cases") || clean.includes("THE RESULT")) return;

                let c = clusters.find(cl => Math.abs(cl.origY - py) < 0.28);
                if (!c) {
                    c = { origY: py, shapes: [] };
                    clusters.push(c);
                }
                c.shapes.push({ s, py, px, pw, xml: s.xml });
            });

            if (clusters.length >= 2) {
                clusters.sort((a, b) => a.origY - b.origY);

                // Find max estimated text height across all rows
                let maxRowH = 0.40;
                clusters.forEach(c => {
                    c.shapes.forEach(sh => {
                        const estH = calculateTextShapeHeight(sh.xml, sh.pw);
                        if (estH > maxRowH) maxRowH = estH;
                    });
                });

                const standardRowH = Math.max(0.62, parseFloat((maxRowH + 0.16).toFixed(2)));
                let curY = 1.85 + movedSublineShift;
                if (gapCompressShift > 0) {
                    curY = Math.max(1.68 + movedSublineShift, curY - gapCompressShift);
                }

                clusters.forEach((c, idx) => {
                    const rowContentY = curY;
                    const rowDividerY = parseFloat((curY + standardRowH).toFixed(2));

                    c.shapes.forEach(sh => {
                        rowLayoutMap[sh.py.toFixed(2)] = { finalY: rowContentY, lineY: rowDividerY };
                    });

                    // Map matching line divider
                    shapes.forEach(lineSh => {
                        if (lineSh.tag === '<p:cxnSp>' || lineSh.xml.includes('prst="line"')) {
                            const lOff = lineSh.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                            if (lOff) {
                                const lY = parseFloat((parseInt(lOff[2]) / 914400).toFixed(3));
                                if (Math.abs(lY - (c.origY + 0.8)) < 0.45 || Math.abs(lY - c.origY) < 0.45) {
                                    rowLineYMap[lY.toFixed(2)] = rowDividerY;
                                }
                            }
                        }
                    });

                    curY = parseFloat((rowDividerY + 0.08).toFixed(2));
                });

                // Set bottom section start Y for explicit bottom callout banners
                const bottomStart = curY + 0.05;
                shapes.forEach(sh => {
                    const offM = sh.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                    if (offM) {
                        const py = parseFloat((parseInt(offM[2]) / 914400).toFixed(3));
                        const rawTxt = [...sh.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ');
                        const clean = decodeXmlEntities(rawTxt).trim();
                        if (clean.includes("team structure") || clean.includes("use cases") || clean.includes("THE RESULT")) {
                            rowLayoutMap[py.toFixed(2)] = { finalY: bottomStart };
                        }
                    }
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

            // Skip full-slide background solid shapes that hide the background image
            if (w >= 9.9 && h >= 5.5 && shapeTag !== '<p:pic>') return;

            // Skip wide dark/green decorative banner shapes (wide but short, no text, from old dark slides)
            // These are decorative footer/header bars that become black/green bars on light slides
            if (shapeTag !== '<p:pic>' && !spXml.includes('<a:t>')) {
                const _spPrDecor = spXml.match(/<p:spPr>[\s\S]*?<\/p:spPr>/);
                const _fillDecor = _spPrDecor ? _spPrDecor[0].match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/) : null;
                const _fColorDecor = _fillDecor ? _fillDecor[1].toUpperCase() : null;
                if (_fColorDecor && ["034E48","1D1D1F","121212","000000","0F2C28","1C1C1E","0A3B36"].includes(_fColorDecor)
                    && w > 7.0 && h > 0 && h < 0.65) return; // Skip dark decorative bars
            }

            // Skip title and subline shapes FIRST because they are rendered as part of unified header container
            const rawTxtForCheck = [...spXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').replace(/\s+/g, ' ').trim();
            const cleanTxtForCheck = decodeXmlEntities(rawTxtForCheck);

            if (!isClosingSlide && sIdx > 0 && spXml.includes('<a:t>')) {
                const szMatch = spXml.match(/sz="(\d+)"/);
                const fontSize = szMatch ? parseFloat((parseInt(szMatch[1]) / 100).toFixed(1)) : 0;
                
                // Skip original top header logo text shapes (BOMBAYDC & bombaydc.com at origY < 0.60)
                if (origY < 0.60 && (cleanTxtForCheck === "BOMBAYDC" || cleanTxtForCheck.includes("bombaydc.com") || (origX < 0.5 && fontSize < 12) || (origX > 7.0 && fontSize < 12))) return;

                // Skip title shape
                if (origY < 0.85 && fontSize >= 16 && origX >= 1.8) return;
                if (headerTitleText && cleanTxtForCheck.toUpperCase() === headerTitleText) return;
                if (headerTitleText && cleanTxtForCheck.length > 10 && headerTitleText.includes(cleanTxtForCheck.toUpperCase())) return;

                // Skip full-width header sublines / taglines (prevent duplicate rendering & text overlapping!)
                if (origY >= 0.75 && origY <= 1.30 && w >= 5.0 && fontSize < 14 && fontSize > 0) return;
                if (headerSublineText && cleanTxtForCheck === headerSublineText) return;
                if (headerSublineText && cleanTxtForCheck.length > 8 && (headerSublineText.includes(cleanTxtForCheck) || cleanTxtForCheck.includes(headerSublineText.substring(0, 15)))) return;
            }

            // Match card pill or text shape via fuzzy proximity to cardPills
            let isRowMapElement = false;
            let matchedCardInfo = null;

            const matchingPill = cardPills.find(p => Math.abs(p.x - origX) < 0.40 && Math.abs(p.y - origY) < 1.20);
            if (matchingPill && !isClosingSlide) {
                const r1Y = (matchingPill.y < 2.8) ? 2.05 : 3.50;
                x = matchingPill.x; // SNAP X to exact Pill X!
                if (origY <= matchingPill.y + 0.35) {
                    // Header text or Pill box shape
                    y = r1Y + movedSublineShift;
                    matchedCardInfo = { cardX: matchingPill.x, cardY: r1Y, textY: r1Y + 0.62, isHeader: true };
                } else {
                    // Body text shape below pill
                    y = r1Y + 0.62 + movedSublineShift;
                    matchedCardInfo = { cardX: matchingPill.x, cardY: r1Y, textY: r1Y + 0.62, isHeader: false };
                }
                isRowMapElement = true;
            }

            // USER RULE: Push content below divider down by movedSublineShift on 2-line title slides
            if (!isClosingSlide && origY >= 1.35 && !isRowMapElement) {
                const origYKey = origY.toFixed(2);
                if (rowLayoutMap[origYKey] !== undefined) {
                    y = rowLayoutMap[origYKey].finalY;
                } else if (origYSnapCache[origYKey] !== undefined) {
                    // ROW SNAP: reuse exact same Y as other shapes in this row
                    y = origYSnapCache[origYKey];
                } else {
                    y += movedSublineShift;
                    // GAP COMPRESSION: pull content up if source PPTX had it positioned too far down
                    if (gapCompressShift > 0) {
                        y -= gapCompressShift;
                        const minAllowed = 1.58 + movedSublineShift;
                        if (y < minAllowed) y = minAllowed;
                    }
                    y = parseFloat(y.toFixed(3));
                    origYSnapCache[origYKey] = y; // Cache so all same-row shapes align
                }
            }

            // COLUMN BODY TEXT PUSH-DOWN & BIG NUMBER COLLISION PREVENTION
            // ─────────────────────────────────────────────────────────────
            if (!isClosingSlide && !isRowMapElement && origY >= 1.35 && w >= 1.0 && w < 5.0) {
                const colKey = Math.round(origX * 2) / 2;

                // A) Prevent 60pt Big Numbers (1 2 3 4 / B E A M) from colliding with labels/descriptions
                if (bigNumberBottoms[colKey] && origY > 1.70 && y < bigNumberBottoms[colKey]) {
                    y = parseFloat((bigNumberBottoms[colKey] + 0.05).toFixed(3));
                }

                // B) Position body text tightly below column headers (resolves empty gap on Slide 7)
                if (columnHeaderLowerBounds[colKey]) {
                    const isBodyText = (cleanTxtForCheck !== cleanTxtForCheck.toUpperCase() || cleanTxtForCheck.length > 40);
                    if (isBodyText && (origY >= columnHeaderLowerBounds[colKey].origY)) {
                        y = parseFloat(columnHeaderLowerBounds[colKey].bottomY.toFixed(3));
                    }
                }
            }

            // Anti-overlap: Multi-column grid collision prevention (applies ONLY to column shapes w < 5.0)
            if (!isClosingSlide && y > 1.4 && w < 5.0) {
                if (typeof maxHeaderBottom !== 'undefined' && maxHeaderBottom > 0) {
                    if (y < maxHeaderBottom + 0.10) {
                        y = parseFloat((maxHeaderBottom + 0.12).toFixed(2));
                    }
                }

                Object.keys(columnBottoms).forEach(colX => {
                    const colXVal = parseFloat(colX);
                    if (Math.abs(colXVal - x) < 1.0) {
                        const prevB = columnBottoms[colX];
                        if (cleanTxtForCheck.includes("THE RESULT") || cleanTxtForCheck.includes("The result:")) {
                            // Match top divider gap (0.25in)
                            if (y > prevB + 0.25) {
                                y = parseFloat((prevB + 0.25).toFixed(2));
                            }
                        } else if (cleanTxtForCheck.includes("What enterprises now expect")) {
                            // Prevent slide bottom cut-off
                            y = 4.70;
                        } else if (y < prevB + 0.10) {
                            y = parseFloat((prevB + 0.12).toFixed(2));
                        }
                    }
                });
            }

            if (spXml.includes('<a:t>') && !isClosingSlide && y > 1.4) {
                const estH = calculateTextShapeHeight(spXml, w);
                // For column text shapes (w < 5.0), use actual estimated text height so we don't overestimate column bottom
                const effectiveH = w < 5.0 ? Math.min(h, estH + 0.05) : Math.max(estH, h);
                if (w >= 5.0 && origY <= 1.65) {
                    maxHeaderBottom = Math.max(typeof maxHeaderBottom !== 'undefined' ? maxHeaderBottom : 0, y + effectiveH);
                } else if (w < 5.0) {
                    // Always update columnBottoms with the LARGER of current bottom vs new bottom
                    const newBottom = y + effectiveH;
                    if (!columnBottoms[x] || newBottom > columnBottoms[x]) {
                        columnBottoms[x] = newBottom;
                    }
                }
            }

            // Check if shape is a PICTURE (<p:pic>)
            if (shapeTag === '<p:pic>') {
                const rEmbedMatch = spXml.match(/r:embed="([^"]+)"/);
                if (rEmbedMatch && relMap[rEmbedMatch[1]]) {
                    const imgFile = relMap[rEmbedMatch[1]];
                    if (fs.existsSync(imgFile)) {
                        if (w >= 9.9 && h >= 5.5) return; // Skip full-slide background picture shapes
                        const isProfilePic = isClosingSlide && (w < 2.0 && y > 2.0);
                        pptSlide.addImage({ path: imgFile, x, y, w, h, rounding: isProfilePic });
                    }
                }
                return;
            }

            const spPrMatch = spXml.match(/<p:spPr>[\s\S]*?<\/p:spPr>/);
            const spPrXml = spPrMatch ? spPrMatch[0] : "";
            const fillMatch = spPrXml.match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/);
            let shapeBgFill = fillMatch ? fillMatch[1] : null;

            // Standardize primary brand green and closing slide card
            if (!isClosingSlide && shapeBgFill && ["1A3632", "224B12", "4DB89A", "0A3B36", "08322D", "0D524A", "004B44", "024E48", "034D47"].includes(shapeBgFill.toUpperCase())) {
                shapeBgFill = "034E48";
            }

            const hasTxBody = spXml.includes('<a:t>');
            const isLine = (shapeTag === '<p:cxnSp>' || spPrXml.includes('prst="line"') || h === 0 || w === 0) && !hasTxBody;

            if (isLine && (w > 0 || h > 0)) {
                const isVertical = (w === 0 || (h > 0 && w < 0.1));
                const lnColorMatch = spPrXml.match(/<a:ln[\s\S]*?<a:srgbClr val="([^"]+)"/);
                let lnColor = lnColorMatch ? lnColorMatch[1] : "B4B4B4";
                
                // Keep native width if it's the closing slide
                const wMatch = spPrXml.match(/<a:ln w="(\d+)"/);
                const nativeW = wMatch ? parseFloat((parseInt(wMatch[1]) / 12700).toFixed(2)) : 0.5;
                const finalWidth = isClosingSlide ? nativeW : (isVertical ? 0.30 : 0.50);

                let lineY = y;
                // Skip redundant duplicate header line right below the main divider line (y: 1.36 to 2.20 in for wide lines)
                if (!isClosingSlide && !isVertical && origY > 1.35 && origY <= 2.20 && w > 5.0) return;

                // STRICT RULE: Horizontal header divider line MUST ALWAYS be placed at y: 1.35 in!
                if (!isClosingSlide && !isVertical && origY >= 0.90 && origY <= 1.50) {
                    lineY = FIXED_DIVIDER_Y; // Always y: 1.35 in
                } else if (!isClosingSlide && !isVertical && rowLineYMap[origY.toFixed(2)] !== undefined) {
                    // Use dynamic row divider line Y from PRE-PASS 4
                    lineY = rowLineYMap[origY.toFixed(2)];
                } else if (!isClosingSlide && !isVertical && origY >= 2.80 && origY <= 3.00 && w < 4.0) {
                    // Center column separator lines at exact midpoint (y: 2.915 in) for 4-column grids
                    lineY = 2.915;
                }

                if (!isClosingSlide) {
                    // Header line: slightly more visible. Body row separator lines: very subtle.
                    const isHeaderLine = (lineY <= 1.36);
                    lnColor = isDarkThemeSlide
                        ? (isHeaderLine ? "4A4A4E" : "2C2C2E")
                        : (isHeaderLine ? "C8C8CC" : "E0E0E3"); // Crisp but non-intrusive
                }

                pptSlide.addShape(pres.shapes.LINE, {
                    x, y: lineY, 
                    w: isVertical ? 0 : w, 
                    h: isVertical ? h : 0,
                    line: { color: lnColor, width: finalWidth, transparency: isClosingSlide ? 0 : 40 }
                });
                return;
            }

            // STANDARDIZE CARD PILL & CONTAINER DIMENSIONS
            const origHasFill = shapeBgFill && shapeBgFill !== "none";
            const isCardPillShape = (origHasFill && ["034E48", "1C1C1E", "333333", "222222", "1E1E1E", "4DB89A", "2A2A2C"].includes(shapeBgFill.toUpperCase())) ||
                                    allCardPills.some(p => Math.abs(p.x - origX) < 0.35 && Math.abs(p.y - origY) < 0.35);

            if (isCardPillShape && !isClosingSlide) {
                if (isDarkThemeSlide) {
                    shapeBgFill = "1C1C1E"; // Sleek Charcoal box on dark background
                } else {
                    shapeBgFill = "034E48"; // Signature BombayDC Brand Green box on light background
                }
                if (h < 0.55) h = 0.55;
            }

            const isCardPill = isCardPillShape;

            // Draw shape background rectangle ONLY ONCE per card position (PREVENT DOUBLE BOXES!)
            if (shapeBgFill && shapeBgFill !== "none" && isCardPillShape) {
                const boxPosKey = `${Math.round(x * 10)}_${Math.round(y * 10)}`;
                if (!renderedCardBoxes.has(boxPosKey)) {
                    renderedCardBoxes.add(boxPosKey);

                    // Header Pill Box ONLY (no secondary box underneath to prevent overlapping)
                    pptSlide.addShape(pres.shapes.RECTANGLE, {
                        x, y, w, h: 0.55,
                        fill: { color: shapeBgFill },
                        line: { color: shapeBgFill, width: 0 }
                    });
                }
            }

            // Parse and render text runs with STANDARDIZED FONT and CENTERED ALIGNMENTS
            if (hasTxBody) {
                const txBodyMatch = spXml.match(/<p:txBody>[\s\S]*?<\/p:txBody>/) || spXml.match(/<a:txBody>[\s\S]*?<\/a:txBody>/);
                if (txBodyMatch) {
                    const pMatches = [...txBodyMatch[0].matchAll(/<a:p>[\s\S]*?<\/a:p>/g)];
                    const textRuns = [];
                    let maxFontSize = 8.55;

                    const alignMatch = txBodyMatch[0].match(/algn="([^"]+)"/);
                    let align = alignMatch ? (alignMatch[1] === "ctr" ? "center" : "left") : "left";

                    const isInsideCard = (shapeBgFill && shapeBgFill !== "none") || allCardPills.some(p => Math.abs(p.x - origX) < 0.35 && Math.abs(p.y - origY) < 0.35);

                    pMatches.forEach((p, pIdx) => {
                        const pXml = p[0];
                        let isBold = pXml.includes('b="1"') || pXml.includes('b="true"');
                        const isBullet = pXml.includes('<a:buChar') || pXml.includes('<a:buAutoNum');
                        const szMatch = pXml.match(/sz="(\d+)"/);
                        
                        let fontSize = szMatch ? parseFloat((parseInt(szMatch[1]) / 100).toFixed(1)) : 8.55;
                        
                        // TEXT COLOR CONTRAST & READABILITY RULES FOR DARK SLIDES & DARK CARDS
                        const clrMatch = pXml.match(/<a:srgbClr val="([^"]+)"/);
                        const isDarkBg = isDarkThemeSlide;
                        const isDarkCardBg = shapeBgFill && ["034E48", "1C1C1E", "1A3632", "0A3B36", "08322D", "0D524A", "004B44", "024E48"].includes(shapeBgFill.toUpperCase());
                        const isDarkContext = isDarkBg || isDarkCardBg;

                        let color = clrMatch ? clrMatch[1] : (isDarkContext ? "ECE9E4" : "1D1D1F");

                        const greenVariants = ["224B12", "1A3632", "4DB89A", "0A3B36", "08322D", "0D524A", "004B44", "024E48", "034D47", "2C5E3B", "1B5E20", "2E7D32", "355E3B", "1E4620", "3E8D86"];
                        if (color && greenVariants.includes(color.toUpperCase())) {
                            color = isDarkContext ? "4DB89A" : "034E48";
                        } else if (isDarkContext) {
                            // On dark bg: convert any dark/black text to off-white
                            if (!color || ["1D1D1F", "1E1E1E", "000000", "444444", "555555", "333333", "666666", "034E48"].includes(color.toUpperCase())) {
                                color = "ECE9E4";
                            }
                        } else {
                            // On light bg: convert white/off-white (invisible on beige) to readable dark
                            if (color && ["FFFFFF", "ECE9E4", "F5F5F5", "FAFAFA", "F0F0F0", "EEEEEE", "E0E0E0"].includes(color.toUpperCase())) {
                                color = "1D1D1F";
                            }
                        }

                        const rawTxt = [...pXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join('');
                        const cleanTxt = decodeXmlEntities(rawTxt);
                        const isHugeTitle = cleanTxt.includes("LET'S BUILD") || cleanTxt.includes("WHAT'S NEXT");

                        // STRICT 3-STYLE TYPOGRAPHY SYSTEM (Apple-inspired minimal design)
                        // Style 1: Slide Title (18.8pt Inter Medium)
                        // Style 2: Section / Content Heading (11.0pt Inter Medium Bold)
                        // Style 3: Body Text / Paragraphs / Bullets / Sublines (9.0pt Inter Regular)
                        const isColumnCategoryTitle = (!isClosingSlide && origY >= 2.2 && origY <= 3.0 && w <= 3.4 && cleanTxt.length < 40 && !isBullet && !cleanTxt.includes('.'));
                        const isBottomSectionHeader = (!isClosingSlide && (cleanTxt.includes("use cases") || cleanTxt.includes("team structure") || cleanTxt.includes("outcomes")));
                        const isTimelineBadge = (!isClosingSlide && (cleanTxt.endsWith("Days") || cleanTxt.endsWith("Weeks") || cleanTxt.endsWith("Months") || cleanTxt.startsWith("Phase") || cleanTxt.startsWith("Step")));

                        if (isHugeTitle) {
                            fontSize = 36.0;
                            isBold = false;
                            color = "ECE9E4";
                        } else if (cleanTxt && cleanTxt.trim().length <= 2 && (fontSize >= 28 || fontSize >= 50)) {
                            fontSize = 60.0; // BEAM / Acronym big numbers
                            isBold = false;
                            color = isDarkContext ? "ECE9E4" : "034E48";
                        } else if (isTimelineBadge) {
                            fontSize = 10.0;
                            isBold = true;
                            color = isDarkContext ? "4DB89A" : "034E48"; // Signature Green / Mint highlight badge
                        } else if (sIdx === 0 && fontSize >= 24) {
                            fontSize = 28.0; // Cover Title
                            color = "FFFFFF";
                        } else if (!isClosingSlide) {
                            if (fontSize >= 16 || (origY < 1.35 && origX >= 1.8)) {
                                fontSize = 18.8; // Style 1: Slide Title
                                color = isDarkContext ? "FFFFFF" : "1A1A1A";
                            } else if (fontSize >= 10.0 || isColumnCategoryTitle || isBottomSectionHeader || (cleanTxt === cleanTxt.toUpperCase() && cleanTxt.length < 50 && !isBullet && !cleanTxt.includes('.')) || cleanTxt.endsWith(':')) {
                                fontSize = 10.5; // Style 2: Section / Content Heading
                                isBold = true;
                                if (isDarkContext) {
                                    color = isDarkCardBg ? "4DB89A" : "FFFFFF"; // Mint on dark card, White on dark slide
                                } else {
                                    color = "034E48"; // Signature BombayDC Brand Green on light slides
                                }
                            } else {
                                fontSize = 9.0;  // Style 3: Body Text / Descriptions / Bullets
                                if (isDarkContext) {
                                    color = isDarkCardBg ? "E8E8EC" : "8A8A8E"; // Near-white on dark card, muted grey on dark slide
                                } else {
                                    color = "1D1D1F"; // Uniform Dark Charcoal for ALL body text on light slides
                                }
                            }
                        }

                        if (fontSize > maxFontSize) maxFontSize = fontSize;

                        if (cleanTxt) {
                            // Enforce White (#FFFFFF) text inside card boxes
                            const finalRunColor = (isCardPillShape || (matchedCardInfo && matchedCardInfo.isHeader)) ? "FFFFFF" : color;
                            textRuns.push({
                                text: (isBullet ? "• " : "") + cleanTxt,
                                options: {
                                    color: finalRunColor,
                                    fontSize,
                                    bold: isHugeTitle ? false : isBold,
                                    fontFace: isHugeTitle ? "Inter" : (isClosingSlide ? (isBold ? "Inter Bold" : "Inter") : (isBold ? "Inter Medium" : "Inter")),
                                    breakLine: pIdx < pMatches.length - 1,
                                    // Tighter spacing for dense bullet lists, generous for short single-para content
                                    paraSpaceAfter: isBullet ? 3 : (pMatches.length > 5 ? 2 : 5),
                                    paraSpaceBefore: (!isBullet && pIdx > 0 && cleanTxt.length > 5) ? (pMatches.length > 5 ? 2 : 4) : 0
                                }
                            });
                        }
                    });

                    if (textRuns.length > 0) {
                        const isHeaderOnPill = matchedCardInfo && matchedCardInfo.isHeader;
                        let align = isHeaderOnPill ? "center" : "left";
                        let valign = isHeaderOnPill ? "middle" : "top";

                        // Text inside boxes MUST ALWAYS be White (#FFFFFF) or Light Grey (#ECE9E4)
                        if (isHeaderOnPill || isCardPillShape) {
                            color = "FFFFFF";
                        }

                        let finalY = y;
                        let finalH = isHeaderOnPill ? 0.55 : parseFloat(Math.max(0.25, Math.min(h, calculateTextShapeHeight(spXml, w) + 0.05)).toFixed(2));

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
                                // COMPONENT-AWARE COVER SLIDE FLOW: Title height drives Subtitle Y position!
                                const isCoverTitle = (maxFontSize >= 20 || cleanTxtForCheck.length > 25 || origY < 2.3);
                                if (isCoverTitle && coverTitleBottomY === 0) {
                                    finalY = 1.95;
                                    const titleLines = Math.max(1, Math.ceil(cleanTxtForCheck.length / 36));
                                    const realTitleH = parseFloat((titleLines * 0.45).toFixed(2)); // Content-driven title height
                                    finalH = realTitleH;
                                    coverTitleBottomY = finalY + realTitleH;
                                } else {
                                    // Subtitle flows dynamically below Title bottom + 0.08 in gap!
                                    finalY = coverTitleBottomY > 0 ? parseFloat((coverTitleBottomY + 0.08).toFixed(2)) : 2.50;
                                    finalH = 0.50;
                                }
                            }
                        }

                        const MAX_SLIDE_BOTTOM = 5.10;
                        if (!isClosingSlide && finalY + finalH > MAX_SLIDE_BOTTOM) {
                            const maxAllowedH = parseFloat((MAX_SLIDE_BOTTOM - finalY).toFixed(2));
                            if (maxAllowedH >= 0.25) {
                                finalH = maxAllowedH;
                            } else {
                                finalY = parseFloat((MAX_SLIDE_BOTTOM - finalH).toFixed(2));
                                const minAllowedTopY = 1.60 + movedSublineShift;
                                if (finalY < minAllowedTopY) finalY = minAllowedTopY;
                            }
                        }

                        const textOpts = {
                            x, 
                            y: finalY, 
                            w, 
                            h: finalH,
                            align,
                            valign,
                            // Consistent 4pt inset padding on left only for clean text margin
                            margin: [0, 0, 0, (isInsideCard ? 0 : 0)]
                        };

                        // Line spacing: tighter for big font (acronym letters), normal for body
                        if (maxFontSize >= 50) {
                            textOpts.lineSpacingMultiple = 1.0; // Compact for huge acronym characters
                        } else if (maxFontSize >= 16) {
                            textOpts.lineSpacingMultiple = 1.15; // Slightly airy for titles
                        } else {
                            textOpts.lineSpacingMultiple = 1.30; // Comfortable reading for body text
                        }

                        pptSlide.addText(textRuns, textOpts);
                    }
                }
            }
        });

        console.log(`  Slide ${sNum}/${slideEntries.length} processed.`);
    }

    const outPath1 = path.join(OUT_DIR_1, outFileName);
    await pres.writeFile({ fileName: outPath1 });

    // Sync to all matching PPTX file locations across Presentation Automation
    function syncToAllLocations(targetNames, sourcePath) {
        const lowerTargets = targetNames.map(n => n.toLowerCase());
        function searchAndCopy(dir) {
            try {
                const items = fs.readdirSync(dir, { withFileTypes: true });
                for (const item of items) {
                    const fullPath = path.join(dir, item.name);
                    if (item.isDirectory()) {
                        searchAndCopy(fullPath);
                    } else if (item.isFile() && lowerTargets.includes(item.name.toLowerCase())) {
                        if (fullPath !== sourcePath) {
                            try {
                                fs.copyFileSync(sourcePath, fullPath);
                                console.log(`  Synced update to: ${fullPath}`);
                            } catch (e) {
                                console.error(`  Failed to sync to ${fullPath}:`, e.message);
                            }
                        }
                    }
                }
            } catch (e) {}
        }
        searchAndCopy("C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation");
    }

    const baseName = outFileName.replace('_BDC_Styled.pptx', '');
    const syncNames = [
        outFileName,
        refFileName,
        baseName + '.pptx',
        baseName + ' (1).pptx',
        baseName.replace(' (1)', '') + '.pptx'
    ];
    syncToAllLocations(syncNames, outPath1);
}

async function main() {
    for (const refFileName of Object.keys(fileMap)) {
        const outFileName = fileMap[refFileName];
        await processDeck(refFileName, outFileName);
    }
    console.log(`\n🎉 All 11 presentation decks built with standardized taller 3, 4, 5 column card dimensions (0.65" header pills)!`);
}

main().catch(console.error);
