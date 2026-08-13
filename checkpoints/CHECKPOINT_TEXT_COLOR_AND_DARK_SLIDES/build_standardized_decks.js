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
        let fontSize = szMatch ? parseFloat((parseInt(szMatch[1]) / 100).toFixed(1)) : 8.55;
        if (pCount > 4 && fontSize <= 11) fontSize = 8.0;
        else if (fontSize >= 9.2) fontSize = 9.5;
        else if (fontSize >= 7.8) fontSize = 8.55;
        
        const isBullet = pXml.includes('<a:buChar') || pXml.includes('<a:buAutoNum');
        const rawTxt = [...pXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join('');
        
        if (rawTxt.trim().length === 0) return;
        
        const charsPerLine = Math.max(1, Math.floor(w * (fontSize > 12 ? 8 : (fontSize <= 9.5 ? 15 : 12))));
        const lines = Math.ceil(rawTxt.length / charsPerLine);
        
        const lineHeight = (fontSize / 72) * 1.30;
        const paraSpace = (isBullet ? 2 : 4) / 72;
        
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
                const r1Y = Math.max(...row1Pills.map(p => p.y));
                row1Pills.forEach(p => {
                    cardRowMap[`${p.x.toFixed(2)}_${p.y.toFixed(2)}`] = { cardY: r1Y, textY: r1Y + 0.73 };
                });
            }
            if (row2Pills.length > 0) {
                const r1Y = row1Pills.length > 0 ? Math.max(...row1Pills.map(p => p.y)) : 1.75;
                const r2Y = Math.max(Math.max(...row2Pills.map(p => p.y)), r1Y + 1.45);
                row2Pills.forEach(p => {
                    cardRowMap[`${p.x.toFixed(2)}_${p.y.toFixed(2)}`] = { cardY: r2Y, textY: r2Y + 0.73 };
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

        // RENDER HEADER CONTAINER ABOVE DIVIDER (y: 0.57 in)
        if (hasHeaderTitle && !isClosingSlide) {
            // USER RULE: If title wraps to 2 lines, do NOT display subtitle above divider.
            // Only display subtitle above divider if title is 1 line.
            const displaySublineAbove = !isTitle2Lines ? headerSublineText : "";

            const headerRuns = [
                {
                    text: headerTitleText,
                    options: {
                        color: isDarkThemeSlide ? "ECE9E4" : "1E1E1E",
                        fontSize: 18.8,
                        bold: false,
                        fontFace: "Inter Medium",
                        breakLine: displaySublineAbove ? true : false,
                        paraSpaceAfter: displaySublineAbove ? 4 : 0
                    }
                }
            ];

            if (displaySublineAbove) {
                headerRuns.push({
                    text: displaySublineAbove,
                    options: {
                        color: isDarkThemeSlide ? "B4B4B4" : "555555",
                        fontSize: 9.0, // Style 3: Body Text / Subline
                        bold: false,
                        fontFace: "Inter",
                        paraSpaceBefore: 4
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

        const columnBottoms = {};
        let maxHeaderBottom = 0;

        // PRE-PASS 3: Gap compression + column-header lower-bound detection
        // ─────────────────────────────────────────────────────────────────
        // A) Gap compression: if the first real content shape in the source PPTX
        //    sits more than 0.55 in below where we want content to start
        //    (targetY = 1.70 in + movedSublineShift), we shift ALL content shapes
        //    up by gapCompressShift so content flows tight below the header.
        // B) Column-header lower-bounds: detect short uppercase overlay labels
        //    (e.g. "INTELLIGENT PRODUCTS") that live ABOVE the body-text shapes
        //    in a multi-column layout.  We record their post-shift bottom edge so
        //    we can push the body text below them later in the main loop.
        let gapCompressShift = 0;
        const columnHeaderLowerBounds = {}; // rounded x-key -> post-shift bottom-Y of column header

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

                if (p3OrigY <= 1.40 || p3OrigX < 1.0) return; // skip header/index area
                const rawTxt3 = [...shapeObj.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ');
                const cleanTxt3 = decodeXmlEntities(rawTxt3).trim();
                if (!cleanTxt3 || cleanTxt3.length < 3) return;
                if (cleanTxt3.includes('CONFIDENTIAL') || cleanTxt3.includes('bombaydc.com')) return;
                // Skip sublines already handled in header container
                // Use the SAME filter as the main rendering loop so the same shapes are excluded
                if (p3OrigY >= 0.75 && p3OrigY <= 1.70 && p3W >= 5.0) return;

                // A) Track minimum content Y for gap compression
                minContentOrigY = Math.min(minContentOrigY, p3OrigY);

                // B) Detect column header overlay labels
                // They are: short uppercase text, column-width (1.5–3.5 in), origY > 1.70, h <= 0.55
                if (p3OrigY >= 1.70 && p3OrigY <= 3.50 && p3W >= 1.5 && p3W < 3.5 && p3H <= 0.55) {
                    if (cleanTxt3 === cleanTxt3.toUpperCase() && cleanTxt3.length > 3 && cleanTxt3.length < 55) {
                        const colKey = Math.round(p3OrigX * 2) / 2; // round to nearest 0.5
                        // Post-shift bottom edge of this header label
                        const headerPostShiftBottom = (p3OrigY + movedSublineShift) + p3H + 0.12;
                        if (!columnHeaderLowerBounds[colKey] || headerPostShiftBottom > columnHeaderLowerBounds[colKey]) {
                            columnHeaderLowerBounds[colKey] = parseFloat(headerPostShiftBottom.toFixed(3));
                        }
                    }
                }
            });

            // Compute gap compression shift
            if (minContentOrigY < Infinity) {
                // TARGET: first content shape should land at y = 1.80 + movedSublineShift
                // After the main loop applies +movedSublineShift, origY + movedSublineShift - gapCompressShift = target
                // => gapCompressShift = (minContentOrigY + movedSublineShift) - (1.80 + movedSublineShift)
                //                     = minContentOrigY - 1.80
                const rawCompress = minContentOrigY - 1.80;
                if (rawCompress > 0.55) { // Only compress if gap is > 0.55 in
                    gapCompressShift = parseFloat((rawCompress - 0.05).toFixed(3)); // 0.05 in buffer
                }
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

            // Skip title and subline shapes FIRST because they are rendered as part of unified header container
            const rawTxtForCheck = [...spXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').replace(/\s+/g, ' ').trim();
            const cleanTxtForCheck = decodeXmlEntities(rawTxtForCheck);

            if (!isClosingSlide && spXml.includes('<a:t>')) {
                const szMatch = spXml.match(/sz="(\d+)"/);
                const fontSize = szMatch ? parseFloat((parseInt(szMatch[1]) / 100).toFixed(1)) : 0;
                
                // Skip title shape
                if (origY < 0.85 && fontSize >= 16 && origX >= 1.8) return;
                if (headerTitleText && cleanTxtForCheck.toUpperCase() === headerTitleText) return;
                if (headerTitleText && cleanTxtForCheck.length > 10 && headerTitleText.includes(cleanTxtForCheck.toUpperCase())) return;

                // Skip full-width header sublines / taglines (prevent duplicate rendering & text overlapping!)
                if (origY >= 0.75 && origY <= 1.70 && w >= 5.0 && fontSize < 14 && fontSize > 0) return;
                if (headerSublineText && cleanTxtForCheck === headerSublineText) return;
                if (headerSublineText && cleanTxtForCheck.length > 8 && (headerSublineText.includes(cleanTxtForCheck) || cleanTxtForCheck.includes(headerSublineText.substring(0, 15)))) return;
            }

            // Check for card pill row alignment override
            const key = `${origX.toFixed(2)}_${origY.toFixed(2)}`;
            let isRowMapElement = false;

            if (cardRowMap[key]) {
                y = cardRowMap[key].cardY + movedSublineShift;
                isRowMapElement = true;
            } else {
                // Check if text shape sits right below a card pill in cardRowMap
                Object.keys(cardRowMap).forEach(k => {
                    const [cx, cy] = k.split('_').map(Number);
                    if (Math.abs(cx - origX) < 0.35 && origY >= cy + 0.35 && origY <= cy + 1.0) {
                        y = cardRowMap[k].textY + movedSublineShift;
                        isRowMapElement = true;
                    }
                });
            }

            // USER RULE: Push content below divider down by movedSublineShift on 2-line title slides
            if (!isClosingSlide && origY >= 1.35 && !isRowMapElement) {
                y += movedSublineShift;
                // GAP COMPRESSION: pull content up if source PPTX had it positioned too far down
                if (gapCompressShift > 0) {
                    y -= gapCompressShift;
                    const minAllowed = 1.65 + movedSublineShift;
                    if (y < minAllowed) y = minAllowed;
                }
            }

            // COLUMN BODY TEXT PUSH-DOWN: if this is a column-width body text shape that
            // starts at the very top of content area (origY < 1.75) AND there are
            // overlay column header labels registered for this column, push the body
            // text below the header label so they don't overlap.
            if (!isClosingSlide && !isRowMapElement && origY >= 1.35 && origY < 1.75 && w >= 1.0 && w < 5.0) {
                const colKey = Math.round(origX * 2) / 2;
                if (columnHeaderLowerBounds[colKey] && y < columnHeaderLowerBounds[colKey]) {
                    const pushDown = columnHeaderLowerBounds[colKey] - y;
                    y = parseFloat(columnHeaderLowerBounds[colKey].toFixed(3));
                    h = parseFloat(Math.max(0.30, h - pushDown).toFixed(3));
                }
            }

            // Anti-overlap: Full-Width Header / Subtitle Y Push-down (applies to ALL shapes including rowMap)
            if (!isClosingSlide && y > 1.4) {
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
                if (w >= 5.0) {
                    maxHeaderBottom = Math.max(typeof maxHeaderBottom !== 'undefined' ? maxHeaderBottom : 0, y + effectiveH);
                } else {
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
                const lnColor = lnColorMatch ? lnColorMatch[1] : "B4B4B4";
                
                // Keep native width if it's the closing slide
                const wMatch = spPrXml.match(/<a:ln w="(\d+)"/);
                const nativeW = wMatch ? parseFloat((parseInt(wMatch[1]) / 12700).toFixed(2)) : 0.5;
                const finalWidth = isClosingSlide ? nativeW : (isVertical ? 0.30 : 0.50);

                let lineY = y;
                // STRICT RULE: Horizontal header divider line MUST ALWAYS be placed at y: 1.35 in!
                if (!isClosingSlide && !isVertical && origY >= 0.90 && origY <= 1.50) {
                    lineY = FIXED_DIVIDER_Y; // Always y: 1.35 in
                }

                pptSlide.addShape(pres.shapes.LINE, {
                    x, y: lineY, 
                    w: isVertical ? 0 : w, 
                    h: isVertical ? h : 0,
                    line: { color: lnColor, width: finalWidth, transparency: isClosingSlide ? 0 : 40 }
                });
                return;
            }

            // STANDARDIZE CARD PILL & CONTAINER DIMENSIONS ONLY FOR SHAPES THAT ORIGINALLY HAD CARDS/FILLS
            const origHasFill = shapeBgFill && shapeBgFill !== "none";
            const isCardPillShape = (origHasFill && ["034E48", "1C1C1E", "333333", "222222", "1E1E1E", "4DB89A", "2A2A2C"].includes(shapeBgFill.toUpperCase())) ||
                                    allCardPills.some(p => Math.abs(p.x - origX) < 0.35 && Math.abs(p.y - origY) < 0.35);

            if (isCardPillShape && !isClosingSlide) {
                // USER RULE: Grey/Charcoal box on dark background slides, Green box on light background slides!
                if (isDarkThemeSlide) {
                    shapeBgFill = "1C1C1E"; // Sleek Charcoal / Grey box on dark background
                } else {
                    shapeBgFill = "034E48"; // Signature BombayDC Brand Green box on light background
                }
                if (h < 0.65) h = 0.65;
            }

            const isCardPill = isCardPillShape;

            // Draw shape background rectangle ONLY ONCE per card position (PREVENT DOUBLE BOXES!)
            if (shapeBgFill && shapeBgFill !== "none" && isCardPillShape) {
                const boxPosKey = `${Math.round(x)}_${Math.round(y * 1.5)}`;
                if (!renderedCardBoxes.has(boxPosKey)) {
                    renderedCardBoxes.add(boxPosKey);

                    // Intercept vertical dividers drawn as thin rectangles
                    if (!isClosingSlide && !hasTxBody && w > 0 && w <= 0.08 && h > 0.5) {
                        w = 0.025; // Make thick rectangle dividers much thinner
                    }
                    pptSlide.addShape(pres.shapes.RECTANGLE, {
                        x, y, w, h: 0.65,
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
                            // Ensure all dark text switches to high-contrast off-white on dark backgrounds
                            if (!color || ["1D1D1F", "1E1E1E", "000000", "444444", "555555", "333333", "666666", "034E48"].includes(color.toUpperCase())) {
                                color = "ECE9E4";
                            }
                        }

                        const rawTxt = [...pXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join('');
                        const cleanTxt = decodeXmlEntities(rawTxt);
                        const isHugeTitle = cleanTxt.includes("LET'S BUILD") || cleanTxt.includes("WHAT'S NEXT");

                        // STRICT 3-STYLE TYPOGRAPHY SYSTEM (Apple-inspired minimal design)
                        // Style 1: Slide Title (18.8pt Inter Medium)
                        // Style 2: Section / Content Heading (11.0pt Inter Medium Bold)
                        // Style 3: Body Text / Paragraphs / Bullets / Sublines (9.0pt Inter Regular)
                        if (isHugeTitle) {
                            fontSize = 36.0;
                            isBold = false;
                        } else if (cleanTxt && cleanTxt.trim().length <= 2 && (fontSize >= 28 || fontSize >= 50)) {
                            fontSize = 65.0; // BEAM / Acronym big numbers
                            isBold = true;
                        } else if (sIdx === 0 && fontSize >= 24) {
                            fontSize = 28.0; // Cover Title (Style 1 variant)
                        } else if (!isClosingSlide) {
                            if (fontSize >= 16) {
                                fontSize = 18.8; // Style 1: Slide Title
                            } else if (fontSize >= 10.0 || (cleanTxt === cleanTxt.toUpperCase() && cleanTxt.length < 50 && !isBullet && !cleanTxt.includes('.'))) {
                                fontSize = 11.0; // Style 2: Section / Content Heading
                                isBold = true;
                            } else {
                                fontSize = 9.0;  // Style 3: Body Text / Descriptions / Bullets
                            }
                        }

                        if (fontSize > maxFontSize) maxFontSize = fontSize;

                        if (cleanTxt) {
                            textRuns.push({
                                text: (isBullet ? "• " : "") + cleanTxt,
                                options: {
                                    color,
                                    fontSize,
                                    bold: isHugeTitle ? false : isBold,
                                    fontFace: isHugeTitle ? "Inter" : (isClosingSlide ? (isBold ? "Inter Bold" : "Inter") : (isBold ? "Inter Medium" : "Inter")),
                                    breakLine: pIdx < pMatches.length - 1,
                                    paraSpaceAfter: isBullet ? 2 : (pMatches.length > 4 ? 3 : 4),
                                    paraSpaceBefore: (!isBullet && pIdx > 0 && cleanTxt.length > 5) ? (pMatches.length > 4 ? 3 : 4) : 0
                                }
                            });
                        }
                    });

                    if (textRuns.length > 0) {
                        // CENTERED ALIGNMENT IN CARDS & BOXES
                        if (isInsideCard) {
                            align = "center";
                        }

                        // VERTICAL ALIGNMENT IN CARDS & ROWS
                        let valign = "top";
                        if (isInsideCard) {
                            valign = "middle";
                        } else if (y > 1.5 && h <= 1.0 && maxFontSize < 30) {
valign = "middle"; // Center vertically in row cells
                        } else if (maxFontSize >= 30) {
                            valign = "middle"; // Center big numbers vertically
                        }

                        let finalY = isInsideCard ? (cardRowMap[key] ? cardRowMap[key].cardY + movedSublineShift : y) : y;
                        let finalH = isInsideCard ? 0.65 : parseFloat(Math.max(0.25, Math.min(h, calculateTextShapeHeight(spXml, w) + 0.05)).toFixed(2));

                        // CONTENT SAFETY GUARD: Clamp all text & card shapes so they NEVER spill past y: 5.10 in
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
                            margin: [0, 0, 0, 0]
                        };

                        // Use PowerPoint default line height for slide titles (fontSize >= 16)
                        if (maxFontSize < 16) {
                            textOpts.lineSpacing = "100%";
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
    function syncToAllLocations(targetName, sourcePath) {
        function searchAndCopy(dir) {
            try {
                const items = fs.readdirSync(dir, { withFileTypes: true });
                for (const item of items) {
                    const fullPath = path.join(dir, item.name);
                    if (item.isDirectory()) {
                        searchAndCopy(fullPath);
                    } else if (item.isFile() && item.name.toLowerCase() === targetName.toLowerCase()) {
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

    syncToAllLocations(outFileName, outPath1);
}

async function main() {
    for (const refFileName of Object.keys(fileMap)) {
        const outFileName = fileMap[refFileName];
        await processDeck(refFileName, outFileName);
    }
    console.log(`\n🎉 All 11 presentation decks built with standardized taller 3, 4, 5 column card dimensions (0.65" header pills)!`);
}

main().catch(console.error);
