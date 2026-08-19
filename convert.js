/**
 * BombayDC Style Converter — convert.js
 *
 * Drop any PPTX, DOCX, or PDF into ../input/
 * Run: node convert.js
 * Results appear in: ../output/
 */

'use strict';

const fs   = require('fs');
const path = require('path');

const ENGINE_DIR  = __dirname;
const ROOT        = path.join(ENGINE_DIR, '..');
const INPUT_DIR   = path.join(ROOT, 'input');
const OUTPUT_DIR  = path.join(ROOT, 'output');
const ASSETS_DIR  = path.join(ROOT, 'assets');
const BG_DIR      = path.join(ASSETS_DIR, 'cover_backgrounds');
const CLOSING_DIR = path.join(ASSETS_DIR, 'closing_media');
const CLOSING_REF = path.join(ASSETS_DIR, 'closing_slide.pptx');

[OUTPUT_DIR].forEach(d => { if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true }); });

// ─── Load modules ────────────────────────────────────────────────────────────
function req(name) {
    try { return require(name); } catch(e) {
        try { return require(path.join(ENGINE_DIR, 'node_modules', name)); } catch(e2) {
            throw new Error(`Missing: '${name}'. Run: npm install (inside engine/)`);
        }
    }
}
const pptxgen = req('pptxgenjs');
const AdmZip  = req('adm-zip');

// ─── BDC Design Constants ─────────────────────────────────────────────────────
const BDC = {
    FONT_TITLE:     'Inter Medium',
    FONT_BODY:      'Inter',
    FONT_BOLD:      'Inter Bold',
    ACCENT:         '034E48',
    ACCENT_LIGHT:   '4DB89A',
    DARK_BG:        '121212',
    LIGHT_BG:       'ECE9E4',
    WHITE:          'FFFFFF',
    DARK_TEXT:      '1A1A1A',
    BODY_DARK:      'B0B0B4',
    BODY_LIGHT:     '2C2C2E',
    SUB_LIGHT:      '5A5A5E',
    DIVIDER_DARK:   '444448',
    DIVIDER_LIGHT:  'C0C0C4',
    HEADER_Y:       1.35,
    CONTENT_Y:      1.60,
    MAX_Y:          5.08,
    SLIDE_W:        10.0,
    SLIDE_H:        5.625,
};

// ─── Dark slide detection (every 3rd content slide alternates dark) ───────────
function isDark(slideIdx) {
    if (slideIdx === 0) return true;
    return (slideIdx % 3) === 0;
}

// ─── Random cover background ─────────────────────────────────────────────────
function randomCoverBg() {
    if (!fs.existsSync(BG_DIR)) return null;
    const files = fs.readdirSync(BG_DIR).filter(f => /\.(jpg|jpeg|png)$/i.test(f));
    if (!files.length) return null;
    return path.join(BG_DIR, files[Math.floor(Math.random() * files.length)]);
}

// ─── XML helpers ─────────────────────────────────────────────────────────────
function decodeXml(s) {
    if (!s) return '';
    return s.replace(/&apos;/g,"'").replace(/&amp;/g,'&').replace(/&quot;/g,'"')
            .replace(/&lt;/g,'<').replace(/&gt;/g,'>').trim();
}

// ─── Layout auto-detection ────────────────────────────────────────────────────
function detectLayout(slide) {
    const { items } = slide;
    if (slide.isCover) return 'COVER';
    if (slide.isClosing) return 'CLOSING';

    const cols = items.filter(i => i.type === 'column');
    if (cols.length === 3) return 'THREE_COLUMN';
    if (cols.length === 2) return 'TWO_COLUMN';

    const cards = items.filter(i => i.type === 'card');
    if (cards.length >= 4 && cards.length <= 6) return 'CARD_GRID';

    const rows = items.filter(i => i.type === 'row');
    if (rows.length >= 3) return 'TIMELINE_LIST';

    const acronymItems = items.filter(i => i.type === 'body' && i.key && i.key.length <= 2);
    if (acronymItems.length === 4) return 'ACRONYM_GRID';

    return 'NARRATIVE';
}

// ─────────────────────────────────────────────────────────────────────────────
// PPTX EXTRACTOR
// ─────────────────────────────────────────────────────────────────────────────
function extractPptx(filePath) {
    const zip    = new AdmZip(filePath);
    const entries = zip.getEntries();

    const slideEntries = entries
        .filter(e => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
        .sort((a,b) => {
            const na = parseInt(a.entryName.match(/(\d+)/)[1]);
            const nb = parseInt(b.entryName.match(/(\d+)/)[1]);
            return na - nb;
        });

    const slides = [];

    slideEntries.forEach((entry, sIdx) => {
        const xml  = zip.readAsText(entry);
        const relPath = `ppt/slides/_rels/${path.basename(entry.entryName)}.rels`;
        const relEntry = entries.find(e => e.entryName === relPath);
        const relXml = relEntry ? zip.readAsText(relEntry) : '';

        const mediaRelMap = {};
        if (relXml) {
            [...relXml.matchAll(/<Relationship[^>]+Id="([^"]+)"[^>]+Target="\.\.\/media\/([^"]+)"/g)]
                .forEach(m => { mediaRelMap[m[1]] = m[2]; });
        }

        const shapeBlocks = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)].map(m => m[1]);
        const imgBlocks   = [...xml.matchAll(/<p:pic>([\s\S]*?)<\/p:pic>/g)].map(m => m[1]);

        let title = '';
        let subtitle = '';

        const shapesData = shapeBlocks.map(sh => {
            const offM = sh.match(/<a:off x="(\d+)" y="(\d+)"/);
            const extM = sh.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
            if (!offM || !extM) return null;
            const x = parseFloat((parseInt(offM[1])/914400).toFixed(3));
            const y = parseFloat((parseInt(offM[2])/914400).toFixed(3));
            const w = parseFloat((parseInt(extM[1])/914400).toFixed(3));
            const h = parseFloat((parseInt(extM[2])/914400).toFixed(3));
            const texts = [...sh.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => decodeXml(m[1]));
            const fullText = texts.join(' ').replace(/\s+/g,' ').trim();
            const szM = sh.match(/<a:sz>(\d+)<\/a:sz>/);
            const fontSize = szM ? parseInt(szM[1])/100 : 0;
            const fillM = sh.match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/);
            const fill  = fillM ? fillM[1].toUpperCase() : null;
            const isBold = sh.includes('<a:b/>') || sh.includes('<a:b val="1"');

            return { x, y, w, h, fullText, fontSize, fill, isBold, texts };
        }).filter(Boolean).filter(s => s.fullText.length > 0);

        const titleShape = shapesData
            .filter(s => s.y < 1.30 && s.x > 1.5)
            .sort((a,b) => b.fontSize - a.fontSize)[0];
        if (titleShape) title = titleShape.fullText;

        const subShape = shapesData.find(s => s.y >= 0.75 && s.y < 1.45 && s.x > 1.5 && s.fullText !== title);
        if (subShape) subtitle = subShape.fullText;

        const contentShapes = shapesData.filter(s => s.y >= 1.4 && s.fullText.length > 0);

        const items = [];

        const byX = {};
        contentShapes.forEach(s => {
            const bucket = Math.round(s.x * 2) / 2;
            if (!byX[bucket]) byX[bucket] = [];
            byX[bucket].push(s);
        });
        const xBuckets = Object.keys(byX).map(Number).sort((a,b)=>a-b);

        if (xBuckets.length >= 2 && xBuckets.length <= 3) {
            xBuckets.forEach(xk => {
                const colShapes = byX[xk].sort((a,b)=>a.y-b.y);
                const header = colShapes.find(s => s.isBold || s.fill || s.fontSize >= 9)?.fullText || '';
                const bullets = colShapes.filter(s => s.fullText !== header).map(s => s.fullText);
                items.push({ type: 'column', header, bullets });
            });
        } else {
            contentShapes.forEach(s => {
                const isCard = s.fill && ['034E48','4DB89A','1C1C1E','224B12','1A3632',
                    '0A3B36','08322D','0D524A','004B44','024E48','034D47','333333'].includes(s.fill);
                if (isCard) {
                    items.push({ type: 'card', header: s.fullText, desc: '' });
                } else {
                    const prevCard = items.length && items[items.length-1].type === 'card' && !items[items.length-1].desc;
                    if (prevCard) {
                        items[items.length-1].desc = s.fullText;
                    } else {
                        items.push({ type: 'body', text: s.fullText });
                    }
                }
            });
        }

        const footnoteShape = contentShapes.find(s => s.y > 4.0 && s.w > 5.0);
        const footnote = footnoteShape ? footnoteShape.fullText : '';

        const images = [];
        imgBlocks.forEach(img => {
            const rEmbedM = img.match(/r:embed="([^"]+)"/);
            if (!rEmbedM || !mediaRelMap[rEmbedM[1]]) return;
            const mediaFile = mediaRelMap[rEmbedM[1]];
            const mediaEntry = entries.find(e => e.entryName === `ppt/media/${mediaFile}`);
            if (!mediaEntry) return;
            const tmpPath = path.join(OUTPUT_DIR, `_tmp_${path.basename(filePath,'pptx')}_${mediaFile}`);
            fs.writeFileSync(tmpPath, mediaEntry.getData());
            images.push(tmpPath);
        });

        slides.push({ index: sIdx, title, subtitle, items, footnote, images });
    });

    return slides;
}

// ─────────────────────────────────────────────────────────────────────────────
// DOCX EXTRACTOR
// ─────────────────────────────────────────────────────────────────────────────
async function extractDocx(filePath) {
    let mammoth;
    try { mammoth = req('mammoth'); } catch(e) {
        console.warn('  mammoth not found, skipping DOCX. Run: npm install');
        return [];
    }

    const result = await mammoth.extractRawText({ path: filePath });
    const text = result.value;
    const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

    const slides = [];
    let current = null;

    lines.forEach(line => {
        const isHeading = (line === line.toUpperCase() && line.length <= 60 && line.length >= 3)
                       || line.startsWith('# ');
        if (isHeading) {
            if (current) slides.push(current);
            current = {
                title: line.replace(/^#+\s*/, '').toUpperCase(),
                subtitle: '',
                items: [],
                footnote: '',
                images: []
            };
        } else if (current) {
            const isBullet = /^[•\-\*]\s/.test(line) || /^\d+\.\s/.test(line);
            const cleanLine = line.replace(/^[•\-\*\d\.]\s+/, '').trim();
            if (!current.subtitle && !isBullet && current.items.length === 0) {
                current.subtitle = cleanLine;
            } else {
                const lastItem = current.items[current.items.length - 1];
                if (lastItem && lastItem.type === 'body') {
                    lastItem.bullets = lastItem.bullets || [];
                    lastItem.bullets.push(cleanLine);
                } else {
                    current.items.push({ type: 'body', text: cleanLine, bullets: [] });
                }
            }
        } else {
            if (!current) {
                current = { title: path.basename(filePath, path.extname(filePath)).toUpperCase(), subtitle: line, items: [], footnote: '', images: [] };
            }
        }
    });
    if (current) slides.push(current);

    return slides.map((s, i) => ({ index: i, ...s }));
}

// ─────────────────────────────────────────────────────────────────────────────
// PDF EXTRACTOR
// ─────────────────────────────────────────────────────────────────────────────
async function extractPdf(filePath) {
    let pdfParse;
    try { pdfParse = req('pdf-parse'); } catch(e) {
        console.warn('  pdf-parse not found, skipping PDF. Run: npm install');
        return [];
    }

    const buffer = fs.readFileSync(filePath);
    const data   = await pdfParse(buffer);

    const pages = data.text.split(/\f|\n{3,}/).map(p => p.trim()).filter(p => p.length > 10);

    return pages.map((pageText, i) => {
        const lines = pageText.split('\n').map(l => l.trim()).filter(Boolean);
        const title = lines[0] ? lines[0].toUpperCase().substring(0, 80) : `PAGE ${i+1}`;
        const subtitle = lines[1] && lines[1] !== title ? lines[1] : '';
        const bodyLines = lines.slice(subtitle ? 2 : 1);
        const items = bodyLines.length
            ? [{ type: 'body', text: bodyLines[0], bullets: bodyLines.slice(1) }]
            : [];
        return { index: i, title, subtitle, items, footnote: '', images: [] };
    });
}

// ─────────────────────────────────────────────────────────────────────────────
// CLOSING SLIDE — copied from reference deck
// ─────────────────────────────────────────────────────────────────────────────
function addClosingSlide(pres) {
    if (!fs.existsSync(CLOSING_REF)) {
        const s = pres.addSlide();
        s.background = { color: BDC.ACCENT };
        s.addText("LET'S BUILD\nWHAT'S NEXT", {
            x: 1.5, y: 1.8, w: 7.0, h: 2.0,
            fontSize: 36, bold: true, color: BDC.WHITE, fontFace: BDC.FONT_BOLD,
            align: 'left', valign: 'middle'
        });
        return;
    }

    const zip = new AdmZip(CLOSING_REF);
    const slides = zip.getEntries().filter(e => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
        .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));
    const lastSlide = slides[slides.length - 1];
    if (!lastSlide) return;

    const xml = zip.readAsText(lastSlide);
    const pptSlide = pres.addSlide();
    pptSlide.background = { color: BDC.ACCENT };

    [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)].forEach(m => {
        const sh = m[1];
        const offM = sh.match(/<a:off x="(\d+)" y="(\d+)"/);
        const extM = sh.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
        if (!offM || !extM) return;
        const x = parseFloat((parseInt(offM[1])/914400).toFixed(3));
        const y = parseFloat((parseInt(offM[2])/914400).toFixed(3));
        const w = parseFloat((parseInt(extM[1])/914400).toFixed(3));
        const h = parseFloat((parseInt(extM[2])/914400).toFixed(3));
        const texts = [...sh.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m2=>decodeXml(m2[1]));
        const fullText = texts.join(' ').trim();
        if (!fullText) return;
        const szM = sh.match(/<a:sz>(\d+)<\/a:sz>/);
        const fontSize = szM ? parseInt(szM[1])/100 : 10;
        const isBold = sh.includes('<a:b/>') || sh.includes('<a:b val="1"');
        const colorM = sh.match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/);
        const color = colorM ? colorM[1] : BDC.WHITE;

        pptSlide.addText(fullText, {
            x, y, w, h: Math.max(h, 0.2),
            fontSize: Math.min(Math.max(Math.round(fontSize/2)*2, 8), 48),
            bold: isBold,
            color,
            fontFace: isBold ? BDC.FONT_BOLD : BDC.FONT_TITLE,
            valign: 'top'
        });
    });

    if (fs.existsSync(CLOSING_DIR)) {
        const logo = path.join(CLOSING_DIR, 'image-10-1.png');
        const person = path.join(CLOSING_DIR, 'image-10-2.png');
        if (fs.existsSync(logo))   pptSlide.addImage({ path: logo,   x: 0.16, y: 0.10, w: 1.50, h: 0.40 });
        if (fs.existsSync(person)) pptSlide.addImage({ path: person, x: 7.50, y: 2.20, w: 2.10, h: 2.10, rounding: true });
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE RENDERER — applies BDC design to one slide
// ─────────────────────────────────────────────────────────────────────────────
function renderSlide(pres, slideData, slideNum, totalSlides) {
    const dark = isDark(slideData.index);
    const bg   = dark ? BDC.DARK_BG : BDC.LIGHT_BG;
    const titleColor  = dark ? BDC.WHITE : BDC.DARK_TEXT;
    const bodyColor   = dark ? BDC.BODY_DARK : BDC.BODY_LIGHT;
    const subColor    = dark ? 'A0A0A4' : BDC.SUB_LIGHT;
    const accentColor = BDC.ACCENT_LIGHT;
    const lineColor   = dark ? BDC.DIVIDER_DARK : BDC.DIVIDER_LIGHT;
    const cardBg      = BDC.ACCENT;
    const cardText    = BDC.WHITE;

    const layout = detectLayout(slideData);

    // Cover slide
    if (layout === 'COVER') {
        const pptSlide = pres.addSlide();
        const coverBg = randomCoverBg();
        if (coverBg) pptSlide.addImage({ path: coverBg, x: 0, y: 0, w: BDC.SLIDE_W, h: BDC.SLIDE_H, sizing: { type: 'cover' }});
        pptSlide.background = { color: '0A0A0A' };

        pptSlide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: BDC.SLIDE_W, h: BDC.SLIDE_H,
            fill: { color: '000000', transparency: 50 }, line: { color: '000000', transparency: 100 } });

        pptSlide.addText('BOMBAYDC', { x: 0.20, y: 0.14, w: 3.0, h: 0.30,
            fontSize: 10, bold: true, color: BDC.WHITE, fontFace: BDC.FONT_BOLD, charSpacing: 1.5 });
        pptSlide.addText('bombaydc.com', { x: 7.60, y: 0.14, w: 2.20, h: 0.30,
            fontSize: 8, color: 'AAAAAA', fontFace: BDC.FONT_BODY, align: 'right' });
        pptSlide.addShape(pres.shapes.LINE, { x: 0.16, y: 0.50, w: 9.68, h: 0,
            line: { color: '555558', width: 0.5 } });

        pptSlide.addText(slideData.title, {
            x: 1.00, y: 1.80, w: 8.00, h: 1.60,
            fontSize: 36, bold: false, color: BDC.WHITE, fontFace: BDC.FONT_TITLE,
            valign: 'middle', wrap: true
        });
        if (slideData.subtitle) {
            pptSlide.addText(slideData.subtitle, {
                x: 1.00, y: 3.50, w: 7.00, h: 0.60,
                fontSize: 10, color: 'C0C0C4', fontFace: BDC.FONT_BODY
            });
        }
        return;
    }

    const pptSlide = pres.addSlide();
    pptSlide.background = { color: bg };

    pptSlide.addText('BOMBAYDC', { x: 0.20, y: 0.14, w: 3.0, h: 0.30,
        fontSize: 10, bold: true, color: dark ? BDC.WHITE : BDC.DARK_TEXT,
        fontFace: BDC.FONT_BOLD, charSpacing: 1.5 });
    pptSlide.addText('bombaydc.com', { x: 7.60, y: 0.14, w: 2.20, h: 0.30,
        fontSize: 8, color: dark ? '888888' : '888888', fontFace: BDC.FONT_BODY, align: 'right' });
    pptSlide.addShape(pres.shapes.LINE, { x: 0.16, y: 0.50, w: 9.68, h: 0,
        line: { color: dark ? '333336' : 'CCCCCC', width: 0.5 } });

    pptSlide.addText(`${slideNum}`, {
        x: 0.16, y: 0.57, w: 1.68, h: 0.33,
        fontSize: 18, bold: false,
        color: dark ? '5A5A5E' : '9A9A9E',
        fontFace: BDC.FONT_TITLE, valign: 'top'
    });

    const titleIsLong = slideData.title.length > 38;
    pptSlide.addText([
        { text: slideData.title.toUpperCase(), options: {
            color: titleColor, fontSize: 18, bold: false, fontFace: BDC.FONT_TITLE,
            breakLine: titleIsLong || !!slideData.subtitle, charSpacing: 0.3
        }},
        ...(!titleIsLong && slideData.subtitle ? [{ text: slideData.subtitle, options: {
            color: subColor, fontSize: 8, bold: false, fontFace: BDC.FONT_BODY,
            paraSpaceBefore: 2
        }}] : [])
    ], { x: 1.99, y: 0.57, w: 7.85, h: 0.76, valign: 'top', margin: [0,0,0,0] });

    pptSlide.addShape(pres.shapes.LINE, { x: 0.16, y: BDC.HEADER_Y, w: 9.68, h: 0,
        line: { color: lineColor, width: 0.5 } });

    const contentY = BDC.CONTENT_Y + (titleIsLong && slideData.subtitle ? 0.30 : 0);
    if (titleIsLong && slideData.subtitle) {
        pptSlide.addText(slideData.subtitle, {
            x: 1.99, y: BDC.CONTENT_Y, w: 7.85, h: 0.26,
            fontSize: 8, color: subColor, fontFace: BDC.FONT_BODY, valign: 'top', margin: [0,0,0,0]
        });
    }

    if (layout === 'NARRATIVE') {
        const bodyItems = slideData.items.filter(i => i.type === 'body');
        let curY = contentY;
        bodyItems.forEach(item => {
            const runs = [];
            if (item.text) {
                runs.push({ text: item.text, options: { color: bodyColor, fontSize: 8, fontFace: BDC.FONT_BODY, breakLine: true, paraSpaceAfter: 3 }});
            }
            if (item.bullets && item.bullets.length) {
                item.bullets.forEach(b => {
                    runs.push({ text: `• ${b}`, options: { color: bodyColor, fontSize: 8, fontFace: BDC.FONT_BODY, breakLine: true, paraSpaceAfter: 2 }});
                });
            }
            if (!runs.length) return;
            const h = Math.min(0.20 + (item.bullets?.length || 0) * 0.18, BDC.MAX_Y - curY);
            pptSlide.addText(runs, { x: 1.99, y: curY, w: 7.85, h, valign: 'top', margin: [0,0,0,0] });
            curY += h + 0.10;
        });
    }

    else if (layout === 'TWO_COLUMN') {
        const [col1, col2] = slideData.items.filter(i => i.type === 'column');
        const COL_W = 3.75;
        const COL2X = 5.95;
        [
            { col: col1, x: 1.99 },
            { col: col2, x: COL2X }
        ].forEach(({ col, x }) => {
            if (!col) return;
            if (col.header) {
                pptSlide.addText(col.header, {
                    x, y: contentY, w: COL_W, h: 0.22,
                    fontSize: 8, bold: true, color: accentColor,
                    fontFace: BDC.FONT_BOLD, charSpacing: 0.5, valign: 'top', margin:[0,0,0,0]
                });
                pptSlide.addShape(pres.shapes.LINE, { x, y: contentY + 0.26, w: COL_W, h: 0,
                    line: { color: dark ? '38383C' : 'D4D4D8', width: 0.35 }});
            }
            const bodyRuns = col.bullets.map(b => ({
                text: `• ${b}`, options: { color: bodyColor, fontSize: 8, fontFace: BDC.FONT_BODY, breakLine: true, paraSpaceAfter: 2 }
            }));
            if (bodyRuns.length) {
                pptSlide.addText(bodyRuns, { x, y: contentY + 0.32, w: COL_W, h: BDC.MAX_Y - contentY - 0.32,
                    valign: 'top', margin: [0,0,0,0] });
            }
        });
    }

    else if (layout === 'THREE_COLUMN') {
        const cols = slideData.items.filter(i => i.type === 'column').slice(0,3);
        const COL_W = 2.55;
        const startX = 1.99;
        const gap = 0.15;
        cols.forEach((col, ci) => {
            const cx = startX + ci * (COL_W + gap);
            if (col.header) {
                pptSlide.addText(col.header, {
                    x: cx, y: contentY, w: COL_W, h: 0.22,
                    fontSize: 8, bold: true, color: accentColor,
                    fontFace: BDC.FONT_BOLD, charSpacing: 0.4, valign: 'top', margin:[0,0,0,0]
                });
                pptSlide.addShape(pres.shapes.LINE, { x: cx, y: contentY + 0.26, w: COL_W, h: 0,
                    line: { color: dark ? '38383C' : 'D4D4D8', width: 0.35 }});
            }
            const bodyRuns = col.bullets.map(b => ({
                text: `• ${b}`, options: { color: bodyColor, fontSize: 8, fontFace: BDC.FONT_BODY, breakLine: true, paraSpaceAfter: 2 }
            }));
            if (bodyRuns.length) {
                pptSlide.addText(bodyRuns, { x: cx, y: contentY + 0.32, w: COL_W, h: BDC.MAX_Y - contentY - 0.32,
                    valign: 'top', margin:[0,0,0,0] });
            }
        });
    }

    else if (layout === 'CARD_GRID') {
        const cards = slideData.items.filter(i => i.type === 'card');
        const numCols = Math.min(3, cards.length);
        const numRows = Math.ceil(cards.length / numCols);
        const CARD_W  = 2.55;
        const CARD_H  = 0.50;
        const TEXT_OFF = 0.58;
        const COL_GAP = 0.12;
        const ROW_GAP = 0.85;
        const startX  = 1.99;

        cards.forEach((card, ci) => {
            const row = Math.floor(ci / numCols);
            const col = ci % numCols;
            const cx = startX + col * (CARD_W + COL_GAP);
            const cy = contentY + row * (CARD_H + ROW_GAP);

            pptSlide.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: CARD_W, h: CARD_H,
                fill: { color: cardBg }, line: { color: cardBg, width: 0 } });
            pptSlide.addText(card.header, { x: cx, y: cy, w: CARD_W, h: CARD_H,
                fontSize: 8, bold: true, color: cardText, fontFace: BDC.FONT_BOLD,
                align: 'center', valign: 'middle', margin: [0,0,0,0] });

            if (card.desc) {
                pptSlide.addText(card.desc, { x: cx, y: cy + TEXT_OFF, w: CARD_W, h: 0.60,
                    fontSize: 8, color: bodyColor, fontFace: BDC.FONT_BODY,
                    valign: 'top', margin: [0,0,0,0], wrap: true });
            }
        });
    }

    else if (layout === 'TIMELINE_LIST') {
        const rows = slideData.items.filter(i => i.type === 'row' || i.type === 'body');
        const ROW_H  = Math.min(0.50, (BDC.MAX_Y - contentY - 0.20) / Math.max(rows.length, 1));
        rows.forEach((row, ri) => {
            const ry = contentY + ri * (ROW_H + 0.18);
            const label = row.key || row.header || String(ri + 1).padStart(2,'0');
            const text  = row.text || row.desc || (row.bullets && row.bullets[0]) || '';

            pptSlide.addText(label, { x: 1.99, y: ry, w: 0.60, h: ROW_H,
                fontSize: 8, bold: true, color: accentColor, fontFace: BDC.FONT_BOLD,
                valign: 'middle', margin:[0,0,0,0] });
            pptSlide.addText(text, { x: 2.70, y: ry, w: 7.10, h: ROW_H,
                fontSize: 8, color: bodyColor, fontFace: BDC.FONT_BODY,
                valign: 'middle', margin:[0,0,0,0] });

            if (ri < rows.length - 1) {
                pptSlide.addShape(pres.shapes.LINE, { x: 1.99, y: ry + ROW_H + 0.08, w: 7.85, h: 0,
                    line: { color: dark ? '2A2A2E' : 'E0E0E4', width: 0.25 }});
            }
        });
    }

    else if (layout === 'ACRONYM_GRID') {
        const items4 = slideData.items.filter(i => i.type === 'body').slice(0,4);
        const CELL_W = 1.80;
        const CELL_H = 0.90;
        const GAP    = 0.12;
        const startX = 1.99;
        items4.forEach((item, ci) => {
            const cx = startX + ci * (CELL_W + GAP);
            const cy = contentY;
            const letter = item.key || String.fromCharCode(66 + ci);
            pptSlide.addText(letter, { x: cx, y: cy, w: CELL_W, h: CELL_H,
                fontSize: 48, bold: false, color: accentColor, fontFace: BDC.FONT_TITLE,
                align: 'center', valign: 'middle' });
            if (item.text) {
                pptSlide.addText(item.text, { x: cx, y: cy + CELL_H + 0.06, w: CELL_W, h: 1.0,
                    fontSize: 8, color: bodyColor, fontFace: BDC.FONT_BODY,
                    wrap: true, valign: 'top', margin:[0,0,0,0] });
            }
        });
    }

    if (slideData.footnote) {
        pptSlide.addText(slideData.footnote, {
            x: 1.99, y: 4.85, w: 7.85, h: 0.22,
            fontSize: 8, color: subColor, fontFace: BDC.FONT_BODY,
            valign: 'bottom', margin: [0,0,0,0]
        });
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// RENDER FULL DECK
// ─────────────────────────────────────────────────────────────────────────────
function renderDeck(slides, outputPath) {
    const pres = new pptxgen();
    pres.layout = 'LAYOUT_WIDE';
    pres.defineLayout({ name: 'LAYOUT_WIDE', width: 10, height: 5.625 });

    slides.forEach((slide, i) => {
        renderSlide(pres, slide, i + 1, slides.length);
    });

    addClosingSlide(pres);

    return pres.writeFile({ fileName: outputPath });
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIN — scan input/ and convert each file
// ─────────────────────────────────────────────────────────────────────────────
async function main() {
    console.log('\n==================================================');
    console.log('  BombayDC Style Converter');
    console.log('==================================================');

    if (!fs.existsSync(INPUT_DIR)) {
        fs.mkdirSync(INPUT_DIR, { recursive: true });
        console.log(`\n  Created input/ folder. Drop your PPTX, DOCX, or PDF files there and run again.\n`);
        return;
    }

    const inputFiles = fs.readdirSync(INPUT_DIR)
        .filter(f => /\.(pptx|docx|pdf)$/i.test(f))
        .map(f => path.join(INPUT_DIR, f));

    if (!inputFiles.length) {
        console.log(`\n  [!] No PPTX, DOCX, or PDF files found in input/\n`);
        console.log(`  Drop your files into: ${INPUT_DIR}\n`);
        return;
    }

    console.log(`\n  Found ${inputFiles.length} file(s) to convert:\n`);
    inputFiles.forEach(f => console.log(`  * ${path.basename(f)}`));
    console.log('');

    for (const filePath of inputFiles) {
        const ext  = path.extname(filePath).toLowerCase();
        const base = path.basename(filePath, ext);
        const out  = path.join(OUTPUT_DIR, `${base}_BDC_Styled.pptx`);

        console.log(`  >> Converting: ${path.basename(filePath)}`);

        try {
            let slides = [];

            if (ext === '.pptx') {
                slides = extractPptx(filePath);
            } else if (ext === '.docx') {
                slides = await extractDocx(filePath);
            } else if (ext === '.pdf') {
                slides = await extractPdf(filePath);
            }

            if (!slides.length) {
                console.log(`    [x] No content extracted. Skipping.`);
                continue;
            }

            if (slides.length > 1) slides[0].isCover = true;

            await renderDeck(slides, out);
            console.log(`    [v] Done -> ${path.basename(out)}`);

            const tmpFiles = fs.readdirSync(OUTPUT_DIR).filter(f => f.startsWith('_tmp_'));
            tmpFiles.forEach(f => { try { fs.unlinkSync(path.join(OUTPUT_DIR, f)); } catch(e){} });

        } catch (err) {
            console.error(`    [x] Error: ${err.message}`);
        }
    }

    console.log(`\n  [SUCCESS] Output saved to: ${OUTPUT_DIR}`);
    console.log('==================================================\n');
}

main().catch(console.error);
