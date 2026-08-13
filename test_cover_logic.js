const AdmZip = require('adm-zip');
const fs = require('fs');

const zip = new AdmZip("C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images\\09_B2B_POV_Deck.pptx");
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

const xml = zip.readAsText(slides[0]);

function decodeXmlEntities(str) {
    if (!str) return "";
    return str.replace(/&apos;/g, "'").replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&gt;/g, ">").trim();
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
        shapes.push({ tag, xml: content.substring(start, end + closeTag.length) });
        pos = end + closeTag.length;
    }
    return shapes;
}

const shapes = getTopLevelShapes(xml);
shapes.sort((a, b) => {
    const offA = a.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
    const offB = b.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
    return (offA ? parseInt(offA[2]) : 0) - (offB ? parseInt(offB[2]) : 0);
});

let coverTextShapeCount = 0;
let coverTitleBottomY = 0;

shapes.forEach((s, idx) => {
    const spXml = s.xml;
    if (!spXml.includes('<a:t>')) return;
    const off = spXml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
    const ext = spXml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
    if (!off || !ext) return;
    const origY = parseFloat((parseInt(off[2]) / 914400).toFixed(3));
    const rawTxt = [...spXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ');
    const cleanTxtForCheck = decodeXmlEntities(rawTxt);

    let finalY = origY;
    let finalH = 0.5;

    if (cleanTxtForCheck.includes("CONFIDENTIAL") || cleanTxtForCheck.includes("PROPRIETARY")) {
        finalY = 4.96;
    } else if (cleanTxtForCheck.includes("Created By")) {
        finalY = 3.90;
    } else {
        coverTextShapeCount++;
        if (coverTextShapeCount === 1) {
            finalY = (origY > 0 && origY < 2.5) ? origY : 1.95;
            const titleLines = Math.max(1, Math.ceil(cleanTxtForCheck.length / 26));
            const realTitleH = parseFloat((titleLines * 0.55).toFixed(2));
            finalH = realTitleH;
            coverTitleBottomY = finalY + realTitleH;
            console.log(`[TITLE] text="${cleanTxtForCheck}" finalY=${finalY} realTitleH=${realTitleH} coverTitleBottomY=${coverTitleBottomY}`);
        } else if (coverTextShapeCount === 2) {
            finalY = parseFloat((coverTitleBottomY + 0.15).toFixed(2));
            console.log(`[SUBLINE] text="${cleanTxtForCheck}" finalY=${finalY} (coverTitleBottomY was ${coverTitleBottomY})`);
        }
    }
});
