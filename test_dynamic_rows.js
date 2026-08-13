const AdmZip = require('adm-zip');

const pptxPath = "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\new MD\\BombayDC_B2B_POV_Deck-2 (1)_BDC_Styled.pptx";
const zip = new AdmZip(pptxPath);
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

const xml = zip.readAsText(slides[7]); // Slide 8 (FROM IDEA TO VALIDATED PRODUCT...)

// Test dynamic row height algorithm
console.log("=== TESTING DYNAMIC ROW HEIGHT ALGORITHM ON SLIDE 8 ===");

// 1. Extract shapes
const shapes = [];
const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
spM.forEach((m, i) => {
    const s = m[1];
    const off = s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
    const ext = s.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
    if (!off || !ext) return;
    const x = parseFloat((parseInt(off[1])/914400).toFixed(3));
    const y = parseFloat((parseInt(off[2])/914400).toFixed(3));
    const w = parseFloat((parseInt(ext[1])/914400).toFixed(3));
    const h = parseFloat((parseInt(ext[2])/914400).toFixed(3));
    const txt = [...s.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t=>t[1]).join(' ').trim();
    shapes.push({ i, x, y, w, h, txt });
});

// Filter content shapes in row list (y > 1.80)
const rowShapes = shapes.filter(s => s.y >= 1.80 && s.txt.length > 0 && !s.txt.includes('CONFIDENTIAL'));

// Group into rows by y proximity (within 0.25in)
const rowClusters = [];
rowShapes.forEach(s => {
    if (s.y >= 4.5) return; // bottom section
    let cluster = rowClusters.find(c => Math.abs(c.origY - s.y) < 0.30);
    if (!cluster) {
        cluster = { origY: s.y, shapes: [] };
        rowClusters.push(cluster);
    }
    cluster.shapes.push(s);
});

console.log(`Found ${rowClusters.length} row clusters:`);
let maxContentH = 0.40;
rowClusters.forEach((c, idx) => {
    console.log(`  Row ${idx+1} (origY=${c.origY}): ${c.shapes.map(s=>`"${s.txt.substring(0,25)}..."`).join(' | ')}`);
    c.shapes.forEach(s => {
        if (s.w > 3.0) {
            // content shape
            const lines = Math.ceil(s.txt.length / 55);
            const estH = lines * 0.22;
            if (estH > maxContentH) maxContentH = estH;
        }
    });
});

const standardRowH = Math.max(0.65, parseFloat((maxContentH + 0.16).toFixed(2)));
console.log(`Calculated maxContentH = ${maxContentH} in -> standardRowH = ${standardRowH} in`);

let curY = 1.95;
rowClusters.forEach((c, idx) => {
    const lineY = parseFloat((curY + standardRowH).toFixed(2));
    console.log(`  Row ${idx+1}: contentY = ${curY.toFixed(2)} in, lineY = ${lineY.toFixed(2)} in`);
    curY = parseFloat((lineY + 0.08).toFixed(2));
});
console.log(`Bottom section startY = ${curY.toFixed(2)} in`);
