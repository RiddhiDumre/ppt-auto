const AdmZip = require('adm-zip');
const fs = require('fs');

const pptxPath = "C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images\\09_B2B_POV_Deck.pptx";
const zip = new AdmZip(pptxPath);
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

const xml = zip.readAsText(slides[1]); // Slide 2
const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
const shapes = [];
spM.forEach(m => {
    const s = m[1];
    shapes.push({ xml: s });
});

// Sort shapes
shapes.sort((a, b) => {
    const offA = a.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
    const offB = b.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
    const yA = offA ? parseInt(offA[2]) : 0;
    const yB = offB ? parseInt(offB[2]) : 0;
    return yA - yB;
});

console.log("=== TRACING SHAPES FOR SLIDE 2 IN SORT ORDER ===");
shapes.forEach((s, idx) => {
    const off = s.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
    const ext = s.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
    if (!off || !ext) return;
    const origX = parseFloat((parseInt(off[1])/914400).toFixed(3));
    const origY = parseFloat((parseInt(off[2])/914400).toFixed(3));
    const w = parseFloat((parseInt(ext[1])/914400).toFixed(3));
    const txt = [...s.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t=>t[1]).join(' ').trim();
    console.log(`[#${idx}] origY=${origY}, origX=${origX}, w=${w} | "${txt.substring(0,40)}"`);
});
