const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

// Load reference XML for Slide 2 of 09_B2B_POV_Deck.pptx
const AdmZip = require("adm-zip");
const zip = new AdmZip("C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images\\09_B2B_POV_Deck.pptx");
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

const sXml = zip.readAsText(slides[1]); // Slide 2
const spM = [...sXml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
const shapes = [];
spM.forEach(m => shapes.push({ xml: m[1] }));

// PRE-PASS 3 simulation
let minContentOrigY = Infinity;
shapes.forEach(shapeObj => {
    const off2 = shapeObj.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/); 
    const ext2 = shapeObj.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/); 
    if (!off2 || !ext2 || !shapeObj.xml.includes('<a:t>')) return;
    const p3OrigY = parseFloat((parseInt(off2[2]) / 914400).toFixed(3));
    const p3OrigX = parseFloat((parseInt(off2[1]) / 914400).toFixed(3));

    if (p3OrigY <= 1.75 || p3OrigX < 1.0) return;
    const rawTxt3 = [...shapeObj.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').trim();
    if (!rawTxt3) return;
    console.log(`[PRE-PASS 3] Candidate content shape: p3OrigY=${p3OrigY}, p3OrigX=${p3OrigX} | "${rawTxt3.substring(0,30)}"`);
    minContentOrigY = Math.min(minContentOrigY, p3OrigY);
});

console.log(`minContentOrigY = ${minContentOrigY}`);
const rawCompress = minContentOrigY - 1.80;
const gapCompressShift = (rawCompress > 0.15) ? parseFloat((rawCompress - 0.05).toFixed(3)) : 0;
console.log(`gapCompressShift = ${gapCompressShift}`);

// Main loop simulation for shape with text "WHAT THIS CREATES"
shapes.forEach((shapeObj, i) => {
    const off = shapeObj.xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/); 
    const ext = shapeObj.xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/); 
    if (!off || !ext) return;
    const origY = parseFloat((parseInt(off[2]) / 914400).toFixed(3));
    const origX = parseFloat((parseInt(off[1]) / 914400).toFixed(3));
    const rawTxt = [...shapeObj.xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').trim();
    if (!rawTxt.includes("WHAT THIS CREATES")) return;

    console.log(`\nFound shape #${i} with "WHAT THIS CREATES": origY=${origY}, origX=${origX}`);
    let y = origY;
    const movedSublineShift = 0.35;
    if (origY >= 1.35) {
        y += movedSublineShift;
        console.log(`  After subline shift: y=${y}`);
        if (gapCompressShift > 0) {
            y -= gapCompressShift;
            console.log(`  After gapCompressShift (-${gapCompressShift}): y=${y}`);
            if (y < 1.93) y = 1.93;
        }
    }
    console.log(`  FINAL COMPUTED Y = ${y}`);
});
