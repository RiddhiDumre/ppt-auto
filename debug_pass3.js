const AdmZip = require('adm-zip');

const pptxPath = "C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images\\09_B2B_POV_Deck.pptx";
const zip = new AdmZip(pptxPath);
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

const xml = zip.readAsText(slides[1]); // Slide 2
const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
let minContentOrigY = Infinity;
spM.forEach((m, i) => {
    const s = m[1];
    const off = s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
    const ext = s.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
    if (!off || !ext) return;
    const x = parseFloat((parseInt(off[1])/914400).toFixed(3));
    const y = parseFloat((parseInt(off[2])/914400).toFixed(3));
    const w = parseFloat((parseInt(ext[1])/914400).toFixed(3));
    const txt = [...s.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t=>t[1]).join(' ').trim();
    if (y <= 1.70 || x < 1.0) return;
    console.log(`Pre-pass shape #${i}: y=${y}, x=${x}, w=${w} | "${txt.substring(0,30)}"`);
    minContentOrigY = Math.min(minContentOrigY, y);
});
console.log(`minContentOrigY = ${minContentOrigY}`);
const rawCompress = minContentOrigY - 1.80;
const gapCompressShift = (rawCompress > 0.15) ? parseFloat((rawCompress - 0.05).toFixed(3)) : 0;
console.log(`rawCompress = ${rawCompress}, gapCompressShift = ${gapCompressShift}`);
