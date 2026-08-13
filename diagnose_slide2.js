const AdmZip = require('adm-zip');

const pptxPath = "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\new MD\\BombayDC_B2B_POV_Deck-2 (1)_BDC_Styled.pptx";
const zip = new AdmZip(pptxPath);
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

const xml = zip.readAsText(slides[1]); // Slide 2
const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
console.log("=== DIAGNOSING SHAPE ORIGINS IN SLIDE 2 XML ===");
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
    console.log(`Shape #${i}: origX=${x}, origY=${y}, w=${w}, h=${h} | txt="${txt.substring(0,40)}..."`);
});
