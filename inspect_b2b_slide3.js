const AdmZip = require('adm-zip');

const pptxPath = "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\new MD\\BombayDC_B2B_POV_Deck-2 (1)_BDC_Styled.pptx";
const zip = new AdmZip(pptxPath);
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

const xml = zip.readAsText(slides[2]); // Slide 3
console.log("=== ALL SHAPES IN B2B SLIDE 3 ===");
const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
spM.forEach((m, i) => {
    const s = m[1];
    const off = s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
    const ext = s.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
    if (!off || !ext) return;
    const x = (parseInt(off[1])/914400).toFixed(3);
    const y = (parseInt(off[2])/914400).toFixed(3);
    const w = (parseInt(ext[1])/914400).toFixed(3);
    const h = (parseInt(ext[2])/914400).toFixed(3);
    const clrM = s.match(/val="([A-Fa-f0-9]{6})"/);
    const clr = clrM ? clrM[1] : 'none';
    const txt = [...s.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t=>t[1]).join(' | ').substring(0,60);
    console.log(`  [#${i}] y=${y} x=${x} w=${w} h=${h} clr=${clr} | "${txt}"`);
});
