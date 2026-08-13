const AdmZip = require('adm-zip');

const pptxPath = "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\new MD\\BombayDC_B2B_POV_Deck-2 (1)_BDC_Styled.pptx";
const zip = new AdmZip(pptxPath);
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

const xml = zip.readAsText(slides[3]); // Slide 4
console.log("=== EXACT TEXT COLORS IN B2B SLIDE 4 ===");
const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
spM.forEach((m, i) => {
    const s = m[1];
    const srgb = s.match(/<a:srgbClr val="([^"]+)"/g);
    const txt = [...s.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t=>t[1]).join(' | ').substring(0,60);
    if (txt) {
        console.log(`  [#${i}] colors: ${srgb ? srgb.join(', ') : 'none'} | "${txt}"`);
    }
});
