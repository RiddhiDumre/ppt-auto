const AdmZip = require('adm-zip');

const pptxPath = "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\SALES DECKS\\Enterprise Platforms\\BombayDC_B2B_POV_Deck-2 (1)_BDC_Styled.pptx";
const zip = new AdmZip(pptxPath);
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

slides.forEach((entry, idx) => {
    const xml = zip.readAsText(entry);
    const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
    console.log(`\n=== SLIDE ${idx + 1} (${entry.entryName}) ===`);
    spM.forEach((m, i) => {
        const s = m[1];
        const off = s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
        const ext = s.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
        if (!off || !ext) return;
        const x = (parseInt(off[1])/914400).toFixed(3);
        const y = (parseInt(off[2])/914400).toFixed(3);
        const txt = [...s.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t=>t[1]).join(' ').trim();
        if (txt.length > 0) {
            console.log(`  [#${i}] y=${y} x=${x} | "${txt.substring(0,60)}"`);
        }
    });
});
