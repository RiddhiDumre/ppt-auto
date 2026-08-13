const AdmZip = require('adm-zip');

const files = [
    "C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\BombayDC_B2B_POV_Deck-2 (1)_BDC_Styled.pptx",
    "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\new MD\\BombayDC_B2B_POV_Deck-2 (1)_BDC_Styled.pptx",
    "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\SALES DECKS\\Enterprise Platforms\\BombayDC_B2B_POV_Deck-2 (1)_BDC_Styled.pptx",
    "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\SALES DECKS\\Enterprise Platforms\\BombayDC_B2B_POV_Deck-2 (1).pptx"
];

files.forEach(pptxPath => {
    try {
        const zip = new AdmZip(pptxPath);
        const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
            .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));
        const xml = zip.readAsText(slides[1]);
        console.log(`\n=== FILE: ${pptxPath} ===`);
        const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
        spM.forEach((m, i) => {
            const s = m[1];
            const off = s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const ext = s.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            if (!off || !ext) return;
            const x = (parseInt(off[1])/914400).toFixed(3);
            const y = (parseInt(off[2])/914400).toFixed(3);
            const txt = [...s.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t=>t[1]).join(' | ').substring(0,60);
            if (txt.includes("WHAT THIS CREATES") || txt.includes("Adoption")) {
                console.log(`  [#${i}] y=${y} x=${x} | "${txt}"`);
            }
        });
    } catch(e) {
        console.log(`Error reading ${pptxPath}: ${e.message}`);
    }
});
