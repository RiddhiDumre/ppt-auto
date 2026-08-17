const AdmZip = require('C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\node_modules\\adm-zip');

const zip = new AdmZip("C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images\\08_Ecommerce_V5.pptx");
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

[9, 11, 12, 13].forEach(sIdx => {
    const xml = zip.readAsText(slides[sIdx]);
    console.log(`=== ECOMM V5 SLIDE ${sIdx+1} ===`);
    const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
    const xVals = new Set();
    spM.forEach((m, i) => {
        const s = m[1];
        const off = s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
        if (!off) return;
        const x = parseFloat((parseInt(off[1])/914400).toFixed(2));
        const y = parseFloat((parseInt(off[2])/914400).toFixed(2));
        if (y >= 1.5 && x >= 1.5) xVals.add(x.toFixed(1));
    });
    console.log("Distinct X values:", [...xVals].sort());
});
