const AdmZip = require('adm-zip');
const fs = require('fs');

const bdcPath = "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\BDC Deck (Copy).pptx";
const zip = new AdmZip(bdcPath);
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

slides.forEach((s, sIdx) => {
    const xml = zip.readAsText(s);
    const bgCol = xml.match(/<p:bg>[\s\S]*?<a:srgbClr val="([^"]+)"/);
    console.log(`\n=== BDC DECK COPY - SLIDE ${sIdx+1} (bg=${bgCol ? bgCol[1] : 'none'}) ===`);
    const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
    spM.forEach(m => {
        const sp = m[1];
        const off = sp.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
        const ext = sp.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
        if (!off || !ext) return;
        const x = (parseInt(off[1])/914400).toFixed(3);
        const y = (parseInt(off[2])/914400).toFixed(3);
        const w = (parseInt(ext[1])/914400).toFixed(3);
        const h = (parseInt(ext[2])/914400).toFixed(3);
        const clrM = sp.match(/val="([A-Fa-f0-9]{6})"/);
        const clr = clrM ? clrM[1] : 'none';
        const txt = [...sp.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t=>t[1]).join(' ').substring(0,50);
        if (txt.trim()) {
            console.log(`  y=${y} x=${x} w=${w} h=${h} clr=${clr} | "${txt}"`);
        }
    });
});
