const AdmZip = require('adm-zip');
const fs = require('fs');

const pptxPath = "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\new MD\\V5_BDC_Styled.pptx";
const zip = new AdmZip(pptxPath);
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

slides.forEach((s, sIdx) => {
    const xml = zip.readAsText(s);
    if (xml.includes('BEAM') || xml.includes('Behaviour') || xml.includes('Behavior')) {
        console.log(`\n=== FOUND BEAM AT SLIDE ${sIdx+1} ===`);
        const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
        spM.forEach((m, i) => {
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
            const txt = [...sp.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t=>t[1]).join(' | ').substring(0,60);
            console.log(`  [#${i}] y=${y} x=${x} w=${w} h=${h} clr=${clr} | "${txt}"`);
        });
    }
});
