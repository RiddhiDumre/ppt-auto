const AdmZip = require('adm-zip');
const fs = require('fs');

const refDir = "C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images";
const files = fs.readdirSync(refDir).filter(f => f.endsWith('.pptx')).sort();

files.forEach(fname => {
    const zip = new AdmZip(`${refDir}\\${fname}`);
    const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
        .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));
    
    const xml = zip.readAsText(slides[0]);
    console.log(`\n=== COVER SLIDE 1: ${fname} ===`);
    const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
    spM.forEach(m => {
        const s = m[1];
        const off = s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
        const ext = s.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
        if (!off || !ext) return;
        const x = (parseInt(off[1])/914400).toFixed(2);
        const y = (parseInt(off[2])/914400).toFixed(2);
        const w = (parseInt(ext[1])/914400).toFixed(2);
        const h = (parseInt(ext[2])/914400).toFixed(2);
        const txt = [...s.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t=>t[1]).join(' ').substring(0,60);
        if (txt.trim()) {
            console.log(`  y=${y} x=${x} w=${w} h=${h} | "${txt}"`);
        }
    });
});
