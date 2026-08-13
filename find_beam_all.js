const AdmZip = require('adm-zip');
const fs = require('fs');

const refDir = "C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images";
const files = fs.readdirSync(refDir).filter(f => f.endsWith('.pptx')).sort();

files.forEach(fname => {
    const zip = new AdmZip(`${refDir}\\${fname}`);
    const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
        .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));
    
    slides.forEach((s, sIdx) => {
        const xml = zip.readAsText(s);
        if (xml.includes('BEAM') || xml.includes('Behavior') || xml.includes('Behaviour')) {
            console.log(`\n=== FOUND BEAM AT: ${fname} (Slide ${sIdx+1}) ===`);
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
        }
    });
});
