const AdmZip = require('adm-zip');
const fs = require('fs');

const refDir = "C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images";
const zip = new AdmZip(`${refDir}\\09_B2B_POV_Deck.pptx`);
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

console.log(`Total slides: ${slides.length}`);

// Check all slides for 1,2,3,4 type content
slides.forEach((s, sIdx) => {
    const xml = zip.readAsText(s);
    // Look for any slide with SITUATION or numbered content
    const texts = [...xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m=>m[1]).join(' ');
    if (texts.includes('SITUATION') || texts.includes('CONSTRAINT') || texts.includes('INTERVENTION')) {
        console.log(`\n=== Found SITUATION/numbered at slide ${sIdx+1} ===`);
        
        const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
        const picM = [...xml.matchAll(/<p:pic>([\s\S]*?)<\/p:pic>/g)];
        
        const all = [...spM.map(m=>({tag:'sp',s:m[1]})), ...picM.map(m=>({tag:'pic',s:m[1]}))];
        all.sort((a,b) => {
            const ay = a.s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const by = b.s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            return (ay?parseInt(ay[2]):0) - (by?parseInt(by[2]):0);
        });
        all.forEach(({tag,s}) => {
            const off = s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const ext = s.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            if (!off||!ext) return;
            const x = (parseInt(off[1])/914400).toFixed(3);
            const y = (parseInt(off[2])/914400).toFixed(3);
            const w = (parseInt(ext[1])/914400).toFixed(3);
            const h = (parseInt(ext[2])/914400).toFixed(3);
            const fills = [...s.matchAll(/val="([A-Fa-f0-9]{6})"/g)];
            const fill = fills.length ? fills[0][1] : null;
            const txt = [...s.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t=>t[1]).join('').substring(0,60);
            console.log(`  [${tag}] y=${y} x=${x} w=${w} h=${h} fill=${fill||'none'} | "${txt}"`);
        });
    }
});
