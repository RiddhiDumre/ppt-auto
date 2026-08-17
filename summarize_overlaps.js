const AdmZip = require('C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\node_modules\\adm-zip');
const fs = require('fs');
const path = require('path');

const outDir = 'C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\new MD';
const files = fs.readdirSync(outDir).filter(f => f.endsWith('_BDC_Styled.pptx')).sort();

const summaryByDeck = {};

files.forEach(fname => {
    const zip = new AdmZip(path.join(outDir, fname));
    const slideEntries = zip.getEntries().filter(e => /ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
        .sort((a,b) => parseInt(a.entryName.match(/\d+/)[0]) - parseInt(b.entryName.match(/\d+/)[0]));

    summaryByDeck[fname] = [];

    slideEntries.forEach((entry, sIdx) => {
        if (sIdx === 0 || sIdx === slideEntries.length - 1) return;
        const xml = zip.readAsText(entry);
        const sps = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)];
        const boxes = [];

        sps.forEach(sp => {
            const off = sp[0].match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const ext = sp[0].match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            const txts = [...sp[0].matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').trim();
            if (!off || !ext || !txts) return;
            const x = parseFloat((parseInt(off[1]) / 914400).toFixed(2));
            const y = parseFloat((parseInt(off[2]) / 914400).toFixed(2));
            const w = parseFloat((parseInt(ext[1]) / 914400).toFixed(2));
            const h = parseFloat((parseInt(ext[2]) / 914400).toFixed(2));
            if (y < 1.35) return;
            boxes.push({ x, y, w, h, txt: txts.substring(0, 45) });
        });

        // Check for overlaps
        for (let i = 0; i < boxes.length; i++) {
            for (let j = i + 1; j < boxes.length; j++) {
                const b1 = boxes[i];
                const b2 = boxes[j];
                const hOverlap = (b1.x < b2.x + b2.w - 0.2) && (b1.x + b1.w > b2.x + 0.2);
                const vOverlap = (b1.y < b2.y + b2.h - 0.08) && (b1.y + b1.h > b2.y + 0.08);
                if (hOverlap && vOverlap) {
                    summaryByDeck[fname].push({
                        slide: sIdx + 1,
                        b1,
                        b2
                    });
                }
            }
        }
    });
});

Object.keys(summaryByDeck).forEach(f => {
    console.log(`\n========================================`);
    console.log(`Deck: ${f} (${summaryByDeck[f].length} overlaps)`);
    console.log(`========================================`);
    const bySlide = {};
    summaryByDeck[f].forEach(o => {
        if (!bySlide[o.slide]) bySlide[o.slide] = [];
        bySlide[o.slide].push(o);
    });
    Object.keys(bySlide).forEach(s => {
        console.log(`  Slide ${s} (${bySlide[s].length} overlaps):`);
        bySlide[s].forEach((o, i) => {
            console.log(`    [${i+1}] "${o.b1.txt}" (y=${o.b1.y}, h=${o.b1.h}, bottom=${(o.b1.y+o.b1.h).toFixed(2)})`);
            console.log(`        COLLIDES WITH "${o.b2.txt}" (y=${o.b2.y}, h=${o.b2.h}, bottom=${(o.b2.y+o.b2.h).toFixed(2)})`);
        });
    });
});
