const AdmZip = require('C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\node_modules\\adm-zip');
const fs = require('fs');
const path = require('path');

const outDir = 'C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\new MD';
const files = fs.readdirSync(outDir).filter(f => f.endsWith('_BDC_Styled.pptx')).sort();

let overlapCount = 0;

files.forEach(fname => {
    const zip = new AdmZip(path.join(outDir, fname));
    const slideEntries = zip.getEntries().filter(e => /ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
        .sort((a,b) => parseInt(a.entryName.match(/\d+/)[0]) - parseInt(b.entryName.match(/\d+/)[0]));

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
            boxes.push({ x, y, w, h, txt: txts.substring(0, 40) });
        });

        // Check for overlaps
        for (let i = 0; i < boxes.length; i++) {
            for (let j = i + 1; j < boxes.length; j++) {
                const b1 = boxes[i];
                const b2 = boxes[j];
                const hOverlap = (b1.x < b2.x + b2.w - 0.2) && (b1.x + b1.w > b2.x + 0.2);
                const vOverlap = (b1.y < b2.y + b2.h - 0.08) && (b1.y + b1.h > b2.y + 0.08);
                if (hOverlap && vOverlap) {
                    overlapCount++;
                    console.log(`OVERLAP #${overlapCount} in [${fname}] Slide ${sIdx+1}:`);
                    console.log(`  Box 1: y=${b1.y} h=${b1.h} (bottom=${(b1.y+b1.h).toFixed(2)}) x=${b1.x} w=${b1.w} | "${b1.txt}"`);
                    console.log(`  Box 2: y=${b2.y} h=${b2.h} (bottom=${(b2.y+b2.h).toFixed(2)}) x=${b2.x} w=${b2.w} | "${b2.txt}"`);
                }
            }
        }
    });
});

console.log(`\nTotal Overlaps Detected: ${overlapCount}`);
