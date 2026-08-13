const AdmZip = require('adm-zip');
const fs = require('fs');
const path = require('path');

const outDir = 'C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\new MD';
const files = fs.readdirSync(outDir).filter(f => f.endsWith('_BDC_Styled.pptx')).sort();

let trueDoubleBoxCount = 0;

files.forEach(fname => {
    const zip = new AdmZip(path.join(outDir, fname));
    const slideEntries = zip.getEntries().filter(e => /ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
        .sort((a,b) => parseInt(a.entryName.match(/\d+/)[0]) - parseInt(b.entryName.match(/\d+/)[0]));

    slideEntries.forEach((entry, sIdx) => {
        const xml = zip.readAsText(entry);
        const sps = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)];
        let rectFills = [];

        sps.forEach(sp => {
            const off = sp[0].match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const ext = sp[0].match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            const spPr = sp[0].match(/<p:spPr>[\s\S]*?<\/p:spPr>/);
            const fill = spPr ? spPr[0].match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/) : null;
            if (!off || !ext || !fill) return;
            const x = parseFloat((parseInt(off[1]) / 914400).toFixed(2));
            const y = parseFloat((parseInt(off[2]) / 914400).toFixed(2));
            const w = parseFloat((parseInt(ext[1]) / 914400).toFixed(2));
            const h = parseFloat((parseInt(ext[2]) / 914400).toFixed(2));
            if (["F4F4F2", "1C1C1E", "034E48"].includes(fill[1].toUpperCase())) {
                rectFills.push({ x, y, w, h, fill: fill[1] });
            }
        });

        // Check for shape background rectangles whose top-left corner (x, y) is within 0.15 in of another shape background rectangle
        for (let i = 0; i < rectFills.length; i++) {
            for (let j = i + 1; j < rectFills.length; j++) {
                const a = rectFills[i];
                const b = rectFills[j];
                if (Math.abs(a.x - b.x) < 0.15 && Math.abs(a.y - b.y) < 0.15) {
                    trueDoubleBoxCount++;
                    console.log(`[TRUE DOUBLE RECTANGLE] ${fname} Slide ${sIdx + 1}: Box A (${a.x}, ${a.y}) vs Box B (${b.x}, ${b.y})`);
                }
            }
        }
    });
});

console.log(`====== TRUE SHAPE BACKGROUND DOUBLE BOX AUDIT ======`);
console.log(`Total True Duplicate Shape Background Rectangles Found: ${trueDoubleBoxCount}`);
