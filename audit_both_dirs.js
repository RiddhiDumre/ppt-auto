const fs = require('fs');
const path = require('path');
const AdmZip = require('C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\node_modules\\adm-zip');

const DIRS = [
    "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\new MD",
    "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\new MD"
];

DIRS.forEach(DIR => {
    console.log(`\n========================================`);
    console.log(`AUDITING DIR: ${DIR}`);
    console.log(`========================================`);
    const files = fs.readdirSync(DIR).filter(f => f.endsWith('.pptx'));
    let totalCollisions = 0;

    files.forEach(f => {
        const filePath = path.join(DIR, f);
        const zip = new AdmZip(filePath);
        const slideEntries = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'));
        slideEntries.sort((a,b) => parseInt(a.entryName.match(/slide(\d+)\.xml/)[1]) - parseInt(b.entryName.match(/slide(\d+)\.xml/)[1]));

        slideEntries.forEach((sEntry, sIdx) => {
            if (sIdx === 0 || sIdx === slideEntries.length - 1) return; // Skip cover and closing

            const xml = zip.readAsText(sEntry);
            const spMatches = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
            const textNodes = [];

            spMatches.forEach(sp => {
                const spXml = sp[1];
                if (!spXml.includes('<a:t>')) return;
                const offM = spXml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
                const extM = spXml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
                if (!offM || !extM) return;

                const x = parseFloat((parseInt(offM[1])/914400).toFixed(2));
                const y = parseFloat((parseInt(offM[2])/914400).toFixed(2));
                const w = parseFloat((parseInt(extM[1])/914400).toFixed(2));
                const h = parseFloat((parseInt(extM[2])/914400).toFixed(2));
                const rawTxt = [...spXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').trim();

                if (rawTxt.length === 0 || rawTxt.includes('CONFIDENTIAL') || rawTxt.includes('bombaydc.com')) return;
                textNodes.push({ x, y, w, h, bottom: parseFloat((y + h).toFixed(2)), txt: rawTxt.substring(0, 35) });
            });

            for (let i = 0; i < textNodes.length; i++) {
                for (let j = i + 1; j < textNodes.length; j++) {
                    const b1 = textNodes[i];
                    const b2 = textNodes[j];

                    const hOverlap = (b1.x < b2.x + b2.w - 0.05) && (b1.x + b1.w > b2.x + 0.05);
                    const vOverlap = (b1.y < b2.bottom - 0.05) && (b1.bottom > b2.y + 0.05);

                    if (hOverlap && vOverlap) {
                        totalCollisions++;
                        console.log(`COLLISION [${f}] Slide ${sIdx+1}: "${b1.txt}" (y=${b1.y}, h=${b1.h}) vs "${b2.txt}" (y=${b2.y}, h=${b2.h})`);
                    }
                }
            }
        });
    });

    console.log(`DIR RESULT: ${totalCollisions} Total Collisions in ${files.length} decks.`);
});
