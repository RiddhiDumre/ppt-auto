const fs = require('fs');
const path = require('path');
const AdmZip = require('C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\node_modules\\adm-zip');

const DIR = "C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\new MD";
const files = fs.readdirSync(DIR).filter(f => f.endsWith('.pptx'));

console.log("=== AUDITING TEXT CLIPPING & BOTTOM BOUNDARIES ===");

files.forEach(f => {
    const filePath = path.join(DIR, f);
    const zip = new AdmZip(filePath);
    const slideEntries = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'));
    slideEntries.sort((a,b) => parseInt(a.entryName.match(/slide(\d+)\.xml/)[1]) - parseInt(b.entryName.match(/slide(\d+)\.xml/)[1]));

    slideEntries.forEach((sEntry, sIdx) => {
        if (sIdx === 0 || sIdx === slideEntries.length - 1) return;

        const xml = zip.readAsText(sEntry);
        const spMatches = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
        
        spMatches.forEach((sp, i) => {
            const spXml = sp[1];
            if (!spXml.includes('<a:t>')) return;
            const offM = spXml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const extM = spXml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            if (!offM || !extM) return;

            const x = parseFloat((parseInt(offM[1])/914400).toFixed(2));
            const y = parseFloat((parseInt(offM[2])/914400).toFixed(2));
            const w = parseFloat((parseInt(extM[1])/914400).toFixed(2));
            const h = parseFloat((parseInt(extM[2])/914400).toFixed(2));
            const bottom = parseFloat((y + h).toFixed(2));

            const rawTxt = [...spXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').trim();
            if (rawTxt.length === 0 || rawTxt.includes('CONFIDENTIAL') || rawTxt.includes('bombaydc.com')) return;

            if (bottom > 5.08) {
                console.log(`[NEAR/OVER BOTTOM] ${f} Slide ${sIdx+1}: y=${y}, h=${h}, bottom=${bottom} (x=${x}, w=${w}) | "${rawTxt.substring(0,40)}..."`);
            }
        });
    });
});
