const AdmZip = require('adm-zip');
const fs = require('fs');
const path = require('path');

const outDir = 'C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\new MD';
const files = fs.readdirSync(outDir).filter(f => f.endsWith('_BDC_Styled.pptx')).sort();

let slideMarginStats = [];

files.forEach(fname => {
    const zip = new AdmZip(path.join(outDir, fname));
    const slideEntries = zip.getEntries().filter(e => /ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
        .sort((a,b) => parseInt(a.entryName.match(/\d+/)[0]) - parseInt(b.entryName.match(/\d+/)[0]));

    slideEntries.forEach((entry, sIdx) => {
        if (sIdx === 0 || sIdx === slideEntries.length - 1) return; // skip cover and closing slide

        const xml = zip.readAsText(entry);
        const sps = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)];
        let minY = Infinity;
        let maxY = 0;

        sps.forEach(sp => {
            const off = sp[0].match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const ext = sp[0].match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            const txts = [...sp[0].matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').trim();
            if (!off || !ext || !txts) return;
            const y = parseFloat((parseInt(off[2]) / 914400).toFixed(2));
            const h = parseFloat((parseInt(ext[2]) / 914400).toFixed(2));
            const x = parseFloat((parseInt(off[1]) / 914400).toFixed(2));
            if (y < 1.35) return; // skip header area
            minY = Math.min(minY, y);
            maxY = Math.max(maxY, y + h);
        });

        if (minY < Infinity) {
            slideMarginStats.push({
                file: fname,
                slide: sIdx + 1,
                minY: minY.toFixed(2),
                maxY: maxY.toFixed(2)
            });
        }
    });
});

console.log('====== SLIDE MARGIN AUDIT (Top Content Y & Bottom Content Y) ======');
const minYs = slideMarginStats.map(s => parseFloat(s.minY));
const maxYs = slideMarginStats.map(s => parseFloat(s.maxY));

console.log(`Total Content Slides Audited: ${slideMarginStats.length}`);
console.log(`Top Content Y Range: min=${Math.min(...minYs)}in, max=${Math.max(...minYs)}in, avg=${(minYs.reduce((a,b)=>a+b,0)/minYs.length).toFixed(2)}in`);
console.log(`Bottom Content Y Range: min=${Math.min(...maxYs)}in, max=${Math.max(...maxYs)}in, avg=${(maxYs.reduce((a,b)=>a+b,0)/maxYs.length).toFixed(2)}in`);

// Find any outliers
const topOutliers = slideMarginStats.filter(s => parseFloat(s.minY) > 2.0);
const bottomOutliers = slideMarginStats.filter(s => parseFloat(s.maxY) > 5.10);

console.log('\n--- Top Outliers (minY > 2.0in) ---');
topOutliers.forEach(o => console.log(`  [${o.file} Slide ${o.slide}] topY=${o.minY}in bottomY=${o.maxY}in`));

console.log('\n--- Bottom Outliers (maxY > 5.10in) ---');
bottomOutliers.forEach(o => console.log(`  [${o.file} Slide ${o.slide}] topY=${o.minY}in bottomY=${o.maxY}in`));
