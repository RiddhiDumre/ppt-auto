
const AdmZip = require('adm-zip');
const path = require('path');

const file = 'C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images\\01_BombayDC_BFSI_POV_Deck.pptx';
const zip = new AdmZip(file);

const entry = zip.getEntry('ppt/slides/slide7.xml');
const xml = zip.readAsText(entry);
const sps = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)];
console.log('====== SOURCE SLIDE 7 ======');
sps.forEach(sp => {
    const off = sp[0].match(/<a:off x="(\d+)" y="(\d+)"\/>/);
    const ext = sp[0].match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
    const txts = [...sp[0].matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ');
    const szMatch = sp[0].match(/sz="(\d+)"/);
    const fontSize = szMatch ? parseFloat((parseInt(szMatch[1]) / 100).toFixed(1)) : 0;
    if (!off || !ext) return;
    const y = parseFloat((parseInt(off[2]) / 914400).toFixed(2));
    const x = parseFloat((parseInt(off[1]) / 914400).toFixed(2));
    const h = parseFloat((parseInt(ext[2]) / 914400).toFixed(2));
    const w = parseFloat((parseInt(ext[1]) / 914400).toFixed(2));
    console.log(`  x=${x} y=${y} w=${w} h=${h} sz=${fontSize} "${txts.substring(0, 60)}"`);
});
