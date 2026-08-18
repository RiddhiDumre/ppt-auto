const AdmZip = require('C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\node_modules\\adm-zip');
const path = require('path');
const REF_DIR = 'C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images';
const files = [
    '01_BombayDC_BFSI_POV_Deck.pptx',
    '02_BFSI_Platform_Pitch_V4.pptx',
    '03_BFSI_CXO_Intro_V5.pptx',
    '04_Consumer_Platforms_Deck.pptx',
    '05_Corporate_Brand_Platforms_Deck.pptx',
    '06_Corporate_Digital_Platforms_Deck_v2.pptx',
    '07_Ecommerce_POV_Deck_Rebuilt.pptx',
    '08_Ecommerce_V5.pptx',
    '09_B2B_POV_Deck.pptx',
    '10_Enterprise_Platforms_POV.pptx',
    '11_Intro_Deck.pptx'
];

files.forEach(f => {
    try {
        const zip = new AdmZip(path.join(REF_DIR, f));
        const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
            .sort((a,b) => parseInt(a.entryName.match(/(\d+)/)[1]) - parseInt(b.entryName.match(/(\d+)/)[1]));
        slides.forEach((s, i) => {
            const xml = zip.readAsText(s);
            if (xml.includes('BEAM') || xml.toLowerCase().includes('beam')) {
                const txt = [...xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m=>m[1]).filter(t=>t.includes('BEAM')).join(' | ');
                // Check for footer images
                const picMatches = [...xml.matchAll(/<p:pic>([\s\S]*?)<\/p:pic>/g)];
                const footerPics = picMatches.filter(pm => {
                    const off = pm[1].match(/a:off x="(\d+)" y="(\d+)"/);
                    if (!off) return false;
                    return parseInt(off[2])/914400 > 4.5;
                });
                console.log(`  [BEAM FOUND] ${f} slide ${i+1} | footer_pics=${footerPics.length} | "${txt.substring(0,80)}"`);
            }
        });
    } catch(e) { console.log(`  ERROR ${f}: ${e.message}`); }
});
