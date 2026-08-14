const AdmZip = require('C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\node_modules\\adm-zip');

function testDeck(deckPath, deckName) {
    const zip = new AdmZip(deckPath);
    const entries = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
        .sort((a,b) => parseInt(a.entryName.match(/slide(\d+)\.xml/)[1]) - parseInt(b.entryName.match(/slide(\d+)\.xml/)[1]));

    console.log(`\n========================================`);
    console.log(`VERIFICATION TEST: ${deckName}`);
    console.log(`========================================`);

    // Test (c): Cover slide
    const coverXml = zip.readAsText(entries[0]);
    const coverSp = [...coverXml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
    console.log(`[TEST C - Cover Slide]: Total shapes = ${coverSp.length}`);
    coverSp.forEach((m, idx) => {
        const s = m[1];
        const off = s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
        const txt = [...s.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t => t[1]).join(' ').trim();
        if (txt) {
            const x = (parseInt(off[1])/914400).toFixed(2);
            const y = (parseInt(off[2])/914400).toFixed(2);
            console.log(`  Cover Shape ${idx}: [x=${x}, y=${y}] -> "${txt.substring(0, 40)}"`);
        }
    });

    // Test (a): Short content slide (Slide 2)
    const shortXml = zip.readAsText(entries[1]);
    const shortSp = [...shortXml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
    const shortLines = [...shortXml.matchAll(/<p:cxnSp>([\s\S]*?)<\/p:cxnSp>/g)];
    console.log(`\n[TEST A - Short Content Slide 2]:`);
    let shortDividerY = null;
    shortLines.forEach(m => {
        const off = m[1].match(/<a:off x="(\d+)" y="(\d+)"\/>/);
        if (off) shortDividerY = (parseInt(off[2])/914400).toFixed(2);
    });
    console.log(`  Header Divider Line Y = ${shortDividerY} in (Expected 1.35)`);
    shortSp.forEach(m => {
        const s = m[1];
        const off = s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
        const txt = [...s.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t => t[1]).join(' ').trim();
        if (txt && off) {
            const x = (parseInt(off[1])/914400).toFixed(2);
            const y = (parseInt(off[2])/914400).toFixed(2);
            if (y < 1.35) console.log(`  Header element: [x=${x}, y=${y}] -> "${txt.substring(0, 40)}"`);
        }
    });

    // Test (b): Long content slide (Slide 6 or 8)
    const longXml = zip.readAsText(entries[entries.length > 5 ? 5 : 2]);
    const longSp = [...longXml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
    const longLines = [...longXml.matchAll(/<p:cxnSp>([\s\S]*?)<\/p:cxnSp>/g)];
    console.log(`\n[TEST B - Long Content Slide 6]:`);
    let longDividerY = null;
    longLines.forEach(m => {
        const off = m[1].match(/<a:off x="(\d+)" y="(\d+)"\/>/);
        if (off) longDividerY = (parseInt(off[2])/914400).toFixed(2);
    });
    console.log(`  Header Divider Line Y = ${longDividerY} in (Expected 1.35)`);
    longSp.forEach(m => {
        const s = m[1];
        const off = s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
        const txt = [...s.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t => t[1]).join(' ').trim();
        if (txt && off) {
            const x = (parseInt(off[1])/914400).toFixed(2);
            const y = (parseInt(off[2])/914400).toFixed(2);
            if (y < 1.35) console.log(`  Header element: [x=${x}, y=${y}] -> "${txt.substring(0, 40)}"`);
        }
    });
}

testDeck("C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\new MD\\V4_BDC_Styled.pptx", "V4_BDC_Styled.pptx");
testDeck("C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\new MD\\Intro deck_BDC_Styled.pptx", "Intro deck_BDC_Styled.pptx");
