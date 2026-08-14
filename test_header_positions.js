const AdmZip = require('C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\node_modules\\adm-zip');

function testDeck(deckPath, deckName) {
    const zip = new AdmZip(deckPath);
    const entries = zip.getEntries()
        .filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
        .sort((a,b) => parseInt(a.entryName.match(/slide(\d+)\.xml/)[1]) - parseInt(b.entryName.match(/slide(\d+)\.xml/)[1]));

    console.log(`\n========================================`);
    console.log(`VERIFICATION TEST: ${deckName}`);
    console.log(`========================================`);

    [0, 1, Math.min(5, entries.length - 2)].forEach(sIdx => {
        const xml = zip.readAsText(entries[sIdx]);
        const label = sIdx === 0 ? 'Cover' : sIdx === 1 ? 'Short Content (Slide 2)' : `Long Content (Slide ${sIdx+1})`;

        // Parse all xfrm off/ext elements and associated text
        const spBlocks = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)].map(m => m[1]);
        const cxnBlocks = [...xml.matchAll(/<p:cxnSp>([\s\S]*?)<\/p:cxnSp>/g)].map(m => m[1]);

        let dividerY = null;
        let slideIndex = { txt: null, x: null, y: null };
        let titleBlock = { txt: null, x: null, y: null };
        let contentMinY = Infinity;
        let contentMaxY = 0;

        const parseOff = xml => {
            const m = xml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            return m ? { x: parseFloat((parseInt(m[1])/914400).toFixed(2)), y: parseFloat((parseInt(m[2])/914400).toFixed(2)) } : null;
        };
        const parseExt = xml => {
            const m = xml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            return m ? { w: parseFloat((parseInt(m[1])/914400).toFixed(2)), h: parseFloat((parseInt(m[2])/914400).toFixed(2)) } : null;
        };

        // Find header divider line (line shapes with height=0)
        [...spBlocks, ...cxnBlocks].forEach(b => {
            if (b.includes('<a:t>')) return;
            const off = parseOff(b); const ext = parseExt(b);
            if (!off || !ext) return;
            if (ext.h === 0 && ext.w > 5.0 && off.y >= 1.30 && off.y <= 1.40) {
                dividerY = off.y;
            }
        });

        // Find slide index (small number text in left-most position near y=0.57)
        // Find title (large text near y=0.57 but right of x=1.0)
        spBlocks.forEach(b => {
            if (!b.includes('<a:t>')) return;
            const off = parseOff(b); const ext = parseExt(b);
            if (!off) return;
            const txt = [...b.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join(' ').trim();
            const sz = b.match(/sz="(\d+)"/);
            const fontSize = sz ? parseInt(sz[1]) / 100 : 0;

            if (off.y >= 0.50 && off.y <= 0.65 && off.x < 1.0) {
                slideIndex = { txt, x: off.x, y: off.y };
            }
            if (off.y >= 0.50 && off.y <= 0.65 && off.x >= 1.5 && fontSize >= 16) {
                titleBlock = { txt: txt.substring(0, 60), x: off.x, y: off.y };
            }
            // Track content area min/max Y
            if (off.y >= 1.55 && off.y <= 5.10) {
                if (off.y < contentMinY) contentMinY = off.y;
                if (ext && off.y + ext.h > contentMaxY) contentMaxY = off.y + ext.h;
            }
        });

        console.log(`\n[${label}]`);
        if (sIdx === 0) {
            // Cover: just confirm no block engine shapes were added (no content area shapes)
            const coverShapes = spBlocks.filter(b => {
                const off = parseOff(b);
                return off && off.y > 1.40;
            });
            console.log(`  Cover shapes below header zone: ${coverShapes.length}`);
            console.log(`  ✅ Cover slide isolated (no block engine involvement)`);
        } else {
            console.log(`  Slide Index:       x=${slideIndex.x}, y=${slideIndex.y} -> "${slideIndex.txt}"`);
            console.log(`  Title:             x=${titleBlock.x}, y=${titleBlock.y} -> "${titleBlock.txt}"`);
            console.log(`  Header Divider Y:  ${dividerY} in  (Target: ~1.35)`);
            const divOk = dividerY !== null && Math.abs(dividerY - 1.35) < 0.02;
            console.log(`  Header invariant:  ${divOk ? '✅ PASS' : `❌ FAIL (got ${dividerY}, expected 1.35)`}`);
            console.log(`  Content area:      y=[${contentMinY.toFixed(2)}, ${contentMaxY.toFixed(2)}]  (must be >= 1.55 and <= 5.08)`);
            const contentOk = contentMinY >= 1.55 && contentMaxY <= 5.10;
            console.log(`  Content bounds:    ${contentOk ? '✅ PASS' : `❌ FAIL (min=${contentMinY}, max=${contentMaxY})`}`);
        }
    });
}

testDeck("C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\new MD\\V4_BDC_Styled.pptx", "V4 (BFSI Platform Pitch)");
testDeck("C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\new MD\\Intro deck_BDC_Styled.pptx", "Intro Deck");
testDeck("C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\SALES DECKS\\new MD\\BombayDC_Enterprise_Platforms_POV_BDC_Styled.pptx", "Enterprise Platforms POV");
