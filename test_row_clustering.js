const AdmZip = require('C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\node_modules\\adm-zip');

const zip = new AdmZip("C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images\\02_BFSI_Platform_Pitch_V4.pptx");
const slides = zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'))
    .sort((a,b)=>parseInt(a.entryName.match(/(\d+)/)[1])-parseInt(b.entryName.match(/(\d+)/)[1]));

[11, 12, 13].forEach(sIdx => {
    const xml = zip.readAsText(slides[sIdx]);
    console.log(`\n=== V4 SLIDE ${sIdx+1} ===`);
    const spM = [...xml.matchAll(/<p:sp>([\s\S]*?)<\/p:sp>/g)];
    const shapes = [];
    spM.forEach((m, i) => {
        const s = m[1];
        const off = s.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
        const ext = s.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
        if (!off || !ext) return;
        const x = parseFloat((parseInt(off[1])/914400).toFixed(2));
        const y = parseFloat((parseInt(off[2])/914400).toFixed(2));
        const w = parseFloat((parseInt(ext[1])/914400).toFixed(2));
        const txt = [...s.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(t => t[1]).join(' ').trim();
        if (txt && y >= 1.5) shapes.push({ x, y, w, txt: txt.substring(0, 30) });
    });

    const clusters = [];
    shapes.forEach(s => {
        let c = clusters.find(cl => Math.abs(cl.origY - s.y) < 0.28);
        if (!c) {
            c = { origY: s.y, shapes: [] };
            clusters.push(c);
        }
        c.shapes.push(s);
    });

    clusters.forEach((cl, cIdx) => {
        console.log(`Cluster ${cIdx} (origY=${cl.origY}):`, cl.shapes.map(s => `[x=${s.x} w=${s.w} "${s.txt}"]`));
    });
});
