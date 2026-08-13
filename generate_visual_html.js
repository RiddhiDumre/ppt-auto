const AdmZip = require('adm-zip');
const fs = require('fs');
const path = require('path');

const outDir = 'C:\\Users\\Riddhi Dumre\\Desktop\\Presentation Automation\\new MD';
const pptxFiles = fs.readdirSync(outDir).filter(f => f.endsWith('_BDC_Styled.pptx')).sort();

let htmlContent = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>BombayDC Decks - Full Visual QA Audit</title>
<style>
  body { font-family: 'Inter', sans-serif; background: #111; color: #fff; margin: 0; padding: 20px; }
  h1 { text-align: center; color: #4DB89A; margin-bottom: 30px; }
  .deck-section { margin-bottom: 50px; border-bottom: 2px solid #333; padding-bottom: 30px; }
  .deck-title { font-size: 24px; color: #fff; margin-bottom: 20px; border-left: 5px solid #4DB89A; padding-left: 15px; }
  .slides-grid { display: flex; flex-wrap: wrap; gap: 30px; justify-content: center; }
  .slide-card { background: #222; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.5); }
  .slide-header { background: #333; padding: 10px 15px; font-size: 14px; color: #aaa; display: flex; justify-content: space-between; }
  .slide-frame { width: 960px; height: 540px; position: relative; background: #ECE9E4; overflow: hidden; transform-origin: top left; }
  .shape-text { position: absolute; box-sizing: border-box; line-height: 1.25; word-wrap: break-word; }
  .shape-card { position: absolute; border-radius: 6px; box-sizing: border-box; }
  .shape-line { position: absolute; box-sizing: border-box; }
</style>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
</head>
<body>
<h1>BombayDC Presentation Suite — Visual QA Audit</h1>
`;

const SCALE = 96; // 96px per inch -> 960px x 540px for 10" x 5.625"

pptxFiles.forEach(fname => {
    const zip = new AdmZip(path.join(outDir, fname));
    const slideEntries = zip.getEntries()
        .filter(e => /ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
        .sort((a, b) => parseInt(a.entryName.match(/(\d+)/)[1]) - parseInt(b.entryName.match(/(\d+)/)[1]));

    htmlContent += `<div class="deck-section">\n<div class="deck-title">${fname} (${slideEntries.length} slides)</div>\n<div class="slides-grid">\n`;

    slideEntries.forEach((entry, idx) => {
        const xml = zip.readAsText(entry);
        const shapes = [...xml.matchAll(/<(?:p:sp|p:cxnSp)>[\s\S]*?<\/(?:p:sp|p:cxnSp)>/g)];

        // Background color check
        const bgClrMatch = xml.match(/<p:bgPr>[\s\S]*?<a:srgbClr val="([^"]+)"/);
        const bgColor = bgClrMatch ? `#${bgClrMatch[1]}` : (idx === 0 ? '#1D1D1F' : '#ECE9E4');

        htmlContent += `
  <div class="slide-card">
    <div class="slide-header"><span>Slide ${idx + 1}</span><span>${fname.substring(0, 25)}</span></div>
    <div class="slide-frame" style="background-color: ${bgColor};">
`;

        shapes.forEach(sp => {
            const off = sp[0].match(/<a:off x="(\d+)" y="(\d+)"\/>/);
            const ext = sp[0].match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            if (!off || !ext) return;

            const x = (parseInt(off[1]) / 914400) * SCALE;
            const y = (parseInt(off[2]) / 914400) * SCALE;
            const w = (parseInt(ext[1]) / 914400) * SCALE;
            const h = (parseInt(ext[2]) / 914400) * SCALE;

            const spPrMatch = sp[0].match(/<p:spPr>[\s\S]*?<\/p:spPr>/);
            const spPrXml = spPrMatch ? spPrMatch[0] : "";
            const fillMatch = spPrXml.match(/<a:solidFill>[\s\S]*?<a:srgbClr val="([^"]+)"/);
            const fillClr = fillMatch ? `#${fillMatch[1]}` : null;

            const isLine = sp[0].includes('prst="line"') || sp[0].startsWith('<p:cxnSp') || h === 0;

            if (isLine) {
                const lnClrMatch = spPrXml.match(/<a:ln[\s\S]*?<a:srgbClr val="([^"]+)"/);
                const lnColor = lnClrMatch ? `#${lnClrMatch[1]}` : "#B4B4B4";
                htmlContent += `<div class="shape-line" style="left:${x}px; top:${y}px; width:${w || 1}px; height:${h || 1}px; background-color:${lnColor}; opacity: 0.6;"></div>\n`;
                return;
            }

            if (fillClr && w < 900) {
                htmlContent += `<div class="shape-card" style="left:${x}px; top:${y}px; width:${w}px; height:${h}px; background-color:${fillClr};"></div>\n`;
            }

            const pMatches = [...sp[0].matchAll(/<a:p>[\s\S]*?<\/a:p>/g)];
            if (pMatches.length > 0) {
                let innerHtml = "";
                pMatches.forEach(pXml => {
                    const txts = [...pXml[0].matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]).join('');
                    if (!txts.trim()) return;

                    const szMatch = pXml[0].match(/sz="(\d+)"/);
                    const fontSize = szMatch ? (parseInt(szMatch[1]) / 100) * (SCALE / 72) * 1.05 : 12;
                    const bMatch = pXml[0].includes('b="1"');
                    const clrMatch = pXml[0].match(/<a:srgbClr val="([^"]+)"/);
                    const color = clrMatch ? `#${clrMatch[1]}` : '#1E1E1E';
                    const fontFace = pXml[0].includes('Inter Medium') ? 'Inter' : 'Inter';

                    innerHtml += `<div style="font-size:${fontSize.toFixed(1)}px; font-weight:${bMatch ? 700 : 400}; color:${color}; font-family:${fontFace}; margin-bottom: 3px;">${txts}</div>`;
                });

                if (innerHtml) {
                    htmlContent += `<div class="shape-text" style="left:${x}px; top:${y}px; width:${w}px; height:${h}px;">${innerHtml}</div>\n`;
                }
            }
        });

        htmlContent += `    </div>\n  </div>\n`;
    });

    htmlContent += `</div>\n</div>\n`;
});

htmlContent += `</body>\n</html>`;

const outHtmlPath = 'C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\visual_qa_audit.html';
fs.writeFileSync(outHtmlPath, htmlContent);
console.log(`🎉 Visual QA Audit HTML generated successfully at: ${outHtmlPath}`);
