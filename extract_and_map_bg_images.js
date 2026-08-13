const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

const REF_DIR = "C:\\Users\\Riddhi Dumre\\Desktop\\BombayDC_Decks_With_BG_Images";
const OUT_MEDIA_DIR = "C:\\Users\\Riddhi Dumre\\Desktop\\ppt_automation\\extracted_bg_images";

if (!fs.existsSync(OUT_MEDIA_DIR)) {
    fs.mkdirSync(OUT_MEDIA_DIR, { recursive: true });
}

const files = fs.readdirSync(REF_DIR).filter(f => f.endsWith('.pptx'));
files.sort();

const masterMappings = {};

files.forEach((file, fileIdx) => {
    const deckNum = fileIdx + 1;
    const deckKey = file.replace('.pptx', '');
    const filePath = path.join(REF_DIR, file);
    const zip = new AdmZip(filePath);
    const zipEntries = zip.getEntries();

    const deckMediaFolder = path.join(OUT_MEDIA_DIR, `deck_${deckNum}`);
    if (!fs.existsSync(deckMediaFolder)) {
        fs.mkdirSync(deckMediaFolder, { recursive: true });
    }

    // Extract all media
    const mediaEntries = zipEntries.filter(e => e.entryName.startsWith('ppt/media/'));
    const mediaMap = {}; // rId -> filename

    mediaEntries.forEach(mEntry => {
        const baseName = path.basename(mEntry.entryName);
        const targetPath = path.join(deckMediaFolder, baseName);
        fs.writeFileSync(targetPath, mEntry.getData());
    });

    // Map slides to media
    const slideEntries = zipEntries.filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'));
    slideEntries.sort((a, b) => {
        const numA = parseInt(a.entryName.match(/slide(\d+)\.xml/)[1]);
        const numB = parseInt(b.entryName.match(/slide(\d+)\.xml/)[1]);
        return numA - numB;
    });

    const slideMappings = [];

    slideEntries.forEach((sEntry, sIdx) => {
        const slideNum = sIdx + 1;
        const relPath = `ppt/slides/_rels/${path.basename(sEntry.entryName)}.rels`;
        const relEntry = zipEntries.find(e => e.entryName === relPath);
        
        let slideImages = [];
        if (relEntry) {
            const relXml = zip.readAsText(relEntry);
            const matches = [...relXml.matchAll(/Target="\.\.\/media\/([^"]+)"/g)];
            if (matches.length > 0) {
                slideImages = matches.map(m => m[1]);
            }
        }

        // Parse slide title or text
        const slideXml = zip.readAsText(sEntry);
        const textMatches = [...slideXml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map(m => m[1]);

        slideMappings.push({
            slideNum,
            images: slideImages.map(img => path.join(deckMediaFolder, img)),
            textSnippet: textMatches.slice(0, 3).join(' | ')
        });
    });

    masterMappings[file] = {
        deckNum,
        file,
        slideCount: slideEntries.length,
        slides: slideMappings
    };
});

fs.writeFileSync(path.join(__dirname, 'deck_bg_mappings.json'), JSON.stringify(masterMappings, null, 2));
console.log(`Extracted all media and created deck_bg_mappings.json for ${files.length} decks.`);
