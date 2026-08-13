const fs = require('fs');
const path = require('path');

function searchFiles(dir) {
    const results = [];
    const items = fs.readdirSync(dir, { withFileTypes: true });
    for (const item of items) {
        const fullPath = path.join(dir, item.name);
        if (item.isDirectory()) {
            results.push(...searchFiles(fullPath));
        } else if (item.name.toLowerCase().includes('b2b') && item.name.endsWith('.pptx')) {
            results.push({ path: fullPath, size: fs.statSync(fullPath).size, mtime: fs.statSync(fullPath).mtime });
        }
    }
    return results;
}

const files = searchFiles("C:\\Users\\Riddhi Dumre\\Desktop");
console.log("=== ALL B2B PPTX FILES ON DESKTOP ===");
files.forEach(f => console.log(`  [${f.mtime.toISOString()}] ${f.path}`));
