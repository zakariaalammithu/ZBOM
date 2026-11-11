const fs = require('fs');
const path = require('path');

const svgContent = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" width="100" height="100">
  <defs>
    <linearGradient id="grad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#0a6ed1"/>
      <stop offset="100%" stop-color="#053b73"/>
    </linearGradient>
    <linearGradient id="accent" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#107e3e"/>
      <stop offset="100%" stop-color="#31bd87"/>
    </linearGradient>
  </defs>
  <rect width="100" height="100" rx="22" fill="url(#grad)"/>
  <path d="M 25 28 L 75 28 L 38 72 L 75 72" fill="none" stroke="#ffffff" stroke-width="11" stroke-linecap="round" stroke-linejoin="round"/>
  <circle cx="75" cy="28" r="6" fill="url(#accent)"/>
  <circle cx="38" cy="72" r="6" fill="url(#accent)"/>
</svg>`;

const assetsDir = path.join(__dirname, 'renderer', 'assets');
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

fs.writeFileSync(path.join(assetsDir, 'zbom-logo.svg'), svgContent);
fs.writeFileSync(path.join(assetsDir, 'zcode-logo.svg'), svgContent);
console.log('ZBOM logo SVG created successfully');
