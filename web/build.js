// Builds the pages for our own domain from the Apps Script sources.
//
//   node web/build.js                      keeps the address already in web/dist/config.js
//   node web/build.js --api <exec URL>     writes that address into web/dist/config.js
//
// Upload the CONTENTS of web/dist to the subdomain's document root.
'use strict';
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const SOURCE = path.join(__dirname, '..', 'MovingSurfaceERP');
const PAGES = [
  {folder: '', view: 'AttendanceView', title: 'Moving Surface Attendance'},
  {folder: 'staff', view: 'StaffPortalView', title: 'Moving Surface Staff Portal'},
  {folder: 'requisition', view: 'RequisitionView', title: 'Moving Surface Requisition'},
  {folder: 'admin', view: 'AdminPortalView', title: 'Moving Surface ERP'},
  {folder: 'recover', view: 'RecoveryView', title: 'Staff account recovery'}
];
const API_PLACEHOLDER = 'PASTE_EXEC_URL';
const THEME = '#087f68';

const read = file => fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, '').replace(/\r\n/g, '\n');
const source = name => read(path.join(SOURCE, name + '.html'));

/** Replaces exactly one occurrence, so a change in the sources cannot be skipped silently. */
function swap(text, find, replacement, what) {
  const parts = text.split(find);
  if (parts.length !== 2) throw new Error('Build stopped: expected exactly one "' + what + '" in the sources, found ' + (parts.length - 1));
  return parts[0] + replacement + parts[1];
}

function buildPage(page, analytics) {
  const up = page.folder ? '../' : '';
  let view = source(page.view);
  if (page.view === 'RecoveryView') {
    view = swap(view, "location.pathname+'?recover=1'", 'location.pathname', 'recovery address reset');
  }
  const head = [
    '<meta charset="utf-8">',
    '<title>' + page.title + '</title>',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    '<meta name="theme-color" content="' + THEME + '">',
    '<meta name="apple-mobile-web-app-capable" content="yes">',
    '<meta name="apple-mobile-web-app-title" content="MS ERP">',
    '<link rel="manifest" href="' + up + 'manifest.webmanifest">',
    '<link rel="icon" type="image/png" href="' + up + 'icon-192.png">',
    '<link rel="apple-touch-icon" href="' + up + 'apple-touch-icon.png">',
    '<script src="' + up + 'config.js"></script>',
    '<script>\n' + read(path.join(__dirname, 'bridge.js')) + '</script>'
  ].join('\n    ');

  let html = source('Index');
  html = swap(html, '<base target="_top">\n    <meta charset="utf-8">', head, 'page head');
  html = swap(html, "<?!= include_('Styles'); ?>", source('Styles'), 'styles include');
  html = swap(html, '<?= canonicalUrl ?>', '', 'canonical address');
  // The recovery code arrives after '#', which is never sent to the web server.
  html = swap(html, "window.MS_RECOVERY_TOKEN = '<?= recoveryToken ?>';",
    "window.MS_RECOVERY_TOKEN = (/[#&]token=([A-Za-z0-9]+)/.exec(window.location.hash) || [])[1] || '';", 'recovery token');
  html = swap(html, '<?!= include_(viewName); ?>', view, 'view include');
  if (page.view !== 'RecoveryView' && analytics.replace(/<!--[\s\S]*?-->/g, '').trim()) {
    html = swap(html, '</body>', analytics + '\n  </body>', 'end of body');
  }
  if (/<\?[!=]?/.test(html)) throw new Error('Build stopped: an Apps Script template tag is left in ' + page.view);
  return html;
}

/** A square PNG: brand green with a white "screen" mark, so the home-screen icon needs no design tool. */
function icon(size) {
  const raw = Buffer.alloc((size * 3 + 1) * size);
  const green = [8, 127, 104], white = [255, 255, 255];
  const inside = (x, y, x0, y0, x1, y1) => x >= x0 * size && x < x1 * size && y >= y0 * size && y < y1 * size;
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      const frame = inside(x, y, .2, .26, .8, .66) && !inside(x, y, .26, .32, .74, .6);
      const stand = inside(x, y, .46, .66, .54, .74) || inside(x, y, .34, .74, .66, .8);
      const colour = frame || stand ? white : green;
      raw.set(colour, y * (size * 3 + 1) + 1 + x * 3);
    }
  }
  const crcTable = Array.from({length: 256}, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc = buffer => {
    let c = 0xffffffff;
    for (const byte of buffer) c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const body = Buffer.concat([Buffer.from(type), data]);
    const out = Buffer.alloc(body.length + 8);
    out.writeUInt32BE(data.length, 0);
    body.copy(out, 4);
    out.writeUInt32BE(crc(body), body.length + 4);
    return out;
  };
  const header = Buffer.alloc(13);
  header.writeUInt32BE(size, 0);
  header.writeUInt32BE(size, 4);
  header.set([8, 2, 0, 0, 0], 8);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', header),
    chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

const HTACCESS = `Options -Indexes
DirectoryIndex index.html
RewriteEngine On

# Leave certificate renewal files alone
RewriteRule ^\\.well-known/ - [L]

# Always use HTTPS: phones only share their location over it
RewriteCond %{HTTPS} off
RewriteRule ^ https://%{HTTP_HOST}%{REQUEST_URI} [R=302,L]

# Old-style addresses such as ?admin=1
RewriteCond %{QUERY_STRING} (^|&)admin=1(&|$) [NC]
RewriteRule ^$ %{REQUEST_URI}admin/? [R=302,L]
RewriteCond %{QUERY_STRING} (^|&)staff=1(&|$) [NC]
RewriteRule ^$ %{REQUEST_URI}staff/? [R=302,L]
RewriteCond %{QUERY_STRING} (^|&)req=1(&|$) [NC]
RewriteRule ^$ %{REQUEST_URI}requisition/? [R=302,L]
RewriteCond %{QUERY_STRING} (^|&)recover=1(&|$) [NC]
RewriteRule ^$ %{REQUEST_URI}recover/? [R=302,L]

<IfModule mod_mime.c>
  AddType application/manifest+json .webmanifest
</IfModule>

<IfModule mod_headers.c>
  # Pages and settings are re-checked on every visit, so an upload takes effect at once
  <FilesMatch "\\.(html|js|webmanifest)$">
    Header set Cache-Control "no-cache"
  </FilesMatch>
  Header set X-Content-Type-Options "nosniff"
  Header set X-Frame-Options "DENY"
  Header set Referrer-Policy "strict-origin-when-cross-origin"
</IfModule>
`;

function build(options) {
  const out = options.out || path.join(__dirname, 'dist');
  const analytics = read(path.join(__dirname, 'analytics.html'));
  fs.mkdirSync(out, {recursive: true});
  PAGES.forEach(page => {
    const folder = path.join(out, page.folder);
    fs.mkdirSync(folder, {recursive: true});
    fs.writeFileSync(path.join(folder, 'index.html'), buildPage(page, analytics));
  });

  const configFile = path.join(out, 'config.js');
  if (options.api || !fs.existsSync(configFile)) {
    const api = options.api || API_PLACEHOLDER;
    if (options.api && !/^https?:\/\/[^\s'"\\]+$/.test(api)) throw new Error('--api must be a web address.');
    fs.writeFileSync(configFile,
      '// The Apps Script web app address ending in /exec (Deploy > Manage deployments).\n' +
      "window.MS_API_URL = '" + api + "';\n");
  }
  fs.writeFileSync(path.join(out, '.htaccess'), HTACCESS);
  fs.writeFileSync(path.join(out, 'manifest.webmanifest'), JSON.stringify({
    name: 'Moving Surface ERP', short_name: 'MS ERP', start_url: './', scope: './', display: 'standalone',
    background_color: '#f2f6f4', theme_color: THEME,
    icons: [192, 512].map(size => ({src: 'icon-' + size + '.png', sizes: size + 'x' + size, type: 'image/png'}))
  }, null, 2) + '\n');
  fs.writeFileSync(path.join(out, 'apple-touch-icon.png'), icon(180));
  [192, 512].forEach(size => fs.writeFileSync(path.join(out, 'icon-' + size + '.png'), icon(size)));
  return out;
}

module.exports = {build, PAGES};

if (require.main === module) {
  const at = process.argv.indexOf('--api');
  if (at !== -1 && !process.argv[at + 1]) throw new Error('--api needs a web address.');
  const out = build({api: at !== -1 ? process.argv[at + 1] : ''});
  const config = fs.readFileSync(path.join(out, 'config.js'), 'utf8');
  console.log('Built ' + PAGES.length + ' pages into ' + out);
  if (config.includes(API_PLACEHOLDER)) console.log('Next: put your /exec address into config.js, or rebuild with --api <address>.');
}
