// One-time copy: the GitHub pages → Apps Script HTML files in apps-script/ (served by 08_Web.gs).
// The GitHub pages are left exactly as they are, so trailjournal.org keeps working until the
// switch. After this runs, apps-script/{Journal,Respond,Admin,Panel,Followup}.html are the
// source of truth for the Apps Script pages.
//
//   node tools/split-for-apps-script.js
//
// What changes in each copy (everything else is untouched):
//   - the backend call: fetch(APP_URL…) with the public token → window.tjCall_ (google.script.run)
//   - APP_URL / APP_TOKEN come from the page's boot data (only the exit auto-save beacon uses them)
//   - URL parameters (?preview=, ?link=, ?tj=, …) come from the page's boot data, since a page
//     inside Apps Script can't see the address the visitor typed
//   - links between pages and the poster images point at trailjournal.org
//   - the journal's "start over" reload sends the whole window back to trailjournal.org
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const out = path.join(root, 'apps-script');
const SITE = 'https://trailjournal.org';

const pages = [
  ['index.html', 'Journal'], ['respond/index.html', 'Respond'], ['admin/index.html', 'Admin'],
  ['panel/index.html', 'Panel'], ['panel/followup/index.html', 'Followup'],
];

const BOOT = "<script>window.TJ_BOOT = <?!= boot ?>; // page, staff, siteUrl, beaconUrl, beaconToken, params (08_Web.gs)\n" +
  "window.TJ_SEARCH = '?' + Object.keys(TJ_BOOT.params || {}).filter(function (k) { return k !== 'page'; })" +
  ".map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(TJ_BOOT.params[k]); }).join('&');</script>\n" +
  "<?!= include('TjBridge'); ?>";

function rep(s, from, to, page, must) {
  const n = s.split(from).length - 1;
  if (must && !n) throw new Error(page + ': not found: ' + from);
  return s.split(from).join(to);
}

// The Apps Script copies have hand edits since (the 800# box in Journal.html), so never
// overwrite them by accident.
if (!process.argv.includes('--force') && pages.some(p => fs.existsSync(path.join(out, p[1] + '.html')))) {
  console.error('apps-script page files already exist and may have edits. Re-run with --force to overwrite them.');
  process.exit(1);
}

for (const [src, name] of pages) {
  let s =fs.readFileSync(path.join(root, src), 'utf8').replace(/^﻿/, '').replace(/\r\n/g, '\n');
  s = s.replace(/<head>/i, '<head>\n' + BOOT);
  s = rep(s, '<link rel="stylesheet" href="/styles.css">', "<?!= include('TrailCss'); ?>", name);
  s = s.replace(/const APP_URL\s*=\s*'[^']*';/, 'const APP_URL   = TJ_BOOT.beaconUrl;');
  s = s.replace(/const APP_TOKEN\s*=\s*'[^']*';/, 'const APP_TOKEN = TJ_BOOT.beaconToken; // only the exit auto-save beacon sends it');
  const before = s;
  s = s.replace(/const res = await fetch\(APP_URL, \{[\s\S]*?body: JSON\.stringify\(([\s\S]*?)\),\s*\}\);\s*const json = await res\.json\(\);/,
    'const json = await window.tjCall_($1); // google.script.run (08_Web.gs tjCall)');
  if (s === before) throw new Error(name + ': backend call not found');
  // URL parameters.
  s = rep(s, 'window.location.pathname + window.location.search', "TJ_BOOT.siteUrl + '/'", name);
  s = rep(s, 'window.location.href = TJ_BOOT.siteUrl', 'window.top.location.href = TJ_BOOT.siteUrl', name);
  s = s.replace(/new URLSearchParams\((window\.)?location\.search\)/g, 'new URLSearchParams(window.TJ_SEARCH)');
  // Links and images.
  s = rep(s, 'href="../?preview=true"', 'href="' + SITE + '/?preview=true"', name);
  s = rep(s, 'href="../respond/?preview=true"', 'href="' + SITE + '/respond/?preview=true"', name);
  s = rep(s, "'../assets/peaks/'", "'" + SITE + "/assets/peaks/'", name);
  s = rep(s, "('assets/peaks/'", "('" + SITE + "/assets/peaks/'", name);
  s = rep(s, 'href="/panel/followup"', 'href="' + SITE + '/panel/followup/"', name);
  s = rep(s, "'/panel/followup?'", "'" + SITE + "/panel/followup/?'", name);
  s = rep(s, 'href="/panel"', 'href="' + SITE + '/panel/"', name);
  if (/location\.search/.test(s.replace(/TJ_SEARCH/g, ''))) console.warn(name + ': still reads location.search');
  fs.writeFileSync(path.join(out, name + '.html'), s);
  console.log(name + '.html', fs.statSync(path.join(out, name + '.html')).size, 'bytes');
}
fs.writeFileSync(path.join(out, 'TrailCss.html'), '<style>\n' + fs.readFileSync(path.join(root, 'styles.css'), 'utf8').replace(/\r\n/g, '\n') + '\n</style>\n');
console.log('TrailCss.html written');
