// Local look at the Apps Script page copies: renders each page with its includes and a
// stand-in google.script.run that answers every call with an empty success. For checking that
// pages load and look right (not for testing the backend: see tests/backend.test.js).
//   node tests/build-preview.js   → http://localhost:3044/?page=journal  (respond, admin, panel, followup)
//   add &staff=1 to see a page as signed-in staff
const fs = require('fs');
const path = require('path');
const http = require('http');
const dir = path.join(__dirname, '..', 'apps-script');
const read = f => fs.readFileSync(path.join(dir, f), 'utf8');
const PAGES = { journal: 'Journal', respond: 'Respond', admin: 'Admin', panel: 'Panel', followup: 'Followup', signup: 'Signup', checkin: 'Checkin', givget: 'GiveGet', consent: 'Consent', outcomes: 'Outcomes' };

http.createServer((req, res) => {
  const q = new URL(req.url, 'http://x').searchParams;
  const page = PAGES[q.get('page')] ? q.get('page') : 'journal';
  const params = Object.fromEntries(q.entries());
  const boot = { page, staff: q.get('staff') ? 'teacher@easdpa.org' : '', siteUrl: 'https://trailjournal.org', beaconUrl: 'about:blank', beaconToken: '', params };
  const stub = `<script>window.__calls = [];
window.google = { script: { run: (function make(ok) { return new Proxy({}, { get(_, fn) {
  if (fn === 'withSuccessHandler') return h => make(h);
  if (fn === 'withFailureHandler') return () => make(ok);
  return (...a) => { window.__calls.push([fn, a]); setTimeout(() => ok && ok(JSON.stringify({ ok: true, data: [] })), 20); };
} }); })() } };</script>`;
  const html = read(PAGES[page] + '.html')
    .replace(/<\?!= include\('([A-Za-z]+)'\); \?>/g, (m, f) => read(f + '.html'))
    .replace('<?!= boot ?>', JSON.stringify(boot))
    .replace('<head>', '<head>' + stub);
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(html);
}).listen(3044, () => console.log('Preview on http://localhost:3044'));
