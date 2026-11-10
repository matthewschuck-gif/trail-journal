// Builds paste-pack.html: every Apps Script file in paste order, each with a Copy button and a
// done checkbox (remembered in this browser). Open it in Chrome next to the Apps Script editor.
//   node tools/build-paste-pack.js
const fs = require('fs');
const path = require('path');
const dir = path.join(__dirname, '..', 'apps-script');
const NEW = ['11_Trailback.gs', '12_TrailbackCases.gs']; // files that don't exist in the Apps Script project yet
const files = fs.readdirSync(dir).filter(f => /\.(gs|html)$/.test(f)).sort((a, b) => {
  const ga = a.endsWith('.gs') ? 0 : 1, gb = b.endsWith('.gs') ? 0 : 1;
  return ga - gb || a.localeCompare(b);
});
const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const rows = files.map((f, i) => {
  const src = fs.readFileSync(path.join(dir, f), 'utf8');
  const name = f.replace(/\.(gs|html)$/, ''), kind = f.endsWith('.gs') ? 'Script' : 'HTML';
  return `<section id="f${i}"><header><label><input type="checkbox" data-k="${f}"> <b>${i + 1}. ${name}</b></label>
<span class="k">${kind}${NEW.includes(f) ? ' · <em>NEW: + › ' + kind + ', name it ' + name + '</em>' : ''}</span>
<span class="n">${src.split('\n').length} lines</span><button data-i="${i}">Copy</button></header>
<textarea readonly id="t${i}">${esc(src)}</textarea></section>`;
}).join('\n');
const html = `<!doctype html><html><head><meta charset="utf-8"><title>Trail Journal paste pack</title>
<meta name="viewport" content="width=device-width,initial-scale=1"><style>
body{font:15px system-ui,sans-serif;margin:0;background:#14082a;color:#f5f0ff}main{max-width:980px;margin:0 auto;padding:20px 16px 80px}
h1{margin:0 0 6px}p{color:#cbb6ef;line-height:1.5}section{border:1px solid #3b2a5a;border-radius:12px;margin:10px 0;background:#1d0f36}
header{display:flex;gap:12px;align-items:center;padding:10px 14px;flex-wrap:wrap}header label{cursor:pointer;font-size:16px}
.k{color:#b09ec8;font-size:13px}.k em{color:#ffe100;font-style:normal;font-weight:700}.n{margin-left:auto;color:#8f7fa8;font-size:12px}
button{font:700 14px system-ui;padding:7px 16px;border-radius:999px;border:0;background:#ffe100;color:#1a0b2e;cursor:pointer}
button.ok{background:#86efac}textarea{display:none}section.done{opacity:.5}#bar{position:sticky;top:0;background:#14082a;padding:10px 0;font-weight:700;z-index:1}
</style></head><body><main><h1>Trail Journal paste pack</h1>
<p>Built ${new Date().toISOString().slice(0, 16).replace('T', ' ')} from the repo's apps-script folder. For each file: click <b>Copy</b>, open the file of the same name in the Apps Script editor, select all (Ctrl+A), paste (Ctrl+V), then tick the box. Files marked NEW don't exist yet: click + next to Files, pick Script or HTML, and type the name exactly as shown (no extension). Press Ctrl+S when all are done.</p>
<div id="bar">Done: <span id="cnt">0</span> / ${files.length}</div>
${rows}</main><script>
function key(k){return 'tj-paste-'+k}
function count(){var b=document.querySelectorAll('input[data-k]');var n=0;b.forEach(function(x){if(x.checked)n++;x.closest('section').classList.toggle('done',x.checked)});document.getElementById('cnt').textContent=n}
document.querySelectorAll('input[data-k]').forEach(function(x){try{x.checked=localStorage.getItem(key(x.dataset.k))==='1'}catch(e){}x.onchange=function(){try{localStorage.setItem(key(x.dataset.k),x.checked?'1':'')}catch(e){}count()}});
document.querySelectorAll('button[data-i]').forEach(function(b){b.onclick=function(){var t=document.getElementById('t'+b.dataset.i);navigator.clipboard.writeText(t.value).then(function(){b.textContent='Copied ✓';b.classList.add('ok');setTimeout(function(){b.textContent='Copy';b.classList.remove('ok')},2500)},function(){t.style.display='block';t.select();document.execCommand('copy');b.textContent='Copied ✓'})}});
count();
</script></body></html>`;
fs.writeFileSync(path.join(__dirname, '..', 'paste-pack.html'), html);
console.log('paste-pack.html: ' + files.length + ' files');
