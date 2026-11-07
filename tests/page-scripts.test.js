// Google strips comments from page scripts before it serves them, carelessly: on 2026-11-06 it cut a
// 'https://…' string at the // in the Behavior Intervention Center and the whole script failed live
// ("Invalid or unexpected token"). So every script here must still parse after the crudest stripper
// (every // and /* */ removed, quotes ignored): write addresses as https:\/\/ and keep /* out of strings.
// Also: include() pastes files raw, so <? ?> tags only work in the page template, never in included files.
//   node tests/page-scripts.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');
const dir = path.join(__dirname, '..', 'apps-script');

const strip = js => js.replace(/\/\*[\s\S]*?\*\//g, m => m.replace(/[^\n]/g, ''))
  .split('\n').map(l => l.replace(/\/\/.*$/, '')).join('\n');
const html = fs.readdirSync(dir).filter(f => f.endsWith('.html'));
let checked = 0;

html.forEach(f => {
  const src = fs.readFileSync(path.join(dir, f), 'utf8');
  [...src.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/g)].forEach((m, i) => {
    if (/\bsrc=/.test(m[1]) || (/type=/.test(m[1]) && !/javascript|module/.test(m[1]))) return;
    const js = strip(m[2].replace(/<\?[\s\S]*?\?>/g, '0'));
    try { new vm.Script(js); checked++; }
    catch (e) {
      const n = Number((e.stack.match(/evalmachine\.<anonymous>:(\d+)/) || [])[1]);
      throw new Error(f + ' script ' + i + ' breaks once comments are stripped, line ' + n + ' of the script: ' + e.message + ' | ' + (js.split('\n')[n - 1] || '').trim().slice(0, 120));
    }
  });
});
console.log('  ok   ' + checked + ' page scripts still parse after comments are stripped');

const included = new Set(html.flatMap(f => [...fs.readFileSync(path.join(dir, f), 'utf8').matchAll(/include\(\s*['"]([^'"]+)['"]\s*\)/g)].map(m => m[1])));
included.forEach(n => {
  const file = path.join(dir, n + '.html');
  if (fs.existsSync(file)) assert.ok(!/<\?/.test(fs.readFileSync(file, 'utf8')), n + '.html is pasted raw by include(), so a <? ?> tag there is never filled in');
});
console.log('  ok   no server tags in the ' + included.size + ' included files');

assert.throws(() => new vm.Script(strip("var a = {v:'https://'};")), /Invalid or unexpected token/, 'the stripper reproduces the live failure');
console.log('  ok   the stripper reproduces the live failure');
