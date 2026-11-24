/* EMS AI switch: one file, the same in every EMS tool that uses AI.
   Master copy: emsd/ai-switch/AiSwitch.gs. Edit it there, then run `node ai-switch/sync.js`.

   Which AI answers (first match wins):
     1. This tool's own choice, from the Sheet menu "AI: Claude or Gemini" (Script Property AI_PROVIDER_HERE).
     2. This tool's Settings row AI_PROVIDER, when it says CLAUDE or GEMINI (MASTER or blank = no choice).
     3. The master switch: the "EMS AI Switch" Google Sheet (row AI_PROVIDER), read every 5 minutes.
     4. CLAUDE.
   Keys live only in Script Properties, saved from the menu: GEMINI_API_KEY, and CLAUDE_API_KEY
   (ANTHROPIC_API_KEY also works). Never in the Sheet.

   How a tool uses it: build the same Claude Messages request as always and call
   emsAiFetch_(options) instead of UrlFetchApp.fetch(<the Claude Messages URL>, options).
   With Claude on, the request goes to Claude unchanged. With Gemini on, it is translated
   (system, text, images, PDFs, JSON schema output) and Gemini's answer comes back in Claude's
   shape, so the code after the call does not change. */

var EMS_AI_MASTER_SHEET_ID = '1E-MLb26bHgTNE0a4U-PHRVYvyPI7wEKiv8UufocKqrM';   // "EMS AI Switch" in Matt's Drive; the menu can point a tool elsewhere
var EMS_AI_CLAUDE_URL = 'https:\/\/api.anthropic.com/v1/messages';
var EMS_AI_GEMINI_URL = 'https:\/\/generativelanguage.googleapis.com/v1beta/models/';
var EMS_AI_GEMINI_DEFAULT = 'gemini-flash-latest';

function emsAiProps_() { return PropertiesService.getScriptProperties(); }
function emsAiNorm_(v) { var s = String(v === null || v === undefined ? '' : v).trim().toUpperCase(); return s === 'CLAUDE' || s === 'GEMINI' ? s : ''; }
function emsAiSetting_(key) {
  try { if (typeof getSetting_ === 'function') return getSetting_(key); } catch (e) { /* this tool has no such setting */ }
  return '';
}
function emsAiMasterId_() { return String(emsAiProps_().getProperty('AI_MASTER_SHEET_ID') || EMS_AI_MASTER_SHEET_ID || '').trim(); }

/* The master Sheet's key/value rows, cached 5 minutes. If the Sheet can't be read, the last good copy is used. */
function emsAiMaster_() {
  var id = emsAiMasterId_(); if (!id) return {};
  var cache = null; try { cache = CacheService.getScriptCache(); } catch (e) {}
  if (cache) { try { var hit = cache.get('emsAiMaster'); if (hit) return JSON.parse(hit); } catch (e) {} }
  var out = null;
  try {
    var v = SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues();
    out = {};
    v.forEach(function (r) { var k = String(r[0] || '').trim(); if (k && !/^key$/i.test(k)) out[k] = String(r[1] === null || r[1] === undefined ? '' : r[1]).trim(); });
    var json = JSON.stringify(out);
    if (emsAiProps_().getProperty('AI_MASTER_LAST') !== json) emsAiProps_().setProperty('AI_MASTER_LAST', json);
  } catch (e) {
    try { out = JSON.parse(emsAiProps_().getProperty('AI_MASTER_LAST') || '{}'); } catch (e2) { out = {}; }
    out._error = 'Could not open the master switch Sheet: ' + e.message;
  }
  if (cache) { try { cache.put('emsAiMaster', JSON.stringify(out), 300); } catch (e) {} }
  return out;
}

/* 'CLAUDE' or 'GEMINI', and why. */
function emsAiChoice_() {
  var here = emsAiNorm_(emsAiProps_().getProperty('AI_PROVIDER_HERE'));
  if (here) return { provider: here, why: 'this tool is set to ' + here + ' (AI menu)' };
  var setting = emsAiNorm_(emsAiSetting_('AI_PROVIDER'));
  if (setting) return { provider: setting, why: 'this tool\'s Settings row AI_PROVIDER says ' + setting };
  var m = emsAiMaster_(), master = emsAiNorm_(m.AI_PROVIDER);
  if (master) return { provider: master, why: 'the master switch says ' + master + (m._error ? ' (last known; ' + m._error + ')' : '') };
  return { provider: 'CLAUDE', why: emsAiMasterId_() ? 'the master switch is blank, so Claude' + (m._error ? ' (' + m._error + ')' : '') : 'no master switch Sheet is set, so Claude' };
}
function emsAiProvider_() { return emsAiChoice_().provider; }
function emsAiKey_(provider) {
  var p = emsAiProps_();
  return (provider || emsAiProvider_()) === 'GEMINI' ? p.getProperty('GEMINI_API_KEY') || '' : p.getProperty('CLAUDE_API_KEY') || p.getProperty('ANTHROPIC_API_KEY') || '';
}
function emsAiHasKey_() { return !!emsAiKey_(); }
function emsAiGeminiModel_() {
  return String(emsAiSetting_('GEMINI_MODEL') || emsAiMaster_().GEMINI_MODEL || EMS_AI_GEMINI_DEFAULT).trim();
}
/* "Claude (claude-opus-5-5)" or "Gemini (gemini-flash-latest)" for labels like Scored By. */
function emsAiLabel_(claudeModel) {
  return emsAiProvider_() === 'GEMINI' ? 'Gemini (' + emsAiGeminiModel_() + ')' : 'Claude (' + (claudeModel || 'claude') + ')';
}

/* Drop-in for UrlFetchApp.fetch(Claude Messages URL, options). Returns an HTTP-response-like object. */
function emsAiFetch_(options) {
  var opts = {}; Object.keys(options || {}).forEach(function (k) { opts[k] = options[k]; });
  if (emsAiProvider_() !== 'GEMINI') {
    var h = {}; Object.keys(opts.headers || {}).forEach(function (k) { h[k] = opts.headers[k]; });
    if (!h['x-api-key']) h['x-api-key'] = emsAiKey_('CLAUDE');
    if (!h['anthropic-version']) h['anthropic-version'] = '2023-06-01';
    opts.headers = h; opts.method = opts.method || 'post'; opts.contentType = opts.contentType || 'application/json';
    return UrlFetchApp.fetch(EMS_AI_CLAUDE_URL, opts);
  }
  var body = typeof opts.payload === 'string' ? JSON.parse(opts.payload) : (opts.payload || {});
  var fail = function (code, message) {   // like UrlFetchApp: throws unless muteHttpExceptions is on
    if (opts.muteHttpExceptions !== true) throw new Error(message);
    return emsAiResponse_(code, { type: 'error', error: { type: code === 401 ? 'authentication_error' : 'api_error', message: message } });
  };
  var key = emsAiKey_('GEMINI');
  if (!key) return fail(401, 'No Gemini API key saved. In the Sheet: AI: Claude or Gemini > Save Gemini API key.');
  var g;
  try { g = emsAiToGemini_(body); } catch (e) { return fail(400, e.message); }
  var model = emsAiGeminiModel_();
  var res = UrlFetchApp.fetch(EMS_AI_GEMINI_URL + encodeURIComponent(model) + ':generateContent', {
    method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key },
    payload: JSON.stringify(g.body), muteHttpExceptions: true
  });
  var code = res.getResponseCode(), text = res.getContentText();
  if (code !== 200) {
    var msg = text.slice(0, 400); try { var j = JSON.parse(text); if (j.error && j.error.message) msg = j.error.message; } catch (e) {}
    if (opts.muteHttpExceptions !== true) throw new Error('Gemini ' + code + ': ' + msg);
    return emsAiResponse_(code, { type: 'error', error: { type: 'api_error', message: 'Gemini ' + code + ': ' + msg } });
  }
  return emsAiResponse_(200, emsAiFromGemini_(JSON.parse(text), model, g.prefill));
}
function emsAiResponse_(code, obj) {
  var s = JSON.stringify(obj);
  return { getResponseCode: function () { return code; }, getContentText: function () { return s; }, getHeaders: function () { return {}; } };
}

/* Claude Messages request -> Gemini generateContent request. */
function emsAiToGemini_(b) {
  if (b.tools && b.tools.length) throw new Error('This AI feature uses tools, which only Claude supports. Switch this tool to Claude.');
  var sys = typeof b.system === 'string' ? b.system : (b.system || []).map(function (x) { return x.text || ''; }).join('\n\n');
  var msgs = (b.messages || []).slice(), prefill = '';
  var last = msgs[msgs.length - 1];
  if (last && last.role === 'assistant') { prefill = typeof last.content === 'string' ? last.content : (last.content || []).map(function (x) { return x.text || ''; }).join(''); msgs.pop(); }
  var part = function (x) {
    if (typeof x === 'string') return { text: x };
    if (x.type === 'text') return { text: x.text || '' };
    if ((x.type === 'image' || x.type === 'document') && x.source) {
      if (x.source.type === 'base64') return { inlineData: { mimeType: x.source.media_type, data: x.source.data } };
      if (x.source.type === 'text') return { text: x.source.data || '' };
      if (x.source.type === 'url') return { fileData: { fileUri: x.source.url, mimeType: x.source.media_type || (x.type === 'document' ? 'application/pdf' : 'image/jpeg') } };
    }
    if (x.type === 'tool_result' || x.type === 'tool_use') throw new Error('This AI feature uses tools, which only Claude supports. Switch this tool to Claude.');
    return { text: x.text || JSON.stringify(x) };
  };
  var contents = msgs.map(function (m) {
    return { role: m.role === 'assistant' ? 'model' : 'user', parts: (typeof m.content === 'string' ? [m.content] : m.content || []).map(part) };
  });
  if (prefill && contents.length) contents[contents.length - 1].parts.push({ text: 'Begin your reply with exactly: ' + prefill });
  var gc = { maxOutputTokens: Math.max(Number(b.max_tokens) || 8192, 8192) };
  if (b.temperature !== undefined) gc.temperature = b.temperature;
  if (b.top_p !== undefined) gc.topP = b.top_p;
  if (b.stop_sequences) gc.stopSequences = b.stop_sequences;
  var fmt = (b.output_config && b.output_config.format) || b.output_format;
  if (fmt && fmt.type === 'json_schema' && fmt.schema) { gc.responseMimeType = 'application/json'; gc.responseSchema = emsAiGeminiSchema_(fmt.schema); }
  var body = { contents: contents, generationConfig: gc };
  if (sys) body.systemInstruction = { parts: [{ text: sys }] };
  return { body: body, prefill: prefill };
}
/* JSON Schema -> Gemini responseSchema (OpenAPI subset). */
function emsAiGeminiSchema_(s) {
  if (!s || typeof s !== 'object') return { type: 'STRING' };
  var t = s.type, nullable = false;
  if (Array.isArray(t)) { nullable = t.indexOf('null') >= 0; t = t.filter(function (x) { return x !== 'null'; })[0]; }
  if (!t && s.anyOf) { var nn = s.anyOf.filter(function (x) { return x.type !== 'null'; }); var o = emsAiGeminiSchema_(nn[0]); if (nn.length < s.anyOf.length) o.nullable = true; return o; }
  if (!t && s.enum) t = 'string';
  if (!t && s.properties) t = 'object';
  var out = { type: String(t || 'string').toUpperCase() };
  if (nullable) out.nullable = true;
  if (s.description) out.description = s.description;
  if (s.enum) out.enum = s.enum.map(String);
  if (s.properties) { out.properties = {}; Object.keys(s.properties).forEach(function (k) { out.properties[k] = emsAiGeminiSchema_(s.properties[k]); }); out.propertyOrdering = Object.keys(s.properties); }
  if (s.required) out.required = s.required;
  if (s.items) out.items = emsAiGeminiSchema_(s.items);
  if (s.minItems !== undefined) out.minItems = s.minItems;
  if (s.maxItems !== undefined) out.maxItems = s.maxItems;
  return out;
}
/* Gemini answer -> Claude Messages answer. */
function emsAiFromGemini_(g, model, prefill) {
  var usage = { input_tokens: 0, output_tokens: 0 };
  var um = g.usageMetadata || {}; usage.input_tokens = um.promptTokenCount || 0; usage.output_tokens = (um.candidatesTokenCount || 0) + (um.thoughtsTokenCount || 0); if (um.cachedContentTokenCount) usage.cache_read_input_tokens = um.cachedContentTokenCount;
  var base = { id: 'gemini-' + Date.now(), type: 'message', role: 'assistant', model: model, provider: 'GEMINI', usage: usage, stop_sequence: null };
  if (g.promptFeedback && g.promptFeedback.blockReason) { base.content = []; base.stop_reason = 'refusal'; return base; }
  var c = (g.candidates || [])[0] || {};
  var text = ((c.content && c.content.parts) || []).filter(function (p) { return !p.thought; }).map(function (p) { return p.text || ''; }).join('');
  if (prefill && text.indexOf(prefill) === 0) text = text.slice(prefill.length);
  var fr = String(c.finishReason || 'STOP');
  base.stop_reason = fr === 'STOP' ? 'end_turn' : fr === 'MAX_TOKENS' ? 'max_tokens' : /SAFETY|PROHIBITED|BLOCKLIST|SPII|RECITATION|IMAGE_SAFETY/.test(fr) ? 'refusal' : 'end_turn';
  base.content = text ? [{ type: 'text', text: text }] : [];
  return base;
}

/* ---------- Sheet menu ---------- */
/* Add to a tool's onOpen: menu.addSubMenu(emsAiMenu_(ui)), or ui.createMenu(...).addSubMenu(emsAiMenu_(ui)). */
function emsAiMenu_(ui) {
  return (ui || SpreadsheetApp.getUi()).createMenu('AI: Claude or Gemini')
    .addItem('Which AI is on? (and test it)', 'emsAiMenuStatus')
    .addSeparator()
    .addItem('Save Gemini API key', 'emsAiMenuSaveGemini')
    .addItem('Save Claude API key', 'emsAiMenuSaveClaude')
    .addSeparator()
    .addItem('This tool: follow the master switch', 'emsAiMenuFollow')
    .addItem('This tool: always Claude', 'emsAiMenuClaude')
    .addItem('This tool: always Gemini', 'emsAiMenuGemini')
    .addSeparator()
    .addItem('Set the master switch Sheet', 'emsAiMenuSetMaster');
}
/* Menu items have public names (Sheet menus can't call names ending in _), so a web page could try to
   call them through google.script.run. getUi() throws anywhere but the Sheet's own menu, which only
   people who can edit the Sheet see, so every menu item calls this first, before it reads or saves anything. */
function emsAiMenuGuard_() { return SpreadsheetApp.getUi(); }
function emsAiStatusText_() {
  var c = emsAiChoice_(), p = emsAiProps_();
  return 'AI in use: ' + c.provider + (c.provider === 'GEMINI' ? ' (' + emsAiGeminiModel_() + ')' : '') + '\nWhy: ' + c.why +
    '\n\nGemini key saved: ' + (p.getProperty('GEMINI_API_KEY') ? 'yes' : 'no') +
    '\nClaude key saved: ' + (emsAiKey_('CLAUDE') ? 'yes' : 'no') +
    '\nMaster switch Sheet: ' + (emsAiMasterId_() || 'not set');
}
function emsAiTest_() {
  var res = emsAiFetch_({ method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ model: 'claude-haiku-5-5', max_tokens: 50, messages: [{ role: 'user', content: 'Reply with the single word OK.' }] }) });
  var b = {}; try { b = JSON.parse(res.getContentText()); } catch (e) {}
  if (res.getResponseCode() !== 200) return 'Test failed: ' + ((b.error && b.error.message) || ('HTTP ' + res.getResponseCode()));
  return 'Test answer: ' + ((b.content || []).map(function (x) { return x.text || ''; }).join('').trim() || '(empty)');
}
function emsAiMenuStatus() {
  var ui = emsAiMenuGuard_();
  var go = ui.alert('AI: Claude or Gemini', emsAiStatusText_() + '\n\nSend a one-word test message to the AI now? (No student data is sent.)', ui.ButtonSet.YES_NO);
  if (go === ui.Button.YES) ui.alert('AI test', emsAiTest_(), ui.ButtonSet.OK);
}
function emsAiSaveKey_(prop, label) {
  var ui = emsAiMenuGuard_();
  var r = ui.prompt('Save ' + label, 'Paste the key. It is stored in this tool\'s Script Properties only, never in the Sheet. Leave blank and press OK to remove it.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var k = String(r.getResponseText() || '').trim();
  if (k) emsAiProps_().setProperty(prop, k); else emsAiProps_().deleteProperty(prop);
  ui.alert(k ? label + ' saved.' : label + ' removed.');
}
function emsAiMenuSaveGemini() { emsAiMenuGuard_(); emsAiSaveKey_('GEMINI_API_KEY', 'Gemini API key'); }
function emsAiMenuSaveClaude() { emsAiMenuGuard_(); emsAiSaveKey_('CLAUDE_API_KEY', 'Claude API key'); }
function emsAiSetHere_(v) {
  emsAiMenuGuard_();
  if (v) emsAiProps_().setProperty('AI_PROVIDER_HERE', v); else emsAiProps_().deleteProperty('AI_PROVIDER_HERE');
  SpreadsheetApp.getUi().alert(emsAiStatusText_());
}
function emsAiMenuFollow() { emsAiMenuGuard_(); emsAiSetHere_(''); }
function emsAiMenuClaude() { emsAiMenuGuard_(); emsAiSetHere_('CLAUDE'); }
function emsAiMenuGemini() { emsAiMenuGuard_(); emsAiSetHere_('GEMINI'); }
function emsAiMenuSetMaster() {
  var ui = emsAiMenuGuard_();
  var r = ui.prompt('Master switch Sheet', 'Paste the link (or ID) of the "EMS AI Switch" Google Sheet. Blank = use the built-in one.\nNow: ' + (emsAiMasterId_() || 'none'), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var t = String(r.getResponseText() || '').trim(), m = t.match(/\/d\/([a-zA-Z0-9_-]{20,})/), id = m ? m[1] : t;
  if (id) emsAiProps_().setProperty('AI_MASTER_SHEET_ID', id); else emsAiProps_().deleteProperty('AI_MASTER_SHEET_ID');
  try { CacheService.getScriptCache().remove('emsAiMaster'); } catch (e) {}
  ui.alert(emsAiStatusText_());
}
