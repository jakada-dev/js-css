// Keep document navigation and browser saves in this file.
const $ = id => document.getElementById(id);
const sides = ['left', 'right'];
let reports = {left:null, right:null}, selected = '', locations = {}, editing = false;
let saveId = '', dirty = false, busy = false, database, tagTarget;
let comparisonVersion = 0, lastComparison;
const status = text => $('status').textContent = text;
function markDirty(value) {
  dirty = value;
  $('state').textContent = value ? 'Unsaved changes' : saveId ? 'Saved locally' : 'Not saved yet';
}
async function run(action) {
  if (busy) return;
  busy = true; document.body.classList.add('busy');
  try { await action(); } catch(error) { status(error.message); }
  finally { busy = false; document.body.classList.remove('busy'); }
}
// Normalize testing letters so 5.3.5a and 5.3.5.a match.
function normalize(text) { return text.trim().toLowerCase().replace(/\s+/g,'').replace(/(\d)\.?([a-z])$/, '$1.$2'); }
function valid(text) { return /^\d{1,2}(?:\.\d{1,2}){1,5}(?:\.[a-z])?$/.test(text); }
function section(text) {
  const match = text.trim().match(/^(\d{1,2}(?:\.\d{1,2}){1,5}(?:\s*\.?\s*[a-z](?=\s|[):]|$))?)(?=\s|[):]|$)/i);
  return match ? normalize(match[1]) : '';
}
// Include child subsections, stopping at the next sibling, parent, or repeated heading.
function sectionEnd(tags, id, start, count) {
  const next = tags.filter(tag => tag.block > start && !tag.id.startsWith(id + '.'));
  return Math.min(count, ...next.map(tag => tag.block));
}
// Allow only document formatting. This also makes imported backups safe to display.
function clean(html) {
  const source = new DOMParser().parseFromString(html, 'text/html');
  const allowed = 'P H1 H2 H3 H4 H5 H6 TABLE TBODY THEAD TR TD TH BR STRONG EM UL OL LI BLOCKQUOTE'.split(' ');
  function copy(node) {
    if (node.nodeType === 3) return document.createTextNode(node.textContent);
    const out = document.createDocumentFragment();
    if (node.nodeType !== 1 || ['SCRIPT','STYLE','OBJECT','IFRAME','IMG'].includes(node.tagName)) return out;
    const target = allowed.includes(node.tagName) ? document.createElement(node.tagName) : out;
    if (['TD','TH'].includes(node.tagName)) {
      for (const key of ['colspan','rowspan']) {
        const span = Number(node.getAttribute(key));
        if (span > 0 && span <= 100) target.setAttribute(key, span);
      }
    }
    for (const child of node.childNodes) target.append(copy(child));
    return target;
  }
  const result = document.createElement('div');
  for (const node of source.body.childNodes) result.append(copy(node));
  return result.innerHTML;
}
function blocks(root) {
  return [...root.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li,td,th,blockquote')].filter(node => !node.querySelector('p,h1,h2,h3,h4,h5,h6,li,td,th,blockquote'));
}
async function readWord(bytes) {
  const response = await fetch('/read-word', {method:'POST', body:bytes});
  const result = await response.json();
  if (!response.ok) throw new Error(result.error);
  return clean(result.html);
}
// FUTURE PDF: branch here by file extension and return HTML plus section locations.
// A PDF reader/renderer is needed; Python's standard library cannot extract PDF text.
async function load(side, file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.docx')) throw new Error('Choose a Word .docx file. PDF support is planned for later.');
  if (file.size > 150 * 1024 * 1024) throw new Error('Choose a file smaller than 150 MB.');
  if (reports[side] && !confirm('Replace this report and its tags?')) return;
  status('Reading ' + file.name + '…');
  const html = await readWord(await file.arrayBuffer());
  const root = document.createElement('div'); root.innerHTML = html;
  const tags = blocks(root).map((node, block) => ({id:section(node.textContent),block})).filter(tag => tag.id);
  reports[side] = {name:file.name, html, tags}; locations = {};
  markDirty(true); draw(); status('Found ' + tags.length + ' section occurrences. Review automatic tags; Edit tags lets you correct them.');
}
function select(id) { selected = id; locations = {}; draw(); status('Showing section ' + id + ' in each report that has a matching tag.'); }
function draw() {
  const ids = [...new Set(sides.flatMap(side => reports[side]?.tags.map(tag => tag.id) || []))].sort((a,b) => a.localeCompare(b,undefined,{numeric:true}));
  $('sections').replaceChildren();
  for (const id of ids.filter(id => id.includes(normalize($('search').value)))) {
    const button = document.createElement('button'); button.textContent = id;
    const both = sides.map(side => reports[side]?.tags.some(tag => tag.id === id));
    const badge = document.createElement('small'); badge.textContent = both.every(Boolean) ? 'Both' : both[0] ? 'A only' : 'B only';
    button.append(badge); button.classList.toggle('selected', id === selected); button.onclick = () => select(id); $('sections').append(button);
  }
  if (!$('sections').children.length) $('sections').textContent = ids.length ? 'No matching sections.' : 'Add a report to see its sections.';
  for (const side of sides) drawReport(side);
  highlightDifferences();
}
// FUTURE PDF: add a separate page viewer here. PDF tags will need page/position fields.
function drawReport(side) {
  const report = reports[side], panel = $(side), letter = side === 'left' ? 'A' : 'B';
  panel.innerHTML = '<div class="report-head"><strong></strong><label class="file">Choose file<input type="file" accept=".docx"></label><span class="note"></span></div><div class="reader"></div>';
  panel.querySelector('strong').textContent = 'Report ' + letter + (report ? ' — ' + report.name : '');
  panel.querySelector('input').onchange = event => run(() => load(side,event.target.files[0]));
  const reader = panel.querySelector('.reader'), note = panel.querySelector('.note');
  if (!report) { reader.innerHTML = '<p class="empty">Choose a Word (.docx) report above.<br>It stays on this computer.</p>'; return; }
  const matches = report.tags.filter(tag => tag.id === selected), active = matches[locations[side] || 0];
  note.textContent = editing ? 'Click a paragraph or table cell to add, correct, or remove its tag.' : selected ? active ? 'Section ' + selected : 'No tag for ' + selected + ' in this report.' : 'Choose a section from the list.';
  if (matches.length > 1) {
    const choose = document.createElement('select'); choose.setAttribute('aria-label','Section occurrence in report ' + letter);
    matches.forEach((tag,index) => choose.add(new Option('Occurrence ' + (index+1), index)));
    choose.value = locations[side] || 0; choose.onchange = () => { locations[side] = Number(choose.value); draw(); };
    panel.querySelector('.report-head').append(choose);
  }
  const page = document.createElement('div'); page.className = 'page'; page.innerHTML = report.html;
  page.style.width = Math.max(280, reader.clientWidth-30, Number($('width').value)) + 'px';
  page.style.paddingBottom = Math.max(25,reader.clientHeight-90) + 'px'; reader.append(page);
  const items = blocks(page);
  items.forEach((node,index) => { node.dataset.block = index; });
  for (const tag of report.tags) {
    const node = items[tag.block]; if (!node) continue;
    node.classList.add('tagged'); node.title = 'Section ' + tag.id;
    if (tag === active) { node.classList.add('selected'); reader.scrollTop = node.getBoundingClientRect().top-reader.getBoundingClientRect().top+reader.scrollTop-45; }
  }
  page.classList.toggle('editing', editing);
  page.onclick = event => {
    const node = event.target.closest('[data-block]'); if (!node || busy) return;
    const block = Number(node.dataset.block), old = report.tags.find(tag => tag.block === block);
    if (!editing) { if (old) select(old.id); return; }
    tagTarget = {side,block};
    $('tag-location').textContent = 'Report ' + letter + ', text block ' + (block+1);
    $('tag-number').value = old?.id || selected;
    $('tag-number').setCustomValidity('');
    $('tag-dialog').showModal(); $('tag-number').focus();
  };
}
// Collect all text in the chosen section, including its table cells and evidence.
function selectedSection(side) {
  const report = reports[side];
  const tag = report?.tags.filter(tag => tag.id === selected)[locations[side] || 0];
  if (!tag) return null;
  const items = [...$(side).querySelectorAll('[data-block]')];
  const end = sectionEnd(report.tags, selected, tag.block, items.length);
  const result = {text: '', parts: []};
  for (let block = tag.block; block < end; block++) {
    const walker = document.createTreeWalker(items[block], NodeFilter.SHOW_TEXT);
    let node;
    while ((node = walker.nextNode())) {
      const start = result.text.length;
      result.text += node.textContent;
      result.parts.push({node, start, end: result.text.length});
    }
    result.text += '\n';
  }
  return result;
}
// Wrap changed text without replacing table structure or inserting report HTML.
function markChanges(section, ranges, side) {
  let rangeIndex = 0;
  for (const part of section.parts) {
    while (rangeIndex < ranges.length && ranges[rangeIndex][1] <= part.start) rangeIndex++;
    const text = part.node.textContent, fragment = document.createDocumentFragment();
    let position = 0;
    for (let index = rangeIndex; index < ranges.length && ranges[index][0] < part.end; index++) {
      const start = Math.max(0, ranges[index][0] - part.start);
      const end = Math.min(text.length, ranges[index][1] - part.start);
      if (start >= end) continue;
      fragment.append(document.createTextNode(text.slice(position, start)));
      const mark = document.createElement('mark');
      mark.className = 'diff-' + side;
      mark.title = side === 'left' ? 'Different or only present in report A' : 'Different or only present in report B';
      mark.textContent = text.slice(start, end); fragment.append(mark);
      position = end;
    }
    if (position) {
      fragment.append(document.createTextNode(text.slice(position)));
      part.node.replaceWith(fragment);
    }
  }
}
async function highlightDifferences() {
  const version = ++comparisonVersion, summary = $('diff-summary');
  summary.hidden = !selected;
  if (!selected) return;
  if (editing) { summary.textContent = 'Finish editing tags to see section differences.'; return; }
  const left = selectedSection('left'), right = selectedSection('right');
  if (!left || !right) { summary.textContent = 'Differences unavailable: this section needs a matching tag in both reports.'; return; }
  summary.textContent = 'Comparing section text…';
  try {
    let changes;
    if (lastComparison?.left === left.text && lastComparison?.right === right.text) {
      changes = lastComparison.changes;
    } else {
      const response = await fetch('/compare-sections', {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({left: left.text, right: right.text})
      });
      changes = await response.json();
      if (!response.ok) throw new Error(changes.error || 'Section comparison failed.');
    }
    // Ignore a slow response if the user has already selected another section.
    if (version !== comparisonVersion) return;
    lastComparison = {left: left.text, right: right.text, changes};
    markChanges(left, changes.left, 'left'); markChanges(right, changes.right, 'right');
    summary.textContent = '';
    if (!changes.left.length && !changes.right.length) {
      summary.textContent = 'No text differences in this section. Spacing and formatting are ignored.';
    } else {
      const a = document.createElement('span'), b = document.createElement('span');
      a.className = 'diff-left'; a.textContent = 'Report A differences';
      b.className = 'diff-right'; b.textContent = 'Report B differences';
      summary.append(a, '  ', b, ' · Comparing the selected section and its child subsections.');
    }
  } catch (error) {
    if (version === comparisonVersion) summary.textContent = 'Could not compare this section: ' + error.message;
  }
}
// The same dialog adds a new tag, corrects a tag, or removes a blank tag.
$('tag-number').oninput = () => $('tag-number').setCustomValidity('');
$('cancel-tag').onclick = () => $('tag-dialog').close();
$('tag-form').onsubmit = event => {
  event.preventDefault();
  const id = normalize($('tag-number').value), {side,block} = tagTarget;
  if (id && !valid(id)) {
    $('tag-number').setCustomValidity('Use a number such as 5.3.5.a.');
    $('tag-number').reportValidity(); return;
  }
  reports[side].tags = reports[side].tags.filter(tag => tag.block !== block);
  if (id) { reports[side].tags.push({id,block}); selected = id; }
  $('tag-dialog').close(); markDirty(true); draw();
  status('Tag updated. Save the comparison to keep it.');
};
// Browser storage holds readable report content and tags, not the original Word files.
function store(mode, action) {
  return new Promise((resolve,reject) => {
    if (!database) { reject(new Error('Browser storage unavailable. Export a backup instead.')); return; }
    const transaction = database.transaction('comparisons',mode), request = action(transaction.objectStore('comparisons'));
    transaction.oncomplete = () => resolve(request.result);
    transaction.onerror = transaction.onabort = () => reject(transaction.error || new Error('Storage failed.'));
  });
}
function snapshot() {
  if (!sides.some(side => reports[side])) throw new Error('Load a Word report first.');
  return {version:2, id:saveId || crypto.randomUUID(), name:$('name').value.trim() || 'Untitled comparison', updated:Date.now(), selected, reports};
}
async function restore(record) {
  if (![1,2].includes(record?.version) || !record.reports) throw new Error('Invalid comparison backup.');
  const next = {left:null,right:null};
  for (const side of sides) {
    const report = record.reports[side]; if (!report) continue;
    // Older saves contained original files; convert Word saves when opening them.
    let html = report.html;
    if (!html && report.bytes) {
      if (!report.name?.toLowerCase().endsWith('.docx')) throw new Error('This old comparison contains a PDF. Only Word is supported now.');
      const bytes = typeof report.bytes === 'string' ? Uint8Array.from(atob(report.bytes),c => c.charCodeAt(0)) : report.bytes;
      html = await readWord(bytes);
    }
    if (typeof html !== 'string' || typeof report.name !== 'string' || !Array.isArray(report.tags)) throw new Error('Invalid report data.');
    html = clean(html); const root = document.createElement('div'); root.innerHTML = html; const count = blocks(root).length;
    const tags = report.tags.map(tag => ({id:normalize(String(tag.id)), block:tag.block}));
    if (tags.some(tag => !valid(tag.id) || !Number.isInteger(tag.block) || tag.block < 0 || tag.block >= count)) throw new Error('Invalid section tags.');
    next[side] = {name:report.name,html,tags};
  }
  reports = next; saveId = typeof record.id === 'string' ? record.id : ''; selected = String(record.selected || ''); locations = {};
  $('name').value = record.name || ''; markDirty(false); draw(); status('Comparison opened.');
}
function discard() { return !dirty || confirm('Continue without saving your changes?'); }
$('search').oninput = draw; $('name').oninput = () => markDirty(true);
$('width').onchange = () => { draw(); status('Scroll left and right inside each report to see the full width.'); };
$('edit').onclick = () => { editing = !editing; $('edit').setAttribute('aria-pressed',editing); draw(); status(editing ? 'Click text inside a report to edit its section tag.' : 'Tag editing finished.'); };
for (const button of document.querySelectorAll('button[data-view]')) {
  button.onclick = () => { $('workspace').dataset.view = button.dataset.view; document.querySelectorAll('button[data-view]').forEach(node => node.setAttribute('aria-pressed',node === button)); draw(); };
}
$('full').onclick = () => {
  const reading = document.body.classList.toggle('reading'); $('full').textContent = reading ? 'Exit full screen' : 'Full screen'; draw();
  if (reading && document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(() => {});
  else if (document.fullscreenElement) document.exitFullscreen();
};
function exitReading() { document.body.classList.remove('reading'); $('full').textContent = 'Full screen'; }
document.onfullscreenchange = () => { if (!document.fullscreenElement) exitReading(); };
document.onkeydown = event => { if (event.key === 'Escape') exitReading(); };
// Save and reopen comparisons in this browser.
$('save').onclick = () => run(async () => {
  const record = snapshot();
  await store('readwrite', store => store.put(record));
  saveId = record.id; markDirty(false);
  status('Saved locally. Export a backup for a separate copy.');
});
$('show-saved').onclick = () => run(async () => {
  const records = await store('readonly', store => store.getAll());
  $('saved').replaceChildren();
  records.sort((a,b) => b.updated-a.updated).forEach(record => {
    $('saved').add(new Option(record.name,record.id));
  });
  $('saved-dialog').showModal();
});
$('cancel').onclick = () => $('saved-dialog').close();
$('open').onclick = () => run(async () => {
  if (!$('saved').value || !discard()) return;
  await restore(await store('readonly', store => store.get($('saved').value)));
  $('saved-dialog').close();
});
$('delete').onclick = () => run(async () => {
  const id = $('saved').value;
  if (!id || !confirm('Delete this saved comparison?')) return;
  await store('readwrite', store => store.delete(id));
  if (id === saveId) { saveId = ''; markDirty(true); }
  $('saved').selectedOptions[0].remove();
  status('Saved comparison deleted.');
});
$('new').onclick = () => {
  if (!discard()) return;
  reports = {left:null,right:null}; selected = ''; locations = {}; saveId = '';
  $('name').value = ''; $('search').value = ''; markDirty(false); draw();
  status('Choose two Word reports to begin.');
};
// Backups contain the readable report content and the corrected section tags.
$('export').onclick = () => run(async () => {
  const record = snapshot();
  const url = URL.createObjectURL(new Blob([JSON.stringify(record)],{type:'application/json'}));
  const link = document.createElement('a');
  link.href = url; link.download = 'comparison.json'; link.click();
  setTimeout(() => URL.revokeObjectURL(url),1000);
  $('more').open = false; status('Backup exported.');
});
$('import').onchange = event => run(async () => {
  const file = event.target.files[0]; event.target.value = '';
  if (!file || !discard()) return;
  if (file.size > 150*1024*1024) throw new Error('Backup is too large.');
  await restore(JSON.parse(await file.text()));
  saveId = ''; markDirty(true); $('more').open = false;
  status('Backup imported. Save it to keep it in this browser.');
});
window.onbeforeunload = event => { if (dirty) { event.preventDefault(); event.returnValue = ''; } };
// Reflow the current documents after the browser size changes.
let resizeTimer; window.onresize = () => { clearTimeout(resizeTimer); resizeTimer = setTimeout(() => { if (!busy) draw(); },200); };
const request = indexedDB.open('report-compare',1);
request.onupgradeneeded = () => request.result.createObjectStore('comparisons',{keyPath:'id'});
request.onsuccess = () => { database = request.result; };
request.onerror = () => status('Browser storage is unavailable. Use Export backup to keep your work.');
draw();
