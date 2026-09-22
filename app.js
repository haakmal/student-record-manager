import {
  normalize,
  parseFrontmatter,
  studentIdFromMarkdown,
  studentNameFromMarkdown,
  addTagPreservingCurrentFormat,
  hasTag,
  courseBlock,
  insertCourseBeforeGeneralFeedback,
  makeNewStudentMarkdown,
  proposedExistingChange,
  classifyRecord
} from './core.js';

const RECORD_STATE_LABELS = Object.freeze({
  new: 'New student',
  change: 'New course',
  recorded: 'Already recorded',
  review: 'Review'
});

const state = {
  xlsxFile: null,
  studentDir: null,
  tutorDir: null,
  tutorRecords: [],
  moodleStudents: [],
  groupMap: new Map(),
  records: [],
  filter: 'all',
  selectedRecordId: null,
};

const $ = (id) => document.getElementById(id);

function setMessage(message = '', type = '') {
  const el = $('setupMessage');
  el.textContent = message;
  el.className = `inline-message ${type}`;
}

function setSessionStatus(text, cls = '') {
  $('sessionStatus').textContent = text;
  $('sessionStatus').className = `status-pill ${cls}`.trim();
  $('sessionState').textContent = text;
  $('sessionState').className = `status-pill ${cls}`.trim();
}

function recordStateLabel(record) {
  if (record.applied) return 'Applied';
  if (record.rejected) return 'Rejected';
  return RECORD_STATE_LABELS[record.state] || record.state;
}

function recordStateClass(record) {
  if (record.applied) return 'applied';
  if (record.rejected) return 'rejected';
  return record.state;
}

function supportsFSAccess() {
  return typeof window.showDirectoryPicker === 'function';
}

function displayName(handle) {
  return handle?.name || 'Selected';
}

async function readAllMarkdownFiles(dirHandle) {
  const results = [];
  async function walk(handle, relativePath = '') {
    for await (const [name, entry] of handle.entries()) {
      const path = relativePath ? `${relativePath}/${name}` : name;
      if (entry.kind === 'directory') {
        await walk(entry, path);
      } else if (/\.md$/i.test(name)) {
        const file = await entry.getFile();
        results.push({ name, path, handle: entry, content: await file.text(), size: file.size, modified: file.lastModified });
      }
    }
  }
  await walk(dirHandle);
  return results;
}

async function chooseStudentDir() {
  if (!supportsFSAccess()) return compatibilityFallback('folder');
  try {
    state.studentDir = await window.showDirectoryPicker({ mode: 'readwrite' });
    $('studentDirName').textContent = displayName(state.studentDir);
    setMessage('Student directory selected.');
    refreshScanAvailability();
  } catch (error) {
    if (error.name !== 'AbortError') setMessage(`Could not select the student directory: ${error.message}`, 'error');
  }
}

async function chooseTutorDir() {
  if (!supportsFSAccess()) return compatibilityFallback('folder');
  try {
    state.tutorDir = await window.showDirectoryPicker({ mode: 'read' });
    $('tutorDirName').textContent = displayName(state.tutorDir);
    state.tutorRecords = await loadTutorRecords(state.tutorDir);
    setMessage(`${state.tutorRecords.length} tutor records found.`);
    refreshScanAvailability();
  } catch (error) {
    if (error.name !== 'AbortError') setMessage(`Could not select the tutor directory: ${error.message}`, 'error');
  }
}

function compatibilityFallback(kind) {
  $('compatNotice').classList.remove('hidden');
  $('compatNotice').className = 'notice warning';
  $('compatNotice').textContent = `This browser does not expose the folder ${kind} API required by this version. Use Chrome or Edge for the local folder workflow.`;
}

async function loadTutorRecords(dirHandle) {
  const files = await readAllMarkdownFiles(dirHandle);
  const tutors = [];
  for (const file of files) {
    const name = studentNameFromMarkdown(file.name, file.content);
    if (name) tutors.push({ name, path: file.path, content: file.content });
  }
  return [...new Map(tutors.map(t => [t.name, t])).values()].sort((a, b) => a.name.localeCompare(b.name));
}

function refreshScanAvailability() {
  const ready = !!state.xlsxFile && !!state.studentDir && !!state.tutorDir && normalize($('courseInput').value) && normalize($('yearInput').value) && normalize($('termInput').value);
  $('scanBtn').disabled = !ready;
  setSessionStatus(ready ? 'Ready to scan' : 'Setup required', ready ? 'ok' : '');
}

function parseMoodleFile(file) {
  return new Promise((resolve, reject) => {
    if (!window.XLSX) return reject(new Error('The XLSX parser did not load. Check your internet connection or bundled library setup.'));
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const workbook = window.XLSX.read(reader.result, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = window.XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
        resolve(rows);
      } catch (error) { reject(error); }
    };
    reader.onerror = () => reject(reader.error || new Error('Could not read XLSX file.'));
    reader.readAsArrayBuffer(file);
  });
}

function normalizeMoodleRow(row, index) {
  const get = (...names) => {
    const key = Object.keys(row).find(k => names.some(name => k.trim().toLowerCase() === name.toLowerCase()));
    return key ? normalize(row[key]) : '';
  };
  const first = get('First name', 'First Name', 'firstname');
  const last = get('Last name', 'Last Name', 'lastname');
  const id = get('Username', 'username', 'ID', 'Student ID');
  const email = get('Email address', 'Email Address', 'email', 'Email');
  const group = get('Groups', 'Group', 'groups');
  return { rowNumber: index + 2, first, last, id, email, group, name: [first, last].filter(Boolean).join(' ').trim() || id };
}

async function runScan() {
  try {
    setMessage('Scanning Moodle and student records. No files will be written.');
    const rows = await parseMoodleFile(state.xlsxFile);
    state.moodleStudents = rows.map(normalizeMoodleRow);
    const existingFiles = await readAllMarkdownFiles(state.studentDir);
    const existingById = new Map();
    for (const file of existingFiles) {
      const id = studentIdFromMarkdown(file.name, file.content);
      if (id) existingById.set(id, file);
    }
    state._existingById = existingById;
    const groups = [...new Set(state.moodleStudents.map(s => s.group).filter(Boolean))].sort((a, b) => a.localeCompare(b));
    state.groupMap = new Map(groups.map(g => [g, state.groupMap.get(g) || '']));
    renderTutorMapping(groups);
    $('mappingPanel').classList.remove('hidden');
    $('mappingPanel').open = true;
    $('resultsPanel').classList.add('hidden');
    $('mappingStatus').textContent = `${groups.length} group${groups.length === 1 ? '' : 's'}`;
    setSessionStatus(`${state.moodleStudents.length} records found`, 'ok');
    setMessage('Assign the tutors for the detected Moodle groups before building the review queue.');
  } catch (error) {
    setMessage(`Scan failed: ${error.message}`, 'error');
    setSessionStatus('Scan error', 'danger');
  }
}

function renderTutorMapping(groups) {
  const wrap = $('mappingTable');
  if (!groups.length) {
    wrap.innerHTML = '<div class="empty-state">No Moodle groups were found in the export.</div>';
    return;
  }
  const options = ['<option value="">— Unassigned —</option>', ...state.tutorRecords.map(t => `<option value="${escapeHtml(t.name)}">${escapeHtml(t.name)}</option>`)].join('');
  wrap.innerHTML = groups.map(group => {
    const current = state.groupMap.get(group) || '';
    return `<div class="mapping-row">
      <div><div class="mapping-group">${escapeHtml(group)}</div><div class="mapping-note">Moodle group</div></div>
      <label><span class="sr-only">Tutor for ${escapeHtml(group)}</span><select data-group="${escapeAttr(group)}">${options}</select></label>
      <div class="mapping-note">${current ? 'Assigned' : 'Needs assignment'}</div>
    </div>`;
  }).join('');
  for (const select of wrap.querySelectorAll('select')) {
    const group = select.dataset.group;
    select.value = state.groupMap.get(group) || '';
    select.addEventListener('change', () => {
      state.groupMap.set(group, select.value);
      const note = select.closest('.mapping-row').querySelector('.mapping-note:last-child');
      note.textContent = select.value ? 'Assigned' : 'Needs assignment';
    });
  }
}

async function prepareQueue() {
  const course = normalize($('courseInput').value);
  const year = normalize($('yearInput').value);
  const term = normalize($('termInput').value).toUpperCase();
  const level = normalize($('levelInput').value) || 'UG';
  const tag = `${course}/${year}/${term}`;
  const result = [];

  for (const student of state.moodleStudents) {
    const existing = state._existingById.get(student.id);
    const tutor = state.groupMap.get(student.group) || '';
    let classification = classifyRecord({ existingMarkdown: existing?.content, course, tag, tutor });
    let proposed = null;
    let error = '';
    try {
      if (classification.state === 'new') {
        if (!tutor) error = 'Assign a tutor to this Moodle group before applying.';
        proposed = makeNewStudentMarkdown({ id: student.id, name: student.name, email: student.email, tag, level, tutor: tutor || 'UNASSIGNED', course });
      } else if (classification.state === 'change') {
        proposed = proposedExistingChange(existing.content, { course, tutor, tag });
      } else if (classification.state === 'review' && existing?.content) {
        if (tutor && /^## General Feedback\s*$/m.test(existing.content)) {
          proposed = proposedExistingChange(existing.content, { course, tutor, tag });
        }
      }
    } catch (err) {
      error = err.message;
      classification = { state: 'review', reason: err.message };
    }
    result.push({
      key: `${student.id}:${tag}`,
      student,
      tutor,
      course,
      year,
      term,
      tag,
      existing,
      state: classification.state,
      reason: error || classification.reason,
      current: existing?.content || '',
      proposed,
      applied: false,
      rejected: false,
      fileHandle: existing?.handle || null,
      fileSnapshot: existing?.content || null,
      fileModified: existing?.modified || null,
      error
    });
  }

  state.records = result;
  state.filter = 'all';
  $('sessionLabel').textContent = `${course} / ${year} / ${term} · ${level}`;
  setSessionStatus(`${state.records.length} ready`, 'ok');
  document.querySelectorAll('.filter').forEach(b => b.classList.toggle('active', b.dataset.filter === 'all'));
  $('mappingPanel').classList.remove('hidden');
  $('resultsPanel').classList.remove('hidden');
  $('setupPanel').open = false;
  $('mappingPanel').open = false;
  $('resultsPanel').open = true;
  renderStats();
  renderRecords();
  $('resultsPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderStats() {
  const counts = { all: state.records.length, new:0, change:0, recorded:0, review:0, applied:0 };
  for (const r of state.records) counts[r.applied ? 'applied' : r.state] = (counts[r.applied ? 'applied' : r.state] || 0) + 1;
  $('resultStatus').textContent = `${state.records.length} total`;
  $('stats').innerHTML = [
    ['all', counts.all, 'Total'],
    ['new', counts.new, 'New students'],
    ['change', counts.change, 'New course'],
    ['recorded', counts.recorded, 'Already recorded'],
    ['review', counts.review, 'Review'],
    ['applied', counts.applied, 'Applied']
  ].map(([k,v,l]) => `<div class="stat ${k}"><span class="stat-value">${v}</span><span class="stat-label">${l}</span></div>`).join('');
}

function renderRecords() {
  const query = normalize($('searchInput').value).toLowerCase();
  let records = state.records.filter(r => {
    const stateFilter = state.filter === 'all' || (state.filter === 'applied' ? r.applied : r.state === state.filter && !r.applied);
    const textMatch = !query || `${r.student.id} ${r.student.name} ${r.student.email}`.toLowerCase().includes(query);
    return stateFilter && textMatch;
  });

  if (!records.length) {
    $('recordList').innerHTML = '<div class="empty-state">No records match this filter.</div>';
    return;
  }

  $('recordList').innerHTML = records.map(renderRecordCard).join('');
  for (const button of $('recordList').querySelectorAll('[data-preview]')) {
    button.addEventListener('click', (event) => {
      event.stopPropagation();
      openPreview(button.dataset.preview);
    });
  }
  document.querySelectorAll('.record-card').forEach(card => card.classList.toggle('selected', card.dataset.recordKey === state.selectedRecordId));
}

function renderRecordCard(record) {
  const disabled = !record.proposed || (record.state === 'review' && !record.tutor) || record.error;
  const actionLabel = record.applied ? 'Applied' : record.state === 'recorded' ? 'No change' : disabled ? 'Review issue' : 'Preview';
  const stateLabel = recordStateLabel(record);
  const stateClass = recordStateClass(record);
  return `<article class="record-card ${stateClass}" data-record-key="${escapeAttr(record.key)}">
    <div class="record-primary">
      <strong>${escapeHtml(record.student.name || record.student.id)}</strong>
      <span class="record-id">${escapeHtml(record.student.id)}</span>
      <div class="record-state ${stateClass}">${escapeHtml(stateLabel)}</div>
    </div>
    <div class="record-info">
      <div><span>Course:</span> ${escapeHtml(record.tag)}</div>
      <div><span>Group:</span> ${escapeHtml(record.student.group || '—')}</div>
      <div><span>Tutor:</span> ${escapeHtml(record.tutor || 'Unassigned')}</div>
      <div><span>Reason:</span> ${escapeHtml(record.reason || '—')}</div>
    </div>
    <div class="record-actions">
      <button class="button secondary" data-preview="${escapeAttr(record.key)}" ${disabled ? '' : ''}>${actionLabel}</button>
    </div>
  </article>`;
}

function getSelectedRecord(key) {
  return state.records.find(r => r.key === key);
}

function openPreview(key) {
  const record = getSelectedRecord(key);
  if (!record) return;
  state.selectedRecordId = key;
  $('previewTitle').textContent = record.student.name || record.student.id;
  $('previewSubtitle').textContent = `${record.student.id} · ${record.tag}`;
  $('previewMeta').innerHTML = [
    ['State', recordStateLabel(record)],
    ['Group', record.student.group || '—'],
    ['Tutor', record.tutor || 'Unassigned'],
    ['Email', record.student.email || '—']
  ].map(([k,v]) => `<div><strong>${escapeHtml(k)}:</strong> ${escapeHtml(v)}</div>`).join('');
  $('currentPreview').innerHTML = formatMarkdown(record.current || '[ No existing student file ]');
  $('proposedPreview').innerHTML = formatMarkdown(record.proposed || '[ No proposal available ]');
  $('changeSummary').textContent = record.state === 'new'
    ? `This will create ${record.year}/${record.student.id}.md with the current course record.`
    : record.state === 'recorded'
      ? 'No file change is proposed because the course session tag already exists.'
      : record.rejected
        ? 'This proposal has been rejected for this session. No file was changed.'
        : `This will preserve the existing note and add the ${record.tag} course instance before General Feedback.`;

  const warning = $('previewWarning');
  const blocked = !record.proposed || !!record.error || (!record.tutor && record.state !== 'recorded');
  if (blocked && record.state !== 'recorded' && !record.rejected) {
    warning.classList.remove('hidden');
    warning.className = 'notice warning';
    warning.textContent = record.error || 'This record cannot be applied until the review issue is resolved.';
  } else {
    warning.classList.add('hidden');
  }
  $('applyBtn').disabled = blocked || record.applied || record.state === 'recorded';
  $('rejectBtn').disabled = record.state === 'recorded' || record.applied || record.rejected;
  $('previewEmpty').classList.add('hidden');
  $('previewContent').classList.remove('hidden');
  document.querySelectorAll('.record-card').forEach(card => card.classList.toggle('selected', card.dataset.recordKey === key));
}

async function writeBackup(fileHandle, studentId, content) {
  const backupDir = await state.studentDir.getDirectoryHandle('__student-record-manager-backups', { create: true });
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const backupHandle = await backupDir.getFileHandle(`${studentId}-${stamp}.md`, { create: true });
  const writable = await backupHandle.createWritable();
  await writable.write(content);
  await writable.close();
}

async function applySelected() {
  const record = getSelectedRecord(state.selectedRecordId);
  if (!record || !record.proposed || record.state === 'recorded') return;
  try {
    if (record.existing?.handle) {
      const latest = await record.existing.handle.getFile();
      const latestText = await latest.text();
      if (latestText !== record.fileSnapshot) {
        throw new Error('This file changed after the preview was built. Re-scan the session before applying.');
      }
      await writeBackup(record.existing.handle, record.student.id, latestText);
      const writable = await record.existing.handle.createWritable();
      await writable.write(record.proposed);
      await writable.close();
      record.current = record.proposed;
      record.applied = true;
      record.state = 'change';
    } else {
      const latestFiles = await readAllMarkdownFiles(state.studentDir);
      const existingNow = latestFiles.find(file => studentIdFromMarkdown(file.name, file.content) === record.student.id);
      if (existingNow) throw new Error('A student file with this ID now exists. Re-scan before applying.');

      const yearFolder = normalize(record.year);
      if (!yearFolder || /[\\/:*?"<>|]/.test(yearFolder) || yearFolder === '.' || yearFolder === '..') {
        throw new Error('The current year cannot be used as a valid student-folder name. Re-scan with a valid year.');
      }
      const yearDir = await state.studentDir.getDirectoryHandle(yearFolder, { create: true });
      const filename = `${record.student.id}.md`;
      const newHandle = await yearDir.getFileHandle(filename, { create: true });
      const writable = await newHandle.createWritable();
      await writable.write(record.proposed);
      await writable.close();
      record.applied = true;
      record.fileHandle = newHandle;
      record.current = record.proposed;
      record.targetPath = `${yearFolder}/${filename}`;
    }
    $('previewContent').classList.remove('hidden');
    renderStats();
    renderRecords();
    openPreview(record.key);
    setMessage(`Applied ${record.student.id} only. No other student file was changed.`);
  } catch (error) {
    $('previewWarning').classList.remove('hidden');
    $('previewWarning').className = 'notice error';
    $('previewWarning').textContent = `Nothing was written: ${error.message}`;
  }
}

function rejectSelected() {
  const record = getSelectedRecord(state.selectedRecordId);
  if (!record || record.state === 'recorded' || record.applied) return;
  record.rejected = true;
  record.state = 'review';
  record.reason = 'Rejected in this session.';
  renderStats();
  renderRecords();
  openPreview(record.key);
}

function resetSession() {
  window.location.reload();
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function escapeAttr(value) {
  return escapeHtml(value).replace(/`/g, '&#96;');
}

function formatMarkdown(value) {
  const text = String(value ?? '');
  return text.split(/\r?\n/).map(line => {
    const escaped = escapeHtml(line);
    if (line === '---') return `<span class="syntax-delim">${escaped}</span>`;
    if (/^#{1,6}\s/.test(line)) return `<span class="syntax-heading">${escaped}</span>`;
    if (/^>\s*\[!/.test(line)) return `<span class="syntax-callout">${escaped}</span>`;
    const keyMatch = escaped.match(/^([A-Za-z][A-Za-z0-9 _-]*):/);
    if (keyMatch) {
      const rest = escaped.slice(keyMatch[0].length);
      return `<span class="syntax-key">${keyMatch[1]}:</span><span class="syntax-value">${rest}</span>`;
    }
    return escaped.replace(/(\[\[[^\]]+\]\])/g, '<span class="syntax-link">$1</span>');
  }).join('\n');
}

$('xlsxBtn').addEventListener('click', () => $('xlsxInput').click());
$('xlsxInput').addEventListener('change', () => {
  state.xlsxFile = $('xlsxInput').files?.[0] || null;
  $('xlsxName').textContent = state.xlsxFile?.name || 'No file selected';
  setMessage(state.xlsxFile ? 'Moodle export selected.' : '');
  refreshScanAvailability();
});
$('studentDirBtn').addEventListener('click', chooseStudentDir);
$('tutorDirBtn').addEventListener('click', chooseTutorDir);
$('scanBtn').addEventListener('click', runScan);
$('prepareBtn').addEventListener('click', prepareQueue);
$('resetBtn').addEventListener('click', resetSession);
$('searchInput').addEventListener('input', renderRecords);
for (const id of ['courseInput','yearInput','termInput','levelInput']) $(id).addEventListener('input', refreshScanAvailability);

document.querySelectorAll('.filter').forEach(button => button.addEventListener('click', () => {
  state.filter = button.dataset.filter;
  document.querySelectorAll('.filter').forEach(b => b.classList.toggle('active', b === button));
  renderRecords();
}));

$('applyBtn').addEventListener('click', applySelected);
$('rejectBtn').addEventListener('click', rejectSelected);
$('helpBtn').addEventListener('click', () => $('helpDialog').showModal());
$('closeHelpBtn').addEventListener('click', () => $('helpDialog').close());

if (!supportsFSAccess()) {
  $('compatNotice').classList.remove('hidden');
  $('compatNotice').className = 'notice warning';
  $('compatNotice').textContent = 'This browser may not support the local folder read/write workflow required by this version. Chrome or Edge is recommended for testing.';
}
refreshScanAvailability();
