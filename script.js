// Holds all parsed rows from the uploaded file
let rows = [];
// Original column names from the file, used to preserve order on export
let headers = [];
// Stores each post's annotation: { risk_label, keep }
let annotations = [];
// Index of the post currently being shown
let current = 0;
// Tracks whether the uploaded file was CSV or XLSX so we can suggest the right export format
let inputFormat = 'csv';
// Original filename without extension, used when naming the exported file
let inputFileName = 'annotations';

// ── Upload / drag-drop ────────────────────────────────────────────────────

const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

// Clicking the styled drop zone triggers the hidden file input
dropZone.addEventListener('click', () => fileInput.click());

// Highlight the drop zone while a file is being dragged over it
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });

// Remove highlight when the file leaves the drop zone
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));

// Handle file dropped onto the drop zone
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) loadFile(file);
});

// Handle file selected via the file browser
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) loadFile(fileInput.files[0]);
});

// Routes to the correct parser based on file extension
function loadFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  inputFormat = (ext === 'xlsx' || ext === 'xls') ? 'xlsx' : 'csv';
  inputFileName = file.name.replace(/\.[^.]+$/, ''); // strip extension
  if (inputFormat === 'xlsx') {
    loadXLSX(file);
  } else {
    loadCSV(file);
  }
}

// Parses a CSV file using PapaParse
function loadCSV(file) {
  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete(result) {
      if (!result.data.length) {
        alert('The CSV appears to be empty.');
        return;
      }
      initData(result.data, result.meta.fields);
    },
    error(err) {
      alert('Could not parse CSV: ' + err.message);
    }
  });
}

// Parses an XLSX/XLS file using SheetJS, reads the first sheet
function loadXLSX(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const workbook = XLSX.read(e.target.result, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet, { defval: '' }); // defval fills empty cells with ''
      if (!data.length) {
        alert('The spreadsheet appears to be empty.');
        return;
      }
      initData(data, Object.keys(data[0]));
    } catch (err) {
      alert('Could not parse spreadsheet: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// Initializes app state and starts annotation after a file is loaded
function initData(data, fields) {
  rows = data;
  headers = fields;
  // Pre-fill all annotations as unlabeled, defaulting keep to true
  annotations = rows.map(() => ({ risk_label: null, keep: true }));
  current = 0;
  showAnnotationScreen();
  renderPost();
}

// ── Screen switching ──────────────────────────────────────────────────────

function showAnnotationScreen() {
  document.getElementById('upload-screen').style.display = 'none';
  document.getElementById('done-screen').style.display = 'none';
  document.getElementById('annotation-screen').style.display = 'block';
}

function showDoneScreen() {
  document.getElementById('annotation-screen').style.display = 'none';
  document.getElementById('done-screen').style.display = 'flex';
  document.getElementById('done-total').textContent = rows.length;
}

// ── Render a post ─────────────────────────────────────────────────────────

function renderPost() {
  const row = rows[current];
  const ann = annotations[current];

  // Update progress counter and bar
  document.getElementById('progress-text').textContent = `${current + 1} / ${rows.length}`;
  const pct = ((current + 1) / rows.length) * 100;
  document.getElementById('progress-bar').style.width = pct + '%';

  // Build meta badges (subreddit, type, date, teen_label, author)
  const metaEl = document.getElementById('post-meta');
  metaEl.innerHTML = '';

  const metaFields = ['subreddit', 'type', 'created_utc', 'teen_label', 'author'];
  metaFields.forEach(field => {
    if (row[field] !== undefined && row[field] !== '') {
      const badge = document.createElement('span');
      badge.className = 'meta-badge' + (field === 'teen_label' ? ' highlight' : '');
      badge.textContent = field === 'subreddit' ? 'r/' + row[field]
        : field === 'teen_label' ? '🏷 ' + row[field]
        : field === 'created_utc' ? formatDate(row[field])
        : row[field];
      metaEl.appendChild(badge);
    }
  });

  // Prepend row number badge
  const idxBadge = document.createElement('span');
  idxBadge.className = 'meta-badge';
  idxBadge.textContent = '#' + (current + 1);
  metaEl.prepend(idxBadge);

  // Show post text — fall back to common alternative column names
  document.getElementById('post-text').textContent =
    row['text'] || row['body'] || row['content'] || '(no text field found)';

  // Show matched keywords if present
  const mw = row['matched_words'];
  if (mw && mw.trim()) {
    document.getElementById('matched-words').style.display = 'block';
    document.getElementById('matched-words-val').textContent = mw;
  } else {
    document.getElementById('matched-words').style.display = 'none';
  }

  // Restore previously saved risk label selection for this post
  document.querySelectorAll('.radio-option').forEach(opt => opt.classList.remove('selected'));
  document.querySelectorAll('input[name="risk"]').forEach(r => { r.checked = false; });

  if (ann.risk_label) {
    const radio = document.querySelector(`input[name="risk"][value="${ann.risk_label}"]`);
    if (radio) {
      radio.checked = true;
      radio.closest('.radio-option').classList.add('selected');
    }
  }

  // Restore keep toggle state
  const keepToggle = document.getElementById('keep-toggle');
  keepToggle.checked = ann.keep;
  updateToggleLabel(ann.keep);

  // Reset error state and update button labels
  document.getElementById('q1-card').classList.remove('error');
  document.getElementById('btn-back').disabled = current === 0;
  document.getElementById('btn-next').textContent = current === rows.length - 1 ? 'Finish ✓' : 'Next →';
}

// ── Radio click handling ──────────────────────────────────────────────────

// Highlight the selected radio option and clear any error state
document.querySelectorAll('.radio-option').forEach(opt => {
  opt.addEventListener('click', () => {
    document.querySelectorAll('.radio-option').forEach(o => o.classList.remove('selected'));
    opt.classList.add('selected');
    opt.querySelector('input[type="radio"]').checked = true;
    document.getElementById('q1-card').classList.remove('error');
  });
});

// ── Toggle ────────────────────────────────────────────────────────────────

// Update the true/false label whenever the keep toggle changes
document.getElementById('keep-toggle').addEventListener('change', function () {
  updateToggleLabel(this.checked);
});

function updateToggleLabel(val) {
  const el = document.getElementById('toggle-label');
  el.textContent = val ? 'true' : 'false';
  el.className = 'toggle-label-text ' + (val ? 'true' : 'false');
}

// ── Navigation ────────────────────────────────────────────────────────────

// Saves the current post's annotation to the annotations array
function saveCurrentAnnotation() {
  const selected = document.querySelector('input[name="risk"]:checked');
  annotations[current] = {
    risk_label: selected ? selected.value : null,
    keep: document.getElementById('keep-toggle').checked,
  };
}

// Validates Q1, saves annotation, and moves to the next post or done screen
function goNext() {
  const selected = document.querySelector('input[name="risk"]:checked');
  if (!selected) {
    // Highlight Q1 card in red and scroll to it
    document.getElementById('q1-card').classList.add('error');
    document.getElementById('q1-card').scrollIntoView({ behavior: 'smooth', block: 'center' });
    return;
  }
  saveCurrentAnnotation();

  if (current === rows.length - 1) {
    showDoneScreen();
  } else {
    current++;
    renderPost();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
}

// Saves current annotation and goes back one post
function goBack() {
  if (current === 0) return;
  saveCurrentAnnotation();
  // If coming back from the done screen, return to annotation screen first
  if (document.getElementById('done-screen').style.display !== 'none') {
    showAnnotationScreen();
  }
  current = Math.max(0, current - 1);
  renderPost();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── Keyboard shortcuts ────────────────────────────────────────────────────

document.addEventListener('keydown', e => {
  if (document.getElementById('annotation-screen').style.display === 'none') return;

  // Keys 1-5 select risk labels
  const labelMap = { '1': 'INDICATOR', '2': 'IDEATION', '3': 'BEHAVIOR', '4': 'ATTEMPT', '5': 'UNCERTAIN' };
  if (labelMap[e.key]) {
    const opt = document.querySelector(`.radio-option[data-value="${labelMap[e.key]}"]`);
    if (opt) opt.click();
  }

  // K toggles the keep switch
  if (e.key === 'k' || e.key === 'K') {
    const toggle = document.getElementById('keep-toggle');
    toggle.checked = !toggle.checked;
    updateToggleLabel(toggle.checked);
  }

  if (e.key === 'Enter') goNext();
  if (e.key === 'ArrowLeft') goBack();
});

// ── Export ────────────────────────────────────────────────────────────────

// Exports all rows with risk_label and keep columns appended, in the chosen format
function exportFile(source) {
  // Save current post before exporting if still on annotation screen
  if (document.getElementById('annotation-screen').style.display !== 'none') {
    const selected = document.querySelector('input[name="risk"]:checked');
    if (selected) saveCurrentAnnotation();
  }

  // Read format choice from whichever dropdown triggered the export
  const formatSelect = source === 'done'
    ? document.getElementById('export-format-done')
    : document.getElementById('export-format');
  const format = formatSelect ? formatSelect.value : 'csv';

  // Build export rows: original data + risk_label + keep
  const exportRows = rows.map((row, i) => {
    const ann = annotations[i];
    return {
      ...row,
      risk_label: ann.risk_label ?? '',
      keep: ann.risk_label !== null ? String(ann.keep) : '',
    };
  });

  const filename = inputFileName + '_annotated_' + new Date().toISOString().slice(0, 10);

  if (format === 'xlsx') {
    const ws = XLSX.utils.json_to_sheet(exportRows, { header: [...headers, 'risk_label', 'keep'] });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Annotations');
    XLSX.writeFile(wb, filename + '.xlsx');
  } else {
    const csv = Papa.unparse(exportRows, { columns: [...headers, 'risk_label', 'keep'] });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename + '.csv';
    a.click();
    URL.revokeObjectURL(url);
  }
}

// ── Label definitions toggle ──────────────────────────────────────────────

// Shows or hides the label definitions panel
function toggleDefs() {
  document.getElementById('defs-panel').classList.toggle('open');
}

// ── Utilities ─────────────────────────────────────────────────────────────

// Converts a UTC timestamp (unix seconds or ISO string) to a readable date
function formatDate(utc) {
  if (!utc) return '';
  const d = isNaN(utc) ? new Date(utc) : new Date(Number(utc) * 1000);
  if (isNaN(d)) return utc;
  return d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
}
