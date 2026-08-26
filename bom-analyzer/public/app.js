(function () {
  'use strict';

  // ── Configuration ────────────────────────────────────────────────────────

  var STORAGE_KEY = 'bom-analyzer.apiBase';

  var MAP_FIELDS = [
    { key: 'mpn', label: 'Part number', required: true },
    { key: 'quantity', label: 'Quantity' },
    { key: 'reference', label: 'Reference designator' },
    { key: 'manufacturer', label: 'Manufacturer' },
    { key: 'description', label: 'Description' },
    { key: 'footprint', label: 'Package / footprint' },
    { key: 'skip', label: 'Skip to production' },
    { key: 'alternates', label: 'Alternate parts' },
  ];

  var SUPPLIER_COLUMNS = [
    { key: 'stock', label: 'Stock' },
    { key: 'lead', label: 'Lead time' },
    { key: 'unit', label: 'Unit' },
    { key: 'ext', label: 'Extended' },
    { key: 'lifecycle', label: 'Lifecycle' },
  ];

  var FILTERS = [
    { key: 'all', label: 'All' },
    { key: 'issues', label: 'Needs attention' },
    { key: 'lifecycle', label: 'Lifecycle risk' },
    { key: 'stock', label: 'Short on stock' },
    { key: 'missing', label: 'Not found' },
  ];

  // Several BOMs can be loaded at once. Everything that belongs to one BOM —
  // its parsed rows, column mapping, results, and even which filter and search
  // you left it on — lives on that BOM, so switching tabs restores exactly the
  // view you had rather than resetting it.
  var state = {
    apiBase: '',
    health: null,
    boms: [],
    activeId: null,
    // Alternatives found for a part number, keyed by the normalized number.
    // Not per BOM: the same part on two boards has the same replacements.
    alternatives: {},
  };

  var nextBomId = 1;

  function newBom(name, fields) {
    var bom = {
      id: 'bom-' + nextBomId++,
      name: name || 'Untitled',
      parse: null,
      mapping: {},
      lines: [],
      fromPaste: false,
      adhoc: false,
      results: null,
      excluded: [],
      claimed: [],
      filter: 'all',
      search: '',
      expanded: {},
      running: false,
      progress: null,
      error: null,
    };
    Object.keys(fields || {}).forEach(function (key) { bom[key] = fields[key]; });
    return bom;
  }

  // Returns the active BOM, or a throwaway blank one so render functions can
  // read fields without guarding every access.
  var placeholderBom = newBom('');
  function bom() {
    for (var i = 0; i < state.boms.length; i++) {
      if (state.boms[i].id === state.activeId) return state.boms[i];
    }
    return placeholderBom;
  }

  function addBom(name, fields) {
    var entry = newBom(name, fields);
    state.boms.push(entry);
    state.activeId = entry.id;
    return entry;
  }

  function removeBom(id) {
    var index = -1;
    state.boms.forEach(function (b, i) { if (b.id === id) index = i; });
    if (index === -1) return;
    state.boms.splice(index, 1);
    if (state.activeId === id) {
      var next = state.boms[index] || state.boms[index - 1];
      state.activeId = next ? next.id : null;
    }
    renderAll();
  }

  // A name the tabs can show: the filename without its extension, made unique
  // when the same file is loaded twice.
  function uniqueName(base) {
    var name = String(base || 'BOM').replace(/\.[^.]+$/, '') || 'BOM';
    var taken = state.boms.map(function (b) { return b.name; });
    if (taken.indexOf(name) === -1) return name;
    for (var n = 2; ; n++) {
      if (taken.indexOf(name + ' (' + n + ')') === -1) return name + ' (' + n + ')';
    }
  }

  var el = {};
  [
    'statusBar', 'settingsBtn', 'settingsPanel', 'apiBase', 'currencyLabel', 'recheckBtn',
    'clearCacheBtn', 'resetAppBtn', 'dropZone', 'fileInput', 'lookupRows', 'quickBtn', 'addRowBtn',
    'clearRowsBtn', 'sampleBtn',
    'mappingCard', 'mappingSummary', 'mappingGrid', 'previewTable', 'analyzeBtn', 'resetBtn',
    'progressWrap', 'progressBar', 'progressText', 'resultsCard', 'statGrid', 'searchInput',
    'filterChips', 'exportBtn', 'resultsTable', 'resultsHead', 'resultsBody', 'emptyState',
    'setupCard', 'toast', 'attribution',
    'bomBar', 'bomTabs', 'bomCount', 'analyzeAllBtn', 'closeAllBtn',
    'reportBtn', 'skippedNote', 'reportOverlay', 'dmsmsBtn', 'dmsmsOverlay',
    'altBtn', 'altOverlay',
  ].forEach(function (id) {
    el[id] = document.getElementById(id);
  });

  // index.html and app.js are one unit. When they arrive from different
  // versions the first line that touches a missing element throws, and the page
  // dies wherever it stood — historically on "Checking backend…", which says
  // nothing about what went wrong. Check once, up front, and say it plainly.
  var missing = Object.keys(el).filter(function (id) { return !el[id]; });
  if (missing.length) {
    reportStaleShell(missing);
    // This script has taken charge and said what is wrong, so the watchdog in
    // index.html has nothing to add.
    window.__bomAppReady = true;
    return;
  }

  // ── Recovering from a stale copy ─────────────────────────────────────────

  function reportStaleShell(missing) {
    var bar = document.getElementById('statusBar');
    if (bar) {
      bar.innerHTML = '<span class="status-pill err"><span class="dot"></span> ' +
        'Page and script are from different versions</span>';
    }

    var notice = document.createElement('section');
    notice.className = 'card setup';
    notice.innerHTML =
      '<h2><span class="num">!</span> The browser is holding an old copy</h2>' +
      '<p class="sub">This page loaded with a script from an earlier version, so the app ' +
      'could not start. Nothing is wrong with your data or the server.</p>' +
      '<div class="btn-row"><button type="button" id="resetShellBtn" class="btn primary">' +
      'Clear the old copy and reload</button></div>' +
      '<div class="hint">Or reload with <strong>Ctrl+Shift+R</strong> ' +
      '(<strong>Cmd+Shift+R</strong> on a Mac). Missing from this page: <code>' +
      missing.map(function (id) { return String(id).replace(/[^A-Za-z0-9_-]/g, ''); })
        .join('</code>, <code>') + '</code>.</div>';

    var app = document.getElementById('app');
    if (app) app.insertBefore(notice, app.firstChild ? app.firstChild.nextSibling : null);

    var button = document.getElementById('resetShellBtn');
    if (button) {
      button.addEventListener('click', function () {
        button.disabled = true;
        button.textContent = 'Clearing…';
        clearBrowserCopy(true);
      });
    }
  }

  // Unregisters any service worker and empties every cache it left behind.
  // index.html owns the implementation, because it has to work even when this
  // script is the stale thing being cleared.
  function clearBrowserCopy(reload) {
    if (typeof window.__bomClearBrowserCopy === 'function') {
      return window.__bomClearBrowserCopy(reload);
    }
    if (reload) location.reload();
    return Promise.resolve();
  }

  // ── Small helpers ────────────────────────────────────────────────────────

  // The same invisible characters bomlib/spreadsheet.py strips out of an
  // uploaded cell. A part number pasted from a datasheet PDF or a web page
  // routinely carries one, it cannot be seen or deleted in a text box, and it
  // stops the part matching anything. Ordinary spaces are handled by trim().
  var INVISIBLE = new RegExp(
    '[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f-\u009f' +
    '\u00ad\u200b-\u200f\u2028\u2029\u202a-\u202e' +
    '\u2060-\u2064\ufeff\ufff9-\ufffb]', 'g'
  );

  function cleanCell(value) {
    if (value === null || value === undefined) return '';
    return String(value).replace(INVISIBLE, '').replace(/[^\S\n\r]+/g, ' ').trim();
  }

  function esc(value) {
    if (value === null || value === undefined) return '';
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // Product and datasheet URLs come from the supplier APIs, so they are treated
  // as untrusted input before being put in an href.
  function safeUrl(value) {
    if (!value) return null;
    try {
      var url = new URL(String(value), location.href);
      return url.protocol === 'http:' || url.protocol === 'https:' ? url.href : null;
    } catch (err) {
      return null;
    }
  }

  function money(value, currency) {
    if (value === null || value === undefined || !isFinite(value)) return null;
    var code = currency || (bom().results && bom().results.summary.currency) || 'USD';
    try {
      return new Intl.NumberFormat(undefined, {
        style: 'currency',
        currency: code,
        minimumFractionDigits: value < 1 ? 4 : 2,
        maximumFractionDigits: value < 1 ? 5 : 2,
      }).format(value);
    } catch (err) {
      return code + ' ' + value.toFixed(4);
    }
  }

  function truncateText(value, limit) {
    var text = String(value === null || value === undefined ? '' : value);
    return text.length <= limit ? text : text.slice(0, limit - 1) + '…';
  }

  function count(value) {
    if (!isFinite(value) || value === null) return '—';
    return Number(value).toLocaleString();
  }

  var toastTimer = null;
  function toast(message, isBad) {
    el.toast.textContent = message;
    el.toast.className = 'toast' + (isBad ? ' bad' : '');
    el.toast.hidden = false;
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(function () {
      el.toast.hidden = true;
    }, 4200);
  }

  function api(path) {
    var base = state.apiBase.replace(/\/+$/, '');
    return base + path;
  }

  function defaultApiBase() {
    var stored = null;
    try {
      stored = localStorage.getItem(STORAGE_KEY);
    } catch (err) {
      stored = null;
    }
    if (stored) return stored;
    // Served by the backend itself: same origin is the right default. Opened
    // as a static file or from a separate host: fall back to a local server.
    if (location.protocol === 'http:' || location.protocol === 'https:') return location.origin;
    return 'http://localhost:8787';
  }

  // ── Backend health ───────────────────────────────────────────────────────

  function renderStatus() {
    var health = state.health;
    if (!health) {
      el.statusBar.innerHTML =
        '<span class="status-pill err"><span class="dot"></span> Backend unreachable at ' +
        esc(state.apiBase) + '</span>';
      el.setupCard.hidden = false;
      return;
    }

    el.setupCard.hidden = health.suppliers.some(function (s) { return s.configured; });

    var pills = health.suppliers.map(function (supplier) {
      var label = supplier.name + (supplier.configured ? ' connected' : ' not configured');
      if (supplier.configured && supplier.sandbox) label += ' (sandbox)';
      return '<span class="status-pill ' + (supplier.configured ? 'on' : 'off') + '">' +
        '<span class="dot"></span> ' + esc(label) + '</span>';
    });
    pills.push(
      '<span class="status-pill on"><span class="dot"></span> ' +
      count(health.cacheEntries) + ' cached lookups</span>'
    );
    el.statusBar.innerHTML = pills.join('');
    renderAttribution();
  }

  function checkHealth() {
    el.statusBar.innerHTML =
      '<span class="status-pill checking"><span class="dot"></span> Checking backend&hellip;</span>';
    return fetch(api('/api/health'), { cache: 'no-store' })
      .then(function (res) {
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return res.json();
      })
      .then(function (data) {
        state.health = data;
        el.currencyLabel.value = data.currency || 'USD';
        renderStatus();
        return data;
      })
      .catch(function () {
        state.health = null;
        renderStatus();
        return null;
      });
  }

  // ── Loading a BOM ────────────────────────────────────────────────────────

  function handleFiles(files) {
    var list = Array.prototype.slice.call(files || []);
    if (list.length === 0) return;
    toast('Reading ' + (list.length === 1 ? list[0].name : list.length + ' files') + '…');
    // Sequentially, so the tabs appear in the order they were chosen.
    return list.reduce(function (chain, file) {
      return chain.then(function () { return handleFile(file); });
    }, Promise.resolve()).then(function () {
      renderAll();
      var loaded = state.boms.length;
      toast('Loaded ' + loaded + ' BOM' + (loaded === 1 ? '' : 's'));
    });
  }

  function handleFile(file) {
    if (!file) return Promise.resolve();
    return fetch(api('/api/parse'), {
      method: 'POST',
      headers: { 'X-File-Name': file.name, 'Content-Type': 'application/octet-stream' },
      body: file,
    })
      .then(readJsonOrThrow)
      .then(function (data) {
        addBom(uniqueName(file.name), {
          parse: data,
          mapping: data.mapping || {},
          lines: data.lines || [],
          source: file.name,
        });
      })
      .catch(function (err) {
        // One bad file must not abandon the rest of the batch.
        addBom(uniqueName(file.name), { error: err.message || 'Could not read that file' });
        toast(file.name + ': ' + (err.message || 'could not be read'), true);
      });
  }

  function readJsonOrThrow(res) {
    return res.text().then(function (text) {
      var data = null;
      try {
        data = text ? JSON.parse(text) : null;
      } catch (err) {
        data = null;
      }
      if (!res.ok) {
        throw new Error((data && data.error) || 'Request failed with HTTP ' + res.status);
      }
      return data || {};
    });
  }

  // Pasting a column is the one place a part still has to be picked out of a
  // line of text. Which field is which is decided by what the values look like
  // rather than by position: the number is the quantity, the first thing that
  // is not a number is the part, and whatever is left is the description. A
  // spreadsheet copy arrives tab-separated and is split on tabs alone, so a
  // description containing commas survives intact.
  function parsePastedRows(text) {
    var rows = [];
    String(text || '').split(/\r?\n/).forEach(function (raw) {
      var line = raw.trim();
      if (!line || line.charAt(0) === '#') return;

      var fields = line.split(line.indexOf('\t') !== -1 ? /\t/ : /[,;]/)
        .map(cleanCell)
        .filter(function (piece) { return piece; });
      if (!fields.length) return;

      // Never the first field: a part number made only of digits is still a
      // part number, not a quantity.
      var quantity = null;
      for (var i = fields.length - 1; i >= 1; i--) {
        if (isQuantity(fields[i])) {
          quantity = parseQuantity(fields[i]);
          fields.splice(i, 1);
          break;
        }
      }

      rows.push({
        mpn: fields.shift() || '',
        description: fields.join(', '),
        quantity: quantity,
      });
    });
    return rows;
  }

  function isQuantity(value) {
    return /^\d[\d,]*$/.test(String(value || '').trim());
  }

  function parseQuantity(value) {
    var digits = String(value || '').replace(/[^0-9]/g, '');
    var parsed = parseInt(digits, 10);
    return isFinite(parsed) && parsed > 0 ? parsed : null;
  }

  // ── Look up parts ────────────────────────────────────────────────────────

  var STARTING_ROWS = 3;

  function lookupRowElements() {
    return Array.prototype.slice.call(el.lookupRows.querySelectorAll('.lookup-row'));
  }

  function addLookupRow(values, focus) {
    var row = document.createElement('div');
    row.className = 'lookup-row';
    row.innerHTML =
      // Each placeholder names its own field: the header row is the first thing
      // a narrow screen drops, and a bare "1" under nothing means nothing.
      '<input type="text" class="mpn" aria-label="Part number" placeholder="e.g. STM32F103C8T6" ' +
      'autocomplete="off" spellcheck="false" />' +
      '<input type="text" class="desc" aria-label="Description" placeholder="What it is (optional)" ' +
      'autocomplete="off" />' +
      '<input type="text" class="qty" aria-label="Quantity needed" placeholder="Qty (1)" ' +
      'inputmode="numeric" autocomplete="off" />' +
      '<button type="button" class="row-drop" aria-label="Remove this part">&times;</button>';

    if (values) {
      row.querySelector('.mpn').value = values.mpn || '';
      row.querySelector('.desc').value = values.description || '';
      row.querySelector('.qty').value = values.quantity ? String(values.quantity) : '';
    }

    el.lookupRows.appendChild(row);
    syncRowControls();
    if (focus) row.querySelector('.mpn').focus();
    return row;
  }

  // The last row is the one that grows the list, so it is never the one you
  // can delete down to nothing.
  function syncRowControls() {
    var rows = lookupRowElements();
    rows.forEach(function (row) {
      row.querySelector('.row-drop').disabled = rows.length === 1;
    });
  }

  function clearLookupRows() {
    el.lookupRows.innerHTML = '';
    for (var i = 0; i < STARTING_ROWS; i++) addLookupRow();
  }

  // A row with no part number is an empty row, whatever else was typed into it.
  function readLookupRows() {
    var lines = [];
    lookupRowElements().forEach(function (row) {
      var mpn = cleanCell(row.querySelector('.mpn').value);
      if (!mpn) return;
      var description = cleanCell(row.querySelector('.desc').value);
      var quantity = parseQuantity(row.querySelector('.qty').value);
      lines.push({
        row: lines.length + 1,
        mpn: mpn,
        quantity: quantity && quantity > 0 ? quantity : 1,
        reference: null,
        manufacturer: null,
        description: description || null,
      });
    });
    return lines;
  }

  function fillLookupRows(values, startRow) {
    var rows = lookupRowElements();
    var index = startRow ? rows.indexOf(startRow) : 0;
    if (index < 0) index = 0;

    values.forEach(function (entry, offset) {
      var target = lookupRowElements()[index + offset] || addLookupRow();
      target.querySelector('.mpn').value = entry.mpn || '';
      if (entry.description) target.querySelector('.desc').value = entry.description;
      if (entry.quantity) target.querySelector('.qty').value = String(entry.quantity);
    });

    // Keep one spare row at the end so the list can always be extended.
    var last = lookupRowElements()[lookupRowElements().length - 1];
    if (last && last.querySelector('.mpn').value.trim()) addLookupRow();
    syncRowControls();
  }

  // Searching reuses one tab rather than opening a new one each time: a lookup
  // is a question you refine, not a document you collect. The rows stay filled
  // in afterwards, so adding another part and searching again is two clicks.
  function quickSearch() {
    var lines = readLookupRows();
    if (!lines.length) {
      toast('Enter at least one part number', true);
      var first = lookupRowElements()[0];
      if (first) first.querySelector('.mpn').focus();
      return;
    }

    var entry = null;
    state.boms.forEach(function (candidate) {
      if (!entry && candidate.adhoc) entry = candidate;
    });

    if (entry && entry.running) {
      toast('That search is still running', true);
      return;
    }

    if (entry) {
      entry.name = searchName(lines);
      entry.lines = lines;
      entry.results = null;
      entry.excluded = [];
      entry.claimed = [];
      entry.expanded = {};
      entry.filter = 'all';
      entry.search = '';
      entry.error = null;
      entry.progress = null;
      state.activeId = entry.id;
    } else {
      entry = addBom(searchName(lines), { lines: lines, fromPaste: true, adhoc: true });
    }

    renderAll();
    analyze(entry);
  }

  function searchName(lines) {
    if (lines.length === 1) return lines[0].mpn;
    return lines[0].mpn + ' +' + (lines.length - 1);
  }

  // ── Column mapping ───────────────────────────────────────────────────────

  // ── BOM tabs ─────────────────────────────────────────────────────────────

  function bomStatus(entry) {
    if (entry.error) return { text: 'could not be read', cls: 'bad' };
    if (entry.running) return { text: entry.progress || 'analyzing…', cls: '', busy: true };
    if (entry.results) {
      var summary = entry.results.summary;
      var risks = (summary.riskLines || []).length;
      var skipped = (entry.excluded || []).length;
      return {
        text: summary.lines + ' lines · ' + (money(summary.bestMixTotal, summary.currency) || '—') +
          (risks ? ' · ' + risks + ' to review' : '') +
          (skipped ? ' · ' + skipped + ' skipped' : ''),
        cls: risks ? 'warn' : 'ok',
      };
    }
    return { text: entry.lines.length + ' parts · not analyzed', cls: '' };
  }

  function renderBomTabs() {
    var many = state.boms.length > 0;
    el.bomBar.hidden = !many;
    if (!many) return;

    el.bomCount.textContent = state.boms.length === 1 ? '' : '(' + state.boms.length + ')';
    el.bomTabs.innerHTML = state.boms.map(function (entry) {
      var status = bomStatus(entry);
      return '<div class="bom-tab' + (entry.id === state.activeId ? ' active' : '') +
        '" role="tab" tabindex="0" aria-selected="' + (entry.id === state.activeId) +
        '" data-id="' + esc(entry.id) + '">' +
        (status.busy ? '<span class="spinner"></span>' : '') +
        '<span class="tab-body">' +
        '<span class="tab-name">' + esc(entry.name) + '</span>' +
        '<span class="tab-meta ' + status.cls + '">' + esc(status.text) + '</span>' +
        '</span>' +
        '<button type="button" class="tab-close" data-close="' + esc(entry.id) +
        '" aria-label="Close ' + esc(entry.name) + '">×</button>' +
        '</div>';
    }).join('');

    Array.prototype.forEach.call(el.bomTabs.querySelectorAll('.bom-tab'), function (tab) {
      var id = tab.getAttribute('data-id');
      tab.addEventListener('click', function (event) {
        if (event.target.closest('.tab-close')) return;
        selectBom(id);
      });
      tab.addEventListener('keydown', function (event) {
        if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); selectBom(id); }
      });
    });
    Array.prototype.forEach.call(el.bomTabs.querySelectorAll('.tab-close'), function (button) {
      button.addEventListener('click', function (event) {
        event.stopPropagation();
        removeBom(button.getAttribute('data-close'));
      });
    });

    var pending = state.boms.filter(function (b) {
      return !b.results && !b.running && !b.error && b.lines.length;
    }).length;
    el.analyzeAllBtn.hidden = pending < 2;
    el.analyzeAllBtn.textContent = 'Analyze all pending (' + pending + ')';
  }

  function selectBom(id) {
    if (state.activeId === id) return;
    state.activeId = id;
    renderAll();
  }

  // One entry point that repaints every panel from the active BOM, so
  // switching tabs cannot leave a stale fragment of another BOM on screen.
  function renderAll() {
    renderBomTabs();
    var entry = bom();

    if (!state.boms.length) {
      el.mappingCard.hidden = true;
      el.resultsCard.hidden = true;
      el.progressWrap.hidden = true;
      return;
    }

    if (entry.error) {
      el.mappingCard.hidden = true;
      el.resultsCard.hidden = true;
      el.progressWrap.hidden = true;
      return;
    }

    renderMapping();
    el.progressWrap.hidden = !(entry.running || entry.progress);
    if (entry.progress) setProgress(entry.percent || 0, entry.progress);

    if (entry.results) {
      el.resultsCard.hidden = false;
      el.searchInput.value = entry.search || '';
      renderStats();
      renderSkipped();
      renderFilters();
      renderTable();
      renderAttribution();
      var resultsHeading = el.resultsCard.querySelector('h2');
      if (resultsHeading) {
        resultsHeading.innerHTML = '<span class="num">3</span> Supplier comparison' +
          (state.boms.length > 1 ? ' <span class="card-scope">' + esc(bom().name) + '</span>' : '');
      }
    } else {
      el.resultsCard.hidden = true;
    }
  }

  function renderMapping() {
    // A search has no header row to check and no columns to remap, and the
    // search box above is already the way to change it.
    if (bom().adhoc) {
      el.mappingCard.hidden = true;
      return;
    }
    el.mappingCard.hidden = false;
    var heading = el.mappingCard.querySelector('h2');
    if (heading) {
      heading.innerHTML = '<span class="num">2</span> Check the columns' +
        (state.boms.length > 1 ? ' <span class="card-scope">' + esc(bom().name) + '</span>' : '');
    }

    if (bom().fromPaste) {
      el.mappingGrid.innerHTML = '';
      el.mappingSummary.textContent =
        bom().lines.length + ' part number' + (bom().lines.length === 1 ? '' : 's') +
        ' from the pasted list. Quantities default to 1 where none was given.';
    } else {
      var headers = (bom().parse && bom().parse.headers) || [];
      el.mappingGrid.innerHTML = MAP_FIELDS.map(function (field) {
        var selected = bom().mapping[field.key];
        var options = ['<option value="">— none —</option>'].concat(
          headers.map(function (header, index) {
            var label = header || 'Column ' + columnLetter(index);
            return '<option value="' + index + '"' + (selected === index ? ' selected' : '') + '>' +
              esc(label) + '</option>';
          })
        );
        return '<label class="field">' + esc(field.label) + (field.required ? ' *' : '') +
          '<select data-field="' + field.key + '">' + options.join('') + '</select></label>';
      }).join('');

      Array.prototype.forEach.call(el.mappingGrid.querySelectorAll('select'), function (select) {
        select.addEventListener('change', onMappingChange);
      });

      var skipped = (bom().parse && bom().parse.skipped) || 0;
      el.mappingSummary.innerHTML =
        'Detected the header on row ' + (((bom().parse && bom().parse.headerRow) || 0) + 1) +
        '. <strong>' + bom().lines.length + '</strong> part' + (bom().lines.length === 1 ? '' : 's') +
        ' ready' + (skipped ? ', ' + skipped + ' row' + (skipped === 1 ? '' : 's') +
        ' skipped with no part number' : '') + '.';
    }

    var preview = previewScreening(bom());
    if (preview) {
      el.mappingSummary.innerHTML += ' <span class="muted">' + esc(preview) + '</span>';
    }

    renderPreview();
    el.analyzeBtn.disabled = bom().lines.length === 0 || bom().running;
    el.analyzeBtn.textContent = state.boms.length > 1
      ? 'Analyze ' + bom().name
      : 'Analyze BOM';
    el.resetBtn.textContent = state.boms.length > 1 ? 'Close this BOM' : 'Start over';
  }

  function columnLetter(index) {
    var letters = '';
    var n = index;
    do {
      letters = String.fromCharCode(65 + (n % 26)) + letters;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return letters;
  }

  function onMappingChange(event) {
    var field = event.target.getAttribute('data-field');
    var value = event.target.value;
    if (value === '') delete bom().mapping[field];
    else bom().mapping[field] = parseInt(value, 10);

    fetch(api('/api/remap'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        rows: (bom().parse && bom().parse.rows) || [],
        mapping: bom().mapping,
        rowOffset: (bom().parse && bom().parse.rowOffset) || 0,
      }),
    })
      .then(readJsonOrThrow)
      .then(function (data) {
        bom().lines = data.lines || [];
        renderMapping();
        renderBomTabs();
      })
      .catch(function (err) {
        toast(err.message || 'Could not apply that mapping', true);
      });
  }

  function renderPreview() {
    var sample = bom().lines.slice(0, 5);
    if (sample.length === 0) {
      el.previewTable.innerHTML =
        '<tr><td class="muted">No rows with a part number yet — pick the part number column above.</td></tr>';
      return;
    }
    var cols = ['row', 'mpn', 'quantity', 'reference', 'manufacturer', 'description'];
    var labels = ['Row', 'Part number', 'Qty', 'Ref', 'Manufacturer', 'Description'];
    // Only worth a column when the BOM actually has one.
    if (bom().mapping && bom().mapping.skip !== undefined) {
      cols.push('skip');
      labels.push('Skip');
    }
    var head = '<tr>' + labels.map(function (l) { return '<th>' + esc(l) + '</th>'; }).join('') + '</tr>';
    var body = sample.map(function (line) {
      return '<tr>' + cols.map(function (col) {
        var value = line[col];
        return '<td>' + (value === null || value === undefined || value === '' ? '—' : esc(value)) + '</td>';
      }).join('') + '</tr>';
    }).join('');
    var more = bom().lines.length > sample.length
      ? '<tr><td colspan="' + cols.length + '" class="muted">…and ' +
        (bom().lines.length - sample.length) + ' more</td></tr>'
      : '';
    el.previewTable.innerHTML = head + body + more;
  }

  // ── Running the analysis ─────────────────────────────────────────────────

  function analyze(entry) {
    entry = entry || bom();
    if (!entry || entry.running || entry.lines.length === 0) return Promise.resolve();
    if (!state.health) {
      toast('Backend is not reachable — check Settings', true);
      return Promise.resolve();
    }
    if (!state.health.suppliers.some(function (s) { return s.configured; })) {
      toast('No supplier API credentials are configured on the server', true);
      return Promise.resolve();
    }

    var max = state.health.maxPartsPerRequest || 500;
    var parts = entry.lines.slice(0, max);
    if (entry.lines.length > max) {
      toast(entry.name + ': analyzing the first ' + max + ' of ' + entry.lines.length +
        ' parts (server limit)', true);
    }

    entry.running = true;
    entry.error = null;
    entry.progress = 'Contacting suppliers…';
    entry.percent = 0;
    renderAll();
    el.analyzeBtn.disabled = true;

    return streamLookup(parts, entry, function (event, data) {
      if (event === 'start') {
        var expected = data.parts * data.suppliers.length;
        entry.percent = 0;
        entry.progress = 'Looking up ' + data.parts + ' parts across ' +
          data.suppliers.map(function (s) { return s.name; }).join(' and ') +
          ' (' + expected + ' queries)…';
      } else if (event === 'progress') {
        entry.percent = data.total ? Math.round((data.completed / data.total) * 100) : 0;
        entry.progress = data.completed + ' of ' + data.total + ' queries — ' +
          data.apiCalls + ' live, ' + data.cacheHits + ' cached' +
          (data.errors ? ', ' + data.errors + ' failed' : '');
      } else if (event === 'done') {
        finishAnalysis(entry, data);
        return;
      } else if (event === 'error') {
        throw new Error(data.error || 'Supplier lookup failed');
      }
      // Only repaint the tabs mid-run; a full repaint would fight the table.
      if (entry.id === state.activeId) {
        setProgress(entry.percent, entry.progress);
      }
      renderBomTabs();
    })
      .catch(function (err) {
        entry.error = err.message || 'Supplier lookup failed';
        entry.progress = 'Stopped: ' + entry.error;
        toast(entry.name + ': ' + entry.error, true);
      })
      .then(function () {
        entry.running = false;
        el.analyzeBtn.disabled = bom().lines.length === 0;
        renderAll();
      });
  }

  // Runs the un-analyzed BOMs one after another rather than in parallel, so
  // the supplier APIs see the same request rate as a single run.
  function analyzeAll() {
    var pending = state.boms.filter(function (b) {
      return !b.results && !b.running && !b.error && b.lines.length;
    });
    if (pending.length === 0) return;
    toast('Analyzing ' + pending.length + ' BOMs…');
    pending.reduce(function (chain, entry) {
      return chain.then(function () { return analyze(entry); });
    }, Promise.resolve()).then(function () {
      toast('Finished analyzing ' + pending.length + ' BOMs');
    });
  }

  function setProgress(percent, text) {
    el.progressBar.style.width = Math.max(0, Math.min(100, percent)) + '%';
    el.progressText.textContent = text || '';
  }

  function streamLookup(parts, entry, onEvent) {
    return fetch(api('/api/lookup'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Accept: 'text/event-stream' },
      body: JSON.stringify({
        parts: parts,
        stream: true,
        // Hand-entered part numbers are looked up as asked: no in-house prefix
        // is applied and no other BOM's claim answers for them.
        manual: !!entry.adhoc,
        claimed: entry.adhoc ? {} : claimedParts(entry),
      }),
    }).then(function (res) {
      if (!res.ok) return readJsonOrThrow(res);
      if (!res.body || !res.body.getReader) {
        throw new Error('This browser cannot read a streaming response');
      }

      var reader = res.body.getReader();
      var decoder = new TextDecoder();
      var buffer = '';

      function pump() {
        return reader.read().then(function (chunk) {
          if (chunk.done) return;
          buffer += decoder.decode(chunk.value, { stream: true });
          var boundary;
          while ((boundary = buffer.indexOf('\n\n')) !== -1) {
            var frame = buffer.slice(0, boundary);
            buffer = buffer.slice(boundary + 2);
            dispatchFrame(frame, onEvent);
          }
          return pump();
        });
      }
      return pump();
    });
  }

  function dispatchFrame(frame, onEvent) {
    var event = 'message';
    var data = '';
    frame.split('\n').forEach(function (line) {
      if (line.indexOf('event:') === 0) event = line.slice(6).trim();
      else if (line.indexOf('data:') === 0) data += line.slice(5).trim();
    });
    if (!data) return;
    onEvent(event, JSON.parse(data));
  }

  function finishAnalysis(entry, data) {
    entry.results = data;
    entry.excluded = data.excluded || [];
    // The part numbers this BOM now owns. Whichever BOM is analyzed first
    // keeps a shared part, so the others skip it instead of paying for the
    // same lookup again.
    entry.claimed = data.claimed || [];
    entry.expanded = {};
    entry.filter = 'all';
    entry.search = '';
    entry.percent = 100;
    entry.progress = 'Done — ' + data.stats.apiCalls + ' live queries, ' +
      data.stats.cacheHits + ' served from cache' +
      (data.stats.errors ? ', ' + data.stats.errors + ' failed' : '') +
      (entry.excluded.length ? ', ' + entry.excluded.length + ' lines skipped' : '') + '.';
    checkHealth();
    if (entry.id === state.activeId) {
      renderAll();
      el.resultsCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } else {
      renderBomTabs();
    }
  }

  // ── Results: summary tiles ───────────────────────────────────────────────

  function renderStats() {
    var summary = bom().results.summary;
    var suppliers = bom().results.suppliers;
    var tiles = [];

    tiles.push(tile('Lines analyzed', String(summary.lines), count(summary.totalQuantity) + ' units total', ''));

    suppliers.forEach(function (supplier) {
      var totals = summary.supplierTotals[supplier.id];
      if (!totals) return;
      var gaps = [];
      if (totals.linesMissing) gaps.push(totals.linesMissing + ' not carried');
      if (totals.linesShort) gaps.push(totals.linesShort + ' short on stock');
      tiles.push(tile(
        supplier.name + ' cart',
        money(totals.total, summary.currency) || '—',
        gaps.length ? gaps.join(' · ') : 'covers all ' + totals.linesPriced + ' lines',
        totals.complete ? '' : 'warn'
      ));
    });

    if (suppliers.length > 1) {
      var savings = summary.mixSavings;
      tiles.push(tile(
        'Cheapest per line',
        money(summary.bestMixTotal, summary.currency) || '—',
        isFinite(savings) && savings > 0
          ? 'saves ' + money(savings, summary.currency) + ' vs. single-sourcing'
          : 'split across suppliers',
        'accent'
      ));
    }

    var stockRisk = bom().results.rows.filter(function (row) {
      return row.comparison.inStockSuppliers.length === 0;
    }).length;
    tiles.push(tile(
      'Stock risk',
      String(stockRisk),
      stockRisk ? 'no supplier holds the full quantity' : 'every line is coverable today',
      stockRisk ? 'bad' : 'good'
    ));

    var lifecycleRisk = bom().results.rows.filter(function (row) {
      var severity = row.comparison.lifecycleSeverity;
      return severity === 'bad' || severity === 'warn';
    }).length;
    tiles.push(tile(
      'Lifecycle risk',
      String(lifecycleRisk),
      lifecycleRisk ? 'NRND, EOL or obsolete parts' : 'nothing flagged end-of-life',
      lifecycleRisk ? 'warn' : 'good'
    ));

    if (summary.notFoundLines) {
      tiles.push(tile('Not found', String(summary.notFoundLines),
        'no supplier returned a match', 'bad'));
    }

    el.statGrid.innerHTML = tiles.join('');
  }

  // Suppliers can require visible attribution. TrustedParts ask for the words
  // "Powered by" followed by their logo, linked back to them, and the link
  // must stay followable — so no rel="nofollow" here, deliberately.
  function renderAttribution() {
    var suppliers = (state.health && state.health.suppliers) || [];
    var blocks = suppliers.filter(function (s) {
      return s.attribution && s.configured;
    }).map(function (supplier) {
      var a = supplier.attribution;
      // Prefer the per-part link the API returned for this run; otherwise the
      // home page, which their guidance allows for multi-part displays.
      var url = safeUrl(preferredAttributionUrl(supplier.id) || a.url) || a.url;
      // The logo is theirs to supply, so the name renders by default and the
      // image replaces it only once it actually loads. No inline handlers:
      // they are fragile to escape and blocked under a strict CSP.
      var logo = '<span class="attr-name">' + esc(a.name) + '</span>' +
        (a.logo ? '<img class="attr-logo" src="' + esc(a.logo) + '" alt="' + esc(a.name) +
          '" hidden>' : '');
      return '<a class="attr-link" href="' + esc(url) + '" target="_blank" rel="noopener">' +
        '<span class="attr-text">' + esc(a.text) + '</span>' + logo + '</a>';
    });

    el.attribution.innerHTML = blocks.join('');
    el.attribution.hidden = blocks.length === 0;

    // Swap the name for the logo only on a successful load, so a missing file
    // leaves working text attribution rather than a broken image.
    Array.prototype.forEach.call(el.attribution.querySelectorAll('.attr-logo'), function (img) {
      img.addEventListener('load', function () {
        img.hidden = false;
        var name = img.parentNode.querySelector('.attr-name');
        if (name) name.hidden = true;
      });
    });
  }

  // When every TrustedParts row points at the same part page (a single-part
  // run) that page is the better target than the home page.
  function preferredAttributionUrl(supplierId) {
    if (!bom().results) return null;
    var urls = [];
    bom().results.rows.forEach(function (row) {
      var offer = row.offers[supplierId];
      if (offer && offer.attribution && offer.attribution.url) urls.push(offer.attribution.url);
    });
    if (urls.length === 1) return urls[0];
    var unique = urls.filter(function (u, i) { return urls.indexOf(u) === i; });
    return unique.length === 1 ? unique[0] : null;
  }

  function tile(label, value, note, tone) {
    return '<div class="stat ' + tone + '">' +
      '<div class="label">' + esc(label) + '</div>' +
      '<div class="value">' + esc(value) + '</div>' +
      '<div class="note">' + esc(note) + '</div>' +
      '</div>';
  }

  // ── Results: filters ─────────────────────────────────────────────────────

  function matchesFilter(row, filter) {
    switch (filter) {
      case 'issues':
        return row.comparison.flags.length > 0;
      case 'lifecycle':
        return row.comparison.lifecycleSeverity === 'bad' || row.comparison.lifecycleSeverity === 'warn';
      case 'stock':
        return row.comparison.inStockSuppliers.length === 0;
      case 'missing':
        return !bom().results.suppliers.some(function (s) {
          return row.offers[s.id] && row.offers[s.id].found;
        });
      default:
        return true;
    }
  }

  function renderFilters() {
    el.filterChips.innerHTML = FILTERS.map(function (filter) {
      var n = bom().results.rows.filter(function (row) { return matchesFilter(row, filter.key); }).length;
      return '<button type="button" class="chip' + (bom().filter === filter.key ? ' active' : '') +
        '" data-filter="' + filter.key + '">' + esc(filter.label) +
        '<span class="count">' + n + '</span></button>';
    }).join('');

    Array.prototype.forEach.call(el.filterChips.querySelectorAll('.chip'), function (chip) {
      chip.addEventListener('click', function () {
        bom().filter = chip.getAttribute('data-filter');
        renderFilters();
        renderTable();
      });
    });
  }

  function visibleRows() {
    var query = bom().search.trim().toLowerCase();
    return bom().results.rows.filter(function (row) {
      if (!matchesFilter(row, bom().filter)) return false;
      if (!query) return true;
      var haystack = [row.mpn, row.description, row.reference, row.manufacturer];
      bom().results.suppliers.forEach(function (supplier) {
        var offer = row.offers[supplier.id];
        if (offer && offer.found) haystack.push(offer.supplierPartNumber, offer.description, offer.manufacturer);
      });
      return haystack.some(function (value) {
        return value && String(value).toLowerCase().indexOf(query) !== -1;
      });
    });
  }

  // ── Results: table ───────────────────────────────────────────────────────

  function renderTable() {
    var suppliers = bom().results.suppliers;

    var groupRow = '<tr class="supplier-row">' +
      '<th class="spacer sticky-a"></th><th class="spacer sticky-b"></th><th class="spacer"></th>' +
      suppliers.map(function (supplier) {
        return '<th colspan="' + SUPPLIER_COLUMNS.length + '" class="group-start">' +
          esc(supplier.name) + '</th>';
      }).join('') +
      '<th class="spacer group-start sticky-z"></th></tr>';

    var fieldRow = '<tr class="field-row"><th class="sticky-a"></th>' +
      '<th class="sticky-b">Part</th><th class="num">Qty</th>' +
      suppliers.map(function () {
        return SUPPLIER_COLUMNS.map(function (col, i) {
          var cls = (i === 0 ? 'group-start ' : '') + (col.key === 'lifecycle' ? '' : 'num');
          return '<th class="' + cls.trim() + '">' + esc(col.label) + '</th>';
        }).join('');
      }).join('') +
      '<th class="group-start sticky-z">Verdict</th></tr>';

    el.resultsHead.innerHTML = groupRow + fieldRow;

    var rows = visibleRows();
    el.emptyState.hidden = rows.length > 0;
    el.resultsTable.hidden = rows.length === 0;

    el.resultsBody.innerHTML = rows.map(function (row) {
      return renderRow(row, suppliers);
    }).join('');

    syncStickyOffset();

    Array.prototype.forEach.call(el.resultsBody.querySelectorAll('.expander'), function (button) {
      button.addEventListener('click', function () {
        var index = button.getAttribute('data-index');
        if (bom().expanded[index]) delete bom().expanded[index];
        else bom().expanded[index] = true;
        renderTable();
      });
    });
  }

  // Both pins are measured rather than assumed. The part column sits beside
  // the expander column and the second header row sits below the first, so a
  // guessed offset leaves either a gap that scrolled content shows through or
  // an overlap that hides a row of headings.
  function syncStickyOffset() {
    var first = el.resultsBody.querySelector('td.sticky-a');
    if (first) {
      var width = Math.ceil(first.getBoundingClientRect().width);
      if (width > 0) {
        el.resultsTable.style.setProperty('--sticky-b-left', width + 'px');
        el.resultsTable.style.setProperty('--sticky-a-width', width + 'px');
      }
    }
    var supplierRow = el.resultsHead.querySelector('tr.supplier-row');
    if (supplierRow) {
      var height = Math.round(supplierRow.getBoundingClientRect().height);
      if (height > 0) {
        el.resultsTable.style.setProperty('--head-row1-height', height + 'px');
      }
    }
  }

  // ── Approved alternates ──────────────────────────────────────────────────
  //
  // Some BOMs carry a column naming the parts engineering already approved as
  // substitutes. They are looked up alongside the primary, so the answer to
  // "the primary is obsolete, now what" is on the page before anyone asks.

  function alternateEntries(row) {
    return (row.alternates || []).filter(function (entry) { return entry && entry.mpn; });
  }

  function usableAlternates(row) {
    return alternateEntries(row).filter(function (entry) { return entry.usable; });
  }

  // The same question the report's "Needs a decision" section asks: is this
  // line unbuyable, unstocked, or on its way out.
  function rowAtRisk(row, suppliers) {
    var comparison = row.comparison;
    var found = (suppliers || []).some(function (supplier) {
      var offer = row.offers[supplier.id];
      return offer && offer.found;
    });
    var severity = comparison.lifecycleSeverity;
    return !found || severity === 'bad' || severity === 'warn' ||
      comparison.inStockSuppliers.length === 0;
  }

  function alternateSummaryText(row) {
    var entries = alternateEntries(row);
    if (!entries.length) return '';
    return entries.map(function (entry) {
      return entry.mpn + ' (' + (entry.usable ? 'available' : 'not available') + ')';
    }).join('; ');
  }

  function alternateBadge(entry) {
    if (entry.usable) return '<span class="badge ok">available</span>';
    if (!entry.found) return '<span class="badge bad">no match</span>';
    if (!entry.coversQuantity) return '<span class="badge warn">no stock</span>';
    return '<span class="badge warn">ending</span>';
  }

  function renderRow(row, suppliers) {
    var isOpen = !!bom().expanded[row.index];
    var comparison = row.comparison;

    var meta = [];
    if (row.reference) meta.push('<span class="refdes">' + esc(row.reference) + '</span>');
    if (row.manufacturer) meta.push(esc(row.manufacturer));
    if (row.description) meta.push(esc(row.description));

    var cells = suppliers.map(function (supplier) {
      return supplierCells(row.offers[supplier.id], supplier, comparison, row.quantity);
    }).join('');

    var verdict = renderVerdict(comparison, suppliers);

    // A line whose primary is in trouble but whose BOM already names a stocked
    // alternate is not the same emergency as one with nowhere to go, so the
    // table says which it is without waiting to be expanded.
    var entries = alternateEntries(row);
    if (entries.length && rowAtRisk(row, suppliers)) {
      var covered = usableAlternates(row);
      verdict += covered.length
        ? '<div class="rowmeta"><span class="badge ok">alt: ' + esc(covered[0].mpn) + '</span></div>'
        : '<div class="rowmeta"><span class="badge warn">' + entries.length +
          ' alt' + (entries.length === 1 ? '' : 's') + ', none available</span></div>';
    }

    var main = '<tr class="part-row' + (isOpen ? ' expanded' : '') + '">' +
      '<td class="sticky-a"><button type="button" class="expander" data-index="' + row.index +
      '" aria-label="Toggle details">' + (isOpen ? '▾' : '▸') + '</button></td>' +
      '<td class="sticky-b"><div class="part-cell"><div class="mpn">' + esc(row.mpn) + '</div>' +
      (meta.length ? '<div class="rowmeta">' + meta.join(' · ') + '</div>' : '') + '</div></td>' +
      '<td class="num">' + count(row.quantity) + '</td>' +
      cells +
      '<td class="group-start sticky-z"><div class="verdict">' + verdict + '</div></td>' +
      '</tr>';

    if (!isOpen) return main;
    var span = 4 + suppliers.length * SUPPLIER_COLUMNS.length;
    return main + '<tr class="detail-row"><td colspan="' + span + '">' +
      renderDetail(row, suppliers) + '</td></tr>';
  }

  function supplierCells(offer, supplier, comparison, needed) {
    var start = ' class="group-start';
    if (!offer || !offer.found) {
      var message = offer && offer.error
        ? '<span class="err-text">' + esc(offer.reason) + '</span>'
        : '<span class="miss">' + esc((offer && offer.reason) || 'No match') + '</span>';
      return '<td' + start + '" colspan="' + SUPPLIER_COLUMNS.length + '">' + message + '</td>';
    }

    var bestPrice = comparison.bestPriceSupplier === supplier.name;
    var bestLead = comparison.bestLeadTimeSupplier === supplier.name;

    var stockCell = '<td' + start + ' num">' +
      (offer.stock === null ? '<span class="muted">—</span>' : count(offer.stock)) +
      (offer.stockSufficient === false
        ? '<div class="rowmeta"><span class="badge warn">short</span></div>'
        : '') +
      '</td>';

    // Stock on hand ships now, so it is shown ahead of whatever the factory
    // quotes behind it.
    var leadText = offer.stockSufficient === true
      ? 'In stock' + (offer.leadTimeText && offer.leadTimeText !== 'In stock'
        ? '<div class="rowmeta">' + esc(offer.leadTimeText) + ' factory</div>' : '')
      : (offer.leadTimeText ? esc(offer.leadTimeText) : '<span class="muted">—</span>');

    var leadCell = '<td class="num' + (bestLead ? ' best' : '') + '">' + leadText +
      (bestLead ? '<span class="best-tag">best</span>' : '') +
      '</td>';

    var unitCell = '<td class="num">' +
      (money(offer.unitPrice, offer.currency) || '<span class="muted">—</span>') +
      (offer.priceBreakQuantity > 1
        ? '<div class="rowmeta">@ ' + count(offer.priceBreakQuantity) + '+</div>'
        : '') +
      '</td>';

    var extCell = '<td class="num' + (bestPrice ? ' best' : '') + '">' +
      (money(offer.extendedPrice, offer.currency) || '<span class="muted">—</span>') +
      (bestPrice ? '<span class="best-tag">best</span>' : '') +
      // MOQ and packaging multiples can push the purchased quantity above what
      // the BOM asked for; only say so when it actually happens.
      (offer.orderQuantity > needed
        ? '<div class="rowmeta">buy ' + count(offer.orderQuantity) + '</div>'
        : '') +
      '</td>';

    var lifecycleCell = '<td>' + lifecycleBadge(offer) + '</td>';

    // An aggregator's column quotes one distributor, so say which one.
    if (offer.aggregator && offer.distributor) {
      var others = (offer.distributorCount || 1) - 1;
      stockCell = stockCell.replace('</td>',
        '<div class="rowmeta">' + esc(offer.distributor) +
        (others > 0 ? ' <span class="muted">+' + others + ' more</span>' : '') + '</div></td>');
    }

    return stockCell + leadCell + unitCell + extCell + lifecycleCell;
  }

  // "Not Recommended for New Designs" is four times the width of every other
  // status and would set the column's width on its own.
  var LIFECYCLE_SHORT = { 'Not Recommended for New Designs': 'NRND' };

  function lifecycleBadge(offer) {
    var full = offer.lifecycle || 'Unknown';
    var short = LIFECYCLE_SHORT[full] || full;
    return '<span class="badge ' + esc(offer.lifecycleSeverity) + '" title="' + esc(full) + '">' +
      esc(short) + '</span>';
  }

  function renderVerdict(comparison, suppliers) {
    var parts = [];
    if (comparison.recommendedSupplier) {
      parts.push('<span class="badge info">' + esc(comparison.recommendedSupplier) + '</span>');
    }
    var spread = money(comparison.priceSpread);
    if (suppliers.length > 1 && spread && comparison.priceSpread > 0) {
      parts.push('<div class="rowmeta">' + esc(spread) + ' cheaper' +
        (isFinite(comparison.priceSpreadPercent) && comparison.priceSpreadPercent !== null
          ? ' (' + comparison.priceSpreadPercent + '%)' : '') + '</div>');
    }
    if (comparison.flags.length) {
      parts.push('<div class="flags">' + comparison.flags.map(function (flag) {
        return '<span class="flag ' + esc(flag.level) + '">' + esc(flag.text) + '</span>';
      }).join('') + '</div>');
    }
    return parts.join('') || '<span class="muted">—</span>';
  }

  function renderDetail(row, suppliers) {
    var columns = suppliers.map(function (supplier) {
      var offer = row.offers[supplier.id];
      if (!offer || !offer.found) {
        return '<div class="detail-col"><h4>' + esc(supplier.name) + '</h4>' +
          '<div class="' + (offer && offer.error ? 'err-text' : 'miss') + '">' +
          esc((offer && offer.reason) || 'No match') + '</div></div>';
      }

      var facts = [
        ['Supplier P/N', offer.supplierPartNumber],
        ['Manufacturer', offer.manufacturer],
        ['Matched MPN', offer.manufacturerPartNumber + (offer.exactMatch ? '' : ' (closest match)')],
        ['Packaging', offer.packaging],
        ['Min. order', offer.minimumOrderQuantity],
        ['Order multiple', offer.orderMultiple],
        ['You would buy', count(offer.orderQuantity) + ' for ' + row.quantity + ' needed'],
        ['Stock (this pack)', offer.stock === null ? null : count(offer.stock)],
        ['Stock (all packs)', offer.totalStock === null ? null : count(offer.totalStock)],
        ['Factory stock', isFinite(offer.factoryStock) ? count(offer.factoryStock) : null],
        ['Lead time', offer.leadTimeText],
        ['Quoted as', offer.leadTimeRaw],
        ['Lifecycle', offer.lifecycle + (offer.lifecycleRaw && offer.lifecycleRaw !== offer.lifecycle
          ? ' (' + offer.lifecycleRaw + ')' : '')],
        ['RoHS', offer.rohs],
        ['Quoted via', offer.aggregator ? offer.distributor : null],
        ['Distributors', offer.aggregator ? offer.distributorCount : null],
        ['Lifecycle risk', offer.lifecycleRisk],
        ['Supply chain risk', offer.supplyChainRisk],
        ['US tariff', offer.affectedByTariff ? 'Affected' : null],
        ['Replacement', offer.suggestedReplacement],
        ['Description', offer.description],
      ].filter(function (pair) {
        return pair[1] !== null && pair[1] !== undefined && pair[1] !== '';
      });

      var dl = '<dl>' + facts.map(function (pair) {
        return '<dt>' + esc(pair[0]) + '</dt><dd>' + esc(pair[1]) + '</dd>';
      }).join('') + '</dl>';

      var breaks = '';
      if (offer.priceBreaks && offer.priceBreaks.length) {
        breaks = '<div class="breaks"><table><tr><th>Qty</th><th>Unit</th><th>Extended</th></tr>' +
          offer.priceBreaks.map(function (brk) {
            var applied = brk.quantity === offer.priceBreakQuantity;
            return '<tr class="' + (applied ? 'applied' : '') + '"><td>' + count(brk.quantity) +
              '</td><td>' + esc(money(brk.unitPrice, offer.currency)) + '</td><td>' +
              esc(money(brk.unitPrice * offer.orderQuantity, offer.currency)) + '</td></tr>';
          }).join('') + '</table></div>';
      }

      // Every distributor the aggregator found, which is the whole point of
      // including it alongside the single-distributor suppliers.
      var distributors = '';
      if (offer.distributorOffers && offer.distributorOffers.length) {
        distributors = '<div class="breaks dist"><table>' +
          '<tr><th class="l">Distributor</th><th class="l">P/N</th><th>Stock</th>' +
          '<th>MOQ</th><th>Unit</th><th>Extended</th></tr>' +
          offer.distributorOffers.map(function (entry) {
            var covers = entry.stockSufficient;
            return '<tr class="' + (covers ? 'covers' : '') + '">' +
              '<td class="l">' + esc(entry.distributor || '—') + '</td>' +
              '<td class="l">' + esc(entry.supplierPartNumber || '—') + '</td>' +
              '<td>' + (entry.stock === null || entry.stock === undefined
                ? esc(entry.availabilityText || '—') : count(entry.stock)) +
                (covers === false ? ' <span class="badge warn">short</span>' : '') + '</td>' +
              '<td>' + count(entry.minimumOrderQuantity) + '</td>' +
              '<td>' + esc(money(entry.unitPrice, entry.currency) || '—') + '</td>' +
              '<td>' + esc(money(entry.extendedPrice, entry.currency) || '—') + '</td>' +
              '</tr>';
          }).join('') + '</table>' +
          '<div class="rowmeta">' + offer.distributorOffers.length +
          ' authorized distributor' + (offer.distributorOffers.length === 1 ? '' : 's') +
          ' via ' + esc(offer.supplier) + '</div></div>';
      }

      var links = [];
      var productUrl = safeUrl(offer.productUrl);
      var datasheetUrl = safeUrl(offer.datasheetUrl);
      if (productUrl) {
        links.push('<a href="' + esc(productUrl) + '" target="_blank" rel="noopener noreferrer">Product page ↗</a>');
      }
      if (datasheetUrl) {
        links.push('<a href="' + esc(datasheetUrl) + '" target="_blank" rel="noopener noreferrer">Datasheet ↗</a>');
      }
      // The per-part attribution target their guidance asks for. Followable,
      // so no nofollow.
      if (offer.attribution) {
        var attrUrl = safeUrl(offer.attribution.url);
        if (attrUrl) {
          links.push('<a href="' + esc(attrUrl) + '" target="_blank" rel="noopener">' +
            esc(offer.attribution.text) + ' ' + esc(offer.attribution.name) + ' ↗</a>');
        }
      }

      return '<div class="detail-col"><h4>' + esc(supplier.name) + lifecycleBadge(offer) + '</h4>' +
        dl + breaks + distributors +
        (links.length ? '<div class="detail-links">' + links.join('') + '</div>' : '') +
        '</div>';
    }).join('');

    return alternateDetail(row) + '<div class="detail">' + columns + '</div>';
  }

  // What the BOM's alternates column promised, checked against the same
  // suppliers as the primary.
  function alternateDetail(row) {
    var entries = alternateEntries(row);
    if (!entries.length) return '';

    var rows = entries.map(function (entry) {
      var suppliers = (entry.comparison && entry.comparison.inStockSuppliers) || [];
      return '<tr class="' + (entry.usable ? 'covers' : '') + '">' +
        '<td class="l">' + esc(entry.mpn) + '</td>' +
        '<td>' + alternateBadge(entry) + '</td>' +
        '<td>' + (entry.found
          ? lifecycleBadge({
            lifecycle: entry.lifecycle,
            lifecycleSeverity: entry.lifecycleSeverity,
          })
          : '<span class="muted">—</span>') + '</td>' +
        '<td>' + (entry.stock === null || entry.stock === undefined
          ? '<span class="muted">—</span>' : count(entry.stock)) + '</td>' +
        '<td>' + esc(money(entry.bestPrice, bom().results.summary.currency) || '—') +
        (entry.bestPriceSupplier
          ? '<div class="rowmeta">' + esc(entry.bestPriceSupplier) + '</div>' : '') + '</td>' +
        '<td class="l">' + (suppliers.length
          ? esc(suppliers.join(', ')) : '<span class="muted">—</span>') + '</td>' +
        '</tr>';
    }).join('');

    var usable = usableAlternates(row).length;
    return '<div class="alt-detail"><h4>Approved alternates ' +
      '<span class="aside">from this BOM</span></h4>' +
      '<div class="breaks dist"><table>' +
      '<tr><th class="l">Part number</th><th>Verdict</th><th>Lifecycle</th>' +
      '<th>Stock</th><th>Best price</th><th class="l">In stock at</th></tr>' +
      rows + '</table>' +
      '<div class="rowmeta">' + (usable
        ? usable + ' of ' + entries.length + ' could cover this line at ' + count(row.quantity)
        : 'None of the ' + entries.length + ' listed alternate' +
          (entries.length === 1 ? '' : 's') + ' can cover this line today') +
      '</div></div></div>';
  }

  // ── CSV export ───────────────────────────────────────────────────────────

  function exportCsv() {
    if (!bom().results) return;
    var suppliers = bom().results.suppliers;
    var header = ['Row', 'Part Number', 'Quantity', 'Reference', 'Manufacturer', 'Description'];
    suppliers.forEach(function (supplier) {
      header.push(
        supplier.name + ' P/N',
        supplier.name + ' Stock',
        supplier.name + ' Lead Time',
        supplier.name + ' Lead Days',
        supplier.name + ' Unit Price',
        supplier.name + ' Extended Price',
        supplier.name + ' Order Qty',
        supplier.name + ' Lifecycle',
        supplier.name + ' Status'
      );
    });
    header.push('Cheapest Supplier', 'Soonest Supplier', 'Soonest (days)', 'Recommended',
      'Worst Lifecycle', 'Approved Alternates', 'Notes');

    var lines = [header];
    visibleRows().forEach(function (row) {
      var record = [row.row, row.mpn, row.quantity, row.reference, row.manufacturer, row.description];
      suppliers.forEach(function (supplier) {
        var offer = row.offers[supplier.id];
        if (!offer || !offer.found) {
          record.push('', '', '', '', '', '', '', '', (offer && offer.reason) || 'No match');
        } else {
          record.push(
            offer.supplierPartNumber,
            offer.stock,
            offer.leadTimeText,
            offer.leadTimeDays,
            offer.unitPrice,
            offer.extendedPrice,
            offer.orderQuantity,
            offer.lifecycle,
            'Found'
          );
        }
      });
      record.push(
        row.comparison.bestPriceSupplier,
        // Several suppliers can be equally fast, so all of them are listed.
        (row.comparison.bestLeadTimeSuppliers || []).join(' / '),
        row.comparison.bestLeadTimeDays,
        row.comparison.recommendedSupplier,
        row.comparison.lifecycle,
        alternateSummaryText(row),
        row.comparison.flags.map(function (f) { return f.text; }).join('; ')
      );
      lines.push(record);
    });

    var csv = lines.map(function (record) {
      return record.map(csvCell).join(',');
    }).join('\r\n');

    var slug = (bom().name || 'bom').replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '');
    download(
      new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' }),
      (slug || 'bom') + '-supplier-comparison-' + new Date().toISOString().slice(0, 10) + '.csv'
    );
    toast('Exported ' + (lines.length - 1) + ' rows');
  }

  function csvCell(value) {
    if (value === null || value === undefined) return '';
    var text = String(value);
    // A leading =, +, - or @ makes a spreadsheet treat the cell as a formula.
    if (/^[=+\-@]/.test(text)) text = "'" + text;
    if (/[",\r\n]/.test(text)) return '"' + text.replace(/"/g, '""') + '"';
    return text;
  }

  // What the server will screen out, worked out here so it shows up before the
  // Analyze button is pressed rather than after the lookups have run.
  function previewScreening(entry) {
    if (entry.adhoc) return null;
    var prefixes = (state.health && state.health.ignorePrefixes) || [];
    var owned = claimedParts(entry);
    var seen = {};
    var flagged = 0;
    var inhouse = 0;
    var repeats = 0;
    var elsewhere = 0;

    (entry.lines || []).forEach(function (line) {
      var mpn = normalizeMpn(line.mpn);
      if (!mpn) return;
      if (skipRequested(line.skip)) { flagged++; return; }
      var ignored = prefixes.some(function (prefix) { return mpn.indexOf(prefix) === 0; });
      if (ignored) { inhouse++; return; }
      if (owned[mpn]) { elsewhere++; return; }
      if (seen[mpn]) { repeats++; return; }
      seen[mpn] = true;
    });

    var pieces = [];
    if (flagged) pieces.push(flagged + ' marked skip to production');
    if (inhouse) pieces.push(inhouse + ' in-house');
    if (repeats) pieces.push(repeats + ' repeated');
    if (elsewhere) pieces.push(elsewhere + ' already in another BOM');
    if (!pieces.length) return null;
    var total = flagged + inhouse + repeats + elsewhere;
    return pieces.join(', ') + ' — ' + total +
      ' line' + (total === 1 ? '' : 's') + ' will be skipped.';
  }

  // Mirrors SKIP_VALUES in bomlib/prepare.py.
  var SKIP_VALUES = ['YES', 'Y', 'TRUE', 'T', '1', 'X', '\u2713', '\u2714'];

  function skipRequested(value) {
    return SKIP_VALUES.indexOf(cleanCell(value).toUpperCase()) !== -1;
  }

  // ── Cross-BOM part ownership ─────────────────────────────────────────────

  // A part number that turns up in three BOMs is still one part to buy and one
  // lookup to pay for. The first BOM analyzed keeps it; the rest are told it is
  // already covered and skip it. Ownership follows analysis order rather than
  // tab order, so whichever BOM you run first is always the one that resolves
  // the part — never a BOM that has not been looked up yet.
  function normalizeMpn(value) {
    return String(value === null || value === undefined ? '' : value)
      .toUpperCase().trim().replace(/\s+/g, ' ');
  }

  function claimedParts(except) {
    var map = {};
    state.boms.forEach(function (entry) {
      // A search answers a question about a part; it does not own it, or a
      // later BOM containing that part would come back empty.
      if (entry === except || entry.adhoc || !entry.results) return;
      (entry.claimed || []).forEach(function (mpn) {
        if (!map[mpn]) map[mpn] = entry.name;
      });
    });
    return map;
  }

  // Which other loaded BOMs list the same part, so the report can show the
  // demand that was folded into a single row.
  function usageIndex() {
    var index = {};
    state.boms.forEach(function (entry) {
      (entry.lines || []).forEach(function (line) {
        var key = normalizeMpn(line.mpn);
        if (!key) return;
        if (!index[key]) index[key] = [];
        index[key].push({ name: entry.name, quantity: line.quantity || 0, id: entry.id });
      });
    });
    return index;
  }

  // ── Skipped lines ────────────────────────────────────────────────────────

  var REASON_LABEL = {
    flagged: 'Marked skip',
    ignored: 'In-house',
    merged: 'Merged',
    duplicate: 'Other BOM',
  };

  function skippedSummary(list) {
    var counts = {};
    list.forEach(function (entry) {
      counts[entry.reason] = (counts[entry.reason] || 0) + 1;
    });
    var pieces = [];
    if (counts.flagged) pieces.push(counts.flagged + ' marked skip to production');
    if (counts.ignored) pieces.push(counts.ignored + ' in-house part number' + (counts.ignored === 1 ? '' : 's'));
    if (counts.merged) pieces.push(counts.merged + ' duplicate line' + (counts.merged === 1 ? '' : 's') + ' merged');
    if (counts.duplicate) pieces.push(counts.duplicate + ' already covered by another BOM');
    return pieces.join(' · ');
  }

  function renderSkipped() {
    var list = (bom().excluded || []);
    if (!list.length) {
      el.skippedNote.innerHTML = '';
      return;
    }
    var rows = list.map(function (entry) {
      return '<tr>' +
        '<td>' + esc(entry.row === null || entry.row === undefined ? '—' : entry.row) + '</td>' +
        '<td class="mpn-cell">' + esc(entry.mpn) + '</td>' +
        '<td>' + count(entry.quantity) + '</td>' +
        '<td><span class="reason-tag ' + esc(entry.reason) + '">' +
        esc(REASON_LABEL[entry.reason] || entry.reason) + '</span></td>' +
        '<td>' + esc(entry.detail || '') + '</td>' +
        '</tr>';
    }).join('');

    el.skippedNote.innerHTML =
      '<details class="skipped-note"><summary>' +
      esc(list.length + ' line' + (list.length === 1 ? '' : 's') + ' not looked up — ' +
        skippedSummary(list)) +
      '</summary><table class="skipped-table">' +
      '<thead><tr><th>Row</th><th>Part number</th><th>Qty</th><th>Reason</th><th>Detail</th></tr></thead>' +
      '<tbody>' + rows + '</tbody></table></details>';
  }

  // ── Summary report ───────────────────────────────────────────────────────

  function reportModel(entry) {
    var results = entry.results;
    var summary = results.summary;
    var currency = summary.currency || 'USD';
    var usage = usageIndex();

    var stockRisk = results.rows.filter(function (row) {
      return row.comparison.inStockSuppliers.length === 0;
    });
    var lifecycleRisk = results.rows.filter(function (row) {
      var severity = row.comparison.lifecycleSeverity;
      return severity === 'bad' || severity === 'warn';
    });
    var risky = results.rows.filter(function (row) {
      var found = results.suppliers.some(function (s) {
        return row.offers[s.id] && row.offers[s.id].found;
      });
      var severity = row.comparison.lifecycleSeverity;
      return !found || severity === 'bad' || severity === 'warn' ||
        row.comparison.inStockSuppliers.length === 0;
    });

    return {
      entry: entry,
      currency: currency,
      summary: summary,
      suppliers: results.suppliers,
      rows: results.rows,
      excluded: entry.excluded || [],
      usage: usage,
      stockRisk: stockRisk,
      lifecycleRisk: lifecycleRisk,
      risky: risky,
      // Only worth a column when the BOM actually carried one.
      hasAlternates: results.rows.some(function (row) {
        return alternateEntries(row).length > 0;
      }),
      generated: new Date().toLocaleString(),
    };
  }

  // What the other BOMs need of this part, one entry per BOM: a BOM that lists
  // the same part on two lines is still one BOM asking for the total.
  function otherBomDemand(model, mpn) {
    var totals = [];
    var byId = {};
    (model.usage[normalizeMpn(mpn)] || []).forEach(function (use) {
      if (use.id === model.entry.id) return;
      if (byId[use.id]) {
        byId[use.id].quantity += use.quantity;
        return;
      }
      byId[use.id] = { name: use.name, quantity: use.quantity };
      totals.push(byId[use.id]);
    });
    return totals;
  }

  function kpi(label, value, note, tone) {
    return '<div class="kpi ' + (tone || '') + '">' +
      '<div class="k-label">' + esc(label) + '</div>' +
      '<div class="k-value">' + esc(value) + '</div>' +
      '<div class="k-note">' + esc(note || '') + '</div></div>';
  }

  // The offer the verdict actually points at: the recommended supplier already
  // balances "soonest" against "cheapest among the soonest", so the report
  // prices that one rather than re-deciding.
  function recommendedOffer(row) {
    var name = row.comparison.recommendedSupplier;
    if (!name) return null;
    var found = null;
    Object.keys(row.offers).forEach(function (id) {
      var offer = row.offers[id];
      if (!found && offer && offer.found && offer.supplier === name) found = offer;
    });
    return found;
  }

  // One report cell: every approved alternate, each carrying whether it could
  // actually stand in today. A dash means the BOM named none for this line.
  function reportAlternateCell(row) {
    var entries = alternateEntries(row);
    if (!entries.length) return '<span class="muted">—</span>';
    return entries.map(function (entry) {
      return '<div class="alt-line">' + esc(entry.mpn) + ' ' + alternateBadge(entry) + '</div>';
    }).join('');
  }

  function reportHtml(model) {
    var currency = model.currency;
    var summary = model.summary;

    var kpis = [
      kpi('Lines', count(summary.lines), count(summary.totalQuantity) + ' units'),
      kpi('Best-mix total', money(summary.bestMixTotal, currency) || '—',
        summary.bestMixLines + ' of ' + summary.lines + ' priced', 'accent'),
      kpi('Stock risk', String(model.stockRisk.length),
        model.stockRisk.length ? 'no supplier covers the quantity' : 'all coverable today',
        model.stockRisk.length ? 'bad' : 'good'),
      kpi('Lifecycle risk', String(model.lifecycleRisk.length),
        model.lifecycleRisk.length ? 'NRND, EOL or obsolete' : 'nothing flagged',
        model.lifecycleRisk.length ? 'warn' : 'good'),
      kpi('Not found', String(summary.notFoundLines || 0),
        (summary.notFoundLines ? 'no supplier matched' : 'every line matched'),
        summary.notFoundLines ? 'bad' : 'good'),
      kpi('Skipped', String(model.excluded.length),
        model.excluded.length ? 'in-house or duplicate' : 'nothing skipped'),
    ].join('');

    // Supplier carts.
    var cartRows = model.suppliers.map(function (supplier) {
      var totals = summary.supplierTotals[supplier.id];
      if (!totals) return '';
      var winner = summary.cheapestSingleSource === supplier.id;
      return '<tr class="' + (winner ? 'winner' : '') + '">' +
        '<td><strong>' + esc(supplier.name) + '</strong>' +
        (winner ? ' <span class="badge info">cheapest cart</span>' : '') + '</td>' +
        '<td class="num">' + count(totals.linesPriced) + '</td>' +
        '<td class="num">' + count(totals.linesMissing) + '</td>' +
        '<td class="num">' + count(totals.linesShort) + '</td>' +
        '<td class="num">' + esc(money(totals.total, currency) || '—') + '</td></tr>';
    }).join('');

    if (model.suppliers.length > 1) {
      var savings = summary.mixSavings;
      cartRows += '<tr class="winner"><td><strong>Cheapest line by line</strong></td>' +
        '<td class="num">' + count(summary.bestMixLines) + '</td>' +
        '<td class="num">—</td><td class="num">—</td>' +
        '<td class="num">' + esc(money(summary.bestMixTotal, currency) || '—') + '</td></tr>';
      if (isFinite(savings) && savings > 0) {
        cartRows += '<tr><td colspan="5" class="desc">Splitting the order across suppliers ' +
          'saves ' + esc(money(savings, currency)) + ' against the cheapest single cart.</td></tr>';
      }
    }

    // Lines that need a decision.
    var riskHtml;
    if (model.risky.length) {
      riskHtml = '<div class="report-scroll"><table class="report-table">' +
        '<thead><tr><th>Part</th><th class="num">Qty</th><th>Lifecycle</th><th>Issue</th>' +
        (model.hasAlternates ? '<th>Approved alternate</th>' : '') +
        '</tr></thead><tbody>' +
        model.risky.map(function (row) {
          return '<tr><td class="mpn-cell">' + esc(row.mpn) +
            (row.reference ? '<div class="desc">' + esc(row.reference) + '</div>' : '') + '</td>' +
            '<td class="num">' + count(row.quantity) + '</td>' +
            '<td>' + lifecycleBadge({
              lifecycle: row.comparison.lifecycle,
              lifecycleSeverity: row.comparison.lifecycleSeverity,
            }) + '</td>' +
            '<td>' + (row.comparison.flags.length
              ? row.comparison.flags.map(function (flag) {
                return '<span class="flag ' + esc(flag.level) + '">' + esc(flag.text) + '</span>';
              }).join(' ')
              : '<span class="muted">—</span>') + '</td>' +
            (model.hasAlternates ? '<td>' + reportAlternateCell(row) + '</td>' : '') +
            '</tr>';
        }).join('') + '</tbody></table></div>';
    } else {
      riskHtml = '<div class="report-empty">Every line is in stock, priced and in production.</div>';
    }

    // The parts themselves, with only the columns a buyer acts on.
    var multi = state.boms.length > 1;
    var partRows = model.rows.map(function (row) {
      var offer = recommendedOffer(row);
      var lead = '—';
      if (offer) {
        lead = offer.stockSufficient === true ? 'In stock' : (offer.leadTimeText || '—');
      }
      var elsewhere = otherBomDemand(model, row.mpn);
      return '<tr>' +
        '<td class="mpn-cell">' + esc(row.mpn) +
        (row.description ? '<div class="desc">' + esc(row.description) + '</div>' : '') + '</td>' +
        '<td class="num">' + count(row.quantity) + '</td>' +
        (multi ? '<td class="desc">' + (elsewhere.length
          ? esc(elsewhere.map(function (use) {
            return use.name + ' (' + count(use.quantity) + ')';
          }).join(', '))
          : '—') + '</td>' : '') +
        '<td>' + esc(row.comparison.recommendedSupplier || '—') +
        (offer && offer.aggregator && offer.distributor
          ? '<div class="desc">via ' + esc(offer.distributor) + '</div>' : '') + '</td>' +
        '<td class="num">' + esc((offer && money(offer.unitPrice, offer.currency)) || '—') + '</td>' +
        '<td class="num">' + esc((offer && money(offer.extendedPrice, offer.currency)) || '—') + '</td>' +
        '<td>' + esc(lead) + '</td>' +
        '<td>' + lifecycleBadge({
          lifecycle: row.comparison.lifecycle,
          lifecycleSeverity: row.comparison.lifecycleSeverity,
        }) + '</td>' +
        (model.hasAlternates ? '<td>' + reportAlternateCell(row) + '</td>' : '') +
        '</tr>';
    }).join('');

    var partsHtml = '<div class="report-scroll"><table class="report-table">' +
      '<thead><tr><th>Part</th><th class="num">Qty</th>' +
      (multi ? '<th>Also in</th>' : '') +
      '<th>Buy from</th><th class="num">Unit</th><th class="num">Extended</th>' +
      '<th>Lead time</th><th>Lifecycle</th>' +
      (model.hasAlternates ? '<th>Approved alternates</th>' : '') +
      '</tr></thead><tbody>' + partRows +
      '</tbody></table></div>';

    // Skipped lines.
    var skippedHtml = '';
    if (model.excluded.length) {
      skippedHtml = '<section class="report-section"><h3>Not looked up ' +
        '<span class="aside">' + esc(skippedSummary(model.excluded)) + '</span></h3>' +
        '<div class="report-scroll"><table class="report-table">' +
        '<thead><tr><th>Row</th><th>Part number</th><th class="num">Qty</th><th>Reason</th></tr></thead><tbody>' +
        model.excluded.map(function (entry) {
          return '<tr><td>' + esc(entry.row === null || entry.row === undefined ? '—' : entry.row) + '</td>' +
            '<td class="mpn-cell">' + esc(entry.mpn) + '</td>' +
            '<td class="num">' + count(entry.quantity) + '</td>' +
            '<td>' + esc(entry.detail || REASON_LABEL[entry.reason] || '') + '</td></tr>';
        }).join('') + '</tbody></table></div></section>';
    }

    var analyzed = state.boms.filter(function (b) { return b.results; });
    var allButton = analyzed.length > 1
      ? '<button type="button" class="btn ghost small" data-report="excel-all">Excel · all ' +
        analyzed.length + ' BOMs</button>'
      : '';

    return '<div class="report-sheet">' +
      '<div class="report-head">' +
      '<div><h2 id="reportTitle">' + esc(model.entry.name) + '</h2>' +
      '<div class="sub">Supplier report · ' + esc(model.generated) + ' · prices in ' +
      esc(currency) + '</div></div>' +
      '<div class="report-actions">' +
      '<button type="button" class="btn primary small" data-report="excel">Export Excel</button>' +
      allButton +
      '<button type="button" class="btn ghost small" data-report="print">Print / PDF</button>' +
      '<button type="button" class="icon-btn" data-report="close" aria-label="Close report">✕</button>' +
      '</div></div>' +
      '<div class="report-body">' +
      '<section class="report-section"><h3>Overview</h3><div class="kpi-grid">' + kpis + '</div></section>' +
      '<section class="report-section"><h3>What each supplier would cost</h3>' +
      '<div class="report-scroll"><table class="report-table">' +
      '<thead><tr><th>Supplier</th><th class="num">Lines quoted</th><th class="num">Not carried</th>' +
      '<th class="num">Short on stock</th><th class="num">Cart total</th></tr></thead>' +
      '<tbody>' + cartRows + '</tbody></table></div></section>' +
      '<section class="report-section"><h3>Needs a decision ' +
      '<span class="aside">' + model.risky.length + ' of ' + model.rows.length + ' lines</span></h3>' +
      riskHtml + '</section>' +
      '<section class="report-section"><h3>Parts <span class="aside">' + model.rows.length +
      ' lines</span></h3>' + partsHtml + '</section>' +
      skippedHtml +
      '<div class="report-foot">Prices, stock and lifecycle status were read live from the supplier ' +
      'APIs at the time shown above and move constantly &mdash; confirm on the supplier&rsquo;s own page ' +
      'before raising a purchase order.</div>' +
      '</div></div>';
  }

  var reportReturnFocus = null;

  function openReport() {
    var entry = bom();
    if (!entry.results) {
      toast('Analyze this BOM first', true);
      return;
    }
    reportReturnFocus = document.activeElement;
    el.reportOverlay.innerHTML = reportHtml(reportModel(entry));
    el.reportOverlay.hidden = false;
    document.body.style.overflow = 'hidden';

    Array.prototype.forEach.call(
      el.reportOverlay.querySelectorAll('[data-report]'),
      function (button) {
        button.addEventListener('click', function () {
          var action = button.getAttribute('data-report');
          if (action === 'close') closeReport();
          else if (action === 'print') window.print();
          else if (action === 'excel') exportWorkbook([entry]);
          else if (action === 'excel-all') {
            exportWorkbook(state.boms.filter(function (b) { return b.results; }));
          }
        });
      }
    );

    var close = el.reportOverlay.querySelector('[data-report="close"]');
    if (close) close.focus();
  }

  function closeReport() {
    el.reportOverlay.hidden = true;
    el.reportOverlay.innerHTML = '';
    document.body.style.overflow = '';
    if (reportReturnFocus && reportReturnFocus.focus) reportReturnFocus.focus();
    reportReturnFocus = null;
  }

  // ── DMSMS case form ──────────────────────────────────────────────────────
  //
  // A part can sit on several boards and belong to one program, so which parts
  // go on a form is a decision only the analyst can make. Everything at risk is
  // listed, the parts that are actually gone are ticked, and the rest is theirs.

  var DMSMS_FIELDS = [
    { key: 'program', label: 'Program / platform', placeholder: 'e.g. Falcon II', required: true },
    { key: 'caseNumber', label: 'DMSMS case number', placeholder: 'optional' },
    { key: 'preparedBy', label: 'Prepared by', placeholder: 'name' },
    { key: 'organization', label: 'Organization', placeholder: 'group or division' },
    { key: 'contract', label: 'Contract number', placeholder: 'optional' },
    { key: 'cage', label: 'CAGE code', placeholder: 'optional' },
  ];

  // Everything except the program persists: the program is the one field that
  // changes every time, and pre-filling it would be how the wrong name ends up
  // on a form.
  var DMSMS_STORAGE_KEY = 'bom-analyzer.dmsms';

  function loadDmsmsMeta() {
    var stored = {};
    try {
      stored = JSON.parse(localStorage.getItem(DMSMS_STORAGE_KEY) || '{}') || {};
    } catch (err) {
      stored = {};
    }
    stored.program = '';
    return stored;
  }

  function saveDmsmsMeta(meta) {
    try {
      var keep = {};
      DMSMS_FIELDS.forEach(function (field) {
        if (field.key !== 'program' && meta[field.key]) keep[field.key] = meta[field.key];
      });
      localStorage.setItem(DMSMS_STORAGE_KEY, JSON.stringify(keep));
    } catch (err) {
      // Private-mode browsers block storage; the form still generates.
    }
  }

  function dmsmsStatuses() {
    var health = state.health && state.health.dmsms;
    return {
      qualifying: (health && health.statuses) || [],
      ticked: (health && health.defaultSelected) || [],
    };
  }

  // Every at-risk line across every analyzed BOM, because a program's parts do
  // not stop at one board.
  function dmsmsCandidates() {
    var vocabulary = dmsmsStatuses();
    var found = [];
    state.boms.forEach(function (entry) {
      if (!entry.results) return;
      entry.results.rows.forEach(function (row) {
        var status = row.comparison.lifecycle;
        if (vocabulary.qualifying.indexOf(status) === -1) return;
        found.push({
          key: entry.id + '::' + row.index,
          bom: entry.name,
          bomId: entry.id,
          row: row,
          status: status,
          severity: row.comparison.lifecycleSeverity,
          ticked: vocabulary.ticked.indexOf(status) !== -1,
        });
      });
    });
    return found;
  }

  var dmsmsState = null;

  function openDmsms() {
    var candidates = dmsmsCandidates();
    if (!candidates.length) {
      toast(state.boms.some(function (b) { return b.results; })
        ? 'Nothing analyzed is obsolete, end of life or NRND'
        : 'Analyze a BOM first', true);
      return;
    }

    dmsmsState = {
      candidates: candidates,
      selected: {},
      meta: loadDmsmsMeta(),
      returnFocus: document.activeElement,
    };
    candidates.forEach(function (candidate) {
      dmsmsState.selected[candidate.key] = candidate.ticked;
    });

    el.dmsmsOverlay.hidden = false;
    document.body.style.overflow = 'hidden';
    renderDmsms();
  }

  function closeDmsms() {
    el.dmsmsOverlay.hidden = true;
    el.dmsmsOverlay.innerHTML = '';
    document.body.style.overflow = '';
    var focus = dmsmsState && dmsmsState.returnFocus;
    if (focus && focus.focus) focus.focus();
    dmsmsState = null;
  }

  function dmsmsSelectedRows() {
    return dmsmsState.candidates.filter(function (candidate) {
      return dmsmsState.selected[candidate.key];
    });
  }

  function renderDmsms() {
    var byBom = [];
    var index = {};
    dmsmsState.candidates.forEach(function (candidate) {
      if (!index[candidate.bomId]) {
        index[candidate.bomId] = { name: candidate.bom, id: candidate.bomId, rows: [] };
        byBom.push(index[candidate.bomId]);
      }
      index[candidate.bomId].rows.push(candidate);
    });

    var fields = DMSMS_FIELDS.map(function (field) {
      return '<label class="field">' + esc(field.label) + (field.required ? ' *' : '') +
        '<input type="text" data-meta="' + field.key + '" value="' +
        esc(dmsmsState.meta[field.key] || '') + '" placeholder="' + esc(field.placeholder) +
        '" autocomplete="off" /></label>';
    }).join('');

    var groups = byBom.map(function (group) {
      var rows = group.rows.map(function (candidate) {
        var row = candidate.row;
        var checked = dmsmsState.selected[candidate.key] ? ' checked' : '';
        return '<tr data-row="' + esc(candidate.key) + '" class="' + (checked ? 'picked' : '') + '">' +
          '<td class="tick"><input type="checkbox" data-pick="' + esc(candidate.key) + '"' +
          checked + ' aria-label="Include ' + esc(row.mpn) + '" /></td>' +
          '<td class="mpn-cell">' + esc(row.mpn) +
          (row.description ? '<div class="desc">' + esc(row.description) + '</div>' : '') + '</td>' +
          '<td class="desc">' + esc(row.reference || '—') + '</td>' +
          '<td class="num">' + count(row.quantity) + '</td>' +
          '<td>' + lifecycleBadge({
            lifecycle: candidate.status, lifecycleSeverity: candidate.severity,
          }) + '</td>' +
          '<td class="num">' + count(dmsmsStock(row)) + '</td>' +
          '<td>' + dmsmsReplacementCell(row) + '</td>' +
          '<td><span class="risk ' + esc(dmsmsRisk(row).toLowerCase()) + '">' +
          esc(dmsmsRisk(row)) + '</span></td>' +
          '</tr>';
      }).join('');

      return '<div class="dmsms-group">' +
        '<div class="dmsms-group-head">' +
        '<strong>' + esc(group.name) + '</strong>' +
        '<span class="aside" data-count="' + esc(group.id) + '"></span>' +
        '<button type="button" class="btn ghost small" data-group="' + esc(group.id) + '"></button>' +
        '</div>' +
        '<div class="report-scroll"><table class="report-table dmsms-table">' +
        '<thead><tr><th class="tick"></th><th>Part</th><th>Reference</th>' +
        '<th class="num">Qty</th><th>Status</th><th class="num">Stock</th>' +
        '<th>Replacement</th><th>Suggested risk</th></tr></thead><tbody>' +
        rows + '</tbody></table></div></div>';
    }).join('');

    el.dmsmsOverlay.innerHTML =
      '<div class="report-sheet">' +
      '<div class="report-head">' +
      '<div><h2 id="dmsmsTitle">DMSMS case form</h2>' +
      '<div class="sub">One form per program &mdash; tick the parts that belong to it</div></div>' +
      '<div class="report-actions">' +
      '<button type="button" class="btn primary small" data-dmsms="build"></button>' +
      '<button type="button" class="icon-btn" data-dmsms="close" aria-label="Close">&times;</button>' +
      '</div></div>' +
      '<div class="report-body">' +
      '<section class="report-section"><h3>Case details</h3>' +
      '<div class="grid">' + fields + '</div>' +
      '<label class="field">Notes<input type="text" data-meta="notes" value="' +
      esc(dmsmsState.meta.notes || '') + '" placeholder="optional" autocomplete="off" /></label>' +
      '</section>' +
      '<section class="report-section"><h3>Parts at risk ' +
      '<span class="aside" data-count="total"></span></h3>' +
      '<div class="btn-row compact"><button type="button" class="btn ghost small" ' +
      'data-dmsms="all">Select all</button><button type="button" class="btn ghost small" ' +
      'data-dmsms="none">Select none</button><button type="button" class="btn ghost small" ' +
      'data-dmsms="obsolete">Only obsolete &amp; EOL</button></div>' +
      groups + '</section>' +
      '<div class="report-foot">The form lists what the supplier APIs reported today. ' +
      'CAGE code, lifetime buy quantity, last-time-buy date, resolution and disposition are ' +
      'left blank for you &mdash; they are decisions, not lookups.</div>' +
      '</div></div>';

    wireDmsms();
    syncDmsms();
  }

  // Ticking a box changes four things on screen and nothing else. Re-rendering
  // the panel for it would throw away focus and scroll position, which on a
  // list of a hundred at-risk parts is the difference between tabbing through
  // and starting again after every tick.
  function syncDmsms() {
    var chosen = dmsmsSelectedRows().length;
    var total = dmsmsState.candidates.length;

    var build = el.dmsmsOverlay.querySelector('[data-dmsms="build"]');
    if (build) build.textContent = 'Generate form (' + chosen + ')';

    var overall = el.dmsmsOverlay.querySelector('[data-count="total"]');
    if (overall) overall.textContent = chosen + ' of ' + total + ' selected';

    Array.prototype.forEach.call(
      el.dmsmsOverlay.querySelectorAll('[data-row]'),
      function (tr) {
        var on = !!dmsmsState.selected[tr.getAttribute('data-row')];
        tr.classList.toggle('picked', on);
        var box = tr.querySelector('[data-pick]');
        if (box && box.checked !== on) box.checked = on;
      }
    );

    Array.prototype.forEach.call(
      el.dmsmsOverlay.querySelectorAll('[data-group]'),
      function (button) {
        var id = button.getAttribute('data-group');
        var rows = dmsmsState.candidates.filter(function (c) { return c.bomId === id; });
        var picked = rows.filter(function (c) { return dmsmsState.selected[c.key]; }).length;
        button.textContent = picked === rows.length ? 'Clear this BOM' : 'Select this BOM';
        var aside = el.dmsmsOverlay.querySelector('[data-count="' + id + '"]');
        if (aside) aside.textContent = picked + ' of ' + rows.length + ' selected';
      }
    );
  }

  // Whether anything has been found to put in the form's Suggested Replacement
  // column yet, so it is visible that running Find alternatives fills it in.
  // Same order of trust as bomlib/dmsms.py: a BOM alternate is a decision an
  // engineer already made, so it outranks anything an algorithm suggests.
  function dmsmsReplacementCell(row) {
    var approved = usableAlternates(row);
    if (!approved.length) approved = alternateEntries(row);
    if (approved.length) {
      return '<span class="mpn-cell">' + esc(approved[0].mpn) + '</span>' +
        '<div class="desc">BOM alternate' +
        (approved[0].usable ? '' : ', unavailable') +
        (approved.length > 1 ? ' +' + (approved.length - 1) + ' more' : '') + '</div>';
    }
    var found = state.alternatives[normalizeMpn(row.mpn)];
    if (found && found.length) {
      return '<span class="mpn-cell">' + esc(found[0].mpn) + '</span>' +
        (found.length > 1 ? '<span class="muted"> +' + (found.length - 1) + '</span>' : '');
    }
    var offered = null;
    Object.keys(row.offers).forEach(function (id) {
      var offer = row.offers[id];
      if (!offered && offer && offer.suggestedReplacement) offered = offer.suggestedReplacement;
    });
    if (offered) return '<span class="mpn-cell">' + esc(offered) + '</span>';
    return '<span class="muted">—</span>';
  }

  function dmsmsStock(row) {
    var total = null;
    Object.keys(row.offers).forEach(function (id) {
      var offer = row.offers[id];
      if (!offer || !offer.found) return;
      var held = offer.totalStock;
      if (held === null || held === undefined) held = offer.stock;
      if (isFinite(held) && held !== null) total = (total || 0) + held;
    });
    return total;
  }

  // Mirrors bomlib/dmsms.py so the screen and the workbook agree.
  function dmsmsRisk(row) {
    var status = row.comparison.lifecycle;
    var needed = row.quantity || 0;
    var stock = dmsmsStock(row);
    var covered = stock !== null && needed && stock >= needed;

    if (status === 'Obsolete' || status === 'Discontinued' || status === 'End of Life') {
      return covered ? 'Medium' : 'High';
    }
    if (status === 'Last Time Buy') return covered ? 'Medium' : 'High';
    if (status === 'Not Recommended for New Designs') return covered ? 'Low' : 'Medium';
    return 'Low';
  }

  function wireDmsms() {
    Array.prototype.forEach.call(
      el.dmsmsOverlay.querySelectorAll('[data-pick]'),
      function (box) {
        box.addEventListener('change', function () {
          dmsmsState.selected[box.getAttribute('data-pick')] = box.checked;
          syncDmsms();
        });
      }
    );

    Array.prototype.forEach.call(
      el.dmsmsOverlay.querySelectorAll('[data-group]'),
      function (button) {
        button.addEventListener('click', function () {
          var id = button.getAttribute('data-group');
          var rows = dmsmsState.candidates.filter(function (c) { return c.bomId === id; });
          var allOn = rows.every(function (c) { return dmsmsState.selected[c.key]; });
          rows.forEach(function (c) { dmsmsState.selected[c.key] = !allOn; });
          syncDmsms();
        });
      }
    );

    // Typing into a field must not repaint the panel underneath the cursor,
    // so these write straight to the model and never re-render.
    Array.prototype.forEach.call(
      el.dmsmsOverlay.querySelectorAll('[data-meta]'),
      function (input) {
        input.addEventListener('input', function () {
          dmsmsState.meta[input.getAttribute('data-meta')] = input.value;
        });
      }
    );

    Array.prototype.forEach.call(
      el.dmsmsOverlay.querySelectorAll('[data-dmsms]'),
      function (button) {
        button.addEventListener('click', function () {
          var action = button.getAttribute('data-dmsms');
          if (action === 'close') return closeDmsms();
          if (action === 'build') return buildDmsms();
          var ticked = dmsmsStatuses().ticked;
          dmsmsState.candidates.forEach(function (candidate) {
            if (action === 'all') dmsmsState.selected[candidate.key] = true;
            else if (action === 'none') dmsmsState.selected[candidate.key] = false;
            else dmsmsState.selected[candidate.key] = ticked.indexOf(candidate.status) !== -1;
          });
          syncDmsms();
        });
      }
    );
  }

  function buildDmsms() {
    var chosen = dmsmsSelectedRows();
    if (!chosen.length) {
      toast('Tick at least one part for the form', true);
      return;
    }
    if (!String(dmsmsState.meta.program || '').trim()) {
      toast('Name the program this form is for', true);
      var input = el.dmsmsOverlay.querySelector('[data-meta="program"]');
      if (input) input.focus();
      return;
    }

    saveDmsmsMeta(dmsmsState.meta);
    var scope = [];
    chosen.forEach(function (candidate) {
      if (scope.indexOf(candidate.bom) === -1) scope.push(candidate.bom);
    });

    var meta = {};
    Object.keys(dmsmsState.meta).forEach(function (key) { meta[key] = dmsmsState.meta[key]; });
    meta.scope = scope.join(', ');
    meta.date = meta.date || new Date().toISOString().slice(0, 10);

    toast('Building the DMSMS form…');
    fetch(api('/api/dmsms'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        meta: meta,
        // Which board a part sits on is the form's Next Higher Assembly, and
        // only the browser knows which BOM each selected row came from.
        rows: chosen.map(function (candidate) {
          var extra = { assembly: candidate.bom };
          var found = state.alternatives[normalizeMpn(candidate.row.mpn)];
          if (found && found.length) {
            // The first two: a case form wants a lead to follow, not a
            // catalogue. The rest stay in the alternatives panel.
            extra.suggestedReplacement = found.slice(0, 2).map(function (alt) {
              return alt.mpn + (alt.manufacturer ? ' (' + alt.manufacturer + ')' : '');
            }).join('; ');
          }
          return Object.assign({}, candidate.row, extra);
        }),
      }),
    })
      .then(function (res) {
        if (!res.ok) return readJsonOrThrow(res);
        return res.blob().then(function (blob) {
          var slug = String(meta.program).replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '');
          download(blob, (slug || 'program') + '-dmsms-' +
            new Date().toISOString().slice(0, 10) + '.xlsx');
          toast('DMSMS form for ' + meta.program + ' — ' + chosen.length + ' parts');
        });
      })
      .catch(function (err) {
        toast(err.message || 'Could not build the form', true);
      });
  }

  // ── Alternative parts ────────────────────────────────────────────────────
  //
  // Nexar answers a different question from the three suppliers: not "what does
  // this cost" but "what could I use instead". It runs only for parts already
  // found to be in trouble, and only when asked, so a free-tier quota is spent
  // on the parts that need it rather than on every line of a healthy BOM.

  var altState = null;

  // A part is worth asking about when its lifecycle says its supply is ending,
  // or when nobody can supply it today whatever its status says.
  function altCandidates() {
    var found = [];
    state.boms.forEach(function (entry) {
      if (!entry.results) return;
      entry.results.rows.forEach(function (row) {
        var comparison = row.comparison;
        var carried = entry.results.suppliers.some(function (s) {
          return row.offers[s.id] && row.offers[s.id].found;
        });
        var reason = null;
        if (!carried) reason = 'No supplier carries it';
        else if (comparison.lifecycleSeverity === 'bad') reason = comparison.lifecycle;
        else if (comparison.lifecycleSeverity === 'warn') reason = comparison.lifecycle;
        else if (!comparison.inStockSuppliers.length) reason = 'Nobody holds the quantity';
        if (!reason) return;

        found.push({
          key: entry.id + '::' + row.index,
          bom: entry.name,
          bomId: entry.id,
          row: row,
          reason: reason,
          // Obsolete and friends are ticked; a part that is merely short on
          // stock is listed, because that is a judgement call.
          ticked: comparison.lifecycleSeverity === 'bad' || !carried,
        });
      });
    });
    return found;
  }

  function openAlternatives() {
    var provider = (state.health && state.health.alternatives) || {};
    if (!provider.configured) {
      toast((provider.provider || 'Nexar') + ' is not configured — add the credentials to .env', true);
      return;
    }

    var candidates = altCandidates();
    if (!candidates.length) {
      toast(state.boms.some(function (b) { return b.results; })
        ? 'Nothing analyzed needs an alternative'
        : 'Analyze a BOM first', true);
      return;
    }

    altState = {
      provider: provider.provider || 'Nexar',
      maxParts: provider.maxParts || 50,
      candidates: candidates,
      selected: {},
      answers: {},
      running: false,
      stats: null,
      returnFocus: document.activeElement,
    };
    candidates.forEach(function (candidate) {
      altState.selected[candidate.key] = candidate.ticked;
    });

    el.altOverlay.hidden = false;
    document.body.style.overflow = 'hidden';
    renderAlternatives();
  }

  function closeAlternatives() {
    el.altOverlay.hidden = true;
    el.altOverlay.innerHTML = '';
    document.body.style.overflow = '';
    var focus = altState && altState.returnFocus;
    if (focus && focus.focus) focus.focus();
    altState = null;
  }

  function altSelected() {
    return altState.candidates.filter(function (c) { return altState.selected[c.key]; });
  }

  function renderAlternatives() {
    var rows = altState.candidates.map(function (candidate) {
      var row = candidate.row;
      var answer = altState.answers[row.mpn];
      var checked = altState.selected[candidate.key] ? ' checked' : '';
      return '<tr data-alt-row="' + esc(candidate.key) + '" class="' + (checked ? 'picked' : '') + '">' +
        '<td class="tick"><input type="checkbox" data-alt-pick="' + esc(candidate.key) + '"' +
        checked + ' aria-label="Look up alternatives for ' + esc(row.mpn) + '" /></td>' +
        '<td class="mpn-cell">' + esc(row.mpn) +
        '<div class="desc">' + esc(candidate.bom) +
        (row.description ? ' · ' + esc(row.description) : '') + '</div></td>' +
        '<td class="num">' + count(row.quantity) + '</td>' +
        '<td>' + lifecycleBadge({
          lifecycle: row.comparison.lifecycle,
          lifecycleSeverity: row.comparison.lifecycleSeverity,
        }) + '</td>' +
        '<td class="desc">' + esc(candidate.reason) + '</td>' +
        '<td>' + altAnswerCell(answer) + '</td>' +
        '</tr>';
    }).join('');

    var panels = altState.candidates.map(function (candidate) {
      var answer = altState.answers[candidate.row.mpn];
      if (!answer || answer.error || !(answer.alternatives || []).length) return '';
      return renderAltPanel(candidate.row.mpn, answer);
    }).join('');

    el.altOverlay.innerHTML =
      '<div class="report-sheet">' +
      '<div class="report-head">' +
      '<div><h2 id="altTitle">Alternative parts</h2>' +
      '<div class="sub">Asked of ' + esc(altState.provider) +
      ' &mdash; only for the parts you tick</div></div>' +
      '<div class="report-actions">' +
      '<button type="button" class="btn primary small" data-alt="run"' +
      (altState.running ? ' disabled' : '') + '></button>' +
      '<button type="button" class="icon-btn" data-alt="close" aria-label="Close">&times;</button>' +
      '</div></div>' +
      '<div class="report-body">' +
      altFailureNotice() +
      '<section class="report-section"><h3>Parts worth replacing ' +
      '<span class="aside" data-alt-count></span></h3>' +
      '<div class="btn-row compact">' +
      '<button type="button" class="btn ghost small" data-alt="all">Select all</button>' +
      '<button type="button" class="btn ghost small" data-alt="none">Select none</button>' +
      '</div>' +
      '<div class="report-scroll"><table class="report-table dmsms-table">' +
      '<thead><tr><th class="tick"></th><th>Part</th><th class="num">Qty</th>' +
      '<th>Status</th><th>Why</th><th>Alternatives</th></tr></thead>' +
      '<tbody>' + rows + '</tbody></table></div></section>' +
      (panels ? '<section class="report-section"><h3>What ' + esc(altState.provider) +
        ' suggests</h3>' + panels + '</section>' : '') +
      '<div class="report-foot">Alternatives are ' + esc(altState.provider) +
      '&rsquo;s own suggestions, matched on the part it found for your number &mdash; check the ' +
      'specification and the datasheet before designing one in. Nothing here has been checked ' +
      'against your board.</div>' +
      '</div></div>';

    wireAlternatives();
    syncAlternatives();
  }

  // Credentials failing is one problem, not one problem per part. When every
  // answer carries the same message it belongs at the top, said once, where
  // there is room to read it.
  function altFailureNotice() {
    var answers = Object.keys(altState.answers).map(function (k) { return altState.answers[k]; });
    if (!answers.length) return '';
    var errors = answers.filter(function (a) { return a.error; });
    if (errors.length !== answers.length) return '';
    var shared = errors[0].error;
    if (!errors.every(function (a) { return a.error === shared; })) return '';

    return '<section class="report-section"><div class="alt-failure">' +
      '<strong>' + esc(altState.provider) + ' could not answer</strong>' +
      '<p>' + esc(shared) + '</p></div></section>';
  }

  function altAnswerCell(answer) {
    if (!answer) return '<span class="muted">—</span>';
    if (answer.error) {
      return '<span class="err-text">' + esc(truncateText(answer.error, 90)) + '</span>';
    }
    var n = (answer.alternatives || []).length;
    if (!n) return '<span class="miss">none found</span>';
    return '<span class="badge info">' + n + '</span>';
  }

  function renderAltPanel(mpn, answer) {
    var matched = answer.matched;
    var rows = (answer.alternatives || []).map(function (alt) {
      var url = safeUrl(alt.url);
      var datasheet = safeUrl(alt.datasheetUrl);
      var links = [];
      if (url) links.push('<a href="' + esc(url) + '" target="_blank" rel="noopener noreferrer">Part ↗</a>');
      if (datasheet) links.push('<a href="' + esc(datasheet) + '" target="_blank" rel="noopener noreferrer">Datasheet ↗</a>');
      var specs = (alt.specs || []).slice(0, 6).map(function (spec) {
        return '<span class="spec">' + esc(spec.name) + ' <b>' + esc(spec.value) + '</b></span>';
      }).join('');

      return '<tr>' +
        '<td class="mpn-cell">' + esc(alt.mpn) +
        (alt.description ? '<div class="desc">' + esc(alt.description) + '</div>' : '') +
        (specs ? '<div class="specs">' + specs + '</div>' : '') + '</td>' +
        '<td>' + esc(alt.manufacturer || '—') + '</td>' +
        '<td class="num">' + (alt.stock === null || alt.stock === undefined
          ? '<span class="muted">—</span>' : count(alt.stock)) + '</td>' +
        '<td class="num">' + esc(money(alt.medianPrice, alt.currency) || '—') + '</td>' +
        '<td class="num">' + (isFinite(alt.leadDays) && alt.leadDays !== null
          ? count(alt.leadDays) + ' d' : '<span class="muted">—</span>') + '</td>' +
        '<td class="desc">' + (links.join(' · ') || '—') + '</td>' +
        '</tr>';
    }).join('');

    return '<div class="alt-group">' +
      '<div class="alt-head"><strong>' + esc(mpn) + '</strong>' +
      (matched ? '<span class="aside">matched ' + esc(matched.mpn) +
        (matched.manufacturer ? ' · ' + esc(matched.manufacturer) : '') + '</span>' : '') +
      '</div>' +
      '<div class="report-scroll"><table class="report-table">' +
      '<thead><tr><th>Alternative</th><th>Manufacturer</th><th class="num">Stock</th>' +
      '<th class="num">Median price</th><th class="num">Lead</th><th>Links</th></tr></thead>' +
      '<tbody>' + rows + '</tbody></table></div></div>';
  }

  function syncAlternatives() {
    var chosen = altSelected().length;
    var run = el.altOverlay.querySelector('[data-alt="run"]');
    if (run) {
      run.textContent = altState.running
        ? 'Asking ' + altState.provider + '…'
        : 'Find alternatives (' + chosen + ')';
      run.disabled = altState.running;
    }
    var aside = el.altOverlay.querySelector('[data-alt-count]');
    if (aside) {
      aside.textContent = altState.stats
        ? chosen + ' of ' + altState.candidates.length + ' selected · ' +
          altState.stats.apiCalls + ' live, ' + altState.stats.cacheHits + ' cached'
        : chosen + ' of ' + altState.candidates.length + ' selected';
    }
    Array.prototype.forEach.call(
      el.altOverlay.querySelectorAll('[data-alt-row]'),
      function (tr) {
        var on = !!altState.selected[tr.getAttribute('data-alt-row')];
        tr.classList.toggle('picked', on);
        var box = tr.querySelector('[data-alt-pick]');
        if (box && box.checked !== on) box.checked = on;
      }
    );
  }

  function wireAlternatives() {
    Array.prototype.forEach.call(
      el.altOverlay.querySelectorAll('[data-alt-pick]'),
      function (box) {
        box.addEventListener('change', function () {
          altState.selected[box.getAttribute('data-alt-pick')] = box.checked;
          syncAlternatives();
        });
      }
    );
    Array.prototype.forEach.call(
      el.altOverlay.querySelectorAll('[data-alt]'),
      function (button) {
        button.addEventListener('click', function () {
          var action = button.getAttribute('data-alt');
          if (action === 'close') return closeAlternatives();
          if (action === 'run') return runAlternatives();
          altState.candidates.forEach(function (candidate) {
            altState.selected[candidate.key] = action === 'all';
          });
          syncAlternatives();
        });
      }
    );
  }

  function runAlternatives() {
    var chosen = altSelected();
    if (!chosen.length) {
      toast('Tick at least one part', true);
      return;
    }
    if (chosen.length > altState.maxParts) {
      toast('Ask for at most ' + altState.maxParts + ' parts at a time', true);
      return;
    }

    altState.running = true;
    syncAlternatives();

    // One request per distinct part number: the same part on two boards is one
    // question, and Nexar should be asked once.
    var asked = [];
    var seen = {};
    chosen.forEach(function (candidate) {
      var key = normalizeMpn(candidate.row.mpn);
      if (seen[key]) return;
      seen[key] = true;
      asked.push({ mpn: candidate.row.mpn, manufacturer: candidate.row.manufacturer || null });
    });

    fetch(api('/api/alternatives'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ parts: asked }),
    })
      .then(readJsonOrThrow)
      .then(function (data) {
        (data.results || []).forEach(function (result) {
          altState.answers[result.mpn] = result;
          if ((result.alternatives || []).length) {
            state.alternatives[normalizeMpn(result.mpn)] = result.alternatives;
          }
        });
        altState.stats = data.stats || null;
        var withAny = (data.results || []).filter(function (r) {
          return (r.alternatives || []).length;
        }).length;
        var failed = (data.results || []).filter(function (r) { return r.error; });
        toast(failed.length
          ? failed[0].error
          : withAny + ' of ' + asked.length + ' parts have alternatives', !!failed.length);
      })
      .catch(function (err) {
        toast(err.message || 'Could not reach the alternatives provider', true);
      })
      .then(function () {
        altState.running = false;
        renderAlternatives();
      });
  }

  // ── Excel export ─────────────────────────────────────────────────────────

  // The workbook is built on the server: it already owns a dependency-free
  // .xlsx writer, and rebuilding one in the browser would be a second
  // implementation of the same thing, free to drift from the first.
  function exportWorkbook(entries) {
    var books = (entries || []).filter(function (entry) { return entry && entry.results; });
    if (!books.length) {
      toast('Nothing analyzed to report on yet', true);
      return;
    }
    toast('Building the workbook…');

    fetch(api('/api/report'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        books: books.map(function (entry) {
          var model = reportModel(entry);
          return {
            name: entry.name,
            generated: model.generated,
            // Cross-BOM demand is attached here because only the browser knows
            // which other BOMs are open; the workbook then says exactly what
            // the report on screen says.
            rows: entry.results.rows.map(function (row) {
              var elsewhere = otherBomDemand(model, row.mpn);
              return elsewhere.length ? Object.assign({}, row, { alsoIn: elsewhere }) : row;
            }),
            suppliers: entry.results.suppliers,
            summary: entry.results.summary,
            stats: entry.results.stats,
            excluded: entry.excluded || [],
          };
        }),
      }),
    })
      .then(function (res) {
        if (!res.ok) return readJsonOrThrow(res);
        return res.blob().then(function (blob) {
          var slug = books.length === 1
            ? (books[0].name || 'bom').replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '')
            : 'all-boms';
          download(blob, (slug || 'bom') + '-report-' +
            new Date().toISOString().slice(0, 10) + '.xlsx');
          toast('Exported ' + books.length + ' BOM' + (books.length === 1 ? '' : 's') + ' to Excel');
        });
      })
      .catch(function (err) {
        toast(err.message || 'Could not build the workbook', true);
      });
  }

  function download(blob, filename) {
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  // ── Sample data ──────────────────────────────────────────────────────────

  var SAMPLE_BOM = [
    'Item,Reference,Qty,Manufacturer,Manufacturer Part Number,Description',
    '1,"C1,C2,C5",300,Murata,GRM188R71H104KA93D,CAP CER 0.1UF 50V X7R 0603',
    '2,"R1,R2",500,Yageo,RC0603FR-0710KL,RES SMD 10K OHM 1% 1/10W 0603',
    '3,U1,25,STMicroelectronics,STM32F103C8T6,ARM Cortex-M3 MCU 64KB LQFP48',
    '4,U2,25,Texas Instruments,LM358DR,IC OPAMP GP 2 CIRCUIT 8SOIC',
    '5,D1,50,ON Semiconductor,1N4148W-7-F,DIODE GEN PURP 100V 300MA SOD123',
    '6,J1,25,Molex,0533980571,CONN FFC/FPC 5POS 1MM SMD',
    '7,Y1,25,Abracon,ABM8G-12.000MHZ-18-D2Y-T,CRYSTAL 12.0000MHZ 18PF SMD',
    '8,U3,25,Microchip,ATMEGA328P-AU,MCU 8BIT 32KB FLASH TQFP32',
  ].join('\n');

  function loadSample() {
    var blob = new Blob([SAMPLE_BOM], { type: 'text/csv' });
    var file = new File([blob], 'sample-bom.csv', { type: 'text/csv' });
    handleFiles([file]);
  }

  // ── Wiring ───────────────────────────────────────────────────────────────

  el.settingsBtn.addEventListener('click', function () {
    var open = el.settingsPanel.hidden;
    el.settingsPanel.hidden = !open;
    el.settingsBtn.setAttribute('aria-expanded', String(open));
  });

  el.apiBase.addEventListener('change', function () {
    var value = el.apiBase.value.trim().replace(/\/+$/, '');
    state.apiBase = value || defaultApiBase();
    el.apiBase.value = state.apiBase;
    try {
      localStorage.setItem(STORAGE_KEY, state.apiBase);
    } catch (err) {
      // Private-mode browsers block storage; the value still applies this session.
    }
    checkHealth();
  });

  el.recheckBtn.addEventListener('click', function () {
    checkHealth().then(function (health) {
      toast(health ? 'Backend is reachable' : 'Still cannot reach the backend', !health);
    });
  });

  el.clearCacheBtn.addEventListener('click', function () {
    fetch(api('/api/cache'), { method: 'DELETE' })
      .then(readJsonOrThrow)
      .then(function () {
        toast('Server cache cleared');
        checkHealth();
      })
      .catch(function (err) {
        toast(err.message || 'Could not clear the cache', true);
      });
  });

  // "Clear server cache" empties the supplier answers the backend is holding.
  // This one empties what the browser is holding of the app itself.
  el.resetAppBtn.addEventListener('click', function () {
    toast('Clearing the browser copy and reloading…');
    clearBrowserCopy(true);
  });

  el.dropZone.addEventListener('click', function () { el.fileInput.click(); });
  el.dropZone.addEventListener('keydown', function (event) {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      el.fileInput.click();
    }
  });
  ['dragenter', 'dragover'].forEach(function (name) {
    el.dropZone.addEventListener(name, function (event) {
      event.preventDefault();
      el.dropZone.classList.add('over');
    });
  });
  ['dragleave', 'drop'].forEach(function (name) {
    el.dropZone.addEventListener(name, function (event) {
      event.preventDefault();
      el.dropZone.classList.remove('over');
    });
  });
  el.dropZone.addEventListener('drop', function (event) {
    var files = event.dataTransfer && event.dataTransfer.files;
    if (files && files.length) handleFiles(files);
  });
  el.fileInput.addEventListener('change', function () {
    if (el.fileInput.files && el.fileInput.files.length) handleFiles(el.fileInput.files);
    el.fileInput.value = '';
  });

  el.quickBtn.addEventListener('click', quickSearch);
  el.addRowBtn.addEventListener('click', function () { addLookupRow(null, true); });
  el.clearRowsBtn.addEventListener('click', function () {
    clearLookupRows();
    lookupRowElements()[0].querySelector('.mpn').focus();
  });

  // One listener on the container rather than three per row, so rows can be
  // added and removed without any bookkeeping.
  el.lookupRows.addEventListener('click', function (event) {
    var button = event.target.closest('.row-drop');
    if (!button || button.disabled) return;
    button.closest('.lookup-row').remove();
    if (!lookupRowElements().length) addLookupRow();
    syncRowControls();
  });

  el.lookupRows.addEventListener('keydown', function (event) {
    // Enter on the remove button belongs to the button, not to the search.
    if (event.key !== 'Enter' || event.target.tagName === 'BUTTON') return;
    event.preventDefault();
    quickSearch();
  });

  // Typing into the last row's part number opens the next one, so the list
  // grows as it is filled instead of needing a button between every part.
  el.lookupRows.addEventListener('input', function (event) {
    if (!event.target.classList.contains('mpn')) return;
    var rows = lookupRowElements();
    var row = event.target.closest('.lookup-row');
    if (row === rows[rows.length - 1] && event.target.value.trim()) addLookupRow();
  });

  el.lookupRows.addEventListener('paste', function (event) {
    var text = event.clipboardData && event.clipboardData.getData('text');
    // A single value is an ordinary paste into one field; only a list needs
    // spreading across the rows.
    if (!text || !/[\t\r\n]/.test(text)) return;
    var values = parsePastedRows(text);
    if (!values.length) return;
    event.preventDefault();
    fillLookupRows(values, event.target.closest('.lookup-row'));
  });

  el.sampleBtn.addEventListener('click', loadSample);
  el.analyzeBtn.addEventListener('click', function () { analyze(); });
  el.analyzeAllBtn.addEventListener('click', analyzeAll);
  el.closeAllBtn.addEventListener('click', function () {
    if (state.boms.some(function (b) { return b.running; })) {
      toast('Wait for the running analysis to finish first', true);
      return;
    }
    state.boms = [];
    state.activeId = null;
    renderAll();
  });
  el.exportBtn.addEventListener('click', exportCsv);
  el.reportBtn.addEventListener('click', openReport);
  el.dmsmsBtn.addEventListener('click', openDmsms);
  el.altBtn.addEventListener('click', openAlternatives);

  el.altOverlay.addEventListener('click', function (event) {
    if (event.target === el.altOverlay) closeAlternatives();
  });

  el.dmsmsOverlay.addEventListener('click', function (event) {
    if (event.target === el.dmsmsOverlay) closeDmsms();
  });

  // Clicking the backdrop closes; clicking the sheet itself must not.
  el.reportOverlay.addEventListener('click', function (event) {
    if (event.target === el.reportOverlay) closeReport();
  });
  document.addEventListener('keydown', function (event) {
    if (event.key !== 'Escape') return;
    if (!el.altOverlay.hidden) closeAlternatives();
    else if (!el.dmsmsOverlay.hidden) closeDmsms();
    else if (!el.reportOverlay.hidden) closeReport();
  });

  // "Start over" closes the BOM being viewed, leaving the others alone.
  el.resetBtn.addEventListener('click', function () {
    if (bom().running) {
      toast('That BOM is still analyzing', true);
      return;
    }
    if (state.activeId) removeBom(state.activeId);
  });

  var searchTimer = null;
  el.searchInput.addEventListener('input', function () {
    if (searchTimer) clearTimeout(searchTimer);
    searchTimer = setTimeout(function () {
      bom().search = el.searchInput.value;
      renderTable();
    }, 160);
  });

  clearLookupRows();

  state.apiBase = defaultApiBase();
  el.apiBase.value = state.apiBase;
  checkHealth();

  // No service worker is registered any more — see public/sw.js. Any copy left
  // over from an earlier version is removed here as well, for browsers that
  // would not otherwise re-fetch the worker for a while.
  clearBrowserCopy(false);

  // Tells the watchdog in index.html that a current script is running the page.
  // Last line of the startup path on purpose: anything that throws before here
  // leaves the flag unset, and the watchdog explains the blank page.
  window.__bomAppReady = true;
})();
