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

  var state = {
    apiBase: '',
    health: null,
    parse: null,
    mapping: {},
    lines: [],
    fromPaste: false,
    results: null,
    filter: 'all',
    search: '',
    expanded: {},
    running: false,
  };

  var el = {};
  [
    'statusBar', 'settingsBtn', 'settingsPanel', 'apiBase', 'currencyLabel', 'recheckBtn',
    'clearCacheBtn', 'dropZone', 'fileInput', 'pasteInput', 'pasteBtn', 'sampleBtn',
    'mappingCard', 'mappingSummary', 'mappingGrid', 'previewTable', 'analyzeBtn', 'resetBtn',
    'progressWrap', 'progressBar', 'progressText', 'resultsCard', 'statGrid', 'searchInput',
    'filterChips', 'exportBtn', 'resultsTable', 'resultsHead', 'resultsBody', 'emptyState',
    'setupCard', 'toast', 'attribution',
  ].forEach(function (id) {
    el[id] = document.getElementById(id);
  });

  // ── Small helpers ────────────────────────────────────────────────────────

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
    var code = currency || (state.results && state.results.summary.currency) || 'USD';
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

  function handleFile(file) {
    if (!file) return;
    toast('Reading ' + file.name + '…');
    fetch(api('/api/parse'), {
      method: 'POST',
      headers: { 'X-File-Name': file.name, 'Content-Type': 'application/octet-stream' },
      body: file,
    })
      .then(readJsonOrThrow)
      .then(function (data) {
        state.parse = data;
        state.mapping = data.mapping || {};
        state.lines = data.lines || [];
        state.fromPaste = false;
        state.results = null;
        el.resultsCard.hidden = true;
        renderMapping();
        toast('Loaded ' + state.lines.length + ' part' + (state.lines.length === 1 ? '' : 's') +
          ' from ' + file.name);
      })
      .catch(function (err) {
        toast(err.message || 'Could not read that file', true);
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

  // Pasted lists are parsed in the browser: they have no header row and no
  // packaging for the server to interpret.
  function parsePasted(text) {
    var lines = [];
    String(text || '')
      .split(/\r?\n/)
      .forEach(function (raw, index) {
        var line = raw.trim();
        if (!line) return;
        var parts = line.split(/[\t,;]/).map(function (p) { return p.trim(); });
        var mpn = parts[0];
        if (!mpn) return;
        var qty = parseInt(String(parts[1] || '').replace(/[^0-9]/g, ''), 10);
        lines.push({
          row: index + 1,
          mpn: mpn,
          quantity: isFinite(qty) && qty > 0 ? qty : 1,
          reference: null,
          manufacturer: parts[2] || null,
          description: null,
        });
      });
    return lines;
  }

  // ── Column mapping ───────────────────────────────────────────────────────

  function renderMapping() {
    el.mappingCard.hidden = false;

    if (state.fromPaste) {
      el.mappingGrid.innerHTML = '';
      el.mappingSummary.textContent =
        state.lines.length + ' part number' + (state.lines.length === 1 ? '' : 's') +
        ' from the pasted list. Quantities default to 1 where none was given.';
    } else {
      var headers = (state.parse && state.parse.headers) || [];
      el.mappingGrid.innerHTML = MAP_FIELDS.map(function (field) {
        var selected = state.mapping[field.key];
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

      var skipped = (state.parse && state.parse.skipped) || 0;
      el.mappingSummary.innerHTML =
        'Detected the header on row ' + (((state.parse && state.parse.headerRow) || 0) + 1) +
        '. <strong>' + state.lines.length + '</strong> part' + (state.lines.length === 1 ? '' : 's') +
        ' ready' + (skipped ? ', ' + skipped + ' row' + (skipped === 1 ? '' : 's') +
        ' skipped with no part number' : '') + '.';
    }

    renderPreview();
    el.analyzeBtn.disabled = state.lines.length === 0 || state.running;
    el.mappingCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
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
    if (value === '') delete state.mapping[field];
    else state.mapping[field] = parseInt(value, 10);

    fetch(api('/api/remap'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        rows: (state.parse && state.parse.rows) || [],
        mapping: state.mapping,
        rowOffset: (state.parse && state.parse.rowOffset) || 0,
      }),
    })
      .then(readJsonOrThrow)
      .then(function (data) {
        state.lines = data.lines || [];
        renderMapping();
      })
      .catch(function (err) {
        toast(err.message || 'Could not apply that mapping', true);
      });
  }

  function renderPreview() {
    var sample = state.lines.slice(0, 5);
    if (sample.length === 0) {
      el.previewTable.innerHTML =
        '<tr><td class="muted">No rows with a part number yet — pick the part number column above.</td></tr>';
      return;
    }
    var cols = ['row', 'mpn', 'quantity', 'reference', 'manufacturer', 'description'];
    var labels = ['Row', 'Part number', 'Qty', 'Ref', 'Manufacturer', 'Description'];
    var head = '<tr>' + labels.map(function (l) { return '<th>' + esc(l) + '</th>'; }).join('') + '</tr>';
    var body = sample.map(function (line) {
      return '<tr>' + cols.map(function (col) {
        var value = line[col];
        return '<td>' + (value === null || value === undefined || value === '' ? '—' : esc(value)) + '</td>';
      }).join('') + '</tr>';
    }).join('');
    var more = state.lines.length > sample.length
      ? '<tr><td colspan="6" class="muted">…and ' + (state.lines.length - sample.length) + ' more</td></tr>'
      : '';
    el.previewTable.innerHTML = head + body + more;
  }

  // ── Running the analysis ─────────────────────────────────────────────────

  function analyze() {
    if (state.running || state.lines.length === 0) return;
    if (!state.health) {
      toast('Backend is not reachable — check Settings', true);
      return;
    }
    if (!state.health.suppliers.some(function (s) { return s.configured; })) {
      toast('No supplier API credentials are configured on the server', true);
      return;
    }

    var max = state.health.maxPartsPerRequest || 500;
    var parts = state.lines.slice(0, max);
    if (state.lines.length > max) {
      toast('Analyzing the first ' + max + ' of ' + state.lines.length + ' parts (server limit)', true);
    }

    state.running = true;
    el.analyzeBtn.disabled = true;
    el.progressWrap.hidden = false;
    setProgress(0, 'Contacting suppliers…');

    streamLookup(parts, function (event, data) {
      if (event === 'start') {
        var expected = data.parts * data.suppliers.length;
        setProgress(0, 'Looking up ' + data.parts + ' parts across ' +
          data.suppliers.map(function (s) { return s.name; }).join(' and ') +
          ' (' + expected + ' queries)…');
      } else if (event === 'progress') {
        var percent = data.total ? Math.round((data.completed / data.total) * 100) : 0;
        setProgress(percent, data.completed + ' of ' + data.total + ' queries — ' +
          data.apiCalls + ' live, ' + data.cacheHits + ' cached' +
          (data.errors ? ', ' + data.errors + ' failed' : ''));
      } else if (event === 'done') {
        finishAnalysis(data);
      } else if (event === 'error') {
        throw new Error(data.error || 'Supplier lookup failed');
      }
    })
      .catch(function (err) {
        toast(err.message || 'Supplier lookup failed', true);
        setProgress(0, 'Stopped: ' + (err.message || 'lookup failed'));
      })
      .then(function () {
        state.running = false;
        el.analyzeBtn.disabled = state.lines.length === 0;
      });
  }

  function setProgress(percent, text) {
    el.progressBar.style.width = Math.max(0, Math.min(100, percent)) + '%';
    el.progressText.textContent = text || '';
  }

  function streamLookup(parts, onEvent) {
    return fetch(api('/api/lookup'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Accept: 'text/event-stream' },
      body: JSON.stringify({ parts: parts, stream: true }),
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

  function finishAnalysis(data) {
    state.results = data;
    state.expanded = {};
    state.filter = 'all';
    state.search = '';
    el.searchInput.value = '';
    setProgress(100, 'Done — ' + data.stats.apiCalls + ' live queries, ' +
      data.stats.cacheHits + ' served from cache' +
      (data.stats.errors ? ', ' + data.stats.errors + ' failed' : '') + '.');
    el.resultsCard.hidden = false;
    renderStats();
    renderFilters();
    renderTable();
    renderAttribution();
    checkHealth();
    el.resultsCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ── Results: summary tiles ───────────────────────────────────────────────

  function renderStats() {
    var summary = state.results.summary;
    var suppliers = state.results.suppliers;
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

    var stockRisk = state.results.rows.filter(function (row) {
      return row.comparison.inStockSuppliers.length === 0;
    }).length;
    tiles.push(tile(
      'Stock risk',
      String(stockRisk),
      stockRisk ? 'no supplier holds the full quantity' : 'every line is coverable today',
      stockRisk ? 'bad' : 'good'
    ));

    var lifecycleRisk = state.results.rows.filter(function (row) {
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
    if (!state.results) return null;
    var urls = [];
    state.results.rows.forEach(function (row) {
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
        return !state.results.suppliers.some(function (s) {
          return row.offers[s.id] && row.offers[s.id].found;
        });
      default:
        return true;
    }
  }

  function renderFilters() {
    el.filterChips.innerHTML = FILTERS.map(function (filter) {
      var n = state.results.rows.filter(function (row) { return matchesFilter(row, filter.key); }).length;
      return '<button type="button" class="chip' + (state.filter === filter.key ? ' active' : '') +
        '" data-filter="' + filter.key + '">' + esc(filter.label) +
        '<span class="count">' + n + '</span></button>';
    }).join('');

    Array.prototype.forEach.call(el.filterChips.querySelectorAll('.chip'), function (chip) {
      chip.addEventListener('click', function () {
        state.filter = chip.getAttribute('data-filter');
        renderFilters();
        renderTable();
      });
    });
  }

  function visibleRows() {
    var query = state.search.trim().toLowerCase();
    return state.results.rows.filter(function (row) {
      if (!matchesFilter(row, state.filter)) return false;
      if (!query) return true;
      var haystack = [row.mpn, row.description, row.reference, row.manufacturer];
      state.results.suppliers.forEach(function (supplier) {
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
    var suppliers = state.results.suppliers;

    var groupRow = '<tr class="supplier-row">' +
      '<th class="spacer sticky-a"></th><th class="spacer sticky-b"></th><th class="spacer"></th>' +
      suppliers.map(function (supplier) {
        return '<th colspan="' + SUPPLIER_COLUMNS.length + '" class="group-start">' +
          esc(supplier.name) + '</th>';
      }).join('') +
      '<th class="spacer group-start"></th></tr>';

    var fieldRow = '<tr class="field-row"><th class="sticky-a"></th>' +
      '<th class="sticky-b">Part</th><th class="num">Qty</th>' +
      suppliers.map(function () {
        return SUPPLIER_COLUMNS.map(function (col, i) {
          var cls = (i === 0 ? 'group-start ' : '') + (col.key === 'lifecycle' ? '' : 'num');
          return '<th class="' + cls.trim() + '">' + esc(col.label) + '</th>';
        }).join('');
      }).join('') +
      '<th class="group-start">Verdict</th></tr>';

    el.resultsHead.innerHTML = groupRow + fieldRow;

    var rows = visibleRows();
    el.emptyState.hidden = rows.length > 0;
    el.resultsTable.hidden = rows.length === 0;

    el.resultsBody.innerHTML = rows.map(function (row) {
      return renderRow(row, suppliers);
    }).join('');

    Array.prototype.forEach.call(el.resultsBody.querySelectorAll('.expander'), function (button) {
      button.addEventListener('click', function () {
        var index = button.getAttribute('data-index');
        if (state.expanded[index]) delete state.expanded[index];
        else state.expanded[index] = true;
        renderTable();
      });
    });
  }

  function renderRow(row, suppliers) {
    var isOpen = !!state.expanded[row.index];
    var comparison = row.comparison;

    var meta = [];
    if (row.reference) meta.push('<span class="refdes">' + esc(row.reference) + '</span>');
    if (row.manufacturer) meta.push(esc(row.manufacturer));
    if (row.description) meta.push(esc(row.description));

    var cells = suppliers.map(function (supplier) {
      return supplierCells(row.offers[supplier.id], supplier, comparison, row.quantity);
    }).join('');

    var verdict = renderVerdict(comparison, suppliers);

    var main = '<tr class="part-row' + (isOpen ? ' expanded' : '') + '">' +
      '<td class="sticky-a"><button type="button" class="expander" data-index="' + row.index +
      '" aria-label="Toggle details">' + (isOpen ? '▾' : '▸') + '</button></td>' +
      '<td class="sticky-b"><div class="part-cell"><div class="mpn">' + esc(row.mpn) + '</div>' +
      (meta.length ? '<div class="rowmeta">' + meta.join(' · ') + '</div>' : '') + '</div></td>' +
      '<td class="num">' + count(row.quantity) + '</td>' +
      cells +
      '<td class="group-start"><div class="verdict">' + verdict + '</div></td>' +
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

    return '<div class="detail">' + columns + '</div>';
  }

  // ── CSV export ───────────────────────────────────────────────────────────

  function exportCsv() {
    if (!state.results) return;
    var suppliers = state.results.suppliers;
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
    header.push('Cheapest Supplier', 'Soonest Supplier', 'Soonest (days)', 'Recommended', 'Worst Lifecycle', 'Notes');

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
        row.comparison.flags.map(function (f) { return f.text; }).join('; ')
      );
      lines.push(record);
    });

    var csv = lines.map(function (record) {
      return record.map(csvCell).join(',');
    }).join('\r\n');

    var blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.href = url;
    link.download = 'bom-supplier-comparison-' + new Date().toISOString().slice(0, 10) + '.csv';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
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
    handleFile(file);
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
    if (files && files.length) handleFile(files[0]);
  });
  el.fileInput.addEventListener('change', function () {
    if (el.fileInput.files && el.fileInput.files.length) handleFile(el.fileInput.files[0]);
    el.fileInput.value = '';
  });

  el.pasteBtn.addEventListener('click', function () {
    var lines = parsePasted(el.pasteInput.value);
    if (lines.length === 0) {
      toast('Nothing to read — paste one part number per line', true);
      return;
    }
    state.parse = null;
    state.mapping = {};
    state.lines = lines;
    state.fromPaste = true;
    state.results = null;
    el.resultsCard.hidden = true;
    renderMapping();
    toast('Loaded ' + lines.length + ' part' + (lines.length === 1 ? '' : 's'));
  });

  el.sampleBtn.addEventListener('click', loadSample);
  el.analyzeBtn.addEventListener('click', analyze);
  el.exportBtn.addEventListener('click', exportCsv);

  el.resetBtn.addEventListener('click', function () {
    state.parse = null;
    state.mapping = {};
    state.lines = [];
    state.results = null;
    state.fromPaste = false;
    el.mappingCard.hidden = true;
    el.resultsCard.hidden = true;
    el.progressWrap.hidden = true;
    el.pasteInput.value = '';
  });

  var searchTimer = null;
  el.searchInput.addEventListener('input', function () {
    if (searchTimer) clearTimeout(searchTimer);
    searchTimer = setTimeout(function () {
      state.search = el.searchInput.value;
      renderTable();
    }, 160);
  });

  state.apiBase = defaultApiBase();
  el.apiBase.value = state.apiBase;
  checkHealth();

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () {
      navigator.serviceWorker.register('sw.js').catch(function () {});
    });
  }
})();
