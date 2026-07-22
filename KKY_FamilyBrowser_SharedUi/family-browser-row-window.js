(function (w) {
  'use strict';

  var api = w.KKYFB || {};
  var stores = api._stores || {};
  var windowSize = 150;
  var originalFilterRows = w.filterRows;
  var originalSelectRow = w.selectRow;
  var columnWidths = api._columnWidths || {};

  function text(value) {
    return value == null ? '' : String(value);
  }

  function html(value) {
    if (w.escHtml) return w.escHtml(text(value));
    return text(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function attrMap(values) {
    return {
      'data-raw': text(values[1]),
      'data-status': text(values[2]),
      'data-discipline-key': text(values[3]),
      'data-tree-discipline-key': text(values[4]),
      'data-discipline': text(values[5]),
      'data-group': text(values[6]),
      'data-name': text(values[7]),
      'data-category': text(values[8]),
      'data-action': text(values[9]),
      'data-notes': text(values[10]),
      'data-approved': text(values[11]),
      'data-loaded': text(values[12]),
      'data-changes': text(values[13]),
      'data-diffrows': text(values[14]),
      'data-kind': text(values[15]),
      'data-parameters': text(values[16]),
      'data-parameter-types': text(values[17]),
      'data-nested': text(values[18]),
      'data-types': text(values[19]),
      'data-systemlayers': text(values[20]),
      'data-previewdiag': text(values[21]),
      'data-detailkey': text(values[22]),
      'data-detailsource': text(values[23]),
      'data-previewpath': text(values[24]),
      'data-nested-child': text(values[25]),
      'data-preview': ''
    };
  }

  function makeItem(values) {
    var item = {
      key: text(values[0]),
      attrs: attrMap(values),
      compactNotes: text(values[26]),
      statusClass: text(values[27]),
      selectable: values[28] === 1 || values[28] === true,
      selectTitle: text(values[29]),
      typeCountDisplay: text(values[30])
    };
    item.adapter = {
      className: 'data',
      style: { display: '' },
      getAttribute: function (name) {
        if (name === 'data-row-key') return item.key;
        return item.attrs[name] || '';
      }
    };
    item.searchText = ((item.attrs['data-discipline'] || '') + ' ' + (item.attrs['data-group'] || '') + ' ' +
      (item.attrs['data-kind'] || '') + ' ' + (item.attrs['data-status'] || '') + ' ' +
      (item.attrs['data-name'] || '') + ' ' + (item.attrs['data-category'] || '') + ' ' +
      (item.attrs['data-action'] || '') + ' ' + (item.attrs['data-notes'] || '') + ' ' +
      (item.attrs['data-approved'] || '') + ' ' + (item.attrs['data-loaded'] || '') + ' ' +
      (item.attrs['data-changes'] || '')).toLowerCase();
    return item;
  }

  function tableBody(tab) {
    var table = document.getElementById(tab + 'Table');
    if (!table) return null;
    if (table.tBodies && table.tBodies.length) return table.tBodies[0];
    var body = document.createElement('tbody');
    table.appendChild(body);
    return body;
  }

  function renderedRows(tab) {
    var table = document.getElementById(tab + 'Table');
    return table ? table.getElementsByTagName('tr') : [];
  }

  function setCellText(cell, value, title) {
    cell.innerText = text(value);
    if (title != null) cell.title = text(title);
  }

  function addSelectionCell(row, tab, item, store) {
    var cell = row.insertCell(-1);
    cell.className = 'select-col';
    var input = document.createElement('input');
    input.className = 'row-check ' + (tab === 'families' ? 'family-row-check' : 'system-row-check');
    input.type = 'checkbox';
    input.title = item.selectTitle;
    input.setAttribute('data-row-key', item.key);
    input.checked = !!store.checked[item.key];
    if (!item.selectable) {
      input.disabled = true;
      input.setAttribute('disabled', 'disabled');
    }
    input.onclick = function (eventValue) {
      return tab === 'families' ? w.toggleFamilyCheck(eventValue || w.event, input) : w.toggleSystemCheck(eventValue || w.event, input);
    };
    cell.appendChild(input);
  }

  function addStatusCell(row, item) {
    var cell = row.insertCell(-1);
    cell.innerHTML = '<span class="badge ' + html(item.statusClass) + '">' + html(item.attrs['data-status']) + '</span>';
  }

  function renderRow(body, tab, store, item) {
    var row = body.insertRow(-1);
    var name;
    row.className = 'data' + (store.selectedKey === item.key ? ' selected' : '');
    row.setAttribute('data-row-key', item.key);
    for (name in item.attrs) {
      if (Object.prototype.hasOwnProperty.call(item.attrs, name)) row.setAttribute(name, item.attrs[name]);
    }
    row.onclick = (function (target) {
      return function () { return w.selectRow(target); };
    })(row);

    addSelectionCell(row, tab, item, store);
    addStatusCell(row, item);
    if (tab === 'families' && store.mode === 'family-modeler') {
      var modelerName = row.insertCell(-1);
      setCellText(modelerName, item.attrs['data-name'], item.attrs['data-name']);
      var modelerCategory = row.insertCell(-1);
      setCellText(modelerCategory, item.attrs['data-category'], item.attrs['data-category']);
      setCellText(row.insertCell(-1), item.typeCountDisplay);
      return row;
    }

    setCellText(row.insertCell(-1), item.attrs['data-discipline']);
    var itemName = row.insertCell(-1);
    setCellText(itemName, item.attrs['data-name'], item.attrs['data-name']);
    var category = row.insertCell(-1);
    setCellText(category, item.attrs['data-category'], item.attrs['data-category']);
    if (tab === 'families') {
      setCellText(row.insertCell(-1), item.typeCountDisplay);
    } else {
      setCellText(row.insertCell(-1), item.attrs['data-action']);
    }
    var memo = row.insertCell(-1);
    setCellText(memo, item.compactNotes, item.attrs['data-notes']);
    return row;
  }

  function clearBody(body) {
    while (body && body.firstChild) body.removeChild(body.firstChild);
  }

  function getStore(tab) {
    return stores[tab] || null;
  }

  function findRenderedRow(tab, key) {
    var rows = renderedRows(tab);
    for (var i = 0; i < rows.length; i++) {
      if ((rows[i].getAttribute('data-row-key') || '') === key) return rows[i];
    }
    return null;
  }

  function pageIndexes(page, totalPages) {
    var result = [];
    var last = -2;
    for (var i = 0; i < totalPages; i++) {
      if (totalPages <= 9 || i === 0 || i === totalPages - 1 || Math.abs(i - page) <= 2) {
        if (i - last > 1) result.push(-1);
        result.push(i);
        last = i;
      }
    }
    return result;
  }

  function renderPageLinks(container, page, totalPages) {
    if (!container) return;
    while (container.firstChild) container.removeChild(container.firstChild);
    var indexes = pageIndexes(page, totalPages);
    for (var i = 0; i < indexes.length; i++) {
      if (indexes[i] < 0) {
        var gap = document.createElement('span');
        gap.className = 'row-window-ellipsis';
        gap.innerText = '...';
        container.appendChild(gap);
        continue;
      }
      var link = document.createElement('a');
      link.href = '#';
      link.className = 'row-window-page' + (indexes[i] === page ? ' active' : '');
      link.setAttribute('data-page-index', String(indexes[i]));
      link.innerText = String(indexes[i] + 1);
      link.onclick = (function (targetPage) {
        return function () { return w.goToRowWindowPage(targetPage); };
      })(indexes[i]);
      container.appendChild(link);
    }
  }

  function updateWindowChrome(total, page) {
    var maxPage = Math.max(0, Math.ceil(total / windowSize) - 1);
    var totalPages = maxPage + 1;
    var start = total ? page * windowSize + 1 : 0;
    var end = Math.min(total, (page + 1) * windowSize);
    var controls = document.getElementById('rowWindowControls');
    var range = document.getElementById('rowWindowRange');
    var pages = document.getElementById('rowWindowPages');
    var summary = document.getElementById('rowWindowPageSummary');
    var previous = document.getElementById('rowWindowPrev');
    var next = document.getElementById('rowWindowNext');
    if (controls) controls.style.display = total > windowSize ? 'block' : 'none';
    if (range) range.innerText = start + '-' + end + ' / ' + total;
    if (summary) {
      var pageLabel = controls ? (controls.getAttribute('data-page-label') || 'Page') : 'Page';
      summary.innerText = pageLabel + ' ' + (page + 1) + ' / ' + totalPages;
    }
    renderPageLinks(pages, page, totalPages);
    if (previous) previous.className = 'row-window-nav' + (page <= 0 ? ' disabled' : '');
    if (next) next.className = 'row-window-nav' + (page >= maxPage ? ' disabled' : '');
    w.rowWindowTotal = total;
    w.rowWindowPage = page;
    w.rowWindowSize = windowSize;
  }

  function hasClass(element, name) {
    return !!element && (' ' + (element.className || '') + ' ').indexOf(' ' + name + ' ') >= 0;
  }

  function headerTable(tab) {
    var host = document.getElementById(tab + 'FixedHeader');
    if (!host) return null;
    var tables = host.getElementsByTagName('table');
    return tables.length ? tables[0] : null;
  }

  function bodyTable(tab) {
    return document.getElementById(tab + 'Table');
  }

  function tableColumns(table) {
    if (!table) return [];
    var groups = table.getElementsByTagName('colgroup');
    return groups.length ? groups[0].getElementsByTagName('col') : [];
  }

  function widthKey(tab, count) {
    return tab + ':' + count;
  }

  function storageKey(tab, count) {
    return 'kkyfb-column-widths-v1:' + widthKey(tab, count);
  }

  function loadStoredWidths(tab, count) {
    var key = widthKey(tab, count);
    if (columnWidths[key] && columnWidths[key].length === count) return columnWidths[key];
    try {
      if (w.localStorage && w.JSON) {
        var parsed = JSON.parse(w.localStorage.getItem(storageKey(tab, count)) || 'null');
        if (parsed && parsed.length === count) {
          columnWidths[key] = parsed;
          return parsed;
        }
      }
    } catch (ignoredStorageRead) {}
    return null;
  }

  function rememberColumnWidths(tab, widths) {
    var key = widthKey(tab, widths.length);
    columnWidths[key] = widths.slice(0);
    api._columnWidths = columnWidths;
  }

  function saveStoredWidths(tab, widths) {
    rememberColumnWidths(tab, widths);
    try {
      if (w.localStorage && w.JSON) w.localStorage.setItem(storageKey(tab, widths.length), JSON.stringify(columnWidths[widthKey(tab, widths.length)]));
    } catch (ignoredStorageWrite) {}
  }

  function measuredWidths(table, count) {
    var widths = [];
    var cells = table ? table.getElementsByTagName('th') : [];
    for (var i = 0; i < count; i++) {
      var measured = cells[i] ? cells[i].offsetWidth : 0;
      widths.push(Math.max(i === 0 ? 42 : 72, measured || (i === 0 ? 42 : 120)));
    }
    return widths;
  }

  function lockTablePixelWidth(table, total) {
    if (!table) return;
    if (!hasClass(table, 'kkyfb-column-width-locked')) table.className = (table.className || '') + ' kkyfb-column-width-locked';
    table.style.tableLayout = 'fixed';
    table.style.width = total + 'px';
  }

  function applyColumnWidths(tab, widths) {
    var head = headerTable(tab);
    var body = bodyTable(tab);
    var headCols = tableColumns(head);
    var bodyCols = tableColumns(body);
    if (!head || !body || !widths || headCols.length !== widths.length || bodyCols.length !== widths.length) return false;
    var total = 0;
    for (var i = 0; i < widths.length; i++) {
      var width = Math.max(i === 0 ? 42 : 72, parseInt(widths[i], 10) || 0);
      widths[i] = width;
      headCols[i].style.width = width + 'px';
      bodyCols[i].style.width = width + 'px';
      total += width;
    }
    lockTablePixelWidth(head, total);
    lockTablePixelWidth(body, total);
    return true;
  }

  function finishColumnResize(moveHandler, upHandler) {
    if (document.removeEventListener) {
      document.removeEventListener('mousemove', moveHandler, false);
      document.removeEventListener('mouseup', upHandler, false);
    } else if (document.detachEvent) {
      document.detachEvent('onmousemove', moveHandler);
      document.detachEvent('onmouseup', upHandler);
    }
    if (document.body) document.body.className = (document.body.className || '').replace(/\s*kkyfb-column-resizing\b/g, '');
  }

  function startColumnResize(eventValue, tab, index) {
    var ev = eventValue || w.event;
    var head = headerTable(tab);
    var cols = tableColumns(head);
    if (!head || !cols.length || index >= cols.length) return false;
    var key = widthKey(tab, cols.length);
    var widths = loadStoredWidths(tab, cols.length) || measuredWidths(head, cols.length);
    widths = widths.slice(0);
    applyColumnWidths(tab, widths);
    var startX = ev ? ev.clientX : 0;
    var startWidth = widths[index];
    if (document.body && !hasClass(document.body, 'kkyfb-column-resizing')) document.body.className += ' kkyfb-column-resizing';
    var moveHandler = function (moveEvent) {
      var mev = moveEvent || w.event;
      widths[index] = Math.max(index === 0 ? 42 : 72, startWidth + ((mev ? mev.clientX : startX) - startX));
      applyColumnWidths(tab, widths);
      if (mev && mev.preventDefault) mev.preventDefault();
      if (mev) mev.returnValue = false;
      return false;
    };
    var upHandler = function () {
      finishColumnResize(moveHandler, upHandler);
      saveStoredWidths(tab, widths);
      if (w.fbRefreshTableHeaders) w.fbRefreshTableHeaders();
      if (w.refreshOverflowTitles) w.refreshOverflowTitles(bodyTable(tab) || document.body);
      return false;
    };
    if (document.addEventListener) {
      document.addEventListener('mousemove', moveHandler, false);
      document.addEventListener('mouseup', upHandler, false);
    } else if (document.attachEvent) {
      document.attachEvent('onmousemove', moveHandler);
      document.attachEvent('onmouseup', upHandler);
    }
    stopEvent(ev);
    if (ev && ev.preventDefault) ev.preventDefault();
    if (ev) ev.returnValue = false;
    return false;
  }

  function ensureColumnResizers(tab) {
    var head = headerTable(tab);
    var body = bodyTable(tab);
    if (!head || !body) return false;
    var cells = head.getElementsByTagName('th');
    var cols = tableColumns(head);
    if (!cells.length || cells.length !== cols.length || tableColumns(body).length !== cols.length) return false;
    var saved = loadStoredWidths(tab, cols.length);
    var baseline = saved ? saved.slice(0) : measuredWidths(head, cols.length);
    if (!applyColumnWidths(tab, baseline)) return false;
    rememberColumnWidths(tab, baseline);
    for (var i = 0; i < cells.length; i++) {
      var existing = cells[i].getElementsByTagName('span');
      var found = false;
      for (var j = 0; j < existing.length; j++) {
        if (hasClass(existing[j], 'column-resize-handle')) { found = true; break; }
      }
      if (found) continue;
      var handle = document.createElement('span');
      handle.className = 'column-resize-handle';
      handle.title = cells[i].getAttribute('data-resize-title') || '';
      handle.setAttribute('data-column-index', String(i));
      handle.onmousedown = (function (columnIndex) {
        return function (resizeEvent) { return startColumnResize(resizeEvent, tab, columnIndex); };
      })(i);
      cells[i].appendChild(handle);
    }
    return true;
  }

  function renderedDataCount(tab) {
    var rows = renderedRows(tab);
    var count = 0;
    for (var i = 0; i < rows.length; i++) {
      if ((' ' + (rows[i].className || '') + ' ').indexOf(' data ') >= 0) count++;
    }
    return count;
  }

  function syncRenderedState(tab, store) {
    var rows = renderedRows(tab);
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if ((' ' + (row.className || '') + ' ').indexOf(' data ') < 0) continue;
      var key = row.getAttribute('data-row-key') || '';
      row.className = 'data' + (store.selectedKey === key ? ' selected' : '');
      var inputs = row.getElementsByTagName('input');
      for (var j = 0; j < inputs.length; j++) {
        if ((' ' + (inputs[j].className || '') + ' ').indexOf(' row-check ') >= 0) {
          inputs[j].checked = !!store.checked[key];
          break;
        }
      }
    }
  }

  function renderWindow(tab, store) {
    var body = tableBody(tab);
    if (!body) return;
    var start = store.page * windowSize;
    var end = Math.min(store.filtered.length, start + windowSize);
    var keys = [];
    for (var i = start; i < end; i++) keys.push(store.filtered[i].key);
    var signature = store.page + '|' + keys.join('\x1f');
    if (store.renderSignature === signature && renderedDataCount(tab) === keys.length) {
      syncRenderedState(tab, store);
    } else {
      clearBody(body);
      for (var renderIndex = start; renderIndex < end; renderIndex++) renderRow(body, tab, store, store.filtered[renderIndex]);
      store.renderSignature = signature;
    }
    var table = document.getElementById(tab + 'Table');
    if (table) {
      table.setAttribute('data-kkyfb-virtualized', 'true');
      table.setAttribute('data-total-rows', String(store.rows.length));
      table.setAttribute('data-filtered-rows', String(store.filtered.length));
      table.setAttribute('data-rendered-rows', String(end - start));
    }
    updateWindowChrome(store.filtered.length, store.page);
    if (w.fbRefreshTableHeaders) w.fbRefreshTableHeaders();
    ensureColumnResizers(tab);
  }

  function queryText(item) {
    return item.searchText || '';
  }

  function filterSignature(tab, query) {
    return [tab, query, w.currentFilter || 'All', w.currentDiscipline || 'All', w.currentTreeDiscipline || 'All',
      w.currentTreeGroup || '', w.currentTreeCategory || '', w.currentSystemTreeDiscipline || 'All',
      w.currentSystemTreeCategory || '', w.advStatus || 'All', w.advGroup || 'All', w.advCategory || '',
      w.advMismatchOnly ? '1' : '0'].join('\x1f');
  }

  function itemMatches(item, query) {
    var row = item.adapter;
    return (query === '' || queryText(item).indexOf(query) >= 0) &&
      (!w.statusMatches || w.statusMatches(row)) &&
      (!w.disciplineMatches || w.disciplineMatches(row)) &&
      (!w.familyTreeMatches || w.familyTreeMatches(row)) &&
      (!w.systemTreeMatches || w.systemTreeMatches(row)) &&
      (!w.advancedMatches || w.advancedMatches(row));
  }

  function stopEvent(eventValue) {
    var ev = eventValue || w.event;
    if (ev && ev.stopPropagation) ev.stopPropagation();
    if (ev) ev.cancelBubble = true;
  }

  function checkedItems(tab) {
    var store = getStore(tab);
    var result = [];
    if (!store) return result;
    for (var i = 0; i < store.rows.length; i++) {
      if (store.checked[store.rows[i].key]) result.push(store.rows[i].adapter);
    }
    return result;
  }

  function checkedCount(store) {
    var count = 0;
    if (!store) return count;
    for (var key in store.checked) {
      if (Object.prototype.hasOwnProperty.call(store.checked, key) && store.checked[key]) count++;
    }
    return count;
  }

  function updateSelectionControls(tab) {
    var store = getStore(tab);
    if (!store) return;
    var rows = renderedRows(tab);
    var visibleSelectable = 0;
    var visibleChecked = 0;
    for (var i = 0; i < rows.length; i++) {
      var inputs = rows[i].getElementsByTagName('input');
      for (var j = 0; j < inputs.length; j++) {
        if (inputs[j].disabled) continue;
        visibleSelectable++;
        if (inputs[j].checked) visibleChecked++;
        break;
      }
    }
    var totalChecked = checkedCount(store);
    var countBox = document.getElementById(tab === 'families' ? 'selectedFamilyCount' : 'selectedSystemCount');
    if (countBox) countBox.innerText = totalChecked;
    var all = document.getElementById(tab === 'families' ? 'familyCheckAll' : 'systemCheckAll');
    if (all) {
      all.checked = visibleSelectable > 0 && visibleChecked === visibleSelectable;
      try { all.indeterminate = visibleChecked > 0 && visibleChecked < visibleSelectable; } catch (ignored) {}
    }
    var apply = document.getElementById(tab === 'families' ? 'checkedFamilyApply' : 'checkedSystemApply');
    if (apply) {
      var base = apply.getAttribute('data-base-label') || apply.innerText || '';
      apply.innerText = base + (totalChecked > 0 ? ' (' + totalChecked + ')' : '');
    }
  }

  api.setRows = function (tab, payload) {
    if (!payload || !payload.rows || !payload.rows.length) {
      api.clearRows(tab);
      return false;
    }
    var previous = stores[tab];
    var next = {
      tab: tab,
      mode: text(payload.mode),
      rows: [],
      filtered: [],
      checked: {},
      selectedKey: previous ? previous.selectedKey : '',
      page: 0,
      filterSignature: '',
      renderSignature: ''
    };
    var valid = {};
    for (var i = 0; i < payload.rows.length; i++) {
      var item = makeItem(payload.rows[i]);
      next.rows.push(item);
      valid[item.key] = true;
      if (previous && previous.checked[item.key]) next.checked[item.key] = true;
    }
    if (!valid[next.selectedKey]) next.selectedKey = '';
    stores[tab] = next;
    api._stores = stores;
    return true;
  };

  api.setRowsFromJson = function (tab, payloadJson) {
    if (!payloadJson) {
      api.clearRows(tab);
      return false;
    }
    try {
      return api.setRows(tab, typeof payloadJson === 'string' ? JSON.parse(payloadJson) : payloadJson);
    } catch (error) {
      if (w.fbErr) w.fbErr('KKYFB.setRowsFromJson', error);
      api.clearRows(tab);
      return false;
    }
  };

  api.clearRows = function (tab) {
    if (stores[tab]) delete stores[tab];
    return true;
  };

  api.findSavedRow = function (tab, saved) {
    var store = getStore(tab);
    if (!store || !saved) return null;
    var source = store.filtered.length ? store.filtered : store.rows;
    for (var i = 0; i < source.length; i++) {
      var a = source[i].attrs;
      if ((a['data-name'] || '') === (saved.name || '') && (a['data-category'] || '') === (saved.category || '') &&
          (a['data-kind'] || '') === (saved.kind || '') && (!(saved.discipline || '') || (a['data-discipline-key'] || '') === saved.discipline)) {
        store.page = Math.floor(i / windowSize);
        store.selectedKey = source[i].key;
        renderWindow(tab, store);
        return findRenderedRow(tab, source[i].key);
      }
    }
    return null;
  };

  api.stats = function (tab) {
    var store = getStore(tab);
    var rows = renderedRows(tab);
    var rendered = 0;
    var visible = 0;
    for (var i = 0; i < rows.length; i++) {
      if ((' ' + (rows[i].className || '') + ' ').indexOf(' data ') < 0) continue;
      rendered++;
      if (rows[i].style.display !== 'none') visible++;
    }
    return {
      total: store ? store.rows.length : rendered,
      filtered: store ? store.filtered.length : visible,
      rendered: rendered,
      visible: visible,
      page: store ? store.page : 0,
      checked: store ? checkedCount(store) : 0
    };
  };

  w.filterRows = function (source) {
    var tab = w.currentTab || 'families';
    var store = getStore(tab);
    if (!store) return originalFilterRows ? originalFilterRows(source) : false;
    var previousPage = store.page;
    if (source !== 'page') store.page = 0;
    var search = document.getElementById('searchBox');
    var restoreFocus = source === 'search' && search && document.activeElement === search;
    var query = search ? text(search.value).toLowerCase() : '';
    var signature = filterSignature(tab, query);
    if (source !== 'page' && previousPage === 0 && store.filterSignature === signature) {
      if (restoreFocus) {
        try { search.focus(); } catch (ignoredRepeatFocus) {}
      }
      return false;
    }
    store.filtered = [];
    for (var i = 0; i < store.rows.length; i++) {
      var item = store.rows[i];
      if (itemMatches(item, query)) {
        store.filtered.push(item);
      } else {
        delete store.checked[item.key];
        if (store.selectedKey === item.key) store.selectedKey = '';
      }
    }
    store.filterSignature = signature;
    var maxPage = Math.max(0, Math.ceil(store.filtered.length / windowSize) - 1);
    if (store.page > maxPage) store.page = maxPage;
    renderWindow(tab, store);
    var count = document.getElementById('visibleCount');
    if (count) count.innerText = store.filtered.length + ' ' + (w.visibleLabel || 'visible');
    var rows = renderedRows(tab);
    var first = null;
    for (var r = 0; r < rows.length; r++) {
      if ((' ' + (rows[r].className || '') + ' ').indexOf(' data ') >= 0) { first = rows[r]; break; }
    }
    if (first) w.selectRow(first, false, source === 'search');
    else if (w.resetDetailForCurrentTab) w.resetDetailForCurrentTab();
    updateSelectionControls(tab);
    if (restoreFocus) {
      try { search.focus(); } catch (ignoredFocus) {}
    }
    return false;
  };

  w.goToRowWindowPage = function (pageIndex) {
    var tab = w.currentTab || 'families';
    var store = getStore(tab);
    if (!store) return false;
    var maxPage = Math.max(0, Math.ceil(store.filtered.length / windowSize) - 1);
    store.page = Math.max(0, Math.min(maxPage, parseInt(pageIndex, 10) || 0));
    w.filterRows('page');
    var table = bodyTable(tab);
    var parent = table ? table.parentNode : null;
    while (parent) {
      if (hasClass(parent, 'family-tablewrap') || hasClass(parent, 'tablewrap')) { parent.scrollTop = 0; break; }
      parent = parent.parentNode;
    }
    return false;
  };

  w.changeRowWindow = function (delta) {
    var tab = w.currentTab || 'families';
    var store = getStore(tab);
    return store ? w.goToRowWindowPage(store.page + (delta || 0)) : false;
  };

  w.resizeColumnForAudit = function (tab, index, width) {
    var head = headerTable(tab);
    var cols = tableColumns(head);
    if (!head || index < 0 || index >= cols.length) return false;
    var widths = loadStoredWidths(tab, cols.length) || measuredWidths(head, cols.length);
    widths = widths.slice(0);
    widths[index] = Math.max(index === 0 ? 42 : 72, parseInt(width, 10) || widths[index]);
    if (!applyColumnWidths(tab, widths)) return false;
    saveStoredWidths(tab, widths);
    if (w.refreshOverflowTitles) w.refreshOverflowTitles(bodyTable(tab) || document.body);
    return true;
  };

  api.columnWidths = function (tab) {
    var cols = tableColumns(headerTable(tab));
    var widths = loadStoredWidths(tab, cols.length) || measuredWidths(headerTable(tab), cols.length);
    return widths.slice(0);
  };

  w.selectRow = function (row, openDetached, quietDetail) {
    var tab = w.currentTab || 'families';
    var store = getStore(tab);
    if (store && row) store.selectedKey = row.getAttribute('data-row-key') || '';
    if (!originalSelectRow) return false;
    if (!quietDetail) return originalSelectRow(row, openDetached);
    var previousSchedule = w.scheduleDetachedDetailSync;
    var previousPreview = w.requestInlinePreview;
    try {
      w.scheduleDetachedDetailSync = function () { return false; };
      w.requestInlinePreview = function (path, detail) {
        try {
          var preview = document.getElementById('preview');
          if (preview && w.renderPreviewFallback) preview.innerHTML = w.renderPreviewFallback(w.preview3dLabel || '', detail || w.previewInlineLoadingLabel || '');
        } catch (ignoredPreview) {}
        return true;
      };
      return originalSelectRow(row, false);
    } finally {
      w.scheduleDetachedDetailSync = previousSchedule;
      w.requestInlinePreview = previousPreview;
    }
  };

  w.checkedFamilyRows = function () { return checkedItems('families'); };
  w.checkedSystemRows = function () { return checkedItems('systems'); };
  w.updateFamilySelectionControls = function () { updateSelectionControls('families'); };
  w.updateSystemSelectionControls = function () { updateSelectionControls('systems'); };

  w.toggleFamilyCheck = function (eventValue, checkbox) {
    stopEvent(eventValue);
    var store = getStore('families');
    if (store && checkbox) {
      var key = checkbox.getAttribute('data-row-key') || '';
      if (checkbox.checked) store.checked[key] = true; else delete store.checked[key];
    }
    updateSelectionControls('families');
    return true;
  };

  w.toggleSystemCheck = function (eventValue, checkbox) {
    stopEvent(eventValue);
    var store = getStore('systems');
    if (store && checkbox) {
      var key = checkbox.getAttribute('data-row-key') || '';
      if (checkbox.checked) store.checked[key] = true; else delete store.checked[key];
    }
    updateSelectionControls('systems');
    return true;
  };

  function toggleAll(tab, eventValue, checkbox) {
    stopEvent(eventValue);
    var store = getStore(tab);
    var rows = renderedRows(tab);
    if (!store) return true;
    for (var i = 0; i < rows.length; i++) {
      var inputs = rows[i].getElementsByTagName('input');
      for (var j = 0; j < inputs.length; j++) {
        if (inputs[j].disabled) continue;
        inputs[j].checked = !!checkbox.checked;
        var key = inputs[j].getAttribute('data-row-key') || '';
        if (checkbox.checked) store.checked[key] = true; else delete store.checked[key];
        break;
      }
    }
    updateSelectionControls(tab);
    return true;
  }

  w.toggleAllFamilyChecks = function (eventValue, checkbox) { return toggleAll('families', eventValue, checkbox); };
  w.toggleAllSystemChecks = function (eventValue, checkbox) { return toggleAll('systems', eventValue, checkbox); };

  w.clearFamilySelection = function () {
    var store = getStore('families');
    if (store) {
      store.checked = {};
      store.selectedKey = '';
      syncRenderedState('families', store);
    }
    updateSelectionControls('families');
    if (w.resetDetailForCurrentTab) w.resetDetailForCurrentTab();
    return false;
  };

  w.clearSystemSelection = function () {
    var store = getStore('systems');
    if (store) {
      store.checked = {};
      store.selectedKey = '';
      syncRenderedState('systems', store);
    }
    updateSelectionControls('systems');
    if (w.resetDetailForCurrentTab) w.resetDetailForCurrentTab();
    return false;
  };

  var previousOnload = w.onload;
  w.onload = function () {
    if (previousOnload) previousOnload();
    ensureColumnResizers(w.currentTab || 'families');
  };

  w.KKYFB = api;
})(window);
