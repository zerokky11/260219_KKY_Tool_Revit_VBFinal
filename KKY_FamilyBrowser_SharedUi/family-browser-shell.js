/* KKY Family Browser dashboard shell asset. */
/* Single-pane shell owner for Revit 2019 IE WebBrowser. */

(function () {
  var shellStamp = '20260710-kky-tool-theme';
  var tabs = ['home', 'families', 'systems', 'requests', 'audit', 'unregisteredfamilies', 'unregisteredsystems', 'permissions', 'admin', 'help', 'settings'];
  var panes = ['homePane', 'familiesPane', 'systemsPane', 'requestsPane', 'auditPane', 'unregisteredFamiliesPane', 'unregisteredSystemsPane', 'permissionsPane', 'adminPane', 'helpPane', 'settingsPane'];

  function trimClass(value) {
    return (value || '').replace(/^\s+|\s+$/g, '').replace(/\s+/g, ' ');
  }

  function css(el, value) {
    if (!el) return;
    try {
      el.style.cssText = value;
    } catch (e) {
      try {
        el.setAttribute('style', value);
      } catch (ignored) {
      }
    }
  }

  function vp() {
    var d = document.documentElement;
    var b = document.body;
    var w = Math.max(900, (d && d.clientWidth) || 0, (b && b.clientWidth) || 0, window.innerWidth || 0);
    var h = Math.max(620, (d && d.clientHeight) || 0, (b && b.clientHeight) || 0, window.innerHeight || 0);
    return { w: w, h: h };
  }

  function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
  }

  function byClass(cls) {
    var all = document.getElementsByTagName('*');
    var needle = ' ' + cls + ' ';
    for (var i = 0; i < all.length; i++) {
      if ((' ' + (all[i].className || '') + ' ').indexOf(needle) >= 0) return all[i];
    }
    return null;
  }

  function findClass(root, cls) {
    if (!root) return null;
    var all = root.getElementsByTagName('*');
    var needle = ' ' + cls + ' ';
    for (var i = 0; i < all.length; i++) {
      if ((' ' + (all[i].className || '') + ' ').indexOf(needle) >= 0) return all[i];
    }
    return null;
  }

  function paneId(tab) {
    if (tab === 'home') return 'homePane';
    if (tab === 'families') return 'familiesPane';
    if (tab === 'systems') return 'systemsPane';
    if (tab === 'requests') return 'requestsPane';
    if (tab === 'audit') return 'auditPane';
    if (tab === 'unregisteredfamilies') return 'unregisteredFamiliesPane';
    if (tab === 'unregisteredsystems') return 'unregisteredSystemsPane';
    if (tab === 'permissions') return 'permissionsPane';
    if (tab === 'admin') return 'adminPane';
    if (tab === 'help') return 'helpPane';
    if (tab === 'settings') return 'settingsPane';
    return 'homePane';
  }

  function isBrowser(tab) {
    return tab === 'families' || tab === 'systems';
  }

  function hasBrowserPanePair() {
    return !!(document.getElementById('familiesPane') && document.getElementById('systemsPane'));
  }

  function wireLocalBrowserTabs() {
    if (!hasBrowserPanePair()) return;
    var names = ['mainNav', 'dashTab'];
    for (var n = 0; n < names.length; n++) {
      var links = document.getElementsByName(names[n]);
      for (var i = 0; i < links.length; i++) {
        var target = links[i].getAttribute('data-tab') || '';
        if (target !== 'families' && target !== 'systems') continue;
        links[i].setAttribute('href', '#');
        links[i].onclick = (function (tab) {
          return function () {
            return window.setTab(tab);
          };
        })(target);
      }
    }
  }

  function removeShellClasses(value) {
    return trimClass((value || '')
      .replace(/(^|\s)fb-tab-(home|families|systems|requests|audit|unregisteredfamilies|unregisteredsystems|permissions|admin|help|settings)(?=\s|$)/g, ' ')
      .replace(/(^|\s)fb-browser(?=\s|$)/g, ' ')
      .replace(/(^|\s)fb-workflow(?=\s|$)/g, ' ')
      .replace(/(^|\s)fb-detached-detail(?=\s|$)/g, ' ')
      .replace(/(^|\s)fb-single-pane(?=\s|$)/g, ' ')
      .replace(/(^|\s)fb-ui-lock-20260507(?=\s|$)/g, ' ')
      .replace(/(^|\s)fb-shell-20260507(?=\s|$)/g, ' '));
  }

  function setBody(tab) {
    var body = document.body;
    if (!body) return;
    body.className = trimClass(removeShellClasses(body.className) + ' fb-shell-20260507 fb-tab-' + tab + (isBrowser(tab) ? ' fb-browser fb-detached-detail' : ' fb-workflow'));
  }

  function setNav(tab) {
    var i;
    var innerTabs = document.getElementsByName('dashTab');
    for (i = 0; i < innerTabs.length; i++) {
      innerTabs[i].className = 'tab' + (innerTabs[i].getAttribute('data-tab') === tab ? ' active' : '');
    }

    var navs = document.getElementsByName('mainNav');
    for (i = 0; i < navs.length; i++) {
      var disabled = (navs[i].getAttribute('data-disabled') || '') === 'true';
      navs[i].className = 'tool' + (navs[i].getAttribute('data-tab') === tab ? ' primary active' : '') + (disabled ? ' disabled' : '');
    }
  }

  function shellTop() {
    var y = 88;
    var top = byClass('top');
    var bar = byClass('toolbar');
    try {
      if (top) y = Math.max(y, top.offsetTop + top.offsetHeight);
      if (bar && (' ' + (bar.className || '') + ' ').indexOf(' fb-nav-10 ') < 0) y = Math.max(y, bar.offsetTop + bar.offsetHeight);
    } catch (e) {
    }
    return y;
  }

  function shellLeft() {
    var bar = byClass('toolbar');
    try {
      if (bar && (' ' + (bar.className || '') + ' ').indexOf(' fb-nav-10 ') >= 0) {
        var expected = vp().w < 1500 ? 188 : 204;
        return Math.max(expected, bar.offsetWidth || 0);
      }
    } catch (e) {
    }
    return 0;
  }

  function shellGap() {
    return shellLeft() > 0 ? 8 : 0;
  }

  function hidePane(id) {
    var pane = byId(id);
    if (!pane) return;
    css(pane, 'display:none !important;visibility:hidden !important;position:absolute !important;left:-99999px !important;top:-99999px !important;width:1px !important;height:1px !important;overflow:hidden !important;');
  }

  function hideInactivePanes(activeId) {
    for (var i = 0; i < panes.length; i++) {
      if (panes[i] !== activeId) hidePane(panes[i]);
    }
  }

  function baseLayoutBox(topPx) {
    var size = vp();
    var layout = byClass('layout');
    var navW = shellLeft();
    var gap = shellGap();
    var left = navW + gap;
    if (document.body) {
      css(document.body, 'margin:0 !important;overflow:hidden !important;width:' + size.w + 'px !important;height:' + size.h + 'px !important;');
    }
    css(layout, 'display:block !important;position:fixed !important;left:' + left + 'px !important;right:0 !important;top:' + topPx + 'px !important;bottom:0 !important;width:auto !important;height:auto !important;min-width:' + Math.max(320, size.w - left) + 'px !important;min-height:' + Math.max(340, size.h - topPx) + 'px !important;padding:0 !important;overflow:hidden !important;box-sizing:border-box !important;');
    return size;
  }

  function styleBrowserSearch(tab) {
    var search = byId('browserSearch');
    var small = vp().w < 1500;
    var minH = small ? 86 : 84;
    css(search, 'display:block !important;visibility:visible !important;position:absolute !important;left:0 !important;right:0 !important;top:0 !important;height:auto !important;min-height:' + minH + 'px !important;overflow-x:auto !important;overflow-y:hidden !important;border-bottom-width:1px !important;border-bottom-style:solid !important;padding:8px 12px !important;box-sizing:border-box !important;z-index:5 !important;');
    var h = minH;
    try {
      h = Math.min(140, Math.max(minH, Math.ceil(search.scrollHeight || 0) + 2));
    } catch (ignoredSearchMeasure) {
      h = minH;
    }
    css(search, 'height:' + h + 'px !important;');
    return h;
  }

  function syncTableHeader(wrap) {
    if (!wrap) return;
    var fixedHeadId = wrap.getAttribute('data-fixed-head') || '';
    if (fixedHeadId && window.syncFixedHeaderScroll) {
      window.syncFixedHeaderScroll(wrap, fixedHeadId);
      return;
    }
    var tables = wrap.getElementsByTagName('table');
    if (!tables || tables.length === 0) return;
    var rows = tables[0].getElementsByTagName('tr');
    if (!rows || rows.length === 0) return;
    var ths = rows[0].getElementsByTagName('th');
    var top = wrap.scrollTop || 0;
    for (var i = 0; i < ths.length; i++) {
      ths[i].style.position = 'relative';
      ths[i].style.top = top + 'px';
      ths[i].style.zIndex = 30;
      ths[i].setAttribute('data-sticky-active', top > 0 ? 'true' : 'false');
      ths[i].style.boxShadow = top > 0 ? '0 1px 0 rgba(100,116,139,.35)' : 'none';
    }
  }

  function lockTableHeaders(root) {
    var scope = root || document;
    var divs = scope.getElementsByTagName ? scope.getElementsByTagName('div') : [];
    for (var i = 0; i < divs.length; i++) {
      var cls = ' ' + (divs[i].className || '') + ' ';
      if (cls.indexOf(' tablewrap ') >= 0 || cls.indexOf(' family-tablewrap ') >= 0) {
        if (!divs[i].getAttribute('data-sticky-head')) {
          divs[i].setAttribute('data-sticky-head', '1');
          divs[i].onscroll = (function (wrap) {
            return function () { syncTableHeader(wrap); };
          })(divs[i]);
        }
        syncTableHeader(divs[i]);
      }
    }
  }

  window.fbRefreshTableHeaders = function () {
    lockTableHeaders(document.body || document);
  };

  function styleBrowserPane(pane, tab, searchH, statusH) {
    if (!pane) return;
    var headBaseH = tab === 'families' ? 58 : 56;
    css(pane, 'display:block !important;visibility:visible !important;position:absolute !important;left:0 !important;right:0 !important;top:' + searchH + 'px !important;bottom:' + statusH + 'px !important;width:auto !important;height:auto !important;overflow:hidden !important;box-sizing:border-box !important;z-index:2 !important;');

    var head = findClass(pane, 'pane-head');
    var headText = findClass(head, 'pane-head-text');
    var kindToggle = findClass(head, 'family-kind-toggle');
    var statusToggle = findClass(head, 'inline-status-toggle');
    css(head, 'display:-ms-flexbox !important;display:flex !important;visibility:visible !important;position:absolute !important;left:0 !important;right:0 !important;top:0 !important;height:auto !important;min-height:' + headBaseH + 'px !important;overflow:visible !important;border-bottom-width:1px !important;border-bottom-style:solid !important;padding:9px 12px 5px 12px !important;box-sizing:border-box !important;z-index:6 !important;-ms-flex-wrap:wrap !important;flex-wrap:wrap !important;-ms-flex-align:center !important;align-items:center !important;');
    css(headText, 'display:none !important;');
    if (kindToggle) {
      css(kindToggle, (tab === 'families' ? 'display:-ms-inline-flexbox !important;display:inline-flex' : 'display:none') + ' !important;vertical-align:middle !important;margin-left:6px !important;margin-bottom:4px !important;white-space:normal !important;overflow:visible !important;text-overflow:clip !important;position:relative !important;z-index:8 !important;-ms-flex-wrap:wrap !important;flex-wrap:wrap !important;-ms-flex-align:center !important;align-items:center !important;');
    }
    if (statusToggle) {
      css(statusToggle, 'display:-ms-inline-flexbox !important;display:inline-flex !important;vertical-align:middle !important;margin-left:6px !important;margin-bottom:4px !important;max-width:none !important;white-space:normal !important;overflow:visible !important;text-overflow:clip !important;position:relative !important;z-index:8 !important;-ms-flex-wrap:wrap !important;flex-wrap:wrap !important;-ms-flex-align:center !important;align-items:center !important;');
    }
    function styleToggleLinks(toggle) {
      if (!toggle) return;
      var links = toggle.getElementsByTagName('a');
      for (var i = 0; i < links.length; i++) {
        css(links[i], 'display:inline-block !important;vertical-align:middle !important;white-space:nowrap !important;overflow:visible !important;text-overflow:clip !important;max-width:none !important;margin:0 4px 4px 0 !important;box-sizing:border-box !important;');
      }
      var spans = toggle.getElementsByTagName('span');
      for (var j = 0; j < spans.length; j++) {
        if ((' ' + (spans[j].className || '') + ' ').indexOf(' kind-label ') >= 0) {
          css(spans[j], 'display:inline-block !important;vertical-align:middle !important;white-space:nowrap !important;overflow:visible !important;text-overflow:clip !important;margin:0 5px 4px 0 !important;box-sizing:border-box !important;');
        }
      }
    }
    styleToggleLinks(kindToggle);
    styleToggleLinks(statusToggle);
    var headH = headBaseH;
    try {
      headH = Math.max(headBaseH, Math.ceil(Math.max(head.scrollHeight || 0, head.offsetHeight || 0)) + 2);
    } catch (ignoredHeadMeasure) {
      headH = headBaseH;
    }
    css(head, 'display:-ms-flexbox !important;display:flex !important;visibility:visible !important;position:absolute !important;left:0 !important;right:0 !important;top:0 !important;height:' + headH + 'px !important;min-height:' + headBaseH + 'px !important;overflow:visible !important;border-bottom-width:1px !important;border-bottom-style:solid !important;padding:9px 12px 5px 12px !important;box-sizing:border-box !important;z-index:6 !important;-ms-flex-wrap:wrap !important;flex-wrap:wrap !important;-ms-flex-align:center !important;align-items:center !important;');

    var grid = findClass(pane, 'family-browser-grid');
    var treeW = (vp().w < 1500) ? 200 : 220;
    css(grid, 'display:block !important;visibility:visible !important;position:absolute !important;left:0 !important;right:0 !important;top:' + headH + 'px !important;bottom:0 !important;width:auto !important;height:auto !important;overflow:hidden !important;box-sizing:border-box !important;z-index:1 !important;');
    css(findClass(grid, 'family-tree-panel'), 'display:block !important;visibility:visible !important;position:absolute !important;left:0 !important;top:0 !important;bottom:0 !important;width:' + treeW + 'px !important;overflow:auto !important;border-right-width:1px !important;border-right-style:solid !important;padding:9px 8px !important;box-sizing:border-box !important;z-index:2 !important;');
    css(findClass(grid, 'family-list-panel'), 'display:block !important;visibility:visible !important;position:absolute !important;left:' + treeW + 'px !important;right:0 !important;top:0 !important;bottom:0 !important;overflow:hidden !important;box-sizing:border-box !important;z-index:1 !important;');

    var divs = grid ? grid.getElementsByTagName('div') : [];
    for (var i = 0; i < divs.length; i++) {
      var cls = ' ' + (divs[i].className || '') + ' ';
      if (cls.indexOf(' tablewrap ') >= 0 || cls.indexOf(' family-tablewrap ') >= 0) {
        var bodyTop = (cls.indexOf(' has-fixed-head ') >= 0 || divs[i].getAttribute('data-fixed-head')) ? 46 : 0;
        css(divs[i], 'display:block !important;visibility:visible !important;position:absolute !important;left:0 !important;right:0 !important;top:' + bodyTop + 'px !important;bottom:0 !important;overflow:auto !important;box-sizing:border-box !important;');
      }
    }
    lockTableHeaders(grid);
  }

  function layoutBrowser(tab) {
    var topPx = shellTop();
    var size = baseLayoutBox(topPx);
    var pad = 10;
    var center = byId('mainCenter');
    var detail = byId('selectionDetailPanel');
    var status = byClass('statusbar');
    var pane = byId(paneId(tab));
    var navW = shellLeft();
    var gap = shellGap();
    var left = navW + gap + pad;
    var minW = Math.max(520, size.w - left - pad);
    var minH = Math.max(300, size.h - topPx - pad * 2);

    hideInactivePanes(paneId(tab));
    css(center, 'display:block !important;visibility:visible !important;position:fixed !important;left:' + left + 'px !important;right:' + pad + 'px !important;top:' + (topPx + pad) + 'px !important;bottom:' + pad + 'px !important;width:auto !important;height:auto !important;min-width:' + minW + 'px !important;min-height:' + minH + 'px !important;overflow:hidden !important;border-width:1px !important;border-style:solid !important;border-radius:6px !important;box-sizing:border-box !important;z-index:10 !important;');
    css(detail, 'display:none !important;visibility:hidden !important;position:absolute !important;left:-99999px !important;top:-99999px !important;right:auto !important;width:1px !important;height:1px !important;overflow:hidden !important;pointer-events:none !important;');
    css(status, 'display:block !important;visibility:visible !important;position:absolute !important;left:0 !important;right:0 !important;bottom:0 !important;height:50px !important;line-height:normal !important;overflow:hidden !important;border-top-width:1px !important;border-top-style:solid !important;padding:0 12px !important;box-sizing:border-box !important;z-index:4 !important;');
    var searchH = styleBrowserSearch(tab);
    styleBrowserPane(pane, tab, searchH, 50);
    if (window.fbLog) fbLog('dashboard shell browser ' + tab + ' detached detail hidden');
  }

  function styleWorkflowPane(pane, tab) {
    if (!pane) return;
    var scroll = findClass(pane, 'request-pane-scroll');
    var head = findClass(pane, 'pane-head');
    if (tab === 'requests' || tab === 'unregisteredfamilies' || tab === 'unregisteredsystems') {
      css(pane, 'display:block !important;visibility:visible !important;position:absolute !important;left:0 !important;right:0 !important;top:0 !important;bottom:0 !important;width:auto !important;height:auto !important;min-height:0 !important;overflow:auto !important;box-sizing:border-box !important;z-index:2 !important;padding:0 !important;');
      css(head, 'display:block !important;visibility:visible !important;position:relative !important;left:auto !important;right:auto !important;top:auto !important;height:auto !important;overflow:visible !important;border-bottom-width:1px !important;border-bottom-style:solid !important;padding:12px 16px !important;box-sizing:border-box !important;z-index:3 !important;');
      css(scroll, 'display:block !important;visibility:visible !important;position:relative !important;left:auto !important;right:auto !important;top:auto !important;bottom:auto !important;width:auto !important;height:auto !important;min-height:0 !important;overflow:visible !important;padding:22px !important;box-sizing:border-box !important;z-index:2 !important;');
    } else {
      var panePad = (tab === 'admin') ? 14 : 24;
      css(pane, 'display:block !important;visibility:visible !important;position:absolute !important;left:0 !important;right:0 !important;top:0 !important;bottom:0 !important;width:auto !important;height:auto !important;min-height:0 !important;overflow:auto !important;box-sizing:border-box !important;z-index:2 !important;padding:' + panePad + 'px !important;');
      css(head, 'display:none !important;');
      css(scroll, 'display:block !important;visibility:visible !important;position:static !important;left:auto !important;right:auto !important;top:auto !important;bottom:auto !important;width:auto !important;height:auto !important;min-height:0 !important;overflow:visible !important;padding:0 !important;box-sizing:border-box !important;');
    }
  }

  function layoutWorkflow(tab) {
    var topPx = shellTop();
    var size = baseLayoutBox(topPx);
    var pad = 0;
    var center = byId('mainCenter');
    var detail = byId('selectionDetailPanel');
    var search = byId('browserSearch');
    var status = byClass('statusbar');
    var pane = byId(paneId(tab));
    var navW = shellLeft();
    var gap = shellGap();
    var left = navW + gap;

    hideInactivePanes(paneId(tab));
    css(search, 'display:none !important;visibility:hidden !important;');
    css(detail, 'display:none !important;visibility:hidden !important;position:absolute !important;left:-99999px !important;top:-99999px !important;width:1px !important;height:1px !important;overflow:hidden !important;');
    css(status, 'display:none !important;visibility:hidden !important;');
    css(center, 'display:block !important;visibility:visible !important;position:fixed !important;left:' + left + 'px !important;right:0 !important;top:' + topPx + 'px !important;bottom:0 !important;width:auto !important;height:auto !important;min-width:' + Math.max(420, size.w - left) + 'px !important;min-height:' + Math.max(260, size.h - topPx) + 'px !important;overflow:hidden !important;border-width:0 0 0 1px !important;border-style:solid !important;border-radius:0 !important;box-sizing:border-box !important;z-index:10 !important;');
    styleWorkflowPane(pane, tab);
    if (window.fbLog) fbLog('dashboard shell workflow ' + tab + ' pane=' + paneId(tab));
  }

  function resetDetailedFilterState() {
    advStatus = 'All';
    advGroup = 'All';
    advCategory = '';
    advMismatchOnly = false;
    var status = document.getElementById('advStatus');
    var group = document.getElementById('advGroup');
    var category = document.getElementById('advCategory');
    var mismatch = document.getElementById('advMismatchOnly');
    var mask = document.getElementById('filterMask');
    if (status) status.value = 'All';
    if (group) group.value = 'All';
    if (category) category.value = '';
    if (mismatch) mismatch.checked = false;
    if (mask) mask.style.display = 'none';
  }

  function resetWorkflowFilters(previousTab, nextTab) {
    if (isBrowser(previousTab) && isBrowser(nextTab) && previousTab !== nextTab) {
      resetDetailedFilterState();
      return;
    }
    if (isBrowser(nextTab)) return;
    currentFilter = 'All';
    resetDetailedFilterState();
  }

  function layout(tab) {
    currentTab = tab || currentTab || 'home';
    setBody(currentTab);
    setNav(currentTab);
    if (isBrowser(currentTab)) {
      layoutBrowser(currentTab);
    } else {
      layoutWorkflow(currentTab);
    }
  }

  window.fbAssetShellRefresh = function () {
    layout(currentTab);
  };

  window.applyFamilyBrowserTheme = function (themeCode) {
    var body = document.body;
    if (!body) return false;
    var code = String(themeCode || '').toLowerCase() === 'dark' ? 'dark' : 'light';
    body.className = trimClass((body.className || '')
      .replace(/(^|\s)theme-light(?=\s|$)/g, ' ')
      .replace(/(^|\s)theme-dark(?=\s|$)/g, ' ') + ' theme-' + code);
    body.setAttribute('data-theme', code);
    var toggle = document.getElementById('fbThemeToggle');
    if (toggle) {
      var target = code === 'dark' ? 'light' : 'dark';
      var label = toggle.getAttribute('data-' + target + '-label') || target;
      var title = toggle.getAttribute('data-' + target + '-title') || label;
      toggle.setAttribute('data-theme-target', target);
      toggle.setAttribute('title', title);
      toggle.innerHTML = '<span class="theme-toggle-mark">' + (target === 'light' ? 'L' : 'D') + '</span>' + label;
    }
    layout(currentTab);
    if (window.fitVisiblePreviewImages) window.fitVisiblePreviewImages();
    if (window.fbLog) fbLog('theme applied ' + code);
    return true;
  };

  window.isBrowserTab = function (tab) {
    return isBrowser(tab);
  };

  window.workflowPaneId = function (tab) {
    return paneId(tab);
  };

  window.setBodyTabClass = function (tab) {
    setBody(tab);
  };

  window.updateSearchChrome = function () {
    layout(currentTab);
  };

  window.forceWorkflowPane = function (tab) {
    currentTab = tab || currentTab;
    layout(currentTab);
    return false;
  };

  window.setTab = function (tab) {
    try {
      var started = new Date().getTime();
      var previousTab = currentTab || '';
      var nextTab = tab || currentTab || 'home';
      resetWorkflowFilters(previousTab, nextTab);
      currentTab = nextTab;
      if (window.applyViewportDensity) applyViewportDensity();
      layout(currentTab);
      if (currentTab === 'families') {
        safeUi('updateFamilyTreeActive', updateFamilyTreeActive);
        safeUi('filterRows', filterRows);
      } else if (currentTab === 'systems') {
        safeUi('updateSystemTreeActive', updateSystemTreeActive);
        safeUi('filterRows', filterRows);
      }
      if (window.fitVisiblePreviewImages) window.fitVisiblePreviewImages();
      if (window.fbLog) fbLog('dashboard shell setTab ' + currentTab + ' ms=' + (new Date().getTime() - started));
      return false;
    } catch (e) {
      if (window.fbErr) fbErr('dashboard shell setTab ' + tab, e);
      return false;
    }
  };

  window.onresize = function () {
    try {
      if (window.applyViewportDensity) applyViewportDensity();
      layout(currentTab);
      if (window.fitVisiblePreviewImages) window.fitVisiblePreviewImages();
    } catch (e) {
      if (window.fbErr) fbErr('dashboard shell resize', e);
    }
  };

  window.onload = function () {
      if (window.fbLog) fbLog('dashboard shell active');
    if (window.applyViewportDensity) applyViewportDensity();
    if (window.applyPermissionUi) applyPermissionUi();
    wireLocalBrowserTabs();
    window.setTab(currentTab);
    if (window.fitVisiblePreviewImages) window.fitVisiblePreviewImages();
    if (window.applyControlTitles) applyControlTitles();
    if (window.fbLog) fbLog('dashboard shell ready ' + shellStamp);
  };
})();

(function () {
  function installObjectDiffStyles() {
    var cssText = ''
      + '.diff-modal-stage .fingerprint-diff-section-title{margin:0 0 8px 0;padding:8px 10px;border:1px solid #d8e1dc;border-radius:6px;background:#eef6f2;color:#20372f;font-size:13px;font-weight:800;}'
      + '.diff-modal-stage .fingerprint-diff-object-table{width:100%;min-width:1320px;border-collapse:collapse;table-layout:fixed;background:#fff;border:1px solid #d8e1dc;border-radius:7px;overflow:hidden;margin:0 0 14px 0;}'
      + '.diff-modal-stage .fingerprint-diff-object-table th{position:static;background:#f7fbf8;color:#3f534b;font-size:12px;font-weight:800;padding:8px 8px;border-bottom:1px solid #d8e1dc;text-align:left;white-space:nowrap;}'
      + '.diff-modal-stage .fingerprint-diff-object-table td{font-size:12px;line-height:1.32;color:#17231f;padding:8px 8px;border-top:1px solid #edf2ef;vertical-align:top;white-space:normal;word-break:keep-all;overflow-wrap:anywhere;}'
      + '.diff-modal-stage .fingerprint-diff-object-table col.diff-object-area{width:132px}.diff-modal-stage .fingerprint-diff-object-table col.diff-object-id{width:92px}.diff-modal-stage .fingerprint-diff-object-table col.diff-object-category{width:128px}.diff-modal-stage .fingerprint-diff-object-table col.diff-object-family{width:150px}.diff-modal-stage .fingerprint-diff-object-table col.diff-object-value{width:190px}.diff-modal-stage .fingerprint-diff-object-table col.diff-object-note{width:130px}'
      + '.diff-modal-stage .fingerprint-diff-nested-table{width:100%;min-width:980px;border-collapse:collapse;table-layout:fixed;background:#fff;border:1px solid #d8e1dc;border-radius:7px;overflow:hidden;margin:0 0 14px 0;}'
      + '.diff-modal-stage .fingerprint-diff-nested-table th{position:static;background:#f2faf6;color:#25493c;font-size:12px;font-weight:800;padding:8px 8px;border-bottom:1px solid #d8e1dc;text-align:left;white-space:nowrap;}'
      + '.diff-modal-stage .fingerprint-diff-nested-table td{font-size:12px;line-height:1.32;color:#17231f;padding:8px 8px;border-top:1px solid #edf2ef;vertical-align:top;white-space:normal;word-break:keep-all;overflow-wrap:anywhere;}'
      + '.diff-modal-stage .fingerprint-diff-nested-table col.diff-nested-area{width:150px}.diff-modal-stage .fingerprint-diff-nested-table col.diff-nested-family{width:230px}.diff-modal-stage .fingerprint-diff-nested-table col.diff-nested-type{width:230px}.diff-modal-stage .fingerprint-diff-nested-table col.diff-nested-note{width:150px}';
    try {
      var style = document.createElement('style');
      style.type = 'text/css';
      if (style.styleSheet) style.styleSheet.cssText = cssText;
      else style.appendChild(document.createTextNode(cssText));
      (document.getElementsByTagName('head')[0] || document.documentElement).appendChild(style);
    } catch (e) {
    }
  }

  function html(v) {
    if (window.escHtml) return window.escHtml(v || '');
    return (v || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function diffCellHtml(v) {
    return html(v || '').replace(/ \[\[BR\]\] /g, '<br>').replace(/\[\[BR\]\]/g, '<br>');
  }

  function diffCellTitle(v) {
    return html((v || '').replace(/ \[\[BR\]\] /g, '\n').replace(/\[\[BR\]\]/g, '\n'));
  }

  function trim(v) {
    return (v || '').replace(/^\s+|\s+$/g, '');
  }

  function clean(v) {
    return trim((v || '').replace(/\r?\n/g, ' ').replace(/\s+/g, ' '));
  }

  function has(v) {
    v = clean(v);
    return v !== '' && v !== '-';
  }

  function lower(v) {
    return clean(v).toLowerCase();
  }

  function shortText(v, max) {
    v = clean(v);
    if (!has(v)) return '-';
    max = max || 80;
    return v.length > max ? v.substring(0, max - 3) + '...' : v;
  }

  function readLabeledField(text, labels) {
    var normalized = (text || '').replace(/ \[\[BR\]\] /g, '|').replace(/\[\[BR\]\]/g, '|').replace(/\s+·\s+/g, '|').replace(/\s*;\s*/g, '|');
    var parts = normalized.split('|');
    for (var i = 0; i < parts.length; i++) {
      var part = trim(parts[i]);
      var partLower = part.toLowerCase();
      for (var l = 0; l < labels.length; l++) {
        var label = labels[l].toLowerCase();
        if (partLower.indexOf(label + ':') === 0) {
          return clean(part.substring(labels[l].length + 1));
        }
      }
    }
    return '';
  }

  function extractIds(text) {
    text = text || '';
    var ids = [];
    var seen = {};
    var match;
    var re = /ID\s*:?\s*([0-9]+)/ig;
    while ((match = re.exec(text)) !== null) {
      if (!seen[match[1]]) {
        seen[match[1]] = true;
        ids.push(match[1]);
      }
      if (ids.length >= 8) break;
    }
    var more = /\+([0-9]+)/.exec(text);
    return ids.join(', ') + (more ? ' +' + more[1] : '');
  }

  function stripIds(text) {
    text = text || '';
    text = text.replace(/\[[^\]]*ID[^\]]*\]/ig, ' ');
    text = text.replace(/\bID\s*:?\s*[0-9, +]+/ig, ' ');
    return clean(text);
  }

  function splitRawTokens(text) {
    text = stripIds(text).replace(/\s+·\s+/g, '/').replace(/\|/g, '/');
    var raw = text.split('/');
    var out = [];
    for (var i = 0; i < raw.length; i++) {
      var token = clean(raw[i]);
      if (!token || token === '-') continue;
      out.push(token);
    }
    return out;
  }

  function parseRawFallback(text, side) {
    var tokens = splitRawTokens(text);
    if (tokens.length === 0) return side;
    var kind = lower(tokens[0]);
    if (!has(side.category) && tokens.length > 1) side.category = tokens[1];
    if (kind.indexOf('familyinstance') >= 0 || kind.indexOf('하위') >= 0) {
      if (!has(side.family) && tokens.length > 2) side.family = tokens[2];
      for (var i = 0; i < tokens.length; i++) {
        if (lower(tokens[i]).indexOf('familysymbol') >= 0) {
          if (!has(side.typeName) && i + 2 < tokens.length) side.typeName = tokens[i + 2];
          break;
        }
      }
      if (!has(side.typeName) && tokens.length > 3) side.typeName = tokens[tokens.length - 1];
    } else if (kind.indexOf('dimension') >= 0 || kind.indexOf('치수') >= 0) {
      if (!has(side.name) && tokens.length > 2) side.name = tokens[2];
      if (!has(side.typeName) && tokens.length > 3) side.typeName = tokens[tokens.length - 1];
    } else if (kind.indexOf('material') >= 0 || kind.indexOf('재료') >= 0) {
      if (!has(side.material) && tokens.length > 2) side.material = tokens[2];
      if (!has(side.name) && has(side.material)) side.name = side.material;
    } else {
      if (!has(side.name) && tokens.length > 2) side.name = tokens[2];
      if (!has(side.typeName) && tokens.length > 3) side.typeName = tokens[tokens.length - 1];
    }
    return side;
  }

  function parseDiffSide(text) {
    text = clean(text);
    var side = { id: '-', category: '-', family: '-', typeName: '-', name: '-', material: '-', raw: text };
    if (!has(text)) return side;
    side.id = extractIds(text) || '-';
    side.category = readLabeledField(text, ['Category', '카테고리']) || '-';
    side.family = readLabeledField(text, ['Nested Family', 'Family', '하위 패밀리', '패밀리']) || '-';
    side.typeName = readLabeledField(text, ['Placed Type', 'Nested Type', 'Type', '배치 타입', '하위 타입', '타입']) || '-';
    side.name = readLabeledField(text, ['Name', '이름']) || '-';
    side.material = readLabeledField(text, ['Material', '재료']) || '-';
    side = parseRawFallback(text, side);
    return side;
  }

  function readDebugToken(text, key) {
    text = text || '';
    var re = new RegExp(key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '="([^"]*)"', 'i');
    var m = re.exec(text);
    return m ? clean(m[1]) : '';
  }

  function isNestedDiffRow(row) {
    var key = lower(row ? row.area : '');
    return key.indexOf('하위') >= 0 || key.indexOf('nested') >= 0;
  }

  function parseNestedDiffSide(text) {
    text = clean(text);
    var side = { id: extractIds(text) || '-', family: '-', typeName: '-', raw: text };
    side.family = readLabeledField(text, ['Nested Family', 'Family', '하위 패밀리', '패밀리'])
      || readDebugToken(text, 'nestedFamily')
      || readDebugToken(text, 'family')
      || '-';
    side.typeName = readLabeledField(text, ['Type', 'Nested Type', 'Placed Type', '타입', '하위 타입', '배치 타입'])
      || readDebugToken(text, 'nestedType')
      || '-';

    var tokens = splitRawTokens(text);
    if (tokens.length > 0) {
      var kind = lower(tokens[0]);
      if (kind.indexOf('familyinstance') >= 0) {
        if (!has(side.family) && tokens.length > 2) side.family = tokens[2];
        if (!has(side.typeName) && tokens.length > 3) side.typeName = tokens[3];
      } else if (kind.indexOf('familysymbol') >= 0) {
        if (!has(side.typeName) && tokens.length > 2) side.typeName = tokens[2];
      } else if (kind === 'family' || kind.indexOf('family') >= 0) {
        if (!has(side.family) && tokens.length > 2) side.family = tokens[2];
      }
    }

    return side;
  }

  function objectValue(side) {
    if (has(side.material)) return side.material;
    if (has(side.typeName)) return side.typeName;
    if (has(side.name)) return side.name;
    return shortText(side.raw, 80);
  }

  function sideHasFields(side) {
    return has(side.id) || has(side.category) || has(side.family) || has(side.typeName) || has(side.name) || has(side.material);
  }

  function isObjectDiffRow(row) {
    if (isNestedDiffRow(row)) return false;
    var key = lower(row.area);
    if (key.indexOf('치수') >= 0 || key.indexOf('dimension') >= 0) return true;
    if (key.indexOf('재료') >= 0 || key.indexOf('material') >= 0) return true;
    if (key.indexOf('요소') >= 0 || key.indexOf('element') >= 0) return true;
    var s = parseDiffSide(row.standard);
    var p = parseDiffSide(row.project);
    return sideHasFields(s) || sideHasFields(p);
  }

  function renderObjectDiffTable(rows) {
    if (!rows || rows.length === 0) return '';
    var itemLabel = window.diffAreaLabel || 'Item';
    var stdLabel = window.diffStandardLabel || 'Standard';
    var prjLabel = window.diffProjectLabel || 'Current Project';
    var noteLabel = window.diffNoteLabel || 'Difference';
    var isKorean = itemLabel.indexOf('항목') >= 0 || stdLabel.indexOf('표준') >= 0 || prjLabel.indexOf('프로젝트') >= 0;
    var idLabel = 'ID';
    var categoryLabel = (typeof categoryLabelText !== 'undefined') ? categoryLabelText : (isKorean ? '카테고리' : 'Category');
    var familyLabel = (typeof detailFamilyLabel !== 'undefined') ? detailFamilyLabel : (isKorean ? '패밀리' : 'Family');
    var valueLabel = isKorean ? '타입/이름/재료' : 'Type/Name/Material';
    var htmlOut = '<div class="fingerprint-diff-section-title">' + html(itemLabel) + '</div>';
    htmlOut += '<table class="fingerprint-diff-object-table"><colgroup><col class="diff-object-area"/><col class="diff-object-id"/><col class="diff-object-category"/><col class="diff-object-family"/><col class="diff-object-value"/><col class="diff-object-id"/><col class="diff-object-category"/><col class="diff-object-family"/><col class="diff-object-value"/><col class="diff-object-note"/></colgroup>';
    htmlOut += '<tr><th>' + html(itemLabel) + '</th><th>' + html(stdLabel + ' ' + idLabel) + '</th><th>' + html(stdLabel + ' ' + categoryLabel) + '</th><th>' + html(stdLabel + ' ' + familyLabel) + '</th><th>' + html(stdLabel + ' ' + valueLabel) + '</th><th>' + html(prjLabel + ' ' + idLabel) + '</th><th>' + html(prjLabel + ' ' + categoryLabel) + '</th><th>' + html(prjLabel + ' ' + familyLabel) + '</th><th>' + html(prjLabel + ' ' + valueLabel) + '</th><th>' + html(noteLabel) + '</th></tr>';
    for (var i = 0; i < rows.length; i++) {
      var std = parseDiffSide(rows[i].standard);
      var prj = parseDiffSide(rows[i].project);
      htmlOut += '<tr>';
      htmlOut += '<td title="' + html(rows[i].area) + '">' + html(shortText(rows[i].area, 42)) + '</td>';
      htmlOut += '<td title="' + html(std.id) + '">' + html(std.id) + '</td>';
      htmlOut += '<td title="' + html(std.category) + '">' + html(shortText(std.category, 42)) + '</td>';
      htmlOut += '<td title="' + html(std.family) + '">' + html(shortText(std.family, 42)) + '</td>';
      htmlOut += '<td title="' + html(objectValue(std)) + '">' + html(shortText(objectValue(std), 54)) + '</td>';
      htmlOut += '<td title="' + html(prj.id) + '">' + html(prj.id) + '</td>';
      htmlOut += '<td title="' + html(prj.category) + '">' + html(shortText(prj.category, 42)) + '</td>';
      htmlOut += '<td title="' + html(prj.family) + '">' + html(shortText(prj.family, 42)) + '</td>';
      htmlOut += '<td title="' + html(objectValue(prj)) + '">' + html(shortText(objectValue(prj), 54)) + '</td>';
      htmlOut += '<td title="' + html(rows[i].note) + '">' + html(shortText(rows[i].note, 48)) + '</td>';
      htmlOut += '</tr>';
    }
    return htmlOut + '</table>';
  }

  function renderNestedDiffTable(rows) {
    if (!rows || rows.length === 0) return '';
    var itemLabel = window.diffAreaLabel || 'Item';
    var stdLabel = window.diffStandardLabel || 'Standard';
    var prjLabel = window.diffProjectLabel || 'Current Project';
    var noteLabel = window.diffNoteLabel || 'Difference';
    var isKorean = itemLabel.indexOf('항목') >= 0 || stdLabel.indexOf('표준') >= 0 || prjLabel.indexOf('프로젝트') >= 0;
    var familyLabel = isKorean ? '하위 패밀리' : 'Nested Family';
    var typeLabel = isKorean ? '타입' : 'Type';
    var htmlOut = '<div class="fingerprint-diff-section-title">' + html(isKorean ? '하위 패밀리 / 타입' : 'Nested Families / Types') + '</div>';
    htmlOut += '<table class="fingerprint-diff-nested-table"><colgroup><col class="diff-nested-area"/><col class="diff-nested-family"/><col class="diff-nested-type"/><col class="diff-nested-family"/><col class="diff-nested-type"/><col class="diff-nested-note"/></colgroup>';
    htmlOut += '<tr><th>' + html(itemLabel) + '</th><th>' + html(stdLabel + ' ' + familyLabel) + '</th><th>' + html(stdLabel + ' ' + typeLabel) + '</th><th>' + html(prjLabel + ' ' + familyLabel) + '</th><th>' + html(prjLabel + ' ' + typeLabel) + '</th><th>' + html(noteLabel) + '</th></tr>';
    for (var i = 0; i < rows.length; i++) {
      var std = parseNestedDiffSide(rows[i].standard);
      var prj = parseNestedDiffSide(rows[i].project);
      var stdTitle = (std.id !== '-' ? 'ID ' + std.id + '\n' : '') + rows[i].standard;
      var prjTitle = (prj.id !== '-' ? 'ID ' + prj.id + '\n' : '') + rows[i].project;
      htmlOut += '<tr>';
      htmlOut += '<td title="' + html(rows[i].area) + '">' + html(shortText(rows[i].area, 44)) + '</td>';
      htmlOut += '<td title="' + html(stdTitle) + '">' + html(shortText(std.family, 64)) + '</td>';
      htmlOut += '<td title="' + html(stdTitle) + '">' + html(shortText(std.typeName, 64)) + '</td>';
      htmlOut += '<td title="' + html(prjTitle) + '">' + html(shortText(prj.family, 64)) + '</td>';
      htmlOut += '<td title="' + html(prjTitle) + '">' + html(shortText(prj.typeName, 64)) + '</td>';
      htmlOut += '<td title="' + html(rows[i].note) + '">' + html(shortText(rows[i].note, 48)) + '</td>';
      htmlOut += '</tr>';
    }
    return htmlOut + '</table>';
  }

  function renderSimpleDiffTable(rows) {
    if (!rows || rows.length === 0) return '';
    var htmlOut = '<table class="fingerprint-diff-table"><colgroup><col class="diff-area"/><col class="diff-side"/><col class="diff-side"/><col class="diff-note"/></colgroup><tr><th>' + html(window.diffAreaLabel || 'Item') + '</th><th>' + html(window.diffStandardLabel || 'Standard') + '</th><th>' + html(window.diffProjectLabel || 'Current Project') + '</th><th>' + html(window.diffNoteLabel || 'Difference') + '</th></tr>';
    for (var r = 0; r < rows.length; r++) {
      htmlOut += '<tr><td title="' + diffCellTitle(rows[r].area) + '">' + diffCellHtml(rows[r].area) + '</td><td title="' + diffCellTitle(rows[r].standard) + '">' + diffCellHtml(rows[r].standard) + '</td><td title="' + diffCellTitle(rows[r].project) + '">' + diffCellHtml(rows[r].project) + '</td><td title="' + diffCellTitle(rows[r].note) + '">' + diffCellHtml(rows[r].note) + '</td></tr>';
    }
    return htmlOut + '</table>';
  }

  window.renderFingerprintDiffDetailTable = function (rows) {
    rows = window.sortFingerprintDiffRows ? window.sortFingerprintDiffRows(rows || []) : (rows || []);
    var nestedRows = [];
    var objectRows = [];
    var simpleRows = [];
    for (var i = 0; i < rows.length; i++) {
      if (isNestedDiffRow(rows[i])) nestedRows.push(rows[i]);
      else if (isObjectDiffRow(rows[i])) objectRows.push(rows[i]);
      else simpleRows.push(rows[i]);
    }
    return renderNestedDiffTable(nestedRows) + renderObjectDiffTable(objectRows) + renderSimpleDiffTable(simpleRows);
  };

  installObjectDiffStyles();
})();
