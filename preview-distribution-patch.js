(function() {
  var STORAGE_KEY = 'poker_preview_distribution_view';
  var reapplyingDistribution = false;

  function getStoredDistributionView() {
    try {
      return window.localStorage && window.localStorage.getItem(STORAGE_KEY) === '1';
    }
    catch (err) {
      return false;
    }
  }

  function setStoredDistributionView(active) {
    try {
      if (!window.localStorage) { return; }
      if (active) {
        window.localStorage.setItem(STORAGE_KEY, '1');
      }
      else {
        window.localStorage.removeItem(STORAGE_KEY);
      }
    }
    catch (err) {}
  }

  function injectDistributionStyles() {
    if (document.getElementById('preview-distribution-patch-style')) { return; }
    var style = document.createElement('style');
    style.id = 'preview-distribution-patch-style';
    style.textContent = [
      '.stack-distribution-table { width: min(860px, 100%); }',
      '.stack-distribution-table th, .stack-distribution-table td { text-align: center; }',
      '.stack-distribution-table td.total-player { font-size: 15px; font-weight: 700; }'
    ].join('\n');
    document.head.appendChild(style);
  }

  function getTabsHost() {
    return document.getElementById('final-stack-view-tabs');
  }

  function getHost() {
    return document.getElementById('final-stack-host');
  }

  function getNativeTabButtons() {
    var tabsHost = getTabsHost();
    if (!tabsHost) { return []; }
    return Array.prototype.slice.call(tabsHost.querySelectorAll('.stack-view-tab')).filter(function(button) {
      return button.getAttribute('data-preview-distribution') !== '1';
    });
  }

  function isDistributionRendered() {
    return !!document.querySelector('#final-stack-host .stack-distribution-table');
  }

  function buildHero() {
    var hero = document.createElement('section');
    hero.className = 'stack-view-card';

    var kicker = document.createElement('p');
    kicker.className = 'stack-view-kicker';
    kicker.textContent = 'Distribution rapide';

    var title = document.createElement('h3');
    title.className = 'stack-view-title';
    title.textContent = 'Nom du joueur et jetons à remettre';

    var detail = document.createElement('p');
    detail.className = 'stack-view-detail';
    detail.textContent = "Cette vue retire les colonnes de vérification pour ne laisser que le strict minimum au moment de préparer et distribuer les piles de jetons.";

    hero.appendChild(kicker);
    hero.appendChild(title);
    hero.appendChild(detail);
    return hero;
  }

  function buildDistributionTableFromOverview() {
    var host = getHost();
    if (!host) { return false; }

    var overviewTable = host.querySelector('table.stack-view-table');
    if (!overviewTable) { return false; }

    var headerCells = Array.prototype.slice.call(overviewTable.querySelectorAll('thead th'));
    var chipColumns = [];
    headerCells.forEach(function(cell, idx) {
      if ((cell.className || '').indexOf('stack-chip-col-') !== -1) {
        chipColumns.push({
          index: idx,
          label: (cell.textContent || '').trim(),
          className: cell.className || ''
        });
      }
    });

    if (!chipColumns.length) { return false; }

    var bodyRows = Array.prototype.slice.call(overviewTable.querySelectorAll('tbody tr'));

    host.innerHTML = '';
    host.appendChild(buildHero());

    var tableWrap = document.createElement('div');
    tableWrap.className = 'stack-calc-table-wrap';

    var table = document.createElement('table');
    table.className = 'total-table stack-calc-table stack-view-table stack-distribution-table';

    var thead = document.createElement('thead');
    var headRow = document.createElement('tr');
    ['Joueur'].concat(chipColumns.map(function(column) { return column.label; })).forEach(function(label, idx) {
      var th = document.createElement('th');
      th.textContent = label;
      if (idx >= 1) {
        th.className = chipColumns[idx - 1].className;
      }
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);

    var tbody = document.createElement('tbody');
    bodyRows.forEach(function(row) {
      if (row.classList.contains('stack-calc-group-row')) {
        var groupRow = document.createElement('tr');
        groupRow.className = row.className;
        var groupCell = document.createElement('td');
        groupCell.colSpan = 1 + chipColumns.length;
        groupCell.textContent = (row.textContent || '').trim();
        groupRow.appendChild(groupCell);
        tbody.appendChild(groupRow);
        return;
      }

      if (!row.cells || row.cells.length < 2) { return; }
      var tr = document.createElement('tr');
      var nameCell = document.createElement('td');
      nameCell.textContent = (row.cells[1].textContent || '').trim();
      nameCell.className = 'total-player';
      tr.appendChild(nameCell);

      chipColumns.forEach(function(column) {
        var td = document.createElement('td');
        td.textContent = row.cells[column.index] ? (row.cells[column.index].textContent || '').trim() : '0';
        td.className = column.className;
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tableWrap.appendChild(table);
    host.appendChild(tableWrap);
    return true;
  }

  function updateCustomButtonState() {
    var tabsHost = getTabsHost();
    if (!tabsHost) { return; }
    var customButton = tabsHost.querySelector('[data-preview-distribution="1"]');
    if (!customButton) { return; }
    var active = getStoredDistributionView();
    customButton.classList.toggle('active', active);
    if (active) {
      getNativeTabButtons().forEach(function(button) {
        button.classList.remove('active');
      });
    }
  }

  function activateDistributionView() {
    if (reapplyingDistribution || isDistributionRendered()) {
      updateCustomButtonState();
      return;
    }

    var overviewButton = getNativeTabButtons().filter(function(button) {
      return ((button.textContent || '').trim() === 'Vue complète');
    })[0] || null;

    reapplyingDistribution = true;

    function finishRender() {
      window.requestAnimationFrame(function() {
        buildDistributionTableFromOverview();
        reapplyingDistribution = false;
        updateCustomButtonState();
      });
    }

    if (overviewButton && !overviewButton.classList.contains('active')) {
      overviewButton.click();
      window.setTimeout(finishRender, 0);
      return;
    }

    finishRender();
  }

  function ensureCustomButton() {
    var tabsHost = getTabsHost();
    if (!tabsHost) { return; }

    getNativeTabButtons().forEach(function(button) {
      if (button.getAttribute('data-preview-distribution-bound') === '1') { return; }
      button.setAttribute('data-preview-distribution-bound', '1');
      button.addEventListener('click', function() {
        setStoredDistributionView(false);
        updateCustomButtonState();
      });
    });

    var customButton = tabsHost.querySelector('[data-preview-distribution="1"]');
    if (!customButton) {
      customButton = document.createElement('button');
      customButton.type = 'button';
      customButton.className = 'stack-view-tab';
      customButton.setAttribute('data-preview-distribution', '1');
      customButton.textContent = 'Distribution';
      customButton.addEventListener('click', function() {
        setStoredDistributionView(true);
        activateDistributionView();
      });
      tabsHost.appendChild(customButton);
    }

    updateCustomButtonState();
  }

  function refreshDistributionView() {
    injectDistributionStyles();
    ensureCustomButton();

    if (getStoredDistributionView()) {
      activateDistributionView();
    }
  }

  function observeCalculator() {
    var observed = [];
    [getTabsHost(), getHost()].forEach(function(target) {
      if (!target) { return; }
      if (target.getAttribute('data-preview-distribution-observed') === '1') { return; }
      target.setAttribute('data-preview-distribution-observed', '1');
      observed.push(target);
    });

    observed.forEach(function(target) {
      var observer = new MutationObserver(function() {
        refreshDistributionView();
      });
      observer.observe(target, { childList: true, subtree: true });
    });
  }

  function boot() {
    refreshDistributionView();
    observeCalculator();
    var runButton = document.getElementById('final-stack-run');
    if (runButton && runButton.getAttribute('data-preview-distribution-run') !== '1') {
      runButton.setAttribute('data-preview-distribution-run', '1');
      runButton.addEventListener('click', function() {
        window.setTimeout(function() {
          refreshDistributionView();
        }, 0);
      });
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  }
  else {
    boot();
  }
})();
