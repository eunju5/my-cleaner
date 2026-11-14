/*
 * app.js — 화면. 파일을 읽고(SheetJS), 처방을 불러오고, clean.js로 정리하고, 결과를 보여 주고 내려받게 한다.
 * 넣은 파일과 정답 파일은 브라우저 안에서만 읽는다. 서버로 보내는 코드는 없다.
 * 사이트에 요청하는 것은 rules/목록.json과, 목록에 있는 이름의 rules/<이름>.json뿐이다.
 */
(function () {
  'use strict';

  var C = window.Cleaner;
  var state = { base: '', sheet: null, ruleSet: null, ruleSource: '', result: null, logKind: '' };

  function $(id) { return document.getElementById(id); }
  function show(id, on) { $(id).hidden = !on; }
  function text(id, t) { $(id).textContent = t; }
  function msg(id, t) { var e = $(id); e.textContent = t || ''; e.hidden = !t; }
  function n(x) { return Number(x).toLocaleString('ko-KR'); }

  function make(tag, cls, content) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (content !== undefined && content !== null) {
      if (content instanceof Node) e.appendChild(content); else e.textContent = String(content);
    }
    return e;
  }

  // 화면에 보일 값
  function shown(v) {
    if (C.isBlank(v)) return make('span', 'blank', v === null || v === undefined ? '(빈칸)' : '(공백)');
    if (v instanceof Date) return document.createTextNode(C.dateObjToString(v));
    return document.createTextNode(String(v));
  }
  // 파일에 쓸 값
  function outVal(v) {
    if (v instanceof Date) return C.dateObjToString(v);
    return v === undefined ? null : v;
  }

  function fillTable(table, head, rows, numericCols) {
    table.textContent = '';
    var thead = make('thead'), tr = make('tr');
    head.forEach(function (h) { tr.appendChild(make('th', null, h)); });
    thead.appendChild(tr);
    table.appendChild(thead);
    var tbody = make('tbody');
    rows.forEach(function (cells) {
      var r = make('tr');
      cells.forEach(function (c, i) {
        var td = make('td', numericCols && numericCols.indexOf(i) >= 0 ? 'n' : null);
        if (c instanceof Node) td.appendChild(c); else td.textContent = c === null || c === undefined ? '' : String(c);
        r.appendChild(td);
      });
      tbody.appendChild(r);
    });
    table.appendChild(tbody);
  }

  function stat(label, value, sub, cls) {
    var d = make('div', 'stat' + (cls ? ' ' + cls : ''));
    d.appendChild(make('div', 'label', label));
    d.appendChild(make('div', 'value', value));
    if (sub) d.appendChild(make('div', 'sub', sub));
    return d;
  }

  function tag(kind) { return make('span', 'tag ' + kind, kind); }

  // ---------- 엑셀 읽기 ----------
  function readSheet(buf) {
    var wb = XLSX.read(buf, { type: 'array', cellDates: true });
    var name = wb.SheetNames[0];
    var ws = wb.Sheets[name];
    if (!ws || !ws['!ref']) throw new Error('첫 시트가 비어 있습니다.');
    var range = XLSX.utils.decode_range(ws['!ref']);
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
    var h = 0;
    while (h < aoa.length && (aoa[h] || []).every(C.isBlank)) h++;
    if (h >= aoa.length) throw new Error('첫 시트가 비어 있습니다.');
    var headers = (aoa[h] || []).map(function (x) { return x === null || x === undefined ? '' : String(x).trim(); });
    while (headers.length && headers[headers.length - 1] === '') headers.pop();
    var rows = [];
    for (var i = h + 1; i < aoa.length; i++) {
      var raw = aoa[i] || [];
      var values = [];
      for (var j = 0; j < headers.length; j++) values.push(raw[j] === undefined ? null : raw[j]);
      if (values.every(C.isBlank)) continue;
      rows.push({ excelRow: range.s.r + i + 1, values: values });
    }
    return { sheetNames: wb.SheetNames, sheetName: name, headers: headers, rows: rows };
  }

  function readFile(file) {
    return file.arrayBuffer().then(readSheet);
  }

  function baseName(name) { return name.replace(/\.[^.]+$/, ''); }

  // ---------- 처방 찾기 ----------
  function findRules(base) {
    if (location.protocol === 'file:') return Promise.resolve({ ok: false, reason: 'file' });
    return fetch('rules/목록.json', { cache: 'no-store' })
      .then(function (res) { if (!res.ok) throw new Error('list'); return res.json(); })
      .then(function (list) {
        if (!list || !Array.isArray(list.files) || list.files.indexOf(base) < 0) return { ok: false, reason: 'none' };
        var path = 'rules/' + encodeURIComponent(base) + '.json';
        return fetch(path, { cache: 'no-store' })
          .then(function (res) { if (!res.ok) throw new Error('rule'); return res.json(); })
          .then(function (rs) { return { ok: true, ruleSet: rs, source: 'rules/' + base + '.json' }; });
      })
      .catch(function () { return { ok: false, reason: 'error' }; });
  }

  function useRules(rs, source) {
    var errs = C.validateRules(rs);
    if (errs.length) {
      msg('rules-error', '처방 파일을 읽을 수 없습니다. ' + errs.join(' / '));
      show('rules-found', false);
      return;
    }
    msg('rules-error', '');
    state.ruleSet = rs;
    state.ruleSource = source;
    renderRules();
  }

  function renderRules() {
    var rs = state.ruleSet, sh = state.sheet;
    show('rules-none', false);
    show('rules-found', true);
    text('rules-source', '처방 파일: ' + state.ruleSource + ' · 처방 ' + rs.rules.length + '개' +
      (rs.file && rs.file !== state.base ? ' · 이 처방은 «' + rs.file + '» 파일용입니다' : ''));
    var cols = rs.columns || [];
    var missing = cols.filter(function (c) { return sh.headers.indexOf(c) < 0; });
    var extra = sh.headers.filter(function (h) { return cols.indexOf(h) < 0 && h !== rs.flag_column; });
    var warn = [];
    if (missing.length) warn.push('처방에 있지만 파일에 없는 칸: ' + missing.join(', ') + ' (이 칸 처방은 건너뜀)');
    if (extra.length) warn.push('처방에 없는 칸: ' + extra.join(', ') + ' (그대로 둠)');
    if (sh.headers.indexOf(rs.flag_column) >= 0) warn.push('파일에 이미 «' + rs.flag_column + '» 칸이 있어요. 정리할 때 새로 만듭니다.');
    msg('rules-colwarn', warn.join(' · '));
    fillTable($('rules-table'), ['순서', '처방', '하는 일', '이유', '정한 사람'],
      rs.rules.map(function (r, i) {
        return [i + 1, r.title || r.id, tag(C.ACTION_NAME[r.type] || r.type), make('span', 'why', r.why), r.owner || ''];
      }), [0]);
    $('run').focus({ preventScroll: true });
  }

  function noRules(reason) {
    state.ruleSet = null;
    show('rules-found', false);
    show('rules-none', true);
    var why = {
      file: '이 페이지를 파일로 직접 열어서 처방을 자동으로 찾을 수 없어요. 링크로 열거나, 처방 파일을 직접 골라 주세요.',
      none: '«' + state.base + '»에 맞는 처방 파일이 사이트에 없어요. 처방 파일(.json)이 있으면 골라 주세요.',
      error: '처방 목록을 불러오지 못했어요. 처방 파일(.json)이 있으면 골라 주세요.'
    };
    text('rules-none-why', why[reason] || why.none);
  }

  // ---------- 파일 넣기 ----------
  function resetAfterFile() {
    ['step-info', 'step-rules', 'rules-found', 'rules-none', 'step-result', 'step-answer', 'answer-result'].forEach(function (id) { show(id, false); });
    msg('file-error', ''); msg('rules-error', ''); msg('answer-error', '');
    state.sheet = null;
    state.ruleSet = null;
    state.result = null;
  }

  function handleFile(file) {
    if (!file) return;
    resetAfterFile();
    state.base = baseName(file.name);
    readFile(file).then(function (sh) {
      state.sheet = sh;
      text('info-name', file.name);
      text('info-sheet', sh.sheetName);
      text('info-rows', n(sh.rows.length) + '줄');
      text('info-cols', n(sh.headers.length) + '칸');
      text('info-colnames', '칸: ' + sh.headers.join(' · '));
      msg('info-warn', sh.sheetNames.length > 1
        ? '시트가 ' + sh.sheetNames.length + '개 있어요. 첫 시트 «' + sh.sheetName + '»만 정리합니다.' : '');
      show('step-info', true);
      show('step-rules', true);
      return findRules(state.base).then(function (found) {
        if (found.ok) useRules(found.ruleSet, found.source); else noRules(found.reason);
      });
    }).catch(function (e) {
      msg('file-error', '파일을 읽을 수 없습니다. ' + (e && e.message ? e.message : ''));
    });
  }

  // ---------- 정리 ----------
  function run() {
    var sh = state.sheet, rs = state.ruleSet;
    if (!sh || !rs) return;
    var res;
    try {
      res = C.clean(sh.headers, sh.rows, rs);
    } catch (e) {
      msg('rules-error', '정리하다 멈췄습니다. ' + e.message);
      return;
    }
    state.result = res;
    var twice = C.checkTwice(res, rs);
    renderResult(res, twice);
    show('step-result', true);
    show('step-answer', true);
    show('answer-result', false);
    msg('answer-error', '');
    $('step-result').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function renderResult(res, twice) {
    var s = res.stats, box = $('stats');
    box.textContent = '';
    box.appendChild(stat('행 (정리 전 → 후)', n(s.before) + ' → ' + n(s.after)));
    box.appendChild(stat('지운 행', n(s.deleted)));
    box.appendChild(stat('고친 값', n(s.changedCells) + '칸', n(s.changedRows) + '행에서'));
    box.appendChild(stat('표시한 행', n(s.flaggedRows), '«' + res.flagColumn + '» 칸에 이유'));
    box.appendChild(twice.same
      ? stat('두 번 돌려도 같은 결과', '같음 ✓', '두 번째에 바뀐 값 0 · 지운 행 0', 'ok')
      : stat('두 번 돌려도 같은 결과', '다름 ✗', '두 번째에 바뀐 값 ' + twice.changedCells + ' · 지운 행 ' + twice.deleted, 'bad'));

    fillTable($('per-rule'), ['순서', '처방', '하는 일', '바꾼 칸', '표시한 행', '지운 행', '이유'],
      res.perRule.map(function (p, i) {
        var title = p.title + (p.skippedColumns.length ? ' (파일에 없는 칸 건너뜀: ' + p.skippedColumns.join(', ') + ')' : '');
        var kind = p.action;
        if (p.type === 'keep') kind = '그대로 (' + n(p.matched) + '행 해당)';
        return [i + 1, title, kind, p.changed || '', p.flaggedRows || '', p.deleted || '', make('span', 'why', p.why)];
      }), [0, 3, 4, 5]);

    state.logKind = '';
    Array.prototype.forEach.call(document.querySelectorAll('#log-filter .chip'), function (b) {
      b.classList.toggle('on', b.getAttribute('data-kind') === '');
    });
    renderLog();
  }

  var LOG_LIMIT = 300;
  function renderLog() {
    var res = state.result, idName = state.ruleSet.row_id || '행 ID';
    var items = res.log.filter(function (e) { return !state.logKind || e.kind === state.logKind; });
    text('log-note', '기록 ' + n(items.length) + '건' + (items.length > LOG_LIMIT ? ' 중 앞 ' + LOG_LIMIT + '건만 보여요. 전체는 내려받아서 보세요.' : ''));
    fillTable($('log-table'), ['종류', '원래 행', idName, '칸', '원래 값', '바뀐 값 / 표시 내용', '처방'],
      items.slice(0, LOG_LIMIT).map(function (e) {
        return [tag(e.kind), e.excelRow, e.id, e.column, e.kind === '지움' ? '' : shown(e.before),
          e.kind === '바꿈' ? shown(e.after) : e.after, e.ruleId];
      }), [1]);
  }

  // ---------- 내려받기 ----------
  function sheetName(name) {
    var s = String(name || '정리본').replace(/[\[\]:*?\/\\]/g, ' ').slice(0, 31);
    return s || '정리본';
  }

  function downloadClean() {
    var res = state.result;
    var aoa = [res.headers.concat([res.flagColumn])].concat(res.rows.map(function (r) {
      return r.values.map(outVal).concat([r.flags.length ? r.flags.join(' / ') : null]);
    }));
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), sheetName(state.sheet.sheetName));
    XLSX.writeFile(wb, state.base + '_정리본.xlsx');
  }

  function downloadLog() {
    var res = state.result, rs = state.ruleSet;
    var head = ['종류', '원래 행', rs.row_id || '행 ID', '칸', '원래 값', '바뀐 값 / 표시 내용', '처방', '이유', '정한 사람'];
    var aoa = [head].concat(res.log.map(function (e) {
      return [e.kind, e.excelRow, outVal(e.id), e.column, e.kind === '지움' ? null : outVal(e.before),
        outVal(e.after), e.ruleId, e.why, e.owner];
    }));
    var s = res.stats;
    var summary = [
      ['항목', '값'],
      ['파일', state.base],
      ['처방 파일', state.ruleSource],
      ['정리 전 행', s.before],
      ['정리 후 행', s.after],
      ['지운 행', s.deleted],
      ['고친 값(칸)', s.changedCells],
      ['고친 행', s.changedRows],
      ['표시한 행', s.flaggedRows],
      [],
      ['순서', '처방', '하는 일', '바꾼 칸', '표시한 행', '지운 행', '이유', '정한 사람']
    ].concat(res.perRule.map(function (p, i) {
      return [i + 1, p.title, p.type === 'keep' ? '그대로 (' + p.matched + '행 해당)' : p.action,
        p.changed, p.flaggedRows, p.deleted, p.why, p.owner];
    }));
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), '바꾼 기록');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), '요약');
    XLSX.writeFile(wb, state.base + '_바꾼기록.xlsx');
  }

  // ---------- 정답과 대조 ----------
  function norm(v) {
    if (C.isBlank(v)) return '';
    if (v instanceof Date) return C.dateObjToString(v);
    return String(v).trim();
  }

  function compare(ans) {
    var res = state.result, rs = state.ruleSet, flagCol = res.flagColumn;
    var ourH = res.headers, ansH = ans.headers;
    var common = ourH.filter(function (h) { return h && h.charAt(0) !== '_' && ansH.indexOf(h) >= 0; });
    var onlyOurs = ourH.filter(function (h) { return ansH.indexOf(h) < 0; });
    var onlyAns = ansH.filter(function (h) { return h && ourH.indexOf(h) < 0 && h !== flagCol; });

    var key = rs.row_id;
    function keyMap(rows, headers) {
      var i = headers.indexOf(key);
      if (i < 0) return null;
      var m = {};
      for (var r = 0; r < rows.length; r++) {
        var k = norm(rows[r].values[i]);
        if (k === '' || Object.prototype.hasOwnProperty.call(m, k)) return null;
        m[k] = rows[r];
      }
      return m;
    }
    var ourMap = keyMap(res.rows, ourH), ansMap = keyMap(ans.rows, ansH);
    var byKey = !!(ourMap && ansMap);
    var pairs = [], onlyAnsRows = [], onlyOurRows = [];
    if (byKey) {
      Object.keys(ansMap).forEach(function (k) {
        if (ourMap[k]) pairs.push({ label: k, ours: ourMap[k], ans: ansMap[k] }); else onlyAnsRows.push(k);
      });
      Object.keys(ourMap).forEach(function (k) { if (!ansMap[k]) onlyOurRows.push(k); });
    } else {
      var len = Math.max(res.rows.length, ans.rows.length);
      for (var i = 0; i < len; i++) {
        var o = res.rows[i], a = ans.rows[i];
        if (o && a) pairs.push({ label: (i + 1) + '번째 줄', ours: o, ans: a });
        else if (a) onlyAnsRows.push(a.excelRow + '행');
        else onlyOurRows.push(o.excelRow + '행');
      }
    }

    var diffs = [], diffRows = 0, sameRows = 0;
    pairs.forEach(function (p) {
      var bad = false;
      common.forEach(function (c) {
        var ov = norm(p.ours.values[ourH.indexOf(c)]), av = norm(p.ans.values[ansH.indexOf(c)]);
        if (ov !== av) { bad = true; diffs.push([p.label, c, ov, av]); }
      });
      if (bad) diffRows += 1; else sameRows += 1;
    });

    var flagNote = '';
    var ansFlagIdx = ansH.indexOf(flagCol);
    if (ansFlagIdx >= 0) {
      var flagDiff = pairs.filter(function (p) {
        return (p.ours.flags.length > 0) !== (norm(p.ans.values[ansFlagIdx]) !== '');
      }).map(function (p) { return p.label; });
      flagNote = '«' + flagCol + '» 칸이 비었는지 여부가 다른 행: ' + flagDiff.length +
        (flagDiff.length ? ' (' + flagDiff.slice(0, 10).join(', ') + (flagDiff.length > 10 ? ' …' : '') + ')' : '') + '. ';
    }

    var box = $('answer-stats');
    box.textContent = '';
    box.appendChild(stat('정답 행', n(ans.rows.length), '정리본 ' + n(res.rows.length) + '행'));
    box.appendChild(stat('일치한 행', n(sameRows), '짝지은 ' + n(pairs.length) + '행 중', diffRows === 0 && !onlyAnsRows.length && !onlyOurRows.length ? 'ok' : null));
    box.appendChild(stat('어긋난 칸', n(diffs.length), n(diffRows) + '행에서', diffs.length ? 'bad' : 'ok'));
    box.appendChild(stat('정답에만 있는 행', n(onlyAnsRows.length)));
    box.appendChild(stat('정리본에만 있는 행', n(onlyOurRows.length)));

    var notes = [];
    notes.push(byKey ? '«' + key + '»로 짝을 지었어요.' : '«' + key + '»가 없거나 겹쳐서 줄 순서대로 짝을 지었어요.');
    notes.push('비교한 칸: ' + common.join(', ') + '.');
    if (onlyOurs.length) notes.push('정답에 없는 칸: ' + onlyOurs.join(', ') + '.');
    if (onlyAns.length) notes.push('정리본에 없는 칸: ' + onlyAns.join(', ') + '.');
    notes.push('«_»로 시작하는 칸(표시 칸)은 값 비교에서 뺐어요.');
    if (flagNote) notes.push(flagNote);
    if (onlyAnsRows.length) notes.push('정답에만 있는 행: ' + onlyAnsRows.slice(0, 20).join(', ') + (onlyAnsRows.length > 20 ? ' …' : '') + '.');
    if (onlyOurRows.length) notes.push('정리본에만 있는 행: ' + onlyOurRows.slice(0, 20).join(', ') + (onlyOurRows.length > 20 ? ' …' : '') + '.');
    if (diffs.length > LOG_LIMIT) notes.push('어긋난 칸은 앞 ' + LOG_LIMIT + '건만 보여요.');
    text('answer-note', notes.join(' '));

    fillTable($('answer-table'), [byKey ? key : '줄', '칸', '정리본 값', '정답 값'],
      diffs.slice(0, LOG_LIMIT).map(function (d) {
        return [d[0], d[1], d[2] === '' ? make('span', 'blank', '(빈칸)') : d[2], d[3] === '' ? make('span', 'blank', '(빈칸)') : d[3]];
      }));
    show('answer-result', true);
  }

  // ---------- 연결 ----------
  var drop = $('drop');
  ['dragenter', 'dragover'].forEach(function (ev) {
    drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    drop.addEventListener(ev, function () { drop.classList.remove('over'); });
  });
  drop.addEventListener('drop', function (e) {
    e.preventDefault();
    if (e.dataTransfer && e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  // 상자 밖에 떨어뜨려도 브라우저가 파일을 열어 버리지 않게
  ['dragover', 'drop'].forEach(function (ev) {
    window.addEventListener(ev, function (e) { if (!drop.contains(e.target)) e.preventDefault(); });
  });

  $('file-input').addEventListener('change', function (e) {
    handleFile(e.target.files[0]);
    e.target.value = '';
  });

  $('rules-input').addEventListener('change', function (e) {
    var f = e.target.files[0];
    e.target.value = '';
    if (!f) return;
    f.text().then(function (t) {
      var rs;
      try { rs = JSON.parse(t); } catch (err) { msg('rules-error', '처방 파일이 JSON 모양이 아닙니다. ' + err.message); return; }
      show('step-result', false); show('step-answer', false);
      useRules(rs, f.name + ' (직접 고른 파일)');
    });
  });

  $('run').addEventListener('click', run);
  $('dl-clean').addEventListener('click', downloadClean);
  $('dl-log').addEventListener('click', downloadLog);

  $('log-filter').addEventListener('click', function (e) {
    var b = e.target.closest('.chip');
    if (!b) return;
    state.logKind = b.getAttribute('data-kind');
    Array.prototype.forEach.call(this.querySelectorAll('.chip'), function (x) { x.classList.toggle('on', x === b); });
    renderLog();
  });

  $('answer-input').addEventListener('change', function (e) {
    var f = e.target.files[0];
    e.target.value = '';
    if (!f || !state.result) return;
    msg('answer-error', '');
    readFile(f).then(compare).catch(function (err) {
      show('answer-result', false);
      msg('answer-error', '정답 파일을 읽을 수 없습니다. ' + (err && err.message ? err.message : ''));
    });
  });
})();
