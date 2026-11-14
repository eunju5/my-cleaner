/*
 * clean.js — 처방 파일(rules/*.json)을 읽어 그대로 실행하는 정리 엔진.
 * 처방(어떤 값을 무엇으로, 기준 숫자, 이유)은 여기에 적지 않는다. 모두 처방 파일에서 온다.
 * 이 파일은 브라우저 안에서만 돈다. 어디로도 데이터를 보내지 않는다.
 */
(function (global) {
  'use strict';

  var TYPES = ['replace', 'number', 'date', 'keep', 'flag', 'consistent', 'dedupe'];
  var ACTION_NAME = { replace: '바꿈', number: '바꿈', date: '바꿈', keep: '그대로', flag: '표시', consistent: '표시', dedupe: '지움' };

  function hasOwn(o, k) { return Object.prototype.hasOwnProperty.call(o, k); }

  function isBlank(v) {
    return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  }

  function pad(n, w) {
    var s = String(n);
    while (s.length < w) s = '0' + s;
    return s;
  }

  function isValidDate(y, m, d) {
    if (m < 1 || m > 12 || d < 1) return false;
    var dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
  }

  function formatDate(y, m, d, target) {
    return target.replace('YYYY', pad(y, 4)).replace('MM', pad(m, 2)).replace('DD', pad(d, 2));
  }

  function dateObjToString(v, target) {
    return formatDate(v.getFullYear(), v.getMonth() + 1, v.getDate(), target || 'YYYY-MM-DD');
  }

  // "YYYY년 M월 D일" 같은 형식 글을 정규식으로 바꾼다.
  function formatToRegex(fmt) {
    var re = '', order = [], i = 0;
    while (i < fmt.length) {
      var rest = fmt.slice(i);
      if (rest.indexOf('YYYY') === 0) { re += '(\\d{4})'; order.push('y'); i += 4; }
      else if (rest.indexOf('MM') === 0) { re += '(\\d{1,2})'; order.push('m'); i += 2; }
      else if (rest.indexOf('DD') === 0) { re += '(\\d{1,2})'; order.push('d'); i += 2; }
      else if (rest[0] === 'M') { re += '(\\d{1,2})'; order.push('m'); i += 1; }
      else if (rest[0] === 'D') { re += '(\\d{1,2})'; order.push('d'); i += 1; }
      else if (/\s/.test(rest[0])) { re += '\\s*'; i += 1; }
      else { re += rest[0].replace(/[.*+?^${}()|[\]\\\/-]/g, '\\$&'); i += 1; }
    }
    return { re: new RegExp('^' + re + '$'), order: order };
  }

  function testValue(cond, v) {
    switch (cond.op) {
      case 'blank': return isBlank(v);
      case '=': return v === cond.value;
      case '!=': return v !== cond.value;
      case '<': case '>': case '<=': case '>=':
        if (typeof v !== 'number' || typeof cond.value !== 'number') return false;
        if (cond.op === '<') return v < cond.value;
        if (cond.op === '>') return v > cond.value;
        if (cond.op === '<=') return v <= cond.value;
        return v >= cond.value;
      default:
        throw new Error('처방에 모르는 조건이 있습니다: ' + cond.op);
    }
  }

  function testRow(cond, values, idx) {
    if (cond.all) return cond.all.every(function (c) { return testRow(c, values, idx); });
    if (cond.any) return cond.any.some(function (c) { return testRow(c, values, idx); });
    if (!hasOwn(idx, cond.column)) return false;
    return testValue(cond, values[idx[cond.column]]);
  }

  function fill(text, col) {
    return String(text || '').split('{칸}').join(col || '');
  }

  // 처방 파일이 읽을 수 있는 모양인지 확인한다. 문제 목록을 돌려준다.
  function validateRules(rs) {
    var errs = [];
    if (!rs || typeof rs !== 'object') return ['처방 파일이 JSON 객체가 아닙니다.'];
    if (!Array.isArray(rs.rules)) return ['처방 파일에 rules 목록이 없습니다.'];
    rs.rules.forEach(function (r, i) {
      var name = r && r.id ? r.id : (i + 1) + '번째 처방';
      if (!r || typeof r !== 'object') { errs.push(name + ': 처방 모양이 아닙니다.'); return; }
      if (TYPES.indexOf(r.type) < 0) errs.push(name + ': 모르는 종류(type) «' + r.type + '»');
      if (!Array.isArray(r.columns)) errs.push(name + ': columns 목록이 없습니다.');
      if (!r.why) errs.push(name + ': why(이유)가 없습니다.');
      if (r.type === 'date' && !Array.isArray(r.read_formats)) errs.push(name + ': read_formats가 없습니다.');
      if (r.type === 'flag' && !r.each_column && !r.by_column && !r.when) errs.push(name + ': 조건(each_column · by_column · when)이 없습니다.');
      if (r.type === 'keep' && !r.when) errs.push(name + ': 조건(when)이 없습니다.');
      if (r.type === 'consistent' && (!r.group_by || !r.column)) errs.push(name + ': group_by와 column이 없습니다.');
    });
    return errs;
  }

  /*
   * headers: 칸 이름 배열
   * rows: [{ excelRow: 엑셀 행 번호, values: [...] }]
   * ruleSet: 처방 파일 내용
   */
  function clean(headers, rows, ruleSet) {
    var flagCol = ruleSet.flag_column || '_표시';
    var idx = {};
    headers.forEach(function (h, i) { idx[String(h).trim()] = i; });

    // 정리본을 다시 넣은 경우: 이전 표시 칸은 버리고 새로 만든다.
    var dropFlagIdx = hasOwn(idx, flagCol) ? idx[flagCol] : -1;
    if (dropFlagIdx >= 0) {
      headers = headers.filter(function (_, i) { return i !== dropFlagIdx; });
      rows = rows.map(function (r) {
        return { excelRow: r.excelRow, values: r.values.filter(function (_, i) { return i !== dropFlagIdx; }) };
      });
      idx = {};
      headers.forEach(function (h, i) { idx[String(h).trim()] = i; });
    }

    var missing = (ruleSet.columns || []).filter(function (c) { return !hasOwn(idx, c); });
    var extra = headers.filter(function (h) { return (ruleSet.columns || []).indexOf(String(h).trim()) < 0; });
    var idCol = hasOwn(idx, ruleSet.row_id) ? idx[ruleSet.row_id] : -1;

    var work = rows.map(function (r) {
      return { excelRow: r.excelRow, values: r.values.slice(), flags: [], changed: false };
    });
    var log = [];
    var perRule = [];

    function rowId(r) {
      return idCol >= 0 && !isBlank(r.values[idCol]) ? r.values[idCol] : '';
    }
    function entry(kind, r, col, before, after, rule) {
      log.push({
        kind: kind, excelRow: r.excelRow, id: rowId(r), column: col || '',
        before: before, after: after, ruleId: rule.id, why: rule.why || '', owner: rule.owner || ''
      });
    }

    (ruleSet.rules || []).forEach(function (rule) {
      var cols = (rule.columns || []).filter(function (c) { return hasOwn(idx, c); });
      var stat = {
        id: rule.id, title: rule.title || rule.id, type: rule.type, action: ACTION_NAME[rule.type] || rule.type,
        why: rule.why || '', owner: rule.owner || '',
        changed: 0, flaggedRows: 0, deleted: 0, matched: 0,
        skippedColumns: (rule.columns || []).filter(function (c) { return !hasOwn(idx, c); })
      };
      var flaggedSet = [];

      function change(r, col, nv) {
        var i = idx[col], before = r.values[i];
        if (before === nv) return;
        r.values[i] = nv;
        r.changed = true;
        stat.changed += 1;
        entry('바꿈', r, col, before, nv, rule);
      }
      function flag(r, col, text) {
        r.flags.push(text);
        if (flaggedSet.indexOf(r) < 0) flaggedSet.push(r);
        entry('표시', r, col, col ? r.values[idx[col]] : '', text, rule);
      }

      if (rule.type === 'replace') {
        work.forEach(function (r) {
          cols.forEach(function (c) {
            var v = r.values[idx[c]];
            if (typeof v !== 'string') return;
            var nv = rule.trim ? v.trim() : v;
            if (rule.map && hasOwn(rule.map, nv)) nv = rule.map[nv];
            change(r, c, nv);
          });
        });
      } else if (rule.type === 'number') {
        work.forEach(function (r) {
          cols.forEach(function (c) {
            var v = r.values[idx[c]];
            if (typeof v !== 'string' || isBlank(v)) return;
            var s = v.trim(), nv;
            if (rule.map && hasOwn(rule.map, s)) {
              nv = rule.map[s];
            } else {
              (rule.remove || []).forEach(function (t) { s = s.split(t).join(''); });
              s = s.trim();
              if (/^-?\d+(\.\d+)?$/.test(s)) nv = Number(s);
            }
            if (nv === undefined) {
              if (rule.fail_flag) flag(r, c, fill(rule.fail_flag, c));
              return;
            }
            change(r, c, nv);
          });
        });
      } else if (rule.type === 'date') {
        var target = rule.target || 'YYYY-MM-DD';
        var formats = (rule.read_formats || []).map(formatToRegex);
        var patterns = (rule.flag_patterns || []).map(function (p) { return { re: new RegExp(p.match), flag: p.flag }; });
        work.forEach(function (r) {
          cols.forEach(function (c) {
            var v = r.values[idx[c]];
            if (isBlank(v)) return;
            if (v instanceof Date) { change(r, c, dateObjToString(v, target)); return; }
            if (typeof v !== 'string') {
              if (rule.fail_flag) flag(r, c, fill(rule.fail_flag, c));
              return;
            }
            var s = v.trim();
            for (var p = 0; p < patterns.length; p++) {
              if (patterns[p].re.test(s)) { flag(r, c, fill(patterns[p].flag, c)); return; }
            }
            for (var f = 0; f < formats.length; f++) {
              var m = formats[f].re.exec(s);
              if (!m) continue;
              var parts = {};
              formats[f].order.forEach(function (k, j) { parts[k] = Number(m[j + 1]); });
              if (isValidDate(parts.y, parts.m, parts.d)) {
                change(r, c, formatDate(parts.y, parts.m, parts.d, target));
                return;
              }
            }
            if (rule.fail_flag) flag(r, c, fill(rule.fail_flag, c));
          });
        });
      } else if (rule.type === 'keep') {
        work.forEach(function (r) { if (testRow(rule.when, r.values, idx)) stat.matched += 1; });
      } else if (rule.type === 'flag') {
        work.forEach(function (r) {
          if (rule.each_column) {
            cols.forEach(function (c) {
              if (testValue(rule.each_column, r.values[idx[c]])) flag(r, c, fill(rule.flag, c));
            });
          }
          if (rule.by_column) {
            Object.keys(rule.by_column).forEach(function (c) {
              if (hasOwn(idx, c) && testValue(rule.by_column[c], r.values[idx[c]])) flag(r, c, fill(rule.flag, c));
            });
          }
          if (rule.when && testRow(rule.when, r.values, idx)) flag(r, '', fill(rule.flag, ''));
        });
      } else if (rule.type === 'consistent') {
        // group_by 칸이 같은 줄끼리 column 칸 값이 하나인지 본다. 빈칸은 세지 않는다.
        // 둘 이상이면 어느 쪽이 맞는지 모르므로 그 무리의 줄을 모두 표시한다.
        if (hasOwn(idx, rule.group_by) && hasOwn(idx, rule.column)) {
          var gi = idx[rule.group_by], ci = idx[rule.column], groups = {};
          work.forEach(function (r) {
            var g = r.values[gi], v = r.values[ci];
            if (isBlank(g) || isBlank(v)) return;
            var gk = JSON.stringify(g);
            if (!hasOwn(groups, gk)) groups[gk] = {};
            groups[gk][JSON.stringify(v)] = true;
          });
          work.forEach(function (r) {
            var g = r.values[gi];
            if (isBlank(g) || isBlank(r.values[ci])) return;
            if (Object.keys(groups[JSON.stringify(g)]).length > 1) flag(r, rule.column, fill(rule.flag, rule.column));
          });
        }
      } else if (rule.type === 'dedupe') {
        var seen = {};
        var kept = [];
        work.forEach(function (r) {
          var key = JSON.stringify(cols.map(function (c) {
            var v = r.values[idx[c]];
            return v instanceof Date ? ['D', v.getTime()] : v;
          }));
          if (hasOwn(seen, key)) {
            stat.deleted += 1;
            entry('지움', r, '', '', seen[key].excelRow + '행과 같음', rule);
          } else {
            seen[key] = r;
            kept.push(r);
          }
        });
        work = kept;
      }
      stat.flaggedRows = flaggedSet.length;
      perRule.push(stat);
    });

    log.sort(function (a, b) { return a.excelRow - b.excelRow; });

    var outRows = work.map(function (r) {
      return { excelRow: r.excelRow, values: r.values, flags: r.flags, changed: r.changed };
    });
    var deleted = log.filter(function (e) { return e.kind === '지움'; });
    return {
      headers: headers.slice(),
      flagColumn: flagCol,
      rows: outRows,
      log: log,
      perRule: perRule,
      missingColumns: missing,
      extraColumns: extra,
      droppedOldFlagColumn: dropFlagIdx >= 0,
      stats: {
        before: rows.length,
        after: outRows.length,
        deleted: deleted.length,
        changedCells: log.filter(function (e) { return e.kind === '바꿈'; }).length,
        changedRows: outRows.filter(function (r) { return r.changed; }).length,
        flaggedRows: outRows.filter(function (r) { return r.flags.length > 0; }).length
      }
    };
  }

  // 정리본을 같은 처방으로 한 번 더 돌려서 결과가 그대로인지 본다.
  function checkTwice(result, ruleSet) {
    var again = clean(result.headers, result.rows.map(function (r) {
      return { excelRow: r.excelRow, values: r.values.slice() };
    }), ruleSet);
    function snap(res) {
      return JSON.stringify(res.rows.map(function (r) { return [r.values, r.flags]; }));
    }
    return {
      same: snap(again) === snap(result) && again.stats.changedCells === 0 && again.stats.deleted === 0,
      changedCells: again.stats.changedCells,
      deleted: again.stats.deleted,
      rowsBefore: result.rows.length,
      rowsAfter: again.rows.length
    };
  }

  global.Cleaner = {
    clean: clean,
    checkTwice: checkTwice,
    validateRules: validateRules,
    isBlank: isBlank,
    dateObjToString: dateObjToString,
    ACTION_NAME: ACTION_NAME
  };
})(window);
