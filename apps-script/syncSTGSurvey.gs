const STG_CONFIG = {
  surveyId: '74a431a9-5960-41cc-af8f-30ed3270e5c2',
  apiUrl: 'https://api.dooblo.net/newapi',
  dataSheet: 'Raw_Data',
  logSheet: 'Sync_Log',
  dashSheet: 'Dash_Data',   // compact sheet the dashboard reads (1 row per interview)
  qcSheet: 'QC_Data',       // compact sheet the QC page reads (1 row per interview)
  chunkSize: 45000
};

// Fast sync: writes Dash_Data (dashboard) and QC_Data (QC page). Run every 5 minutes.
function syncSTGSurvey() { stgRunSync_(true, false); }

// Raw XML sync: slow (many large cells). Writes only Raw_Data for the QC page.
// Run it manually or on a trigger of every 1 hour or longer.
function syncSTGRaw() { stgRunSync_(false, true); }

function stgRunSync_(doDash, doRaw) {
  const t0 = Date.now();
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('STG_API_KEY');
  const password = props.getProperty('STG_PASSWORD');

  if (!apiKey || !password) {
    throw new Error('Missing API credentials.');
  }

  const auth = 'Basic ' +
    Utilities.base64Encode(apiKey + ':' + password);

  // 1. Get interview IDs
  const idsUrl = STG_CONFIG.apiUrl +
    '/SurveyInterviewIDs?surveyIDs=' +
    encodeURIComponent(STG_CONFIG.surveyId);

  const idsResponse = UrlFetchApp.fetch(idsUrl, {
    method: 'get',
    headers: {
      Authorization: auth,
      Accept: 'application/json'
    },
    muteHttpExceptions: true
  });

  if (idsResponse.getResponseCode() !== 200) {
    throw new Error('Interview ID error: HTTP ' +
      idsResponse.getResponseCode());
  }

  const interviewIds = JSON.parse(idsResponse.getContentText());

  if (!Array.isArray(interviewIds)) {
    throw new Error('Unexpected Interview IDs response.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let dataSheet = ss.getSheetByName(STG_CONFIG.dataSheet);
  if (!dataSheet) {
    dataSheet = ss.insertSheet(STG_CONFIG.dataSheet);
  }

  const syncTime = new Date();
  let tFetch = 0, tParse = 0;
  const rows = [];
  const dashRows = [];
  const qcRows = [];
  const badSubjects = [];

  // 2. Download in batches of 99
  //    (Sheet is NOT cleared yet, so the dashboard keeps reading the previous
  //     data while downloads run.)
  for (let i = 0; i < interviewIds.length; i += 99) {
    const batch = interviewIds.slice(i, i + 99);

    const dataUrl = STG_CONFIG.apiUrl +
      '/SurveyInterviewData?subjectIDs=' +
      encodeURIComponent(batch.join(',')) +
      '&surveyID=' + encodeURIComponent(STG_CONFIG.surveyId) +
      '&onlyHeaders=false&includeNulls=false';

    const tf = Date.now();
    const response = UrlFetchApp.fetch(dataUrl, {
      method: 'get',
      headers: {
        Authorization: auth,
        Accept: 'text/xml'
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('Data download failed: HTTP ' +
        response.getResponseCode());
    }

    const xml = response.getContentText().replace(/^\uFEFF/, '');
    tFetch += Date.now() - tf;
    const tp = Date.now();

    // 2b. Compact rows for the dashboard: parse each <Subject> on its own
    if (doDash) stgSplitSubjects_(xml).forEach(function (seg) {
      try {
        const root = stgParseXml_(seg);
        const r = stgExtractSubject_(root);
        dashRows.push([r.id, r.date, r.done ? 'True' : 'False', r.st, JSON.stringify(r)]);
        const qj = stgQcJson_(stgQcSubject_(root));
        if (qj.length > 49000) throw new Error('QC record too long: ' + qj.length);
        qcRows.push([r.id, qj]);
      } catch (e) {
        const id = /<SubjectNum[^>]*>(\d+)</.exec(seg);
        badSubjects.push(id ? id[1] : '?');
      }
    });

    tParse += Date.now() - tp;

    // 3. Split XML into chunks below 50,000 characters
    const batchNumber = Math.floor(i / 99) + 1;
    const chunkSize = STG_CONFIG.chunkSize;

    for (let j = 0; doRaw && j < xml.length; j += chunkSize) {
      rows.push([
        syncTime,
        batchNumber,
        Math.floor(j / chunkSize) + 1,
        // FIX: leading apostrophe forces Sheets to store the chunk as plain text.
        // Without it, a chunk that happens to start with "=", "+", "-" or "@"
        // is parsed as a formula and becomes #ERROR!, which loses ~45,000
        // characters of XML. The apostrophe itself is not saved in the cell.
        "'" + xml.substring(j, j + chunkSize)
      ]);
    }

    Utilities.sleep(600);
  }

  // 4a. Write compact dashboard sheet
  if (doDash) {
    let dashSheet = ss.getSheetByName(STG_CONFIG.dashSheet);
    if (!dashSheet) {
      dashSheet = ss.insertSheet(STG_CONFIG.dashSheet);
    }
    dashSheet.clearContents();
    dashSheet.getRange(1, 1, dashRows.length + 1, 5).setValues(
      [['SubjectNum', 'Date', 'Completed', 'Status', 'JSON']].concat(dashRows));
    dashSheet.setFrozenRows(1);

    let qcSheet = ss.getSheetByName(STG_CONFIG.qcSheet);
    if (!qcSheet) {
      qcSheet = ss.insertSheet(STG_CONFIG.qcSheet);
    }
    qcSheet.clearContents();
    qcSheet.getRange(1, 1, qcRows.length + 1, 2).setValues(
      [['SubjectNum', 'JSON']].concat(qcRows));
    qcSheet.setFrozenRows(1);
  }

  // 4b. Write raw XML chunks (optional backup)
  //    FIX: clear only now, right before writing, so the sheet is empty for
  //    as short a time as possible.
  if (doRaw) {
  dataSheet.clearContents();
  dataSheet.getRange(1, 1, 1, 4).setValues([[
    'Sync Time',
    'Batch',
    'Chunk',
    'XML Data'
  ]]);

  if (rows.length > 0) {
    dataSheet.getRange(2, 1, rows.length, 4).setValues(rows);
    dataSheet.setFrozenRows(1);
    dataSheet.setColumnWidth(1, 160);
    dataSheet.setColumnWidth(2, 80);
    dataSheet.setColumnWidth(3, 80);
    dataSheet.setColumnWidth(4, 500);
  }
  }

  // FIX: check that no chunk was turned into an error value.
  const written = (doRaw && rows.length > 0)
    ? dataSheet.getRange(2, 4, rows.length, 1).getValues()
    : [];
  const badChunks = [];
  written.forEach(function (r, k) {
    if (String(r[0]).length !== rows[k][3].length - 1) {
      badChunks.push(rows[k][1] + '/' + rows[k][2]);
    }
  });

  // 5. Record sync log
  let logSheet = ss.getSheetByName(STG_CONFIG.logSheet);

  if (!logSheet) {
    logSheet = ss.insertSheet(STG_CONFIG.logSheet);
    logSheet.appendRow([
      'Sync Time',
      'Survey ID',
      'Interview Count',
      'XML Chunks',
      'Status'
    ]);
  }

  logSheet.appendRow([
    syncTime,
    STG_CONFIG.surveyId,
    interviewIds.length,
    rows.length,
    (doDash ? '[dash] ' : '[raw] ') + Math.round((Date.now() - t0) / 1000) + 's (tải ' +
    Math.round(tFetch / 1000) + 's, xử lý ' + Math.round(tParse / 1000) + 's) ' +
    ((badChunks.length || badSubjects.length)
      ? 'WARNING:' + (badChunks.length ? ' bad chunks ' + badChunks.join(', ') : '') +
        (badSubjects.length ? ' bad subjects ' + badSubjects.join(', ') : '')
      : 'SUCCESS')
  ]);

  SpreadsheetApp.flush();

  Logger.log('Total interviews: ' + interviewIds.length);
  Logger.log('XML chunks saved: ' + rows.length);
  if (badChunks.length) Logger.log('Bad chunks: ' + badChunks.join(', '));
  Logger.log('Dashboard rows: ' + dashRows.length);
  Logger.log('Download ' + Math.round(tFetch / 1000) + 's, parse ' + Math.round(tParse / 1000) +
    's, total ' + Math.round((Date.now() - t0) / 1000) + 's');
  Logger.log('Sync completed.');
}

/* ---- Fast XML reader (replaces XmlService.parse, which is too slow for this survey) ----
   Returns an element with the same methods the extractors use:
   getName, getChildren([name]), getChild(name), getChildText(name), getAttribute(name), getText, getValue. */
function stgDecode_(s) {
  return s.indexOf('&') < 0 ? s : s.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, function (m, e) {
    if (e === 'lt') return '<'; if (e === 'gt') return '>'; if (e === 'amp') return '&';
    if (e === 'quot') return '"'; if (e === 'apos') return "'";
    return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
  });
}
function StgEl_(name, attrs) { this.n = name; this.a = attrs; this.c = []; this.t = ''; }
StgEl_.prototype.getName = function () { return this.n; };
StgEl_.prototype.getChildren = function (name) { return name ? this.c.filter(function (x) { return x.n === name; }) : this.c.slice(); };
StgEl_.prototype.getChild = function (name) { for (var i = 0; i < this.c.length; i++) if (this.c[i].n === name) return this.c[i]; return null; };
StgEl_.prototype.getChildText = function (name) { var c = this.getChild(name); return c ? c.t : null; };
StgEl_.prototype.getAttribute = function (name) { return Object.prototype.hasOwnProperty.call(this.a, name) ? this.a[name] : null; };
StgEl_.prototype.getText = function () { return this.t; };
StgEl_.prototype.getValue = function () { return this.t + this.c.map(function (x) { return x.getValue(); }).join(''); };
function stgParseXml_(xml) {
  xml = xml.replace(/\r\n?/g, '\n');        // XML line-ending normalisation (same as XmlService)
  var re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<!DOCTYPE[^>]*>|<(\/?)([A-Za-z_][\w.:-]*)([^>]*?)(\/?)>/g;
  var attrRe = /([\w.:-]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  var top = new StgEl_('#doc', {}), stack = [top], last = 0, m;
  while ((m = re.exec(xml))) {
    var cur = stack[stack.length - 1];
    if (m.index > last) cur.t += stgDecode_(xml.slice(last, m.index));
    last = re.lastIndex;
    if (m[1] != null) { cur.t += m[1]; continue; }
    if (!m[3]) continue;                       // comment, declaration, doctype
    if (m[2] === '/') {
      if (stack.length < 2 || cur.n !== m[3]) throw new Error('Bad XML: </' + m[3] + '>');
      stack.pop(); continue;
    }
    var attrs = {}, am; attrRe.lastIndex = 0;
    while ((am = attrRe.exec(m[4]))) attrs[am[1]] = stgDecode_(am[3] != null ? am[3] : am[4]);
    var el = new StgEl_(m[3], attrs); cur.c.push(el);
    if (!m[5]) stack.push(el);
  }
  if (stack.length !== 1 || top.c.length !== 1) throw new Error('Bad XML: unclosed element');
  return top.c[0];
}

/* ==== STG subject extractor – shared by the Apps Script sync and the dashboard ====
   Works on any element exposing getName(), getChildren([name]), getChild(name), getChildText(name)
   (XmlService in Apps Script; a small DOM adapter in the browser). Keep both copies identical. */
function stgSplitSubjects_(xml) {
  var re = /<Subject[\s>]/g, starts = [], m;
  while ((m = re.exec(xml))) starts.push(m.index);
  return starts.map(function (a, i) {
    var seg = xml.slice(a, i + 1 < starts.length ? starts[i + 1] : xml.length);
    var e = seg.lastIndexOf('</Subject>');
    return e >= 0 ? seg.slice(0, e + 10) : seg;
  });
}
function stgDescendants_(el, name, out) {
  el.getChildren().forEach(function (c) {
    if (c.getName() === name) out.push(c); else stgDescendants_(c, name, out);
  });
  return out;
}
function stgCodes_(el) {
  var a = el.getChild('Answers');
  return a ? a.getChildren('Answer').map(function (x) {
    return { c: Number(x.getChildText('Code')), t: x.getChildText('Text') || '' };
  }) : [];
}
function stgAttr_(el, name) {
  var a = el.getAttribute(name);
  return a == null ? null : (typeof a === 'string' ? a : a.getValue());
}
// value of the Subject child that carries Field="<field>" (e.g. SbjNam, Test)
function stgField_(s, field) {
  var hit = s.getChildren().filter(function (c) { return stgAttr_(c, 'Field') === field; })[0];
  return hit ? hit.getText() : null;
}
function stgTopicKey_(t) { return Number(String(t.getChildText('TopicVariable') || '').split('_').pop()); }
function stgExtractSubject_(s) {
  var d = new Date(String(s.getChildText('Date') || ''));
  var r = { id: Number(s.getChildText('SubjectNum')),
    // interview date in Vietnam time (UTC+7)
    date: isNaN(d) ? '' : new Date(d.getTime() + 7 * 3600e3).toISOString().slice(0, 10),
    done: s.getChildText('Completed') === 'True', st: s.getChildText('Status') || '',
    // interviewer name (SbjNam): the dashboard treats names like STGTester / RMIT as test interviews
    nam: String(stgField_(s, 'SbjNam') || '').trim(),
    // test interviews: interviewer STGTester / RMIT, or Test flag (same rule as the QC dashboard)
    test: /^(STGTester|RMIT)$/i.test(String(stgField_(s, 'SbjNam') || '').trim()) || stgField_(s, 'Test') === 'True',
    A: {}, B1: [], B4: {}, C5: {}, D4: [], E4: {} };
  stgDescendants_(s, 'Question', []).forEach(function (q) {
    var v = q.getChildText('Variable');
    if (['A1', 'A4', 'A7', 'A10', 'A11'].indexOf(v) >= 0) {
      var a = stgCodes_(q)[0]; if (a) r.A[v] = [a.c, a.t];
    } else if (v === 'B1') {
      r.B1 = stgCodes_(q).map(function (a) { return [a.c, a.t]; });
    } else if (v === 'B4' || v === 'C5' || v === 'E4') {
      var tp = q.getChild('Topics'); if (!tp) return;
      tp.getChildren('Topic').forEach(function (t) {
        var cs = stgCodes_(t).map(function (a) { return a.c; }); if (!cs.length) return;
        if (v === 'E4') r.E4[stgTopicKey_(t)] = cs; else r[v][stgTopicKey_(t)] = cs[0];
      });
    } else if (v === 'D4') {
      var fv = q.getChildText('FullVariable') || '';
      var m = /^RF(\d+)_/.exec(fv), o = /khác\s*(\d)/i.exec(fv);
      var rf = m ? Number(m[1]) : (o ? 94 + Number(o[1]) : 0); if (!rf) return;
      var tp2 = q.getChild('Topics'); if (!tp2) return;
      tp2.getChildren('Topic').forEach(function (t) {
        var c = stgCodes_(t)[0]; if (c) r.D4.push([rf, stgTopicKey_(t), c.c]);
      });
    }
  });
  return r;
}
/* ---- QC page: compact per-subject record {m, a, times} (same shape qc.js builds from raw XML) ---- */
var STG_QC_META = ['SbjNum', 'SbjNam', 'Test', 'Complete', 'Status', 'CustomData1', 'VStart', 'NumberOfSessions'];
function stgQcAnswers_(el) {
  var an = el.getChild('Answers'); if (!an) return [];
  return an.getChildren('Answer').map(function (x) {
    return {code: x.getChildText('Code'), text: x.getChildText('Text'), oth: x.getChildText('OtherSpec')};
  });
}
function stgQcSubject_(s) {
  var m = {}, a = {}, times = [];
  s.getChildren().forEach(function (c) {
    if (c.getName() === 'Chapter') return;
    var k = stgAttr_(c, 'Field') || c.getName();
    if (STG_QC_META.indexOf(k) >= 0) m[k] = c.getValue();
  });
  stgDescendants_(s, 'Question', []).forEach(function (q) {
    var v = q.getChildText('Variable'), qt = q.getChildText('QType'), t = q.getChildText('AnsweredTimeStamp');
    if (t) times.push(new Date(t).getTime());
    if (qt === 'eqtAmerican') {
      var an = stgQcAnswers_(q)[0];
      a[v] = {code: an ? an.code : null, text: q.getChildText('QuestionAnswer'), oth: an ? an.oth : null};
    } else if (qt === 'eqtMultiSelect') {
      a[v] = stgQcAnswers_(q);
    } else {
      var d = {}, tps = q.getChild('Topics');
      (tps ? tps.getChildren('Topic') : []).forEach(function (tp) {
        d[tp.getChildText('Text')] = (qt === 'eqtMultiNumeric' || qt === 'eqtMultiText') ? tp.getChildText('TopicAnswer') : stgQcAnswers_(tp);
      });
      a[v] = d;
    }
  });
  times.sort(function (x, y) { return x - y; });
  return {m: m, a: a, times: times.length > 1 ? [times[0], times[times.length - 1]] : []};
}
function stgQcJson_(rec) { return JSON.stringify(rec, function (k, v) { return v === null ? undefined : v; }); }
/* ==== end shared extractor ==== */
