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
    // test interviews: interviewer STGTester or Test flag (same rule as the QC dashboard)
    test: stgField_(s, 'SbjNam') === 'STGTester' || stgField_(s, 'Test') === 'True',
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
/* ==== end shared extractor ==== */
