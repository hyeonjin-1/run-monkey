// ==UserScript==
// @name         건생포인트 자동 차감 (요양기관 정보마당)
// @namespace    wellcheck-tools
// @version      1.5
// @description  차감 대기 파일(CSV·엑셀)을 넣으면 포인트차감신청을 자동으로 진행하고, 잔여 포인트를 기준 일시와 함께 기록합니다.
// @match        https://medicare.nhis.or.kr/*
// @match        https://*.nhis.or.kr/*
// @updateURL    https://raw.githubusercontent.com/hyeonjin-1/run-monkey/main/point.user.js
// @downloadURL  https://raw.githubusercontent.com/hyeonjin-1/run-monkey/main/point.user.js
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @grant        unsafeWindow
// @run-at       document-idle
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// ==/UserScript==

/*
 * 동작 순서 (환자 1명마다)
 *  1) 포인트차감신청 화면: 성명만 입력 → [조회]
 *     동명이인이 여러 명 나오면 파일의 생년월일과 일치하는 사람을 선택
 *  2) [차감신청] → '포인트 차감' 창
 *  3) 진료일은 달력에서 선택, 차감 신청 포인트 = 차감할 금액, 동의 체크 → [차감신청]
 *  4) '차감 후 잔액'을 잔여 포인트로 기록 (기준 일시 함께)
 *  - 잔액이 모자라면 차감하지 않고 '포인트 부족'으로 기록 (상태는 차감 대기 유지)
 *  - 3번 연속 실패하면 자동으로 멈춤 (화면 구조가 바뀐 경우 대비)
 *
 * 화면 요소는 좌표가 아니라 '글자(라벨)'로 찾습니다.
 * 환자 정보는 이 PC의 브라우저 안에서만 처리되며 외부로 전송하지 않습니다.
 */
(function () {
  'use strict';

  const KEY = 'wc_point_state_v1';
  const W = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;
  const PHASE_TIMEOUT = 25000;   // 단계별 최대 대기 (ms)
  const MAX_FAIL_STREAK = 3;     // 연속 실패 시 자동 중지

  // ---------------- 상태 저장 (페이지가 새로고침돼도 이어서 진행) ----------------
  function load() { try { return JSON.parse(GM_getValue(KEY, 'null')); } catch (e) { return null; } }
  function save(s) { s.updated = Date.now(); GM_setValue(KEY, JSON.stringify(s)); }
  function clearState() { GM_deleteValue(KEY); }

  // ---------------- 화면 요소 찾기 도우미 ----------------
  const norm = t => (t || '').replace(/[\s*·:：※]/g, '');
  const isVisible = el => !!el && (el.offsetParent !== null || el.getClientRects().length > 0);
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  // 글자가 정확히 일치하는 라벨 요소 (가장 안쪽 요소, 문서 순서상 첫 번째)
  function findLabel(root, text, exclude) {
    const want = norm(text);
    const els = root.querySelectorAll('th,td,label,span,div,p,strong,b,dt,dd,li,font,h1,h2,h3,h4,em');
    let best = null;
    for (const e of els) {
      if (exclude && exclude.some(x => x && x.contains(e))) continue;
      if (!isVisible(e)) continue;
      if (norm(e.textContent) !== want) continue;
      if (!best) best = e; else if (best.contains(e)) best = e;
    }
    return best;
  }

  const FOLLOWING = 4; // Node.DOCUMENT_POSITION_FOLLOWING
  // 라벨 뒤에 오는 입력칸 (같은 줄 우선)
  function inputAfter(label, root, filter) {
    if (!label) return null;
    const inputs = Array.from(root.querySelectorAll('input,select,textarea'))
      .filter(i => i.type !== 'hidden' && i.type !== 'button' && i.type !== 'submit' && isVisible(i) && (!filter || filter(i)))
      .filter(i => label.compareDocumentPosition(i) & FOLLOWING);
    const row = label.closest('tr');
    if (row) { const inRow = inputs.find(i => row.contains(i)); if (inRow) return inRow; }
    let p = label.parentElement;
    for (let k = 0; k < 3 && p; k++, p = p.parentElement) {
      const near = inputs.find(i => p.contains(i));
      if (near) return near;
    }
    return inputs[0] || null;
  }

  // 라벨 옆의 값 (입력칸 값 또는 다음 칸 글자)
  function valueAfterLabel(root, text, within) {
    const lbl = findLabel(within || root, text);
    if (!lbl) return null;
    const inp = inputAfter(lbl, within || root);
    if (inp && (!lbl.closest('tr') || lbl.closest('tr').contains(inp) || lbl.parentElement.contains(inp))) return inp.value;
    let n = lbl.nextElementSibling;
    while (n && !norm(n.textContent)) n = n.nextElementSibling;
    return n ? n.textContent : null;
  }

  const toNum = v => { const d = String(v == null ? '' : v).replace(/[^\d-]/g, ''); return d === '' ? null : parseInt(d, 10); };

  // 버튼 찾기 (button/a/input 우선, 없으면 클릭 가능한 다른 요소)
  function findButton(root, text, opt) {
    opt = opt || {};
    const want = norm(text);
    const scope = opt.within || root;
    const label = e => norm(e.tagName === 'INPUT' ? (e.value || e.alt || e.title)
                     : e.tagName === 'IMG' ? (e.alt || e.title) : (e.textContent || e.title || ''));
    const ok = e => isVisible(e) && !(opt.exclude || []).some(x => x && x.contains(e)) && label(e) === want;
    const first = scope.querySelectorAll('button,a,input[type=button],input[type=submit],input[type=image],img');
    for (const e of first) if (ok(e)) return e;
    for (const e of scope.querySelectorAll('span,div,td,li')) if (ok(e) && !Array.from(e.children).some(ok)) return e;
    return null;
  }

  // 비활성화된 버튼인지 (WebSquare: disabled 속성 또는 class에 disabled)
  function isDisabled(e) {
    if (!e) return true;
    if (e.disabled || e.getAttribute('disabled') != null || e.getAttribute('aria-disabled') === 'true') return true;
    if (/(^|\s|_)disabled/i.test(e.className || '')) return true;
    return !!e.querySelector('[class*="disabled"]');
  }

  // WebSquare 화면 구성요소 (값을 넣을 때 화면 칸과 내부 데이터를 같이 바꾸기 위함)
  function wsComp(el) {
    const ids = [];
    if (el.id) { ids.push(el.id, el.id.replace(/_input$/, '')); }
    const host = el.parentElement && el.parentElement.closest('[id]');
    if (host) ids.push(host.id);
    for (const id of ids) {
      try { const c = W.$p && W.$p.getComponentById && W.$p.getComponentById(id); if (c && c.setValue) return c; } catch (e) {}
      try { const c = W.WebSquare && W.WebSquare.util && W.WebSquare.util.getComponentById(id); if (c && c.setValue) return c; } catch (e) {}
    }
    return null;
  }

  function setVal(el, v, wsValue) {
    const comp = wsComp(el);
    if (comp) { try { comp.setValue(wsValue != null ? wsValue : v); } catch (e) {} }
    const win = el.ownerDocument.defaultView;
    const proto = el.tagName === 'SELECT' ? win.HTMLSelectElement.prototype
                : el.tagName === 'TEXTAREA' ? win.HTMLTextAreaElement.prototype : win.HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
    try { el.focus(); } catch (e) {}
    setter.call(el, v);
    if (comp && el.value === (wsValue != null ? wsValue : v)) { /* WebSquare가 이미 화면에 반영 */ }
    else setter.call(el, v);
    for (const t of ['input', 'keydown', 'keyup', 'change']) el.dispatchEvent(new win.Event(t, { bubbles: true }));
    try { el.blur(); } catch (e) {}
    el.dispatchEvent(new win.Event('blur', { bubbles: true }));
  }

  // 사람이 키보드로 치는 것처럼 입력 (WebSquare 칸은 값만 바꾸면 조회할 때 빈칸으로 읽는 경우가 있음)
  function typeText(el, text) {
    const doc = el.ownerDocument, win = doc.defaultView;
    try { el.focus(); el.click(); } catch (e) {}
    try { el.select(); } catch (e) {}
    let ok = false;
    try { ok = doc.execCommand('selectAll', false, null) && doc.execCommand('insertText', false, text); } catch (e) { ok = false; }
    if (!ok || el.value !== text) setVal(el, text);
    else {
      const comp = wsComp(el);
      if (comp) { try { comp.setValue(text); } catch (e) {} }
      for (const t of ['keyup', 'change']) el.dispatchEvent(new win.Event(t, { bubbles: true }));
    }
    return el.value === text;
  }

  function click(el) { try { el.click(); } catch (e) { el.dispatchEvent(new MouseEvent('click', { bubbles: true })); } }

  // ---------------- 화면 판별 ----------------
  // '포인트 차감' 창 (같은 화면 레이어, iframe, 새 창 모두 대응)
  function findPopup(doc) {
    const lbl = findLabel(doc, '차감 신청 포인트');
    if (!lbl) return null;
    let el = lbl;
    while (el && el !== doc.documentElement) {
      if (el.querySelector('input[type=checkbox]') && findButton(doc, '차감신청', { within: el })) return el;
      el = el.parentElement;
    }
    return null;
  }

  // 포인트차감신청 메인 화면
  function isMainForm(doc) {
    const popup = findPopup(doc);
    const ex = [popup, doc.getElementById('wc-point-panel')];
    return !!(findLabel(doc, '환자등록번호', ex) && findButton(doc, '조회', { exclude: ex }));
  }

  // ---------------- 알림창/확인창 자동 처리 (실행 중일 때만) ----------------
  function hookDialogs() {
    try {
      if (W.__wcPointHooked) return;
      W.__wcPointHooked = true;
      const oa = W.alert, oc = W.confirm;
      W.alert = function (m) {
        const s = load();
        if (s && s.running) { s.lastAlert = String(m); s.lastAlertAt = Date.now(); save(s); return; }
        return oa.call(W, m);
      };
      W.confirm = function (m) {
        const s = load();
        if (s && s.running) { s.lastConfirm = String(m); s.lastConfirmAt = Date.now(); save(s); return true; }
        return oc.call(W, m);
      };
    } catch (e) { /* 무시 */ }
  }

  // 사이트 자체 메시지 창 (WebSquare 팝업 등): 실행 중이면 내용을 기록하고 [확인]
  function handleMessageBoxes(doc) {
    const s = load();
    if (!s || !s.running) return;
    const panel = doc.getElementById('wc-point-panel');
    const popup = findPopup(doc);
    const boxes = Array.from(doc.querySelectorAll('[class*="w2window"],[class*="w2popup"],[class*="alert"],[class*="Alert"],[class*="msgbox"],[class*="message"],[role="dialog"],[role="alertdialog"]'))
      .filter(b => isVisible(b) && !(panel && panel.contains(b)) && !(popup && (b.contains(popup) || popup.contains(b))));
    for (const b of boxes) {
      if (boxes.some(o => o !== b && b.contains(o) && isVisible(o))) continue;   // 가장 안쪽 상자만
      const ok = findButton(doc, '확인', { within: b }) || findButton(doc, '예', { within: b });
      if (!ok) continue;
      const text = (b.innerText || '').replace(/확인|취소|닫기|예|아니오/g, '').replace(/\s+/g, ' ').trim();
      if (!text || text.length > 300) continue;
      const s2 = load(); s2.lastAlert = text; s2.lastAlertAt = Date.now(); save(s2);
      click(ok);
      return;
    }
  }

  // ---------------- 날짜/시간 ----------------
  const pad = n => String(n).padStart(2, '0');
  function nowStamp() { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`; }
  function today() { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
  function asOfLabel(stamp) { const m = /(\d+)-(\d+)-(\d+) (\d+):(\d+)/.exec(stamp); return m ? `${+m[2]}월 ${+m[3]}일 ${m[4]}:${m[5]} 기준` : stamp; }
  // 화면의 기존 날짜 형식(2025.12.01 / 2025-12-01 / 20251201)에 맞춰 변환
  function formatLike(sample, ymd) {
    const [y, m, d] = ymd.split(/[-./]/);
    const sep = (/\d{4}([.\-/])\d{2}/.exec(sample || '') || [])[1];
    if (sample && /^\d{8}$/.test(sample.trim())) return `${y}${m}${d}`;
    return [y, m, d].join(sep || '.');
  }

  // ---------------- CSV ----------------
  function parseCSV(text) {
    const rows = []; let row = [], cur = '', q = false;
    text = text.replace(/^﻿/, '');
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; }
        else cur += c;
      } else if (c === '"') q = true;
      else if (c === ',') { row.push(cur); cur = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(cur); cur = ''; if (row.some(x => x !== '')) rows.push(row); row = [];
      } else cur += c;
    }
    row.push(cur); if (row.some(x => x !== '')) rows.push(row);
    return rows;
  }
  const csvCell = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';

  async function readFileText(file) {
    const buf = await file.arrayBuffer();
    let t = new TextDecoder('utf-8').decode(buf);
    if (t.includes('�')) { try { t = new TextDecoder('euc-kr').decode(buf); } catch (e) {} }
    return t;
  }

  // 엑셀(.xlsx/.xls) → CSV와 같은 2차원 표. 날짜 칸은 YYYY-MM-DD, 생년월일은 6자리로 맞춤
  async function readExcelRows(file) {
    if (typeof XLSX === 'undefined') throw new Error('엑셀 읽기 도구를 불러오지 못했어요. 인터넷 연결 확인 후 새로고침하거나 CSV로 넣어 주세요.');
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    const fmtDate = v => {
      if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
      if (typeof v === 'number' && v > 20000 && v < 80000) { const d = new Date(Math.round((v - 25569) * 86400000)); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; }
      return String(v);
    };
    // 머리글 줄 찾기 (위에 제목 줄이 있어도 됨)
    let hi = raw.findIndex(r => r.some(c => String(c).trim() === '이름') && r.some(c => String(c).trim() === '진료일'));
    if (hi < 0) hi = 0;
    const head = raw[hi].map(c => String(c).trim());
    const di = head.indexOf('진료일'), bi = head.indexOf('생년월일'), ci = head.indexOf('차감한 날');
    return raw.slice(hi).map((r, k) => r.map((c, j) => {
      if (k === 0) return String(c).trim();
      if (j === di || j === ci) return fmtDate(c);
      if (j === bi) { const d = String(c).replace(/\D/g, ''); return d && d.length < 6 ? d.padStart(6, '0') : d; }
      return c instanceof Date ? fmtDate(c) : String(c);
    })).filter(r => r.some(x => String(x).trim() !== ''));
  }

  function buildState(text, testOne, rowsIn) {
    const rows = rowsIn || parseCSV(text);
    if (rows.length < 2) throw new Error('파일에 데이터가 없습니다.');
    const head = rows[0].map(h => h.trim());
    const col = name => head.indexOf(name);
    const need = ['진료일', '이름', '생년월일', '차감할 금액', '상태'];
    const miss = need.filter(n => col(n) < 0);
    if (miss.length) throw new Error('파일에 다음 칸이 없습니다: ' + miss.join(', '));
    const items = [];
    rows.slice(1).forEach((r, i) => {
      const amount = toNum(r[col('차감할 금액')]);
      if ((r[col('상태')] || '').replace(/\s/g, '') === '차감대기' && amount > 0) {
        items.push({ row: i, date: r[col('진료일')].trim(), name: r[col('이름')].trim(), birth: r[col('생년월일')].trim(), amount });
      }
    });
    return { head, rows: rows.slice(1), items, idx: 0, phase: 'search', phaseAt: Date.now(), running: false,
             results: {}, failStreak: 0, testOne: !!testOne, log: [] };
  }

  // ---------------- 결과 파일 ----------------
  function buildResultCSV(s) {
    const head = s.head.slice();
    const ci = n => head.indexOf(n);
    const extra = ['처리 결과', '처리 메모', '잔여 포인트', '잔여 포인트 기준'];
    const out = [head.concat(extra)];
    s.rows.forEach((r, i) => {
      const row = r.slice(); while (row.length < head.length) row.push('');
      const res = s.results[i];
      if (res) {
        if (res.status === '성공') {
          if (ci('상태') >= 0) row[ci('상태')] = '차감 완료';
          if (ci('차감한 날') >= 0) row[ci('차감한 날')] = res.day;
        }
        out.push(row.concat([res.status, res.msg || '', res.remain == null ? '' : res.remain, res.at ? asOfLabel(res.at) : '']));
      } else out.push(row.concat(['', '', '', '']));
    });
    return '﻿' + out.map(r => r.map(csvCell).join(',')).join('\r\n');
  }

  function download(name, text) {
    const blob = new Blob([text], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  // ---------------- 진행 기록 ----------------
  function logLine(s, msg) { s.log.push(`[${nowStamp().slice(11)}] ${msg}`); if (s.log.length > 200) s.log.shift(); }

  function record(s, status, msg, remain) {
    const it = s.items[s.idx];
    const at = nowStamp();
    s.results[it.row] = { status, msg, remain: remain == null ? null : remain, at: remain == null ? null : at, day: today() };
    const r = remain == null ? '' : ` · 잔여 ${remain.toLocaleString()}P (${asOfLabel(at)})`;
    logLine(s, `${it.name}(${it.date}) ${it.amount.toLocaleString()}P → ${status}${msg ? ' · ' + msg : ''}${r}`);
    s.failStreak = (status === '실패') ? (s.failStreak || 0) + 1 : 0;
  }

  // 조회만 테스트가 끝나면 실제 진행 상태(몇 번째까지 했는지, 결과)를 원래대로 돌려 놓음
  function endLook(s) {
    const b = s.lookBackup || { idx: 0, results: {}, failStreak: 0 };
    s.idx = b.idx; s.results = b.results; s.failStreak = b.failStreak; s.lookOnly = false; s.lookBackup = null; s.finished = false;
  }

  function nextItem(s) {
    s.idx++; s.phase = 'search'; s.phaseAt = Date.now();
    s.lastAlert = null; s.popupFor = null;
    if (s.failStreak >= MAX_FAIL_STREAK) {
      s.running = false; s.stopReason = `${MAX_FAIL_STREAK}번 연속 실패해서 멈췄어요. 화면 구조를 확인해 주세요.`;
      logLine(s, '⛔ ' + s.stopReason);
    } else if (s.testOne && s.idx > (s.testStart || 0)) {
      s.running = false; s.stopReason = '1명 테스트 완료';
      logLine(s, '✔ 1명 테스트 완료');
    } else if (s.idx >= s.items.length) {
      s.running = false; s.finished = true;
      logLine(s, '✔ 모두 처리했어요');
    }
    if (s.lookOnly && !s.running) {
      const n = Object.values(s.results).filter(r => r.status === '조회 확인').length;
      logLine(s, `🔍 조회만 테스트 끝: ${n}/${Object.keys(s.results).length}명 조회 성공 (차감은 하지 않았어요)`);
      endLook(s);
      s.stopReason = '조회만 테스트 끝 · 문제 없으면 [1명만 테스트]나 [전체 자동 차감]을 눌러 주세요';
    }
  }

  // 팝업 닫기 (X, 닫기, 취소)
  function closePopup(doc, popup) {
    const scopes = [];
    let el = popup;
    for (let k = 0; k < 6 && el; k++, el = el.parentElement) scopes.push(el);
    for (const sc of scopes) {
      for (const t of ['닫기', '×', 'X', 'x', '취소']) {
        const b = findButton(doc, t, { within: sc });
        if (b) { click(b); return true; }
      }
      const c = sc.querySelector('[title*="닫기"],[aria-label*="닫기"],[class*="close"],[class*="Close"]');
      if (c && isVisible(c)) { click(c); return true; }
    }
    try { if (doc.defaultView !== doc.defaultView.top && !doc.defaultView.frameElement) doc.defaultView.close(); } catch (e) {}
    return false;
  }

  // 이름 일치 (포털이 '강*호'처럼 가려서 보여줘도 인정)
  function nameRegex(name) {
    return new RegExp(name.split('').map(c => `[${c.replace(/[\\\]^-]/g, '\\$&')}*＊○●Oo]`).join(''));
  }

  // 생년월일(720305)이 글자 안에 있는지 (720305, 1972-03-05, 1972.03.05, 72.03.05, 19720305 등)
  function birthForms(b6) {
    const yy = b6.slice(0, 2), mm = b6.slice(2, 4), dd = b6.slice(4, 6);
    const cy = new Date().getFullYear() % 100;
    const yyyy = (parseInt(yy, 10) > cy ? '19' : '20') + yy;
    return [b6, yyyy + mm + dd, `${yyyy}-${mm}-${dd}`, `${yyyy}.${mm}.${dd}`, `${yyyy}/${mm}/${dd}`, `${yy}.${mm}.${dd}`, `${yy}-${mm}-${dd}`,
            `${yyyy}년${mm}월${dd}일`, `${yyyy}년${+mm}월${+dd}일`];
  }
  const hasBirth = (text, b6) => { const t = (text || '').replace(/\s/g, ''); return birthForms(b6).some(f => t.includes(f)); };
  const birthLike = text => /(?<!\d)\d{6}(?!\d)|\d{4}[.\-/]\d{2}[.\-/]\d{2}|(?<!\d)\d{2}[.\-/]\d{2}[.\-/]\d{2}(?!\d)/.test((text || '').replace(/\s/g, ''));

  // 조회 결과 중 이름이 들어 있는 줄(행)들
  function rowsWithName(doc, name, ex) {
    const re = nameRegex(name);
    return Array.from(doc.querySelectorAll('tr,li')).filter(r =>
      isVisible(r) && !ex.some(x => x && x.contains(r)) && re.test(r.innerText || '') &&
      !r.querySelector('input[type=text],input:not([type])') &&             // 검색 입력칸이 있는 줄 제외
      !Array.from(r.querySelectorAll('tr,li')).some(c => re.test(c.innerText || '')));   // 가장 안쪽 줄만
  }
  // 행 안에서 누를 것 (선택 버튼/링크/라디오 → 없으면 행 자체)
  function pickInRow(row) {
    return row.querySelector('input[type=radio],input[type=checkbox]') ||
           Array.from(row.querySelectorAll('button,a,input[type=button]')).find(isVisible) || row;
  }

  // ---- 달력에서 날짜 선택 ----
  function calendarBox(doc, before) {
    const cands = Array.from(doc.querySelectorAll('div,table')).filter(e => isVisible(e) && !before.has(e) &&
      /datepicker|calendar|(^|[\s_-])cal([\s_-]|$)/i.test((e.id || '') + ' ' + (e.className || '')));
    // 가장 바깥 달력 상자
    return cands.find(c => !cands.some(o => o !== c && o.contains(c))) || null;
  }
  function calHeader(box) {
    const sels = box.querySelectorAll('select');
    if (sels.length >= 2) {
      const y = toNum(sels[0].value), m = toNum(sels[1].value);
      if (y > 1900) return { y, m: m >= 1 && m <= 12 ? m : m + 1 };
    }
    const m = /(\d{4})\s*[년.\-/]?\s*(\d{1,2})\s*월?/.exec(box.innerText || '');
    return m ? { y: +m[1], m: +m[2] } : null;
  }
  function calNav(box, dir) {
    const re = dir < 0 ? /prev|이전|◀|‹|«|^<$/i : /next|다음|▶|›|»|^>$/i;
    const els = Array.from(box.querySelectorAll('a,button,span,img,input[type=button],th,td')).filter(isVisible);
    return els.find(e => re.test(norm(e.textContent)) && norm(e.textContent).length <= 3) ||
           els.find(e => re.test((e.title || '') + ' ' + (e.className || '') + ' ' + (e.alt || '') + ' ' + (e.value || '')));
  }
  async function pickDateByCalendar(doc, input, ymd) {
    const [ty, tm, td] = ymd.split(/[-./]/).map(n => parseInt(n, 10));
    const before = new Set(Array.from(doc.querySelectorAll('div,table')).filter(isVisible));
    // 달력 아이콘: 입력칸 근처의 이미지/버튼
    let scope = input.parentElement, icon = null;
    for (let k = 0; k < 3 && scope && !icon; k++, scope = scope.parentElement) {
      icon = Array.from(scope.querySelectorAll('img,button,a,input[type=button],input[type=image],span'))
        .filter(e => e !== input && isVisible(e))
        .find(e => /달력|calendar|cal|date/i.test((e.alt || '') + (e.title || '') + (e.className || '') + (e.src || '') + (e.textContent || '')) || e.tagName === 'IMG');
    }
    if (icon) click(icon); else click(input);
    await sleep(500);
    let box = calendarBox(doc, before);
    if (!box) return 'NOCAL';
    for (let i = 0; i < 40; i++) {
      const h = calHeader(box);
      if (!h) return 'NOHEADER';
      const diff = (ty - h.y) * 12 + (tm - h.m);
      if (diff === 0) break;
      const sels = box.querySelectorAll('select');
      if (sels.length >= 2 && i === 0) {
        setVal(sels[0], String(ty));
        const mOpt = Array.from(sels[1].options).find(o => toNum(o.textContent) === tm || toNum(o.value) === tm) || sels[1].options[tm - 1];
        if (mOpt) setVal(sels[1], mOpt.value);
        await sleep(300); box = calendarBox(doc, before) || box; continue;
      }
      const nav = calNav(box, diff);
      if (!nav) return 'NONAV';
      click(nav); await sleep(250);
      box = calendarBox(doc, before) || box;
    }
    const h2 = calHeader(box);
    if (!h2 || h2.y !== ty || h2.m !== tm) return 'MONTH_FAIL';
    const days = Array.from(box.querySelectorAll('a,td,button,span')).filter(e =>
      isVisible(e) && norm(e.textContent) === String(td) &&
      !/other|disabled|unselectable|prev|next|off|dim/i.test((e.className || '') + ' ' + ((e.parentElement && e.parentElement.className) || '')) &&
      !Array.from(e.children).some(c => norm(c.textContent) === String(td)));
    if (!days.length) return 'NODAY';
    // 다른 달 날짜가 같이 보이는 달력: 16일 이후는 마지막 칸, 15일 이전은 첫 칸이 이번 달
    click(td > 15 ? days[days.length - 1] : days[0]);
    await sleep(400);
    return 'OK';
  }

  // 페이지 변화 감지 (조회 결과가 새로 그려졌는지 확인용)
  const pageLoadedAt = Date.now();
  let lastMutationAt = 0;
  function watchMutations(doc) {
    if (doc.__wcObs) return;
    doc.__wcObs = new MutationObserver(muts => {
      const panel = doc.getElementById('wc-point-panel');
      if (muts.some(m => !(panel && panel.contains(m.target)))) lastMutationAt = Date.now();
    });
    doc.__wcObs.observe(doc.body, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ['value', 'style', 'class'] });
  }

  // ---------------- 메인 화면 단계 ----------------
  let busy = false;
  async function mainTick(doc) {
    if (busy) return;
    const s = load();
    if (!s || !s.running) return;
    const it = s.items[s.idx];
    if (!it) { s.running = false; s.finished = true; save(s); return; }
    const panel = doc.getElementById('wc-point-panel');
    const popup = findPopup(doc);
    const ex = [panel, popup];
    busy = true;
    try {
      // 단계별 시간 초과
      if (Date.now() - s.phaseAt > PHASE_TIMEOUT && s.phase !== 'search') {
        record(s, '실패', `시간 초과(${s.phase})${s.lastAlert ? ' · ' + s.lastAlert : ''}`);
        if (popup) closePopup(doc, popup);
        nextItem(s); save(s); return;
      }

      if (s.phase === 'search') {
        if (popup) { closePopup(doc, popup); await sleep(500); }
        const nameIn = doc.getElementById('txt_patFnm') || inputAfter(findLabel(doc, '성명', ex), doc, i => !ex.some(x => x && x.contains(i)));
        const regIn = inputAfter(findLabel(doc, '환자등록번호', ex), doc, i => !ex.some(x => x && x.contains(i)));
        const btn = findButton(doc, '조회', { exclude: ex });
        if (!nameIn || !btn) {
          record(s, '실패', '성명 입력칸 또는 조회 버튼을 찾지 못함');
          nextItem(s); save(s); return;
        }
        if (regIn && regIn !== nameIn && !regIn.readOnly && regIn.value) setVal(regIn, '');   // 이름만으로 조회 (먼저 비우고)
        typeText(nameIn, it.name);
        await sleep(400);
        // 조회 누르기 직전에 성명이 정말 들어가 있는지 확인 (비어 있으면 한 번 더 입력)
        if (nameIn.value.trim() !== it.name) { typeText(nameIn, it.name); await sleep(400); }
        if (nameIn.value.trim() !== it.name) {
          record(s, '실패', `성명 칸에 이름이 입력되지 않음 (칸 값: "${nameIn.value}")`);
          nextItem(s); save(s); return;
        }
        try { nameIn.blur(); } catch (e) {}
        s.lastAlert = null; s.phase = 'searched'; s.phaseAt = Date.now(); s.searchedAt = Date.now(); s.picked = null;
        save(s);
        await sleep(300);
        // WebSquare 버튼은 <a> 안에 <span>조회</span> 구조라서, 바깥 <a>를 누르고 마우스 신호도 함께 보냄
        const target = (btn.closest && btn.closest('a,button')) || btn;
        const win = target.ownerDocument.defaultView;
        ['mousedown', 'mouseup'].forEach(t => { try { target.dispatchEvent(new win.MouseEvent(t, { bubbles: true, cancelable: true })); } catch (e) {} });
        click(target);
        return;
      }

      if (s.phase === 'searched') {
        if (s.lastAlert && s.lastAlertAt >= s.searchedAt) {
          record(s, '실패', '조회: ' + s.lastAlert);
          nextItem(s); save(s); return;
        }
        if (Date.now() - s.searchedAt < 800) return;
        // 조회 후 화면이 새로 그려졌고(또는 새로고침됐고) 참여자 정보에 이 환자 이름이 보여야 진행
        // → 이전 환자 화면이 남아 있는 상태에서 차감하는 사고 방지
        const refreshed = lastMutationAt >= s.searchedAt || pageLoadedAt > s.searchedAt;
        if (!refreshed) return;
        const text = (doc.body.innerText || '').replace(panel ? panel.innerText : '', '');
        if (!nameRegex(it.name).test(text)) return;

        // 동명이인: 같은 이름 줄이 여러 개면 생년월일이 일치하는 사람 선택
        if (s.picked !== s.idx) {
          const rows = rowsWithName(doc, it.name, ex);
          if (rows.length >= 2) {
            const match = rows.filter(r => hasBirth(r.innerText, it.birth));
            if (match.length === 0) { record(s, '실패', `동명이인 ${rows.length}명 중 생년월일(${it.birth}) 일치 없음`); nextItem(s); save(s); return; }
            if (match.length > 1) { record(s, '실패', `같은 이름·생년월일 ${match.length}명 · 직접 확인 필요`); nextItem(s); save(s); return; }
            s.picked = s.idx; s.searchedAt = Date.now(); s.phaseAt = Date.now();
            logLine(s, `${it.name}: 동명이인 ${rows.length}명 중 생년월일 ${it.birth} 선택`);
            save(s); click(pickInRow(match[0])); return;
          }
          // 한 줄짜리 목록이면 눌러서 선택해야 하는 경우
          if (rows.length === 1 && !findButton(doc, '차감신청', { exclude: ex }) && pickInRow(rows[0]) !== rows[0]) {
            if (birthLike(rows[0].innerText) && !hasBirth(rows[0].innerText, it.birth)) {
              record(s, '실패', `조회된 환자의 생년월일이 파일(${it.birth})과 다름`); nextItem(s); save(s); return;
            }
            s.picked = s.idx; s.searchedAt = Date.now(); s.phaseAt = Date.now(); save(s);
            click(pickInRow(rows[0])); return;
          }
        }
        // 참여자 정보에 생년월일이 표시되면 한 번 더 확인
        const bVal = valueAfterLabel(doc, '생년월일', null);
        if (bVal && birthLike(bVal) && !hasBirth(bVal, it.birth)) {
          record(s, '실패', `참여자 생년월일(${bVal.trim()})이 파일(${it.birth})과 다름`); nextItem(s); save(s); return;
        }
        let btn = doc.getElementById('btn_deduAply');
        if (!btn || !isVisible(btn)) btn = findButton(doc, '차감신청', { exclude: ex });
        if (!btn || isDisabled(btn)) {   // 조회가 끝나 버튼이 켜질 때까지 대기 (시간 초과 시 실패 처리)
          if (s.lookOnly && Date.now() - s.searchedAt > 5000) {
            const can0 = valueAfterLabel(doc, '차감가능여부', null);
            record(s, '조회 확인', '환자 조회됨 · 차감신청 버튼 꺼져 있음' + (can0 ? ` · 차감가능여부: ${can0.trim()}` : ''));
            nextItem(s); save(s);
          }
          return;
        }

        // 차감가능여부 / 차감가능 포인트
        const can = valueAfterLabel(doc, '차감가능여부', null);
        if (can && /불가|아니|N\b/.test(can)) { record(s, '실패', `차감가능여부: ${can.trim()}`); nextItem(s); save(s); return; }
        const avail = toNum(valueAfterLabel(doc, '차감가능 포인트', null));
        if (s.lookOnly) {
          record(s, '조회 확인', `차감가능 ${avail == null ? '?' : avail.toLocaleString()}P${avail != null && avail < it.amount ? ' (부족)' : ''}`, avail);
          nextItem(s); save(s); return;
        }
        if (avail != null && avail < it.amount) {
          record(s, '포인트 부족', `차감가능 ${avail.toLocaleString()}P < 차감 ${it.amount.toLocaleString()}P`, avail);
          nextItem(s); save(s); return;
        }
        s.lastAlert = null; s.phase = 'popup'; s.phaseAt = Date.now(); s.popupFor = null;
        save(s);
        click(btn);
        return;
      }

      if (s.phase === 'submitted') {
        const since = s.submittedAt || s.phaseAt;
        if (s.lastAlert && s.lastAlertAt >= since) {
          const ok = /완료|정상|되었습니다|성공/.test(s.lastAlert) && !/실패|오류|불가|초과|부족|없습니다|확인하/.test(s.lastAlert);
          if (ok) record(s, '성공', '', s.expectRemain);
          else record(s, '실패', '차감신청: ' + s.lastAlert);
          if (popup) closePopup(doc, popup);
          nextItem(s); save(s); return;
        }
        // 알림 없이 창이 닫힌 경우
        if (Date.now() - since > 5000 && !popup && !anyPopupFrame(doc)) {
          record(s, '성공', '완료 메시지 없음 · 차감내역 확인 권장', s.expectRemain);
          nextItem(s); save(s); return;
        }
        return;
      }
    } finally { busy = false; renderPanel(doc); }
  }

  function anyPopupFrame(doc) {
    for (const f of doc.querySelectorAll('iframe,frame')) {
      try { if (f.contentDocument && findPopup(f.contentDocument)) return true; } catch (e) {}
    }
    return false;
  }

  // ---------------- '포인트 차감' 창 단계 ----------------
  let popupBusy = false;
  async function popupTick(doc) {
    if (popupBusy) return;
    const s0 = load();
    if (!s0 || !s0.running || s0.phase !== 'popup' || s0.popupFor === s0.idx) return;
    const popup = findPopup(doc);
    if (!popup) return;
    popupBusy = true;
    try {
      let s = load(); s.popupFor = s.idx; save(s);
      const it = s.items[s.idx];
      await sleep(400);

      // 진료일
      // 팝업에 표시된 환자 이름 확인 ('○○○님의 포인트 차감을 신청합니다')
      const head = (/([^\s]*)\s*님의/.exec(popup.innerText || '') || [])[1] || '';
      if (head && !nameRegex(it.name).test(head)) {
        s = load(); record(s, '실패', `차감 창의 환자(${head})가 파일과 다름`); closePopup(doc, popup); nextItem(s); save(s); return;
      }

      const dateIn = inputAfter(findLabel(popup, '진료일'), popup);
      if (!dateIn) { s = load(); record(s, '실패', '진료일 입력칸을 찾지 못함'); closePopup(doc, popup); nextItem(s); save(s); return; }
      // 달력에서 선택 → 안 되면 직접 입력
      const want8 = it.date.replace(/\D/g, '');
      let how = await pickDateByCalendar(doc, dateIn, it.date);
      if (dateIn.value.replace(/\D/g, '') !== want8) {
        setVal(dateIn, formatLike(dateIn.value, it.date), want8);
        how = '직접입력(' + how + ')';
      }
      s = load(); logLine(s, `${it.name}: 진료일 ${dateIn.value} (${how === 'OK' ? '달력 선택' : how})`); save(s);

      // 잔액 확인
      const before = toNum(valueAfterLabel(doc, '건강생활실천지원금제 잔액', popup));
      if (before != null && before < it.amount) {
        s = load(); record(s, '포인트 부족', `잔액 ${before.toLocaleString()}P < 차감 ${it.amount.toLocaleString()}P`, before);
        closePopup(doc, popup); nextItem(s); save(s); return;
      }

      // 차감 신청 포인트
      const amtIn = inputAfter(findLabel(popup, '차감 신청 포인트'), popup);
      if (!amtIn) { s = load(); record(s, '실패', '차감 신청 포인트 칸을 찾지 못함'); closePopup(doc, popup); nextItem(s); save(s); return; }
      setVal(amtIn, String(it.amount));
      await sleep(400);

      // 동의 체크
      const cbs = Array.from(popup.querySelectorAll('input[type=checkbox]')).filter(isVisible);
      const cb = cbs.find(c => /동의/.test((c.parentElement && c.parentElement.textContent) || '') || /동의/.test((c.labels && c.labels[0] && c.labels[0].textContent) || '')) || cbs[cbs.length - 1];
      if (!cb) { s = load(); record(s, '실패', '동의 체크박스를 찾지 못함'); closePopup(doc, popup); nextItem(s); save(s); return; }
      if (!cb.checked) click(cb);
      await sleep(200);

      // 입력값 재확인 (진료일, 금액)
      if (toNum(amtIn.value) !== it.amount) { s = load(); record(s, '실패', `금액 입력 확인 실패(${amtIn.value})`); closePopup(doc, popup); nextItem(s); save(s); return; }
      if (dateIn.value.replace(/\D/g, '') !== it.date.replace(/\D/g, '')) { s = load(); record(s, '실패', `진료일 입력 확인 실패(${dateIn.value})`); closePopup(doc, popup); nextItem(s); save(s); return; }

      let after = toNum(valueAfterLabel(doc, '차감 후 잔액', popup));
      if (after == null && before != null) after = before - it.amount;

      const btn = findButton(doc, '차감신청', { within: popup });
      s = load();
      s.expectRemain = after; s.phase = 'submitted'; s.phaseAt = Date.now(); s.submittedAt = Date.now(); s.lastAlert = null;
      save(s);
      click(btn);
    } finally { popupBusy = false; }
  }

  // ---------------- 화면 패널 ----------------
  function renderPanel(doc) {
    const p = doc.getElementById('wc-point-panel');
    if (!p) return;
    const s = load();
    const st = p.querySelector('#wc-status'), lg = p.querySelector('#wc-log');
    const bRun = p.querySelector('#wc-run'), bOne = p.querySelector('#wc-one'), bStop = p.querySelector('#wc-stop'), bDown = p.querySelector('#wc-down'), bLook = p.querySelector('#wc-look');
    if (!s) { st.textContent = '차감 대기 파일(CSV 또는 엑셀)을 넣어 주세요.'; lg.textContent = ''; bRun.disabled = bOne.disabled = bLook.disabled = true; bStop.disabled = true; bDown.disabled = true; return; }
    const done = Object.values(s.results);
    const cnt = k => done.filter(r => r.status === k).length;
    st.innerHTML = `대상 <b>${s.items.length}</b>건 · 성공 <b style="color:#1a7f37">${cnt('성공')}</b> · 실패 <b style="color:#c62828">${cnt('실패')}</b> · 포인트 부족 <b style="color:#b26a00">${cnt('포인트 부족')}</b>`
      + (s.lookOnly ? '<br>🔍 조회만 테스트 모드 (차감 안 함)' : '')
      + (s.running ? `<br>▶ 진행 중: ${Math.min(s.idx + 1, s.items.length)}/${s.items.length} ${s.items[s.idx] ? s.items[s.idx].name : ''}` : '')
      + (s.stopReason && !s.running ? `<br>${s.stopReason}` : '')
      + (s.finished ? '<br>✔ 완료 · [결과 파일 받기]를 눌러 주세요' : '');
    lg.textContent = s.log.slice().reverse().join('\n');
    bRun.disabled = s.running || s.idx >= s.items.length;
    bOne.disabled = s.running || s.idx >= s.items.length;
    bLook.disabled = s.running || s.items.length === 0;
    bStop.disabled = !s.running;
    bDown.disabled = done.length === 0;
  }

  function makePanel(doc) {
    if (doc.getElementById('wc-point-panel')) return;
    const p = doc.createElement('div');
    p.id = 'wc-point-panel';
    p.innerHTML = `
      <div style="font-weight:700;font-size:14px;margin-bottom:6px">건생포인트 자동 차감</div>
      <input type="file" id="wc-file" accept=".csv,.xlsx,.xls" style="font-size:12px;width:100%">
      <div id="wc-status" style="margin:8px 0;font-size:12px;line-height:1.6"></div>
      <div style="display:flex;gap:4px;flex-wrap:wrap">
        <button id="wc-look" title="성명 입력 → 조회 → 환자·차감가능 포인트 확인까지만 하고 차감은 하지 않아요">조회만 테스트</button>
        <button id="wc-one">1명만 테스트</button>
        <button id="wc-run" style="background:#e8620c;color:#fff;border-color:#e8620c">전체 자동 차감</button>
        <button id="wc-stop">중지</button>
        <button id="wc-down">결과 파일 받기</button>
        <button id="wc-reset" title="저장된 진행 정보 지우기">초기화</button>
      </div>
      <pre id="wc-log" style="margin-top:8px;max-height:220px;overflow:auto;background:#f6f6f6;padding:6px;font-size:11px;white-space:pre-wrap"></pre>`;
    p.style.cssText = 'position:fixed;top:12px;right:12px;width:340px;z-index:2147483647;background:#fff;border:2px solid #e8620c;border-radius:10px;padding:12px;box-shadow:0 6px 20px rgba(0,0,0,.2);font-family:"Malgun Gothic",sans-serif;color:#222';
    for (const b of p.querySelectorAll('button')) b.style.cssText += ';font-size:12px;padding:5px 8px;border:1px solid #bbb;border-radius:6px;background:' + (b.id === 'wc-run' ? '#e8620c' : '#fff') + ';cursor:pointer';
    doc.body.appendChild(p);

    p.querySelector('#wc-file').addEventListener('change', async e => {
      const f = e.target.files[0]; if (!f) return;
      try {
        const isXl = /\.xlsx?$/i.test(f.name);
        const s = isXl ? buildState('', false, await readExcelRows(f)) : buildState(await readFileText(f), false);
        logLine(s, `파일 불러옴: ${f.name} · 차감 대기 ${s.items.length}건`);
        save(s);
      } catch (err) { alert('파일을 읽지 못했어요: ' + err.message); }
      renderPanel(doc);
    });
    const start = testOne => {
      const s = load(); if (!s) return;
      const total = s.items.slice(s.idx).reduce((a, b) => a + b.amount, 0);
      const n = testOne ? 1 : s.items.length - s.idx;
      if (!W.confirm.call(W, `${n}건, 총 ${(testOne ? s.items[s.idx].amount : total).toLocaleString()}P를 실제로 차감 신청합니다.\n진행할까요?`)) return;
      if (s.lookOnly) endLook(s);
      s.testOne = testOne; s.running = true; s.finished = false; s.stopReason = null; s.failStreak = 0;
      s.phase = 'search'; s.phaseAt = Date.now(); s.popupFor = null;
      if (testOne) { s.testStart = s.idx; }
      logLine(s, testOne ? '▶ 1명 테스트 시작' : '▶ 전체 자동 차감 시작');
      save(s); renderPanel(doc);
    };
    p.querySelector('#wc-one').onclick = () => start(true);
    // 조회만 테스트: 처음부터 모든 대상자를 조회만 해 보고, 결과는 진행 기록에만 남김 (차감·결과 파일에는 영향 없음)
    p.querySelector('#wc-look').onclick = () => {
      const s = load(); if (!s || !s.items.length) return;
      s.lookBackup = { idx: s.idx, results: s.results, failStreak: s.failStreak };
      s.lookOnly = true; s.testOne = false; s.idx = 0; s.results = {}; s.running = true; s.finished = false; s.stopReason = null; s.failStreak = 0;
      s.phase = 'search'; s.phaseAt = Date.now(); s.popupFor = null;
      logLine(s, '🔍 조회만 테스트 시작 (차감 안 함)');
      save(s); renderPanel(doc);
    };
    p.querySelector('#wc-run').onclick = () => start(false);
    p.querySelector('#wc-stop').onclick = () => { const s = load(); if (!s) return; s.running = false; s.stopReason = '사용자가 중지함'; logLine(s, '■ 중지'); if (s.lookOnly) endLook(s); save(s); renderPanel(doc); };
    p.querySelector('#wc-down').onclick = () => { const s = load(); if (s) download(`차감결과_${today()}.csv`, buildResultCSV(s)); };
    p.querySelector('#wc-reset').onclick = () => { if (W.confirm.call(W, '불러온 파일과 진행 기록을 지울까요?\n(결과 파일을 먼저 받아 두세요)')) { clearState(); p.querySelector('#wc-file').value = ''; renderPanel(doc); } };
    renderPanel(doc);
  }

  // 포털 첫 화면 등: [포인트차감신청 열기] 버튼
  function menuShortcut(doc) {
    if (W.top !== W) return;
    const ex = doc.getElementById('wc-open-menu');
    let onForm = isMainForm(doc);
    for (const f of doc.querySelectorAll('iframe')) { try { if (f.contentDocument && isMainForm(f.contentDocument)) onForm = true; } catch (e) {} }
    if (onForm) { if (ex) ex.remove(); return; }
    if (ex) return;
    const b = doc.createElement('button');
    b.id = 'wc-open-menu'; b.textContent = '포인트차감신청 열기';
    b.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:2147483647;padding:10px 14px;border:0;border-radius:8px;background:#e8620c;color:#fff;font:700 13px "Malgun Gothic",sans-serif;cursor:pointer;box-shadow:0 4px 12px rgba(0,0,0,.25)';
    b.onclick = () => {
      const link = Array.from(doc.querySelectorAll('a')).find(a => norm(a.textContent) === '포인트차감신청');
      if (link) link.click(); else alert('메뉴에서 [포인트차감신청] 링크를 찾지 못했어요. 건강생활실천지원금제 메뉴에서 직접 열어 주세요.');
    };
    doc.body.appendChild(b);
  }

  // ---------------- 참여정보조회및신청 화면으로 넘어간 경우 ----------------
  // 포인트차감신청에서 이름으로 조회했는데 참여자를 바로 찾지 못하면 포털이 이 화면으로 넘겨 버림
  // → 결과를 기록하고 왼쪽 메뉴의 [포인트차감신청]으로 돌아가 다음 사람을 진행
  function isJoinInfoPage(doc) {
    return !isMainForm(doc) && !!findLabel(doc, '참여정보조회및신청') && !!findLabel(doc, '케어플랜 수립일');
  }
  let joinBusy = false;
  async function joinInfoTick(doc) {
    if (joinBusy) return;
    const s = load();
    if (!s || !s.running || (s.phase !== 'searched' && s.phase !== 'search')) return;
    const it = s.items[s.idx];
    if (!it) return;
    joinBusy = true;
    try {
      await sleep(800);
      const s2 = load();
      if (!s2 || !s2.running || s2.idx !== s.idx) return;
      const rows = rowsWithName(doc, it.name, [doc.getElementById('wc-point-panel')]);
      const noData = s2.lastAlert && /자료가\s*존재하지\s*않/.test(s2.lastAlert);
      if (noData || rows.length === 0) record(s2, '실패', '정보마당에 이 이름의 참여자가 없음 (참여정보조회 화면으로 넘어감 · 이름 확인 필요)');
      else record(s2, '실패', `참여정보조회 화면으로 넘어감 (같은 이름 ${rows.length}명) · 직접 확인 필요`);
      nextItem(s2); save(s2);
      backToDeductMenu(doc);
    } finally { joinBusy = false; }
  }
  // 왼쪽 업무메뉴의 [포인트차감신청] 누르기 (이 화면이 iframe 안이면 바깥 문서에서 찾음)
  function backToDeductMenu(doc) {
    const docs = [doc];
    try { if (W.parent && W.parent.document && W.parent.document !== doc) docs.push(W.parent.document); } catch (e) {}
    try { if (W.top && W.top.document && docs.indexOf(W.top.document) < 0) docs.push(W.top.document); } catch (e) {}
    for (const d of docs) {
      const link = Array.from(d.querySelectorAll('a,span,li,div')).find(a => isVisible(a) && norm(a.textContent) === '포인트차감신청' && !a.querySelector('a'));
      if (link) { click(link.closest('a') || link); return true; }
    }
    return false;
  }

  // ---------------- 시작 ----------------
  function boot() {
    hookDialogs();
    setInterval(() => {
      const doc = document;
      if (!doc.body) return;
      const s = load();
      // testOne 시작점 보정: 1명 테스트는 시작한 위치에서 1명만
      if (s && s.running && s.testOne && s.testStart != null && s.idx > s.testStart) { s.running = false; s.stopReason = '1명 테스트 완료'; save(s); }
      watchMutations(doc);
      handleMessageBoxes(doc);
      try { menuShortcut(doc); } catch (e) {}
      if (isMainForm(doc)) {
        const top = (() => { try { return W.top === W || !W.parent.document || !isMainForm(W.parent.document); } catch (e) { return true; } })();
        if (top) { makePanel(doc); mainTick(doc); renderPanel(doc); }
      }
      if (findPopup(doc)) popupTick(doc);
      try { if (isJoinInfoPage(doc)) joinInfoTick(doc); } catch (e) {}
    }, 400);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();