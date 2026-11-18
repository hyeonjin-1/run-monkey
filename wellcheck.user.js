// ==UserScript==
// @name         🐵 칼퇴몽키 – 웰체크용
// @namespace    wellcheck-hira-automation
// @version      11.3
// @description  클릭은 몽키가, 퇴근은 내가 · 웰체크 교육상담·포괄평가·점검평가 대기열 자동화(심평원 작성완료는 별도 스크립트). SPA 페이지 이동 감지 포함.
// @match        https://web.well-check.co.kr/*
// @updateURL    https://raw.githubusercontent.com/hyeonjin-1/run-monkey/main/wellcheck.user.js
// @downloadURL  https://raw.githubusercontent.com/hyeonjin-1/run-monkey/main/wellcheck.user.js
// @run-at       document-idle
// @grant        unsafeWindow
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// ==/UserScript==

(function () {
  'use strict';

  // 샌드박스 실행 시에도 페이지 쪽 창(confirm/alert, 프레임)을 다루기 위해 사용
  const PAGE = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
  const ADMIN = false; // 관리자용(점검·화면 변경 감지)은 true
  const GM_VERSION = (typeof GM_info !== 'undefined' && GM_info.script && GM_info.script.version) || '';

  // ---- 모드: 심평원까지 / 임시저장까지만 ----
  const MODE_KEY = 'kaltoe-mode';
  function saveOnly() { try { return localStorage.getItem(MODE_KEY) === 'save'; } catch (e) { return false; } }
  function setMode(m) {
    try { localStorage.setItem(MODE_KEY, m); } catch (e) {}
    applyModeUI();
  }
  function applyModeUI() {
    const so = saveOnly();
    [['wcq-m-hira', !so], ['wcq-m-save', so], ['wcp-m-hira', !so], ['wcp-m-save', so]].forEach(([id, on]) => {
      const b = document.getElementById(id);
      if (!b) return;
      b.style.background = on ? '#fff3e0' : '';
      b.style.fontWeight = on ? '700' : '';
      b.style.borderColor = on ? '#ef6c00' : '';
    });
  }
  const MODE_ROW = (p) => '<div style="display:flex;gap:4px;margin-bottom:4px">' +
    '<button id="' + p + '-m-hira" type="button" style="flex:1">심평원까지</button>' +
    '<button id="' + p + '-m-save" type="button" style="flex:1">임시저장까지만</button></div>';


  console.log('[자동화] 스크립트 로드됨:', location.href);

  // ---- 공용 유틸 ----

  // [변경] 글자 사이 공백 개수/종류가 달라도 찾도록 공백을 모두 지우고 비교
  function norm(t) {
    return (t || '').replace(/\s+/g, '');
  }

  function findByText(selectors, text) {
    for (const sel of selectors) {
      const el = Array.from(document.querySelectorAll(sel))
        .find((e) => norm(e.textContent) === norm(text));
      if (el) return el;
    }
    return null;
  }

  // [추가] 콘솔 대신 화면 왼쪽 아래에 현재 상태를 표시 (샌드박스에서는 콘솔 로그가 안 보일 수 있음)
  // 포괄평가 환자 화면 판별: '계획 수립 일자' 라벨 가까이에 날짜 칸이 있는 화면만 (교육상담 탭과 구분)
  function planDateButton(enabledOnly) {
    let btns = Array.from(document.querySelectorAll(DATE_SEL));
    if (enabledOnly) btns = btns.filter((b) => !b.disabled);
    for (const b of btns) {
      let el = b;
      for (let i = 0; i < 5 && el; i++, el = el.parentElement) {
        const t = norm(el.textContent);
        if (t.includes('계획수립일자') && t.length < 120) return b;
      }
    }
    return null;
  }
  // 날짜 칸 버튼: 개편 전(day-picker-input_date_picker_button)과 개편 후(date-picker_trigger) 둘 다 찾음
  const DATE_SEL = 'button[class*="date_picker_button"], button[class*="date-picker_trigger"]';
  // 교육상담 날짜 칸: '교육상담 반영내용' 창 안의 날짜 칸을 먼저 찾음 (화면의 다른 날짜 칸과 헷갈리지 않게)
  function eduDateBtn() {
    const d = openDialogs().find((x) => norm(x.textContent).includes('교육상담반영내용'));
    if (d) return d.querySelector(DATE_SEL);
    return document.querySelector('button[class*="date_picker_button"]'); // 개편 전 화면
  }

  // ---- 교육자료 체크칸 ----
  // 개편 후 화면에는 숨겨진 목록(추천 콘텐츠 등)에도 같은 모양의 체크칸이 있어서, 화면에 보이는 체크칸만 씀
  function isShown(el) {
    if (!el || el.getClientRects().length === 0) return false;
    if (el.closest('[hidden], [aria-hidden="true"]:not(body)') && !el.closest('[role="dialog"][aria-label="발송 바구니"]')) {
      // 모달이 열리면 다른 부분에 aria-hidden이 붙으므로, hidden 속성만 확실한 숨김으로 봄
      if (el.closest('[hidden]')) return false;
    }
    const cs = getComputedStyle(el);
    return cs.visibility !== 'hidden' && cs.display !== 'none';
  }
  function materialCheckboxes(checkedOnly) {
    let list = Array.from(document.querySelectorAll('td button[role="checkbox"]')).filter(isShown)
      .filter((b) => !b.closest('#wc-queue-panel, #wc-pcp-panel'));
    const basketOnes = list.filter((b) => /발송\s*바구니/.test(b.getAttribute('aria-label') || '')); // 개편 후 자료 체크칸 이름: "○○ 발송바구니 담기"
    if (basketOnes.length) list = basketOnes;
    return checkedOnly ? list.filter((b) => b.getAttribute('aria-checked') === 'true') : list;
  }

  // ---- 달력 도우미 ----
  // 개편 전 달력: td[data-day="YYYY-MM-DD"] 안의 날짜 버튼
  // 개편 후 달력: 날짜 버튼(rdp-day_button)의 이름이 "2026년 11월 4일 수요일"(오늘이면 앞에 "오늘, ") 형식 → 이 이름으로 찾음
  function calDayButtons() { return Array.from(document.querySelectorAll('button.rdp-day_button, button[class*="day_button"]')).filter((b) => b.getClientRects().length > 0); }
  function calReady() { return document.querySelector('.rdp-month_grid td[data-day]') || calDayButtons()[0] || null; }
  function calLabelYmd(b) {
    const m = (b.getAttribute('aria-label') || '').match(/(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/);
    return m ? m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0') : '';
  }
  function calDayButton(date) {
    const td = document.querySelector('.rdp-month_grid td[data-day="' + date + '"]');
    const old = td && td.querySelector('button.rdp-day_button, button');
    if (old) return old;
    return calDayButtons().find((b) => calLabelYmd(b) === date) || null;
  }
  function calDisabled(b) {
    const td = b.closest('td');
    return b.disabled || b.getAttribute('aria-disabled') === 'true' ||
      (td && (td.classList.contains('rdp-disabled') || td.getAttribute('data-disabled') === 'true' || td.getAttribute('aria-disabled') === 'true'));
  }
  function calShownYm() {
    const td = document.querySelector('.rdp-month_grid td.rdp-day[data-day]:not(.rdp-outside)') || document.querySelector('.rdp-month_grid td[data-day]');
    if (td) return td.getAttribute('data-day').slice(0, 7);
    const cnt = {};
    calDayButtons().forEach((b) => { const d = calLabelYmd(b); if (d) cnt[d.slice(0, 7)] = (cnt[d.slice(0, 7)] || 0) + 1; });
    const keys = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a]);
    return keys[0] || '';
  }
  function calNavButtons() { // 달력의 이전/다음 달 버튼
    const caption = document.querySelector('[class*="day-picker-modal_caption"]');
    const cb = caption ? Array.from(caption.querySelectorAll(':scope > button')) : [];
    let prev = cb.find((b) => b.className.includes('rotate-180'));
    let next = cb.find((b) => !b.className.includes('rotate-180'));
    if (prev && next) return { prev, next };
    const day = calDayButtons()[0];
    let box = day;
    for (let k = 0; k < 8 && box && box.parentElement; k++) {
      box = box.parentElement;
      if (Array.from(box.querySelectorAll('button')).filter((b) => !/day_button/.test(b.className)).length >= 2) break;
    }
    const nb = box ? Array.from(box.querySelectorAll('button')).filter((b) => !/day_button/.test(b.className) && b.getClientRects().length > 0) : [];
    const lab = (b) => (b.getAttribute('aria-label') || '') + ' ' + b.className + ' ' + (b.textContent || '');
    prev = document.querySelector('.rdp-button_previous') || nb.find((b) => /previous|prev|이전|rotate-180/i.test(lab(b))) || nb[0];
    next = document.querySelector('.rdp-button_next') || nb.find((b) => /next|다음/i.test(lab(b))) || nb[nb.length - 1];
    return { prev, next };
  }
  async function calGotoMonth(ym) {
    for (let i = 0; i < 60; i++) {
      const cur = calShownYm();
      if (!cur) return false;
      if (cur === ym) return true;
      const { prev, next } = calNavButtons();
      const target = cur > ym ? prev : next;
      if (!target || target.disabled) return false;
      target.click();
      await new Promise((r) => setTimeout(r, 200));
    }
    return false;
  }

  // ---- 열린 팝업 찾기 ----
  // 개편된 화면에는 role="dialog"이지만 팝업이 아닌 상시 요소(AI 추천 안내·메시지 창 등)가 있어서
  // 화면에 보이는 것만 보고, 닫으려 해도 안 닫히고 확인/취소 버튼도 없는 상시 요소는 이후 무시함
  const ignoredDialogs = new WeakSet();
  function openDialogs() {
    return Array.from(document.querySelectorAll('[role="dialog"], [role="alertdialog"]')).filter((d) => {
      if (ignoredDialogs.has(d)) return false;
      if (d.closest('#wc-queue-panel, #wc-pcp-panel')) return false;
      if (d.getAttribute('aria-hidden') === 'true') return false;
      const r = d.getBoundingClientRect();
      if (r.width < 2 || r.height < 2) return false;
      const cs = getComputedStyle(d);
      return cs.display !== 'none' && cs.visibility !== 'hidden' && cs.opacity !== '0';
    });
  }
  function ignoreStuckWidgets() { // 확인·취소·예·아니오 버튼이 없는 '팝업'은 화면 상시 요소로 보고 무시
    let n = 0;
    for (const d of openDialogs()) {
      const hasAnswer = Array.from(d.querySelectorAll('button')).some((b) => /^(확인|취소|예|아니오|닫기)$/.test((b.textContent || '').replace(/\s+/g, '')));
      if (!hasAnswer) { ignoredDialogs.add(d); n++; }
    }
    return n;
  }

  // 교육상담(내원) 화면의 환자 목록 상자. 예전 클래스(style_patient_list)가 없으면
  // 머리글에 '환자명·생년월일·작성완료'가 있는 표(대상자 목록 표 제외)를 찾아 그 목록 상자를 씀
  function eduListRoot() {
    const old = document.querySelector('[class*="style_patient_list"]');
    if (old) return old;
    const t = Array.from(document.querySelectorAll('table')).find((tb) => {
      const head = (tb.querySelector('thead') || tb.rows[0] || {}).textContent || '';
      const h = head.replace(/\s+/g, '');
      return h.includes('환자명') && h.includes('생년월일') && h.includes('작성완료') && !tb.querySelector('td[class*="assessment-table"]');
    });
    if (!t) return null;
    let el = t;
    for (let i = 0; i < 8 && el.parentElement; i++) {
      el = el.parentElement;
      const tx = (el.textContent || '').replace(/\s+/g, '');
      if (tx.includes('건씩보기') || tx.includes('관리목록')) return el;
    }
    return t.parentElement || t;
  }
  function isPcpPatientPage() { return !!planDateButton(false); }
  function isAssessPage() { return !!document.querySelector('form[class*="monitoring-assessment"]'); }

  let lastEvent = '대기';
  // 칼퇴몽키 상태 표시: 진행 중에는 🐵, 다 끝나면 🙈 (끝난 문구는 1분 동안 유지)
  let doneText = '';
  let doneUntil = 0;
  function badgeDone(text) {
    doneText = '🙈 칼퇴몽키 · ' + text;
    doneUntil = Date.now() + 60000;
    badge('');
  }
  function badgeReset() { doneUntil = 0; }
  function badge(text) {
    text = Date.now() < doneUntil ? doneText : '🐵 칼퇴몽키 · ' + text;
    try {
      let el = document.getElementById('wc-auto-badge');
      if (!el) {
        el = document.createElement('div');
        el.id = 'wc-auto-badge';
        el.style.cssText =
          'position:fixed;left:8px;bottom:8px;z-index:2147483647;background:rgba(30,30,30,.85);' +
          'color:#fff;font:12px/1.4 sans-serif;padding:4px 8px;border-radius:6px;' +
          'pointer-events:none;max-width:70vw;';
        (document.body || document.documentElement).appendChild(el);
      }
      el.textContent = text;
    } catch (e) {}
  }

  function realClick(el) {
    if (!el) return false;
    ['mousedown', 'mouseup', 'click'].forEach((type) => {
      const ev = new MouseEvent(type, { bubbles: true, cancelable: true });
      el.dispatchEvent(ev);
    });
    return true;
  }

  // [추가] "확인/취소" 팝업 안에 '발송' 문구가 있을 때만 확인 버튼을 돌려줌
  // (웰체크 전체에서 돌기 때문에, 삭제 등 다른 확인 팝업을 누르지 않도록 하는 안전장치)
  function findSendConfirmBtn() {
    const btns = Array.from(document.querySelectorAll('button'));
    const oks = btns.filter((b) => norm(b.textContent) === '확인');
    for (const ok of oks) {
      // 가장 가까운 팝업(dialog) 전체를 기준으로 문구 확인, 없으면 위로 몇 단계 올라가며 확인
      const dlg = ok.closest('[role="dialog"], [role="alertdialog"]');
      const boxes = [];
      if (dlg) {
        boxes.push(dlg);
      } else {
        let el = ok.parentElement;
        for (let i = 0; i < 6 && el; i++, el = el.parentElement) boxes.push(el);
      }
      for (const box of boxes) {
        const hasCancel = Array.from(box.querySelectorAll('button')).some(
          (b) => norm(b.textContent) === '취소'
        );
        if (hasCancel && norm(box.textContent).includes('발송하시겠습니까')) return ok;
      }
    }
    return null;
  }

  function confirmDialogOpen() {
    const texts = Array.from(document.querySelectorAll('button')).map((b) =>
      b.textContent.trim()
    );
    return texts.includes('확인') && texts.includes('취소');
  }

  // ============================================================
  // 1) 웰체크 교육상담 (edu-consult / patientEducation) 자동화
  // ============================================================
  function startEduConsult() {
    const MARK = 'wcAutoClicked';

    function tick() {
      if (isPcpPatientPage() || isAssessPage()) return; // 포괄평가·점검평가 화면에서는 교육상담 자동화를 쓰지 않음
      const dateBtn = eduDateBtn();
      // 날짜는 사용자가 직접 선택함 → 날짜가 채워졌는지(글자에 '--'가 없는지)만 확인
      const dateFilled = dateBtn && !dateBtn.textContent.includes('--');

      // 날짜가 선택되었고 확인 팝업이 아직 안 떠있을 때만 "발송" 클릭
      // 개편 후 화면 오른쪽 아래에도 '발송'(바구니) 버튼이 있어서, '교육상담 반영내용' 창 안의 '발송'을 먼저 찾음
      const rdlg = openDialogs().find((x) => norm(x.textContent).includes('교육상담반영내용'));
      const sendBtn = rdlg ? Array.from(rdlg.querySelectorAll('button')).find((b) => norm(b.textContent) === '발송') : findByText(['button'], '발송');
      if (dateFilled && sendBtn && !sendBtn.dataset[MARK] && !confirmDialogOpen()) {
        sendBtn.click();
        sendBtn.dataset[MARK] = 'true';
        console.log('[자동화] 날짜 선택 확인 → 발송 버튼 클릭 완료');
        lastEvent = '발송 클릭함';
      }

      const confirmBtn = findSendConfirmBtn();
      if (confirmBtn && !confirmBtn.dataset[MARK]) {
        confirmBtn.click();
        confirmBtn.dataset[MARK] = 'true';
        console.log('[자동화] 발송 확인 팝업 확인 클릭 완료');
        lastEvent = '발송 확인 클릭함';
      }

      const moveBtn = findByText(['button', 'div'], '심평원 이동');
      if (moveBtn && !moveBtn.dataset[MARK]) {
        moveBtn.dataset[MARK] = 'true';
        if (saveOnly()) {
          lastEvent = '발송 완료(심평원 이동 생략)';
        } else {
          moveBtn.click();
          console.log('[자동화] 심평원 이동 클릭 완료');
          lastEvent = '심평원 이동 클릭함';
        }
      }
    }

    const observer = new MutationObserver(tick);
    observer.observe(document.body, { childList: true, subtree: true });

    // [변경] 자료 선택 감지: 어디를 눌러 선택하든(체크박스/행/제목) 상관없이
    //        표 안의 체크된 자료 개수가 늘어나면 "교육자료 발송" 버튼을 한 번 자동 클릭
    let prevChecked = null;
    const pollId = setInterval(() => {
      const checkedNow = materialCheckboxes(true).length;

      badge(
        'v11.3 · 선택된 자료 ' + checkedNow + '개 · 교육자료 발송 버튼 ' +
        (findByText(['button'], '교육자료 발송') ? '있음' : '없음') + ' · ' + lastEvent
      );

      if (prevChecked === null) {
        prevChecked = checkedNow; // 처음 상태는 기준값으로만 저장
        return;
      }

      if (checkedNow > prevChecked) {
        const sendTabBtn = findByText(['button'], '교육자료 발송');
        if (sendTabBtn) {
          sendTabBtn.click();
          console.log('[자동화] 자료 선택 감지 → 교육자료 발송 버튼 자동 클릭 완료');
          lastEvent = '교육자료 발송 클릭함';
        } else {
          console.log('[자동화] 자료 선택은 감지했지만 교육자료 발송 버튼을 찾지 못함');
          lastEvent = '선택 감지, 발송 버튼 못 찾음';
        }
      }
      prevChecked = checkedNow;
      tick();
    }, 300);

    console.log('[자동화] 교육상담 자동화 시작됨');

    return function teardown() {
      observer.disconnect();
      clearInterval(pollId);
    };
  }

  // ============================================================
  // 3) 교육상담 대기열 패널
  //    목록(이름,생년월일,날짜)을 붙여넣고 "다음 건 실행"을 누르면
  //    환자 선택 → 자료 선택(즐겨찾기, 발송됨 제외) → 날짜 선택까지 자동으로 하고,
  //    날짜가 선택된 뒤는 위의 자동화(발송→확인→심평원 이동)가 이어받음
  // ============================================================
  function startQueuePanel() {
    const DATE_BTN = DATE_SEL;
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const digits = (s) => (s || '').replace(/\D/g, '');
    const pad2 = (n) => String(n).padStart(2, '0');

    async function waitFor(fn, timeout, interval) {
      const t0 = Date.now();
      while (Date.now() - t0 < (timeout || 8000)) {
        let v = null;
        try { v = fn(); } catch (e) {}
        if (v) return v;
        await sleep(interval || 150);
      }
      return null;
    }

    // ---- 입력 목록 해석 ----
    function normBirth(s) {
      const d = digits(s);
      if (d.length === 8) return d.slice(2); // YYYYMMDD → YYMMDD
      if (d.length === 6) return d;          // YYMMDD
      if (d.length === 13) return d.slice(0, 6); // 주민번호 형식이면 앞 6자리만 사용
      return '';
    }
    function normDate(s) {
      const t = (s || '').trim();
      let m = t.match(/^(\d{4})(\d{2})(\d{2})$/);
      if (!m) m = t.match(/^(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
      if (!m) return '';
      return m[1] + '-' + pad2(m[2]) + '-' + pad2(m[3]);
    }
    // ---- 처방통계 엑셀 → 업무별 명단 ----
    // 파일은 이 브라우저 안에서만 읽고 어디로도 보내지 않음
    // 머리글 이름은 차트사마다 달라서 여러 이름을 함께 찾음 (새 양식은 여기에 이름만 추가)
    const XL_HEAD = {
      name: ['환자명', '수진자명', '수진자', '성명', '이름', '환자성명'],
      birth: ['주민등록번호', '주민번호', '생년월일', '생일'],
      date: ['처방일자', '처방일', '진료일자', '진료일', '내원일자', '내원일', '접수일자'],
      code: ['코드', '처방코드', '청구코드', '수가코드', '행위코드', '처방명코드'],
    };
    const XL_TASK = { AX001: 'pcp', AX002: 'pcp', AX004: 'edu', AX003: 'assess' };
    function xlBirth6(v) {
      const d = digits(String(v || ''));
      if (/^\d{4}[-./]/.test(String(v || '').trim()) && d.length >= 8) return d.slice(2, 8); // 1959-04-22 형식
      return d.length >= 6 ? d.slice(0, 6) : '';
    }
    function xlDate(v) {
      if (v instanceof Date && !isNaN(v)) return v.getFullYear() + '-' + pad2(v.getMonth() + 1) + '-' + pad2(v.getDate());
      if (typeof v === 'number' && v > 20000 && v < 80000) { // 엑셀 날짜 숫자
        const dt = new Date(Math.round((v - 25569) * 86400000));
        return dt.getUTCFullYear() + '-' + pad2(dt.getUTCMonth() + 1) + '-' + pad2(dt.getUTCDate());
      }
      return normDate(String(v || ''));
    }
    function xlFindHeader(rows) {
      for (let r = 0; r < Math.min(rows.length, 30); r++) {
        const cells = rows[r].map((c) => norm(String(c || '')));
        const col = {};
        for (const k of Object.keys(XL_HEAD)) {
          let i = -1;
          for (const h of XL_HEAD[k]) { i = cells.indexOf(h); if (i >= 0) break; }
          col[k] = i;
        }
        if (col.name >= 0 && col.date >= 0) return { row: r, col }; // 생년월일·코드 열은 없어도 됨
      }
      return null;
    }
    async function readPrescriptionExcel(file, codeChoice) {
      if (typeof XLSX === 'undefined') throw new Error('엑셀 읽기 도구를 불러오지 못했어요. 인터넷 연결을 확인하고 새로고침해 주세요.');
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
      let rows = null, hd = null;
      for (const sn of wb.SheetNames) {
        const r = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: '' });
        const h = xlFindHeader(r);
        if (h) { rows = r; hd = h; break; }
      }
      if (!hd) throw new Error('엑셀에서 환자명·처방일자(내원일) 열을 찾지 못했어요. 이 양식의 머리글 이름을 알려 주세요.');
      // 코드 열 고르기: '코드'가 들어간 열이 여러 개일 수 있어서(진료과 코드·처방코드 등) 실제로 AX001~AX004 값이 가장 많은 열을 씀
      {
        const body = rows.slice(hd.row + 1);
        const axCount = (ci) => body.reduce((n, r) => n + (/AX00[1-4]/i.test(String(r[ci] || '')) ? 1 : 0), 0);
        const width = rows[hd.row].length;
        let best = -1, bestN = 0;
        for (let ci = 0; ci < width; ci++) {
          const n = axCount(ci);
          if (n > bestN) { best = ci; bestN = n; }
        }
        if (best >= 0) hd.col.code = best; // 머리글 이름과 상관없이 AX 코드가 들어 있는 열
        else if (hd.col.code >= 0) { hd.codeEmpty = String(rows[hd.row][hd.col.code] || '코드'); hd.col.code = -1; } // 코드 열에 AX 코드가 없으면 코드 열이 없는 파일처럼 다룸
      }
      // 코드 열이 없는 파일(코드 하나만 골라서 뽑은 파일): 고른 코드 → 파일 이름·제목 줄 순서로 코드를 정함
      let fileCode = null;
      if (hd.col.code < 0) {
        // 처방코드가 없는 파일(코드 하나만 골라서 뽑은 파일): 지금 선택된 탭의 업무로 넣음
        if (!codeChoice) throw new Error('이 엑셀에는 처방코드가 없어요. 넣을 업무 탭(포괄평가·교육상담·점검평가)을 먼저 고른 뒤 불러와 주세요.');
        fileCode = codeChoice;
      }
      const noBirthCol = hd.col.birth < 0;
      const out = { pcp: [], edu: [], assess: [] };
      const seen = new Set();
      const skip = { noName: 0, badBirth: 0, badDate: 0, otherCode: 0, dup: 0 };
      for (let r = hd.row + 1; r < rows.length; r++) {
        const row = rows[r];
        const codeM = fileCode ? [fileCode] : String(row[hd.col.code] || '').toUpperCase().match(/AX00[1-4]/);
        if (!codeM) { if (row.some((c) => String(c).trim())) skip.otherCode++; continue; }
        const name = String(row[hd.col.name] || '').trim();
        if (/^(합계|총계|소계|계)$/.test(norm(name))) continue;
        const birth6 = noBirthCol ? '?' : xlBirth6(row[hd.col.birth]); // '?' = 웰체크 대상자 목록에서 이름으로 찾아 채움
        const date = xlDate(row[hd.col.date]);
        if (!name) { if (row.some((c) => String(c).trim())) skip.noName++; continue; }
        if (!birth6) { skip.badBirth++; continue; }
        if (!date) { skip.badDate++; continue; }
        const t = XL_TASK[codeM[0]];
        const key = t + '|' + name + '|' + birth6 + '|' + date;
        if (seen.has(key)) { skip.dup++; continue; }
        seen.add(key);
        out[t].push({ name, birth6, date });
      }
      for (const t of Object.keys(out)) out[t].sort((a, b) => (a.name + a.birth6 + a.date < b.name + b.birth6 + b.date ? -1 : 1));
      const toText = (list) => list.map((x) => x.name + ',' + x.birth6 + ',' + x.date).join('\n');
      const notes = [];
      if (skip.noName) notes.push('건너뜀: 이름 없음 ' + skip.noName + '줄');
      if (skip.badBirth) notes.push('생년월일 읽기 실패 ' + skip.badBirth + '줄');
      if (skip.badDate) notes.push('날짜 읽기 실패 ' + skip.badDate + '줄');
      if (skip.dup) notes.push('중복 ' + skip.dup + '줄');
      if (fileCode) notes.unshift('처방코드가 없어 지금 선택한 ' + ({ pcp: '포괄평가', edu: '교육상담', assess: '점검평가' })[XL_TASK[fileCode]] + ' 탭에 넣음');
      if (!out.pcp.length && !out.edu.length && !out.assess.length && skip.otherCode) notes.unshift('AX001~AX004가 아닌 줄 ' + skip.otherCode + '줄은 뺐어요');
      if (noBirthCol) notes.unshift('생년월일 없음 → 웰체크에서 이름으로 찾음(동명이인은 확인 필요로 남김)');
      const summary = '포괄평가 ' + out.pcp.length + '건 · 교육상담 ' + out.edu.length + '건 · 점검평가 ' + out.assess.length + '건' +
        (notes.length ? ' (' + notes.join(', ') + ')' : '');
      const only = fileCode ? XL_TASK[fileCode] : null; // 코드 하나짜리 파일이면 그 업무 명단만 바꿈
      return { pcp: toText(out.pcp), edu: toText(out.edu), assess: toText(out.assess), summary, only };
    }
    // 두 패널이 함께 쓰는 엑셀 불러오기: 세 대기열을 한 번에 채움
    async function importExcel(file, say, codeChoice) {
      try {
        if (running || (pcpShared.isRunning && pcpShared.isRunning())) { say('실행 중에는 엑셀을 불러올 수 없어요.'); return; }
        say('엑셀 읽는 중…', '#1565c0');
        const r = await readPrescriptionExcel(file, codeChoice);
        if (!r.only || r.only === 'edu') { q('#wcq-text').value = r.edu; items = parseQueue(r.edu); render(); }
        if (pcpShared.loadLists) pcpShared.loadLists(!r.only || r.only === 'pcp' ? r.pcp : null, !r.only || r.only === 'assess' ? r.assess : null);
        say('🐵 엑셀에서 명단을 만들었어요 · ' + r.summary, '#2e7d32');
      } catch (e) {
        console.error('[자동화] 엑셀 불러오기 실패:', e);
        say('엑셀 불러오기 실패: ' + (e && e.message ? e.message : e));
      }
    }

    function parseQueue(text) {
      const out = [];
      text.split(/\r?\n/).forEach((line) => {
        // 쉼표/탭이 있으면 그걸로, 없으면 공백으로 나눔 (예: "홍길동 960824 2026-10-05")
        const sep = /[\t,]/.test(line) ? /[\t,]/ : /\s+/;
        const parts = line.split(sep).map((x) => x.trim()).filter((x) => x !== '');
        if (parts.length < 3) return;
        const noBirth = parts[1] === '?'; // 엑셀에 생년월일이 없던 줄
        const birth6 = noBirth ? '' : normBirth(parts[1]);
        if (!birth6 && !noBirth) return; // 머리글 줄 등은 건너뜀
        parts.slice(2).forEach((d) => {
          const date = normDate(d);
          out.push({ name: parts[0], birth6, date, status: date ? '대기' : '날짜 형식 오류' });
        });
      });
      return out;
    }

    // ---- 화면 읽기 ----
    function patientRows() { const r = eduListRoot(); return r ? Array.from(r.querySelectorAll('table tbody tr')) : []; }
    function rowInfo(tr) {
      const tds = tr.querySelectorAll('td');
      return {
        name: ((tds[1] && tds[1].textContent) || '').trim(),
        birth6: digits(((tds[2] && tds[2].textContent) || '').split('/')[0]),
      };
    }
    function isSame(tr, it) {
      const i = rowInfo(tr);
      return i.name === it.name && i.birth6 === it.birth6;
    }
    // 목록에서 선택된 줄인지: 개편으로 클래스 이름이 바뀔 수 있어서 여러 표시를 함께 봄
    function rowSelected(tr) {
      if (!tr) return false;
      if (/selected|active|current/i.test(tr.className || '')) return true;
      if (tr.getAttribute('aria-selected') === 'true' || tr.getAttribute('data-selected') === 'true') return true;
      const st = tr.getAttribute('data-state') || '';
      return st === 'selected' || st === 'active' || st === 'checked';
    }
    // 오른쪽에 그 환자 화면이 열렸는지: 머리줄 생년월일 + '○○○님' 표시(메시지 창 제목 등)
    function patientOpened(it) {
      return headerBirth6() === it.birth6 && (document.body.innerText || '').indexOf(it.name + '님') >= 0;
    }
    function headerBirth6() {
      const m = document.body.innerText.match(/(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일\s*\(\d+세\)/);
      return m ? m[1].slice(2) + pad2(m[2]) + pad2(m[3]) : '';
    }

    // ---- 환자 검색(목록에 없을 때만 사용) ----
    function searchInput() {
      const r = eduListRoot();
      return document.querySelector('[class*="style_search_container"] input') ||
        (r && r.querySelector('input[placeholder*="환자명"]')) || null;
    }
    function setInputValue(input, v) {
      const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
      setter.call(input, v);
      input.dispatchEvent(new Event('input', { bubbles: true }));
    }
    // 검색창에 값을 넣고 엔터(와 돋보기 버튼)로 검색 실행
    async function pressEnterSearch(input, v) {
      setInputValue(input, v);
      ['keydown', 'keypress', 'keyup'].forEach((type) => {
        input.dispatchEvent(new KeyboardEvent(type, { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true }));
      });
      const holder = input.closest('div');
      const scope = (holder && holder.parentElement) || holder;
      const icon = scope && Array.from(scope.querySelectorAll('button')).find((b) => !(b.textContent || '').trim() && b.querySelector('svg, img'));
      if (icon) icon.click();
      await sleep(1500);
    }
    // 쪽 번호 버튼 찾기: 숫자만 적힌 버튼이 한 묶음에 2개 이상 있는 곳
    function pageNumberButtons(scope) {
      const root = scope || document;
      const els = Array.from(root.querySelectorAll('button, a, [role="button"]')).filter(
        (e) => /^\d{1,3}$/.test((e.textContent || '').trim()) && !e.closest('table')
      );
      const groups = new Map();
      els.forEach((e) => {
        const key = e.closest('ul, ol, nav') || e.parentElement;
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push(e);
      });
      let best = [];
      groups.forEach((arr) => { if (arr.length >= 2 && arr.length > best.length) best = arr; });
      return best;
    }
    function isActivePage(el) {
      const cls = (el.className || '') + ' ' + ((el.parentElement && el.parentElement.className) || '');
      return el.getAttribute('aria-current') === 'page' || el.getAttribute('aria-current') === 'true' ||
        el.getAttribute('aria-selected') === 'true' || !!el.disabled ||
        /(^|[\s_-])(active|selected|current|on)([\s_-]|$)/i.test(cls);
    }
    // 화살표(처음 « / 이전 ‹ / 다음 › / 마지막 ») 버튼인지 판별: 글자가 없는 아이콘 버튼이거나 화살표 글자
    function isArrow(e) {
      const t = (e.textContent || '').trim();
      const label = ((e.getAttribute('aria-label') || '') + ' ' + (e.getAttribute('title') || '')).toLowerCase();
      if (/next|prev|first|last|다음|이전|처음|마지막/.test(label)) return true;
      if (/^[‹›«»<>≪≫＜＞]+$/.test(t)) return true;
      return !t && !!e.querySelector('svg, img');
    }
    // 쪽 이동 묶음 분석: 숫자 버튼들과 그 앞뒤의 화살표
    function paginationParts(scope) {
      const nums = pageNumberButtons(scope);
      if (!nums.length) return null;
      let box = nums[0].closest('ul, ol, nav') || nums[0].parentElement;
      let arrows = [];
      for (let i = 0; i < 4 && box; i++, box = box.parentElement) {
        arrows = Array.from(box.querySelectorAll('button, a, [role="button"]')).filter(isArrow);
        if (arrows.length) break;
      }
      const lastNum = nums[nums.length - 1];
      const before = arrows.filter((a) => nums[0].compareDocumentPosition(a) & Node.DOCUMENT_POSITION_PRECEDING);
      const after = arrows.filter((a) => lastNum.compareDocumentPosition(a) & Node.DOCUMENT_POSITION_FOLLOWING);
      return {
        nums,
        first: before.length >= 2 ? before[0] : null,
        next: after.length ? after[0] : null,
      };
    }
    const btnOff = (b) => !b || b.disabled || b.getAttribute('aria-disabled') === 'true';
    async function clickAndWait(btn, getInfos) {
      const sig = () => getInfos().map((r) => r.name + r.birth6).join('|');
      const before = sig();
      btn.click();
      const changed = await waitFor(() => sig() !== before, 3500, 150);
      await sleep(300);
      return !!changed;
    }
    // 첫 쪽으로 이동
    async function goFirstPage(scope, getInfos) {
      const parts = paginationParts(scope);
      if (!parts) return;
      const act = parts.nums.find(isActivePage);
      if (act && act.textContent.trim() === '1') return;
      const one = parts.nums.find((b) => b.textContent.trim() === '1');
      const btn = one || parts.first;
      if (btnOff(btn)) return;
      await clickAndWait(btn, getInfos);
    }
    // 다음 쪽으로 이동: 현재 쪽 번호+1 이 보이면 그걸, 아니면 다음(›) 화살표. 끝이면 false
    async function goNextPage(scope, getInfos) {
      const parts = paginationParts(scope);
      if (!parts) return false;
      const act = parts.nums.find(isActivePage);
      const cur = act ? parseInt(act.textContent, 10) : NaN;
      let target = isNaN(cur) ? null : parts.nums.find((b) => parseInt(b.textContent, 10) === cur + 1);
      if (!target) target = parts.next;
      if (btnOff(target)) return false;
      return clickAndWait(target, getInfos);
    }
    // 첫 쪽부터 마지막 쪽까지(최대 40쪽) 한 쪽씩 넘기며 환자 찾기 (getInfos: [{name, birth6}] 반환)
    async function scanAllPages(scope, getInfos, match) {
      // 검색을 비운 직후에는 목록과 쪽 번호가 다시 그려질 때까지 기다림
      await waitFor(() => getInfos().length > 0, 5000, 200);
      await waitFor(() => paginationParts(scope), 3000, 200);
      if (!paginationParts(scope)) return { found: false, tried: 0, pages: 0 };
      await goFirstPage(scope, getInfos);
      const seen = new Set();
      let tried = 0;
      for (let i = 0; i < 40; i++) {
        await sleep(200);
        const sig = getInfos().map((r) => r.name + r.birth6).join('|');
        if (!sig) { await sleep(800); continue; } // 목록이 아직 안 떴으면 잠깐 더 기다림
        if (seen.has(sig)) break; // 같은 화면이 반복되면 끝
        seen.add(sig);
        tried++;
        if (getInfos().some(match)) return { found: true, tried, pages: tried };
        if (!(await goNextPage(scope, getInfos))) break;
      }
      return { found: false, tried, pages: tried };
    }

    async function clearSearch() {
      const input = searchInput();
      if (input && input.value) await pressEnterSearch(input, '');
    }
    async function searchPatient(name) {
      const input = searchInput();
      if (!input) return false;
      await pressEnterSearch(input, name);
      return true;
    }

    // ---- 자료 선택 ----
    async function ensureFavoritesOnly() {
      // "즐겨찾기만" 토글: button[aria-pressed] (표 안의 행별 즐겨찾기 토글은 제외)
      const sw = Array.from(document.querySelectorAll('button[aria-pressed]')).find(
        (b) =>
          !b.closest('td') &&
          (norm(b.textContent).includes('즐겨찾기만') ||
            norm((b.parentElement || {}).textContent).includes('즐겨찾기만'))
      );
      if (!sw) return false;
      if (sw.getAttribute('aria-pressed') !== 'true') {
        sw.click();
        await sleep(800);
      }
      return true;
    }
    function findMaterial(used) {
      const cbs = materialCheckboxes(false);
      return cbs.find((cb) => {
        if (cb.getAttribute('aria-checked') === 'true') return false;
        if (used && used.has(cb.getAttribute('aria-label'))) return false; // 같은 환자에게 방금 보낸 자료 제외
        const tr = cb.closest('tr');
        return !(tr && tr.textContent.includes('발송됨'));
      }) || null;
    }

    // ---- 달력 ----
    function shownYm() { return calShownYm(); }
    async function gotoMonth(ym) { return calGotoMonth(ym); }

    // 실패 원인 파악용: 지금 화면에 떠 있는 짧은 팝업/알림 문구
    function screenNote() {
      try {
        const clean = (t) => (t || '').replace(/\s+/g, ' ').trim();
        const parts = [];
        const dtexts = openDialogs()
          .map((d) => clean(d.innerText || d.textContent))
          .filter(Boolean)
          .sort((a, b) => a.length - b.length);
        if (dtexts[0] && dtexts[0].length <= 120) parts.push('팝업: ' + dtexts[0]);
        const toasts = Array.from(document.querySelectorAll('[role="alert"], [role="status"], [class*="toast" i]'))
          .map((e) => clean(e.innerText || e.textContent))
          .filter((t) => t && t.length <= 120);
        if (toasts[0]) parts.push('알림: ' + toasts[0]);
        return parts.length ? ' · 화면 문구 [' + parts.join(' / ') + ']' : '';
      } catch (e) { return ''; }
    }
    // 화면에 보이는 첫 번째 기간(관리주기로 추정)과 비교한 설명
    function cycleNote(date) {
      try {
        const m = (document.body.innerText || '').match(/(\d{4}-\d{2}-\d{2})\s*[~\-–]\s*(\d{4}-\d{2}-\d{2})/);
        if (!m) return '';
        if (date < m[1]) return ' · 화면에 보이는 기간(' + m[1] + '~' + m[2] + ') 시작일보다 이른 날짜';
        if (date > m[2]) return ' · 화면에 보이는 기간(' + m[1] + '~' + m[2] + ') 종료일보다 늦은 날짜';
        return '';
      } catch (e) { return ''; }
    }

    // ---- 한 건 실행 ----
    async function runItem(it, dry) {
      if (openDialogs()[0]) {
        throw new Error('열려 있는 팝업이 있어요. 닫고 다시 눌러 주세요');
      }
      const now = new Date();
      const todayStr = now.getFullYear() + '-' + pad2(now.getMonth() + 1) + '-' + pad2(now.getDate());
      if (it.date > todayStr) throw new Error('오늘(' + todayStr + ') 이후 날짜라 선택할 수 없음');
      step('1/5 환자 찾는 중…');
      if (!it.birth6) {
        // 엑셀에 생년월일이 없던 줄: 전체 대상자 목록에서 이름으로 찾아 한 명일 때만 진행
        step('1/5 생년월일 찾는 중… (전체 대상자 목록에서 이름 검색)');
        const allMenu = findByText(['a', 'button', 'span', 'li', 'div'], '일만사대상자목록');
        if (!allMenu || !pcpShared.resolveBirth) throw new Error('생년월일이 없어 전체 대상자 목록에서 찾아야 하는데 메뉴를 찾지 못함');
        if (!pcpShared.onListPage()) { allMenu.click(); if (!(await waitFor(() => pcpShared.onListPage(), 12000, 300))) throw new Error('전체 대상자 목록으로 이동하지 못함'); await sleep(600); }
        await pcpShared.resolveBirth(it);
        render();
      }
      // 이전 건(전체 목록에서 연 환자 화면 등)이 남아 있으면 교육상담(내원) 목록으로 먼저 돌아옴
      if (!eduListRoot()) {
        const eduMenu = findByText(['a', 'button', 'span', 'li', 'div'], '교육상담(내원)');
        if (eduMenu) { eduMenu.click(); await waitFor(() => eduListRoot(), 10000, 300); await sleep(500); }
      }
      await clearSearch();

      // 1. 환자 선택 (이름 + 생년월일이 모두 같은 행만)
      let rows = patientRows().filter((tr) => isSame(tr, it));
      const preAll = patientRows().map(rowInfo); // 검색 전 목록(힌트용)
      let scanNote = '';
      if (rows.length === 0 && (await searchPatient(it.name))) {
        rows = patientRows().filter((tr) => isSame(tr, it));
      }
      if (rows.length === 0) {
        // 검색으로도 안 나오면 필터를 풀고 다른 쪽을 차례로 확인
        await clearSearch();
        const sc = await scanAllPages(
          eduListRoot(),
          () => patientRows().map(rowInfo),
          (i) => i.name === it.name && i.birth6 === it.birth6
        );
        rows = patientRows().filter((tr) => isSame(tr, it));
        scanNote = sc.pages ? ' · ' + sc.tried + '쪽까지 확인함' : ' · 쪽 이동 버튼을 찾지 못해 현재 쪽만 확인함';
      }
      let viaAll = false;
      if (rows.length === 0 && pcpShared.findPatientRow) {
        // 교육상담(내원) 목록에 없으면 '일만사 대상자 목록'(전체)에서 찾아 그 환자의 교육상담 탭으로 들어감
        step('1/5 교육상담 목록에 없어 전체 대상자 목록에서 찾는 중…');
        const allMenu = findByText(['a', 'button', 'span', 'li', 'div'], '일만사대상자목록');
        if (!allMenu) throw new Error('목록에서 환자를 찾지 못함' + scanNote + ' · 전체 대상자 목록 메뉴("일만사 대상자 목록")도 찾지 못함');
        allMenu.click();
        if (!(await waitFor(() => pcpShared.onListPage(), 12000, 300))) throw new Error('목록에서 환자를 찾지 못함' + scanNote + ' · 전체 대상자 목록으로 이동하지 못함');
        await sleep(600);
        let prow;
        try { prow = await pcpShared.findPatientRow(it); } catch (e) {
          throw new Error('교육상담 목록과 전체 대상자 목록 모두에서 환자를 찾지 못함 (' + (e && e.message ? e.message : e) + ')' + scanNote);
        }
        prow.nameCell.click();
        let opened = await waitFor(() => headerBirth6() === it.birth6 && (document.body.innerText || '').indexOf(it.name) >= 0, 8000, 300);
        if (!opened) {
          const inner = prow.nameCell.querySelector('a, button, span');
          if (inner) { inner.click(); opened = await waitFor(() => headerBirth6() === it.birth6, 8000, 300); }
        }
        if (!opened) throw new Error('전체 대상자 목록에서 이름을 눌러도 환자 화면으로 이동하지 못함(또는 환자 정보가 다름)');
        await sleep(600);
        const tabs = Array.from(document.querySelectorAll('button, a, [role="tab"], li, div, span')).filter((e) => e.children.length <= 3 && norm(e.textContent) === '교육상담');
        const eduTab = tabs.filter((e) => !tabs.some((o) => o !== e && e.contains(o)))[0];
        if (!eduTab) throw new Error('환자 화면에서 "교육 상담" 탭을 찾지 못함');
        eduTab.click();
        if (!(await waitFor(() => materialCheckboxes(false)[0], 12000, 300))) throw new Error('환자의 교육상담 화면(자료 목록)이 열리지 않음');
        viaAll = true;
        scanNote += ' · 전체 대상자 목록에서 찾음';
      }
      if (rows.length === 0 && !viaAll) {
        const sameBirth = preAll.filter((i) => i.birth6 === it.birth6).map((i) => i.name);
        const sameName = patientRows().map(rowInfo).concat(preAll).filter((i) => i.name === it.name).map((i) => i.birth6);
        let hint = '';
        if (sameBirth.length) hint += ' · 같은 생년월일의 이름: ' + Array.from(new Set(sameBirth)).join(', ');
        if (sameName.length) hint += ' · 같은 이름의 생년월일: ' + Array.from(new Set(sameName)).join(', ');
        if (!hint && preAll.length >= 20) hint = ' · 목록에 ' + preAll.length + '명만 보임(다음 쪽 또는 다른 목록에 있을 수 있음)';
        throw new Error('목록에서 환자를 찾지 못함' + hint + scanNote);
      }
      if (rows.length > 1) throw new Error('이름·생년월일이 같은 환자가 2명 이상');
      if (!viaAll) {
        if (!rowSelected(rows[0])) {
          const tds = rows[0].querySelectorAll('td');
          (tds[1] || rows[0]).click();
        }
        const picked = await waitFor(() => {
          const r = patientRows().find((tr) => isSame(tr, it));
          return r && headerBirth6() === it.birth6 && (rowSelected(r) || patientOpened(it));
        }, 10000);
        if (!picked) throw new Error('환자 선택 확인 실패(선택 표시 또는 생년월일 불일치)');
      } else if (headerBirth6() !== it.birth6) {
        throw new Error('환자 선택 확인 실패(생년월일 불일치)');
      }
      await sleep(800);

      // 2. 자료 선택 (즐겨찾기 + 발송됨 제외, 위에서부터 첫 번째)
      step('2/5 자료 고르는 중…');
      const favOk = await ensureFavoritesOnly();
      const leftover = materialCheckboxes(true);
      for (const b of leftover) { b.click(); await sleep(250); }
      const key = it.name + '|' + it.birth6;
      const used = usedByPatient.get(key) || new Set();
      const cb = await waitFor(() => findMaterial(used), 6000);
      if (!cb) throw new Error('선택할 수 있는 자료가 없음(발송됨 제외)');
      const title = ((cb.closest('tr') || {}).textContent || '').trim().slice(0, 20);
      const label = cb.getAttribute('aria-label') || title;
      // 개편 후: 사람이 누를 때처럼 마우스·포인터 신호를 모두 보내야 바구니에 담기고 바구니가 저절로 열림
      try { cb.scrollIntoView({ block: 'center' }); } catch (e) {}
      const humanClick = (el) => {
        const r = el.getBoundingClientRect();
        const o = { bubbles: true, cancelable: true, clientX: r.left + r.width / 2, clientY: r.top + r.height / 2, button: 0, buttons: 1, pointerId: 1, pointerType: 'mouse', isPrimary: true };
        ['pointerover', 'pointerenter', 'mouseover', 'mouseenter', 'pointermove', 'mousemove'].forEach((t) => el.dispatchEvent(t.startsWith('pointer') ? new PointerEvent(t, o) : new MouseEvent(t, o)));
        el.dispatchEvent(new PointerEvent('pointerdown', o));
        el.dispatchEvent(new MouseEvent('mousedown', o));
        try { el.focus(); } catch (e) {}
        el.dispatchEvent(new PointerEvent('pointerup', Object.assign({}, o, { buttons: 0 })));
        el.dispatchEvent(new MouseEvent('mouseup', Object.assign({}, o, { buttons: 0 })));
        el.dispatchEvent(new MouseEvent('click', Object.assign({}, o, { buttons: 0 })));
      };
      humanClick(cb); // 위쪽 자동화가 바구니의 "교육자료 발송"을 눌러 날짜 창을 열어 줌
      if (!(await waitFor(() => cb.getAttribute('aria-checked') === 'true', 1500, 100))) { cb.click(); await sleep(300); }
      if (cb.getAttribute('aria-checked') !== 'true') throw new Error('자료 체크칸을 눌렀는데 체크되지 않음 (' + label.slice(0, 20) + ')');

      // 3. 날짜 창 대기
      step('3/5 날짜 창 기다리는 중…');
      let dateBtn = await waitFor(() => eduDateBtn(), 3000);
      if (!dateBtn) {
        // 개편 후: 자료를 고르면 오른쪽 아래 '발송 바구니'가 뜨고 그 안의 '교육자료 발송'을 눌러야 함.
        // 바구니가 늦게 뜨거나 숨겨진 버튼이 먼저 잡힐 수 있어서, 화면에 보이는 버튼이 나올 때까지 기다렸다가 누름
        const visible = (el) => !!el && el.getClientRects().length > 0 && !el.closest('[hidden]') && !el.disabled;
        const findSendTab = () => Array.from(document.querySelectorAll('button')).find((b) => norm(b.textContent) === '교육자료발송' && visible(b));
        let sendTab = await waitFor(findSendTab, 2500, 250);
        if (!sendTab) {
          // 바구니가 닫혀 있으면 오른쪽 아래 '발송'(바구니) 버튼을 눌러 엶
          step('3/5 발송 바구니 여는 중…');
          const grp = document.querySelector('[aria-label="교육자료 발송 버튼"]');
          // 바구니 버튼이 <button>이 아닐 수도 있어서 누를 수 있는 요소를 넓게 찾음
          const clickable = (root) => root ? Array.from(root.querySelectorAll('button, [role="button"], a, [tabindex]')).filter(visible) : [];
          const cands = clickable(grp).filter((c) => !/추천/.test((c.getAttribute('aria-label') || '') + c.textContent))
            .sort((a, b) => (/발송/.test(b.getAttribute('aria-label') || '') ? 1 : 0) - (/발송/.test(a.getAttribute('aria-label') || '') ? 1 : 0));
          if (grp && visible(grp)) cands.push(grp);
          Array.from(document.querySelectorAll('button, [role="button"]')).forEach((b) => {
            if (norm(b.textContent) === '발송' && visible(b) && !b.closest('[role="dialog"]') && cands.indexOf(b) < 0) cands.push(b);
          });
          const tried = [];
          for (const c of cands.slice(0, 4)) {
            tried.push((c.tagName || '').toLowerCase() + (c.getAttribute('aria-label') ? '[' + c.getAttribute('aria-label') + ']' : '') + (norm(c.textContent) ? '"' + norm(c.textContent).slice(0, 8) + '"' : ''));
            try { c.click(); } catch (e) {}
            sendTab = await waitFor(findSendTab, 3000, 250);
            if (sendTab) break;
            realClick(c);
            sendTab = await waitFor(findSendTab, 3000, 250);
            if (sendTab) break;
          }
          if (!sendTab) {
            const basket = document.querySelector('[role="dialog"][aria-label="발송 바구니"]');
            it.diag = '바구니 버튼 묶음 ' + (grp ? '있음' : '없음') + ' · 눌러 본 것: ' + (tried.join(', ') || '없음') +
              ' · 발송 바구니 ' + (basket ? (basket.hasAttribute('hidden') || !visible(basket) ? '닫힘' : '열림') : '없음');
          }
        }
        if (sendTab && !eduDateBtn()) { sendTab.click(); }
        dateBtn = await waitFor(() => eduDateBtn(), 10000);
        if (!dateBtn && sendTab && sendTab.isConnected && visible(sendTab)) { realClick(sendTab); dateBtn = await waitFor(() => eduDateBtn(), 6000); }
      }
      if (!dateBtn) throw new Error('날짜 선택 창(교육상담 반영내용)이 열리지 않음 · 바구니의 교육자료 발송 버튼 ' + (findByText(['button'], '교육자료 발송') ? '있음' : '없음') + (it.diag ? ' · ' + it.diag : ''));

      // 4. 달력에서 날짜 찾기
      step('4/5 달력에서 날짜 찾는 중…');
      let grid = calReady();
      if (!grid) { // 개편된 날짜 칸은 한 번 눌러서 안 열릴 때가 있어 방법을 바꿔 가며 최대 3번
        const opens = [() => dateBtn.click(), () => realClick(dateBtn), () => { ['pointerdown', 'pointerup'].forEach((t) => dateBtn.dispatchEvent(new PointerEvent(t, { bubbles: true, cancelable: true, pointerType: 'mouse' }))); dateBtn.click(); }];
        for (const o of opens) {
          try { dateBtn.scrollIntoView({ block: 'center' }); o(); } catch (e) {}
          grid = await waitFor(() => calReady(), 2000);
          if (grid) break;
        }
      }
      if (!grid) throw new Error('달력이 열리지 않음 (날짜 칸을 3번 눌러 봄)');
      if (!(await gotoMonth(it.date.slice(0, 7)))) throw new Error('달력에서 해당 달로 이동하지 못함');
      const dayBtn = calDayButton(it.date);
      if (!dayBtn) throw new Error('달력에서 날짜를 찾지 못함 (' + it.date + ' · 달력에 보이는 달: ' + (calShownYm() || '모름') + ')');
      if (calDisabled(dayBtn)) {
        throw new Error('선택할 수 없는 날짜(달력에서 비활성)' + cycleNote(it.date));
      }

      // 5. 연습 모드: 날짜는 누르지 않고 멈춤
      if (dry) {
        dayBtn.style.outline = '3px solid #e53935';
        return '연습 완료: 환자·자료·날짜 위치까지 확인(날짜 클릭 안 함) · ' + title + (favOk ? '' : ' · 즐겨찾기 토글 확인 못함');
      }

      // 6. 최종 확인 후 날짜 클릭 → 이후는 자동 진행
      step('5/5 날짜 클릭');
      if (headerBirth6() !== it.birth6) throw new Error('최종 확인 실패: 화면의 환자가 다름');
      dayBtn.click();
      const filled = await waitFor(() => {
        const b = eduDateBtn();
        return !b || !b.textContent.includes('--');
      }, 4000);
      if (!filled) throw new Error('날짜가 입력되지 않음');
      used.add(label);
      usedByPatient.set(key, used);
      it.materialLabel = label;
      return '날짜 선택됨 → 자동 진행 · ' + title + (favOk ? '' : ' · 즐겨찾기 토글 확인 못함');
    }

    // ---- 패널 ----
    const panel = document.createElement('div');
    panel.id = 'wc-queue-panel';
    panel.style.cssText =
      'position:fixed;right:12px;bottom:44px;width:330px;max-height:70vh;overflow:auto;' +
      'z-index:2147483000;background:#fff;color:#222;border:1px solid #888;border-radius:8px;' +
      'box-shadow:0 4px 16px rgba(0,0,0,.25);font:12px/1.5 sans-serif;padding:8px;display:none;';
    panel.innerHTML =
      '<div style="font-weight:700;margin-bottom:0">🐵 칼퇴몽키 · 교육상담 v11.3 <span id="wcq-count"></span>' +
      '<span id="wcq-min" style="float:right;cursor:pointer;padding:0 6px">−</span></div>' +
      '<div style="color:#777;font-size:11px;margin-bottom:4px">클릭은 몽키가, 퇴근은 내가</div>' +
      '<div id="wcq-body">' + MODE_ROW('wcq') +
      '<div style="display:flex;gap:4px;margin-bottom:4px">' +
      '<button id="wcq-t-pcp" type="button" style="flex:1">포괄평가</button>' +
      '<button id="wcq-t-edu" type="button" style="flex:1;font-weight:700;background:#e3f2fd">교육상담</button>' +
      '<button id="wcq-t-assess" type="button" style="flex:1">점검평가</button></div>' +
      '<textarea id="wcq-text" rows="5" style="width:100%;box-sizing:border-box;font:11px monospace" ' +
      'placeholder="이름,생년월일,날짜 (쉼표·탭·공백 모두 가능)\n홍길동,960824,2026-10-05"></textarea>' +
      '<div style="display:flex;gap:4px;margin:4px 0"><button id="wcq-xl" type="button" style="flex:1">엑셀 불러오기</button>' +
      '<input id="wcq-xl-file" type="file" accept=".xlsx,.xls,.csv" style="display:none"></div>' +
      '<div style="display:flex;gap:4px;margin-bottom:4px">' +
      '<button id="wcq-apply" type="button" style="flex:1">목록 적용</button>' +
      '<button id="wcq-run" type="button" style="flex:1;font-weight:700">다음 건 실행</button>' +
      '<button id="wcq-clear" type="button" style="flex:1">비우기</button></div>' +
      '<div style="display:flex;gap:4px;margin-bottom:4px">' +
      '<button id="wcq-all" type="button" style="flex:1;font-weight:700">연속 실행</button>' +
      '<button id="wcq-stop" type="button" style="flex:1">정지</button></div>' +
      '<label style="display:block;margin:2px 0">심평원 이동 후 대기 <input id="wcq-wait" type="number" value="2" min="0" style="width:52px;border:1px solid #aaa"> 초 (심평원 창은 혼자 마무리해요)</label>' +
      '<div id="wcq-msg" style="color:#c62828;min-height:16px"></div>' +
      '<div id="wcq-report" style="display:none;border:1px solid #e65100;border-radius:4px;padding:4px;margin:4px 0">' +
      '<div id="wcq-report-body" style="white-space:pre-wrap;user-select:text"></div>' +
      '<button id="wcq-copy" type="button" style="margin-top:4px">확인 필요 목록 복사</button></div>' +
      '<div id="wcq-list"></div></div>';
    const oldPanel = document.getElementById('wc-queue-panel');
    if (oldPanel) oldPanel.remove();
    document.body.appendChild(panel);

    const pcpShared = {}; // 포괄평가 쪽 목록 검색 기능을 교육상담에서도 쓰기 위한 연결
    let items = [];
    let running = false;
    let confirmedPrev = false;
    let stopFlag = false;
    let reportText = '';
    const usedByPatient = new Map(); // 환자별로 이번에 이미 보낸 자료
    const q = (sel) => panel.querySelector(sel);
    // 업무 탭: 포괄평가·점검평가를 누르면 일만사 대상자 목록으로 이동해 그 패널을 띄움
    function goPcpTask(t) {
      if (running) { msg('실행 중에는 업무를 바꿀 수 없어요.'); return; }
      const allMenu = findByText(['a', 'button', 'span', 'li', 'div'], '일만사대상자목록');
      if (!allMenu) { msg('"일만사 대상자 목록" 메뉴를 찾지 못했어요. 메뉴에서 직접 열어 주세요.'); return; }
      if (pcpShared.setTask) pcpShared.setTask(t);
      allMenu.click();
    }
    panel.querySelector('#wcq-t-pcp').addEventListener('click', () => goPcpTask('pcp'));
    panel.querySelector('#wcq-t-assess').addEventListener('click', () => goPcpTask('assess'));
    const msg = (t, color) => {
      const m = q('#wcq-msg');
      m.textContent = t || '';
      m.style.color = color || '#c62828';
    };
    function step(t) { msg(t, '#1565c0'); }

    function render() {
      const pending = items.filter((x) => x.status === '대기').length;
      q('#wcq-count').textContent = items.length ? '(' + items.length + '건, 대기 ' + pending + ')' : '';
      const list = q('#wcq-list');
      list.innerHTML = '';
      items.forEach((it, idx) => {
        const d = document.createElement('div');
        d.style.cssText = 'border-top:1px solid #eee;padding:2px 0;';
        d.textContent = (idx + 1) + '. ' + it.name + ' ' + it.birth6 + ' ' + it.date + ' · ' + it.status + (it.note ? ' (' + it.note + ')' : '');
        list.appendChild(d);
      });
    }

    q('#wcq-min').addEventListener('click', () => {
      const b = q('#wcq-body');
      b.style.display = b.style.display === 'none' ? '' : 'none';
    });
    q('#wcq-m-hira').addEventListener('click', () => { if (running) { msg('실행 중에는 모드를 바꿀 수 없어요.'); return; } setMode('hira'); });
    q('#wcq-m-save').addEventListener('click', () => { if (running) { msg('실행 중에는 모드를 바꿀 수 없어요.'); return; } setMode('save'); });
    applyModeUI();
    q('#wcq-xl').addEventListener('click', () => q('#wcq-xl-file').click());
    q('#wcq-xl-file').addEventListener('change', (e) => {
      const f = e.target.files && e.target.files[0];
      e.target.value = '';
      if (f) importExcel(f, msg, 'AX004'); // 교육상담 패널에서 넣으면 코드 없는 파일은 교육상담으로
    });
    q('#wcq-apply').addEventListener('click', () => {
      items = parseQueue(q('#wcq-text').value);
      msg(items.length ? '' : '해석할 수 있는 줄이 없어요 (이름,생년월일,날짜)');
      render();
    });
    q('#wcq-clear').addEventListener('click', () => {
      q('#wcq-report').style.display = 'none';
      items = [];
      q('#wcq-text').value = '';
      msg('');
      render();
    });
    let runningSince = 0;
    q('#wcq-run').addEventListener('click', async () => {
      msg('실행 버튼이 눌렸어요…', '#1565c0');
      try {
        if (running && Date.now() - runningSince < 60000) {
          msg('이전 실행이 아직 진행 중이에요 (최대 1분). 계속 그대로면 페이지를 새로고침해 주세요.');
          return;
        }
        const dry = false; // 연습 모드는 없앰
        const it = items.find((x) => x.status === '대기');
        if (!it) { msg('대기 중인 건이 없어요. 먼저 "목록 적용"을 눌러 주세요.'); return; }

        const prev = items.find((x) => x.status.startsWith('날짜 선택됨'));
        if (!dry && prev && !confirmedPrev) {
          msg('이전 건의 심평원 작성완료가 끝났는지 확인하고, 끝났으면 한 번 더 눌러 주세요.');
          confirmedPrev = true;
          return;
        }
        confirmedPrev = false;
        if (prev) prev.status = '진행됨(완료 신호는 못 받음)';

        running = true; badgeReset();
        runningSince = Date.now();
        it.note = '';
        it.status = '진행 중…';
        render();
        try {
          const r = await runItem(it, dry);
          if (dry) {
            it.status = '대기';
            it.note = '연습 성공';
            msg('✔ ' + r, '#2e7d32');
          } else {
            it.status = r;
            msg('✔ ' + r, '#2e7d32');
          }
        } catch (e) {
          const m = e && e.message ? e.message : String(e);
          console.error('[자동화] 대기열 실행 오류:', e);
          if (dry) { it.status = '대기'; it.note = '연습 실패: ' + m; } else { it.status = '실패: ' + m; }
          msg('✘ ' + m + ' · 팝업이 열려 있으면 직접 닫아 주세요');
        }
        running = false;
        render();
      } catch (e) {
        running = false;
        console.error('[자동화] 대기열 버튼 오류:', e);
        msg('오류: ' + (e && e.message ? e.message : e));
      }
    });

    // ---- 연속 실행 ----
    // 발송 중 뜨는 오류 팝업 찾기: 짧은 팝업 + "확인" 버튼만 있고 "취소"가 없는 것
    // (예: "임시저장 실패", "교육상담 인력 중복"). 발송 확인 팝업·심평원 이동 안내는 제외
    function findErrorDialog() {
      const dlgs = openDialogs().reverse(); // 맨 위 팝업부터
      for (const d of dlgs) {
        const btns = Array.from(d.querySelectorAll('button'));
        const ok = btns.find((b) => norm(b.textContent) === '확인');
        if (!ok) continue;
        if (btns.some((b) => norm(b.textContent) === '취소')) continue;
        if (btns.some((b) => norm(b.textContent).includes('심평원이동'))) continue;
        const raw = d.innerText || d.textContent || '';
        if (norm(raw).includes('발송하시겠습니까')) continue;
        const lines = raw.split(/\n+/).map((t) => t.replace(/\s+/g, ' ').trim()).filter((t) => t && t !== '확인');
        const text = lines.join(' · ').replace(/\s*확인\s*$/, '');
        if (!text) continue;
        // 큰 창(교육상담 반영내용 등)은 제외. 단 '실패'·'필수값' 팝업은 누락 항목이 길게 나열돼도 오류로 봄
        if (text.length > 160 && !/실패|필수/.test(norm(text))) continue;
        return { ok, text };
      }
      return null;
    }

    async function waitCompletion(it) {
      // 1) 웰체크 쪽: 발송 → 확인 → 심평원 이동 클릭까지. 중간에 오류 팝업이 뜨면 바로 확인을 눌러 사유를 기록
      const first = await waitFor(() => {
        if (stopFlag) return 'stop';
        if (lastEvent === '심평원 이동 클릭함') return 'moved';
        if (lastEvent === '발송 완료(심평원 이동 생략)') return 'saved';
        const err = findErrorDialog();
        return err || null;
      }, 45000, 400);
      if (first === 'stop') return 'stop';
      if (!first) return 'stuck';
      if (first === 'saved') { // 임시저장까지만: 심평원 이동 안내창을 닫고 바로 다음 건
        await sleep(400);
        await closeDialogs();
        it.status = '발송 완료(심평원 생략)';
        return 'ok';
      }
      if (first !== 'moved') {
        it.errorText = first.text;
        first.ok.click();
        await sleep(500);
        return 'error';
      }
      // 2) 심평원 쪽은 스크립트가 알아서 작성완료로 바꿈 → 완료 신호가 오면 바로, 없으면 대기 시간 뒤 다음 건
      const wv = parseInt(q('#wcq-wait').value, 10);
      const waitMs = Math.max(0, isNaN(wv) ? 2 : wv) * 1000; // 심평원 탭은 따로 작성완료하고 스스로 닫히므로 오래 기다릴 필요 없음
      const t1 = Date.now();
      while (Date.now() - t1 < waitMs) {
        runningSince = Date.now();
        if (stopFlag) return 'stop';
        if (it.status === '심평원 작성완료 확인됨' || it.status.startsWith('발송 완료')) return 'ok';
        await sleep(500);
      }
      if (it.status.startsWith('날짜 선택됨')) it.status = '발송 완료(심평원은 자동 진행)';
      return 'ok';
    }

    // 열린 팝업의 오른쪽 위 닫기(X) 버튼 찾기
    function findCloseButton(dlg) {
      const rect = dlg.getBoundingClientRect();
      const cands = Array.from(dlg.querySelectorAll('button')).filter((b) => {
        if (b.getAttribute('role') === 'checkbox') return false;
        if (/day-picker|date-picker|rdp-/.test(b.className || '')) return false; // 달력 이동 버튼 제외
        const label = (b.getAttribute('aria-label') || '') + ' ' + (b.textContent || '');
        if (/닫기|close/i.test(label)) return true;
        return !(b.textContent || '').trim() && !!b.querySelector('svg');
      });
      if (!cands.length) return null;
      const score = (b) => {
        const r = b.getBoundingClientRect();
        return Math.abs(r.right - rect.right) + Math.abs(r.top - rect.top);
      };
      cands.sort((x, y) => score(x) - score(y));
      return cands[0];
    }
    // 열린 팝업을 모두 닫기 (Esc와 X 버튼을 번갈아 시도, 맨 위 팝업부터)
    async function closeDialogs() {
      for (let i = 0; i < 8; i++) {
        const dlgs = openDialogs();
        if (!dlgs.length) return true;
        const dlg = dlgs[dlgs.length - 1];
        if (i % 2 === 0) {
          document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true }));
        } else {
          const btn = findCloseButton(dlg);
          if (btn) btn.click();
        }
        await sleep(600);
      }
      if (openDialogs()[0]) ignoreStuckWidgets(); // 끝까지 안 닫히는 상시 요소는 무시
      return !openDialogs()[0];
    }
    async function uncheckAll() {
      const checked = materialCheckboxes(true);
      for (const b of checked) { b.click(); await sleep(300); }
    }
    // 정리: 팝업 닫고 체크해 둔 자료 해제 (실패한 건이나 연습 모드 뒤에 사용)
    async function dryCleanup() {
      if (!(await closeDialogs())) return false;
      await uncheckAll();
      return true;
    }

    function showReport(problems) {
      const box = q('#wcq-report');
      if (!problems.length) { box.style.display = 'none'; reportText = ''; return; }
      const lines = problems.map((p) => p.it.name + ' ' + p.it.birth6 + ' ' + p.it.date + ' — ' + p.reason);
      reportText = '확인이 필요한 건 ' + problems.length + '건\n' + lines.join('\n');
      q('#wcq-report-body').textContent = reportText;
      box.style.display = 'block';
    }
    q('#wcq-copy').addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(reportText);
        msg('확인 필요 목록을 복사했어요.', '#2e7d32');
      } catch (e) {
        msg('복사하지 못했어요. 위 글자를 직접 선택해서 복사해 주세요.');
      }
    });

    async function runAll() {
      if (running && Date.now() - runningSince < 60000) { msg('이미 실행 중이에요.'); return; }
      if (!items.some((x) => x.status === '대기')) { msg('대기 중인 건이 없어요. 먼저 "목록 적용"을 눌러 주세요.'); return; }
      running = true; badgeReset();
      runningSince = Date.now();
      stopFlag = false;
      confirmedPrev = false;
      q('#wcq-report').style.display = 'none';
      let done = 0;
      let failed = 0; // 연습 모드에서만 사용
      const problems = []; // 실제 실행에서 확인이 필요한 건 {it, reason}
      let stopReason = '';
      const dry = false; // 연습 모드는 없앰
      const queue = dry ? items.filter((x) => x.status === '대기') : null; // 연습은 한 바퀴만 돔
      let qi = 0;
      try {
        while (!stopFlag) {
          const it = dry ? queue[qi++] : items.find((x) => x.status === '대기');
          if (!it) break;
          runningSince = Date.now();
          let clean = await waitFor(() => !openDialogs()[0], 3000, 300);
          if (!clean) clean = await closeDialogs();
          if (!clean) { stopReason = '팝업이 닫히지 않아 멈췄어요. 팝업을 닫고 다시 시작해 주세요.'; break; }

          lastEvent = '다음 건 시작';
          it.note = '';
          it.status = '진행 중…';
          render();
          if (dry) {
            try {
              await runItem(it, true);
              it.note = '연습 성공';
              done++;
            } catch (e) {
              it.note = '연습 실패: ' + (e && e.message ? e.message : String(e));
              failed++;
            }
            it.status = '대기';
            render();
            if (!(await dryCleanup())) {
              stopReason = '팝업을 닫지 못해 멈췄어요. 팝업을 직접 닫고 다시 시작해 주세요.';
              break;
            }
            await sleep(500);
            continue;
          }

          // 1) 환자·자료·날짜 선택
          try {
            it.status = await runItem(it, false);
          } catch (e) {
            const m = (e && e.message ? e.message : String(e)) + (openDialogs()[0] ? screenNote() : '');
            console.error('[자동화] 연속 실행 중 실패:', e);
            it.status = '실패: ' + m;
            problems.push({ it, reason: m });
            render();
            await sleep(500);
            if (!(await dryCleanup())) {
              stopReason = '실패한 건을 정리하다 팝업을 닫지 못해 멈췄어요 (' + m + ')';
              break;
            }
            continue; // 실패해도 다음 건으로 진행
          }
          render();

          // 2) 발송 → 심평원 작성완료까지 기다림
          const res = await waitCompletion(it);
          if (res === 'ok') { done++; render(); await sleep(700); continue; }
          if (res === 'stop') { stopReason = '정지했어요.'; break; }

          if (res === 'error') {
            // 오류 팝업의 "확인"은 이미 눌렀음 → 사유를 기록하고, 남은 창을 정리한 뒤 다음 건으로
            const why = '발송 실패: ' + (it.errorText || '오류 팝업');
            // 발송이 실패했으니 그 자료는 아직 안 보낸 것 → 같은 환자의 다음 건에서 다시 쓸 수 있게 되돌림
            const usedSet = usedByPatient.get(it.name + '|' + it.birth6);
            if (usedSet && it.materialLabel) usedSet.delete(it.materialLabel);
            it.status = '실패: ' + why;
            problems.push({ it, reason: why });
            render();
            if (!(await dryCleanup())) {
              stopReason = '발송 실패 후 남은 팝업을 닫지 못해 멈췄어요 (' + why + ')';
              break;
            }
            continue;
          }

          // 발송 단계(심평원 이동 전)에서 멈춘 경우만 확인 필요로 기록
          const why = '발송 단계에서 멈춤 (마지막 상태: ' + lastEvent + ')' + screenNote();
          it.status = '확인 필요(' + why + ')';
          problems.push({ it, reason: why });
          render();
          if (!(await dryCleanup())) {
            stopReason = '발송 단계에서 멈췄고 팝업을 닫지 못해 중단했어요. 화면을 확인해 주세요.';
            break;
          }
        }
      } catch (e) {
        console.error('[자동화] 연속 실행 오류:', e);
        stopReason = '오류: ' + (e && e.message ? e.message : e);
      }
      running = false;
      render();
      showReport(problems);
      if (dry) {
        const sm = ' (연습 성공 ' + done + '건, 실패 ' + failed + '건)';
        if (stopReason) msg(stopReason + sm);
        else msg('✔ 연습이 끝났어요. 실제 발송은 하지 않았어요.' + sm, '#2e7d32');
        return;
      }
      const summary = ' (완료 ' + done + '건, 확인 필요 ' + problems.length + '건)';
      if (stopReason) msg(stopReason + summary);
      else if (problems.length) msg('끝까지 진행했어요. 아래 ' + problems.length + '건은 확인이 필요해요.' + summary, '#e65100');
      else { msg('🙈 모든 건을 처리했어요. 이제 퇴근하세요!' + summary, '#2e7d32'); badgeDone(done + '건 완료, 이제 퇴근하세요'); }
    }
    q('#wcq-all').addEventListener('click', () => {
      runAll().catch((e) => { running = false; msg('오류: ' + (e && e.message ? e.message : e)); });
    });
    q('#wcq-stop').addEventListener('click', () => {
      stopFlag = true;
      msg('정지 요청: 지금 진행 중인 건이 끝나면 멈춰요.', '#1565c0');
    });

    // 심평원 탭이 작성완료를 눌렀다는 신호 (받을 수 있을 때만)
    function onMsg(e) {
      if (e.origin !== 'https://biz.hira.or.kr') return;
      const t = e.data && e.data.wcAuto;
      if (t !== 'done' && t !== 'unconfirmed') return;
      const it = items.find((x) => x.status.startsWith('날짜 선택됨') || x.status.startsWith('발송 완료'));
      if (!it) return;
      it.status = t === 'done' ? '심평원 작성완료 확인됨' : '발송 완료(심평원 확인 신호 없음)';
      render();
    }
    window.addEventListener('message', onMsg);

    // ============================================================
    // 4) 포괄평가 대기열
    //    대상자 목록 → (문진·검사 버튼) → 임상검사값 입력하기 → 환자 포괄평가 화면
    //    → 계획수립일자 → EMR 검사 결과 일괄 반영 → 임시저장(심평원) → 심평원 바로가기
    //    심평원 탭에서는 별도 스크립트가 작성완료를 눌러 줌
    // ============================================================
    function startPcp() {
      const SAVE_TEXT = '임시저장(심평원)';
      let pItems = [];
      let task = 'pcp';
      const store = { pcp: { items: [], text: '' }, assess: { items: [], text: '' } };
      let pRunning = false;
      let pStop = false;
      let pSince = 0;
      let pSignal = null;
      let pReportText = '';

      // ---- 화면 읽기 ----
      const saveBtn = () => findByText(['button'], SAVE_TEXT) || findByText(['button'], '임시저장');
      const onPatientPage = () => isPcpPatientPage();
      function pcpRows() {
        const out = [];
        document.querySelectorAll('tr').forEach((tr) => {
          // 대상자 목록의 행만: assessment-table 칸이 있고, 이름 칸에 글자가 있어야 함 (교육상담 이력 표와 구분)
          if (!tr.querySelector('td[class*="assessment-table"]')) return;
          const cells = Array.from(tr.children).filter((c) => c.tagName === 'TD');
          const bi = cells.findIndex((td) => /^\d{2}\.\d{2}\.\d{2}$/.test(td.textContent.trim()));
          if (bi < 1) return;
          const nm = cells[bi - 1].textContent.trim();
          if (!/[가-힣A-Za-z]/.test(nm)) return;
          out.push({ tr, cells, nameCell: cells[bi - 1], name: nm, birth6: digits(cells[bi].textContent) });
        });
        return out;
      }
      const onListPage = () => !onPatientPage() && pcpRows().length > 0;
      // 문진·검사 버튼: 개편 전 클래스로 먼저 찾고, 없으면 글자에 '문진'과 '검사'가 들어간 버튼을 찾음
      const evalButton = (row) => row.tr.querySelector('button[class*="PcpStatusButtonRenderer_new_evaluation_button"]') ||
        Array.from(row.tr.querySelectorAll('button, [role="button"]')).find((b) => /문진.*검사/.test(norm(b.textContent)) && !b.disabled) ||
        Array.from(row.tr.querySelectorAll('button, [role="button"]')).find((b) => /문진/.test(norm(b.textContent)) && !b.disabled) || null;
      function rowSummary(row) {
        return row.cells.slice(-9).map((c) => c.textContent.replace(/\s+/g, ' ').trim()).filter(Boolean).join(' | ');
      }
      const searchInput = () => document.querySelector('input[placeholder*="환자명"]');
      async function clearSearch() {
        const inp = searchInput();
        if (inp && inp.value) await pressEnterSearch(inp, '');
      }
      const findPlanDateBtn = () => planDateButton(true);

      function emrTrigger() {
        const scope = document.querySelector('[class*="detail_header"]') || document;
        const cbs = Array.from(scope.querySelectorAll('button[role="combobox"]'));
        return cbs.find((b) => norm(b.textContent).includes('EMR')) || cbs[0] || null;
      }

      // ---- 클릭 도우미 (Radix 선택 상자는 포인터 이벤트가 필요함) ----
      function firePointer(el, type) {
        const init = { bubbles: true, cancelable: true, pointerType: 'mouse', isPrimary: true, button: 0, buttons: type === 'pointerup' ? 0 : 1 };
        let ev;
        try { ev = new PointerEvent(type, init); } catch (e) { ev = new MouseEvent(type, init); }
        el.dispatchEvent(ev);
      }
      function fireMouse(el, type) {
        el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, button: 0, buttons: type === 'mousedown' ? 1 : 0 }));
      }
      function pointerClick(el) {
        firePointer(el, 'pointerdown'); fireMouse(el, 'mousedown');
        firePointer(el, 'pointerup'); fireMouse(el, 'mouseup'); fireMouse(el, 'click');
      }
      const findOptionByText = (n) =>
        Array.from(document.querySelectorAll('[role="option"], [role="menuitem"]')).find((e) => norm(e.textContent) === n) || null;
      async function selectOption(trigger, optionNorm) {
        pointerClick(trigger);
        let opt = await waitFor(() => findOptionByText(optionNorm), 2500, 100);
        if (!opt) {
          trigger.focus();
          trigger.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
          opt = await waitFor(() => findOptionByText(optionNorm), 2500, 100);
        }
        if (!opt) return false;
        pointerClick(opt);
        await sleep(400);
        if (findOptionByText(optionNorm)) { // 아직 열려 있으면 키보드로 다시 시도
          if (opt.focus) opt.focus();
          opt.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
          await sleep(400);
        }
        return true;
      }

      // ---- 팝업 도우미 ----
      function dialogsTopFirst() { return openDialogs().reverse(); }
      function findDialogButton(textNorm) {
        for (const d of dialogsTopFirst()) {
          const b = Array.from(d.querySelectorAll('button')).find((x) => norm(x.textContent) === textNorm);
          if (b) return b;
        }
        return null;
      }
      function findConfirm(phrases) { // 확인+취소가 함께 있고 문구가 맞는 팝업의 확인 버튼
        for (const d of dialogsTopFirst()) {
          const btns = Array.from(d.querySelectorAll('button'));
          const ok = btns.find((b) => norm(b.textContent) === '확인');
          const cancel = btns.some((b) => norm(b.textContent) === '취소');
          if (!ok || !cancel) continue;
          const raw = d.innerText || d.textContent || '';
          if (phrases.some((p) => norm(raw).includes(norm(p)))) return { btn: ok, text: raw.replace(/\s+/g, ' ').trim() };
        }
        return null;
      }
      function withLabel(label, text) { return text.indexOf(label) === 0 ? text : label + ': ' + text; }
      // 환자 화면에 들어가면 뜨는 "케어플랜을 출력해주세요" 안내창: '이미 출력했습니다.'를 눌러 닫음
      function dismissCarePlan() {
        const dlg = openDialogs().find((d) => norm(d.textContent).indexOf('케어플랜을출력해주세요') >= 0);
        if (!dlg) return false;
        const link = Array.from(dlg.querySelectorAll('div, span, a, button, p'))
          .filter((e) => norm(e.textContent) === '이미출력했습니다.')
          .pop();
        if (!link) return false;
        realClick(link);
        return true;
      }
      async function clearCarePlan() {
        for (let i = 0; i < 4; i++) {
          if (!dismissCarePlan()) return;
          await sleep(500);
        }
      }

      function classifyPopup(text) {
        const t = norm(text);
        if (/실패|불가|오류|에러|누락|필수|초과|이상|일치하지|없습니다|않습니다|않았|못했|수없|종료|입력해|확인해|차이|찾을수/.test(t)) return 'fail';
        if (/완료|성공|되었습니다|저장되었/.test(t)) return 'success';
        return 'unknown';
      }
      async function watchPopups(ms, confirmPhrases) {
        const t0 = Date.now();
        while (Date.now() - t0 < ms) {
          if (confirmPhrases && confirmPhrases.length) {
            const cf = findConfirm(confirmPhrases);
            if (cf) { cf.btn.click(); await sleep(600); continue; }
          }
          const er = findErrorDialog();
          if (er) {
            const kind = classifyPopup(er.text);
            er.ok.click();
            await sleep(500);
            return { kind, text: er.text };
          }
          await sleep(300);
        }
        return { kind: null, text: '' };
      }

      // ---- 목록으로 돌아가기 ----
      async function backToList() {
        if (onListPage()) return true;
        const btn = findByText(['button'], '목록');
        if (!btn) return false;
        btn.click();
        for (let i = 0; i < 25; i++) {
          await sleep(400);
          if (onListPage()) return true;
          const cf = findConfirm(['나가', '이동', '변경', '저장되지', '작성중', '사라']);
          if (cf) cf.btn.click();
        }
        return onListPage();
      }
      async function cleanup() {
        if (!(await closeDialogs())) return false;
        document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true })); // 달력 닫기
        await sleep(300);
        return backToList();
      }

      // ---- 한 건 실행 ----
      // ---- LDL 자동계산 도우미 ----
      // 임상검사 표에서 검사 이름 칸 → 결과 칸(입력칸·자동계산 버튼) → 검사일자 칸(달력 버튼)
      function labCells(prefixNorm) {
        const cells = Array.from(document.querySelectorAll('td, th')).filter((c) => {
          const t = norm(c.textContent);
          return t.startsWith(prefixNorm) && t.length < 40 && !c.querySelector('input');
        });
        for (const c of cells) {
          let input = null, calc = null, dateBtn = null;
          let n = c.nextElementSibling;
          for (let i = 0; i < 5 && n; i++, n = n.nextElementSibling) {
            if (!input) {
              const inp = n.querySelector('input:not([type="checkbox"]):not([type="radio"]):not([type="hidden"])');
              if (inp) {
                input = inp;
                // 개편 후 자동계산 버튼은 아이콘+글자라서 글자에 '자동계산'이 들어 있으면 됨
                calc = Array.from(n.querySelectorAll('button, [role="button"]')).find((b) => norm(b.textContent).indexOf('자동계산') >= 0) || null;
                const sameCellDate = n.querySelector(DATE_SEL);
                if (sameCellDate) { dateBtn = sameCellDate; break; }
                continue;
              }
            }
            if (input && !dateBtn) { dateBtn = n.querySelector(DATE_SEL); if (dateBtn) break; }
          }
          if (input && !calc) { const tr = c.closest('tr'); if (tr) calc = Array.from(tr.querySelectorAll('button, [role="button"]')).find((b) => norm(b.textContent).indexOf('자동계산') >= 0) || null; }
          if (input && !dateBtn) { const tr = c.closest('tr'); if (tr) dateBtn = tr.querySelector(DATE_SEL); }
          if (input) return { input, calc, dateBtn };
        }
        return null;
      }
      function btnDate(btn) { // 달력 버튼의 날짜 → 'YYYY-MM-DD' (비어 있으면 '')
        const d = btn ? digits(btn.textContent) : '';
        const f = d.length === 8 ? d : d.length === 6 ? '20' + d : '';
        return f ? f.slice(0, 4) + '-' + f.slice(4, 6) + '-' + f.slice(6, 8) : '';
      }
      // 총콜레스테롤·HDL·LDL 중 2개가 있으면 빠진 1개의 '자동계산'을 누르고, 그 검사일자를 다른 지질 검사일자와 똑같이 넣음
      // (원장님에 따라 HDL이나 총콜레스테롤을 안 하는 경우가 있음. 계산에는 TG 값도 필요)
      const LIPIDS = [['총콜레스테롤', '총콜레스테롤'], ['HDL-콜레스테롤', 'HDL'], ['LDL-콜레스테롤', 'LDL']];
      async function fillLipid(it) {
        const cells = LIPIDS.map(([k, short]) => ({ k, short, c: labCells(k) }));
        const lost = cells.filter((x) => !x.c);
        if (lost.length) return lost.map((x) => x.short).join(', ') + ' 칸을 찾지 못해 자동계산은 건너뜀';
        const empty = cells.filter((x) => !x.c.input.value);
        if (empty.length === 0) return '';
        if (empty.length > 1) return '지질 검사 값이 ' + (3 - empty.length) + '개뿐이라 자동계산 못 함(' + empty.map((x) => x.short).join(', ') + ' 없음)';
        const tg = labCells('TG중성지방');
        const target = empty[0];
        if (!tg || !tg.input.value) return target.short + ' 자동계산 못 함(TG 값 없음)';
        if (!target.c.calc) return target.short + ' 자동계산 버튼을 찾지 못해 건너뜀';
        try { target.c.calc.scrollIntoView({ block: 'center' }); } catch (e) {}
        target.c.calc.click();
        const calcDone = () => { const er = findErrorDialog(); if (er) return er; const c = labCells(target.k); return c && c.input.value ? 'ok' : null; };
        let got = await waitFor(calcDone, 3000, 200);
        if (!got) { const c2 = labCells(target.k); if (c2 && c2.calc) realClick(c2.calc); got = await waitFor(calcDone, 3000, 200); }
        if (got && got !== 'ok') { const t = got.text; got.ok.click(); await sleep(400); return target.short + ' 자동계산 안 됨: ' + t; }
        if (!got) return target.short + ' 자동계산을 눌렀지만 값이 들어오지 않음';
        const notes = [target.short + ' 자동계산함'];
        const cur = labCells(target.k);
        if (!cur.dateBtn) return notes.concat(target.short + ' 검사일자 칸을 찾지 못함').join(' · ');
        let ref = '';
        for (const k of ['총콜레스테롤', 'HDL-콜레스테롤', 'LDL-콜레스테롤', 'TG중성지방']) {
          if (k === target.k) continue;
          const c = labCells(k); ref = c && btnDate(c.dateBtn); if (ref) break;
        }
        if (!ref) return notes.concat('참고할 지질 검사일자가 없어 ' + target.short + ' 날짜는 비워 둠').join(' · ');
        if (btnDate(cur.dateBtn) === ref) return notes.join(' · ');
        if (btnDate(cur.dateBtn)) return notes.concat(target.short + ' 검사일자가 이미 있어 그대로 둠').join(' · ');
        const dayBtn = await pickDate(cur.dateBtn, ref, target.short + ' 검사일자');
        dayBtn.click();
        let setOk = await waitFor(() => dateIs(labCells(target.k).dateBtn, ref), 2500);
        if (!setOk && dayBtn.isConnected) { realClick(dayBtn); setOk = await waitFor(() => dateIs(labCells(target.k).dateBtn, ref), 2500); }
        if (!setOk) throw new Error(target.short + ' 검사일자가 입력되지 않음 (넣으려던 날짜 ' + ref + ')');
        document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true }));
        return notes.concat(target.short + ' 검사일자 ' + ref).join(' · ');
      }

      // ---- EMR로 불러온 값 중 필수(*)가 아닌 검사 지우기 ----
      // 임상검사 표에서 이름 칸에 '*'가 없는 검사에 값이 들어 있으면 그 줄의 휴지통 버튼을 누름
      // 혈압은 항상 필수: 혈압 줄이 '수축기'·'이완기'로 나뉘어 이름 칸에 '혈압'이 없을 수 있어서 이 말들이 있으면 항상 남김
      const KEEP_RE = /혈압|수축기|이완기|mmHg/i;
      function labTableRows() {
        const tables = Array.from(document.querySelectorAll('table')).filter((t) => {
          const head = norm((t.querySelector('thead') || t.rows[0] || {}).textContent || '');
          return /검사항목|임상검사/.test(head) && t.querySelector(DATE_SEL);
        });
        const out = [];
        for (const t of tables) {
          for (const tr of Array.from(t.querySelectorAll('tbody tr'))) {
            const cells = Array.from(tr.children);
            const fi = cells.findIndex((c) => c.querySelector('input:not([type="hidden"])'));
            if (fi < 1) continue;
            const labelCell = cells[fi - 1];
            const label = norm(labelCell.textContent);
            if (!label || label.length > 40) continue;
            const before = cells.slice(0, fi).map((c) => c.textContent).join(' '); // 이름 칸 앞쪽(묶음 칸 포함) 글자
            out.push({ tr, label, required: before.indexOf('*') >= 0, keep: KEEP_RE.test(before) });
          }
        }
        return out;
      }
      function rowHasValue(tr) {
        return Array.from(tr.querySelectorAll('input:not([type="hidden"])')).some((i) =>
          i.type === 'radio' || i.type === 'checkbox' ? i.checked : !!i.value.trim());
      }
      function trashButton(tr) {
        const btns = Array.from(tr.querySelectorAll('button')).filter((b) =>
          !/date_picker_button|date-picker_trigger/.test(b.className) && b.getAttribute('role') !== 'checkbox' && norm(b.textContent) !== '자동계산');
        return btns.find((b) => /delete|trash|remove/i.test(b.className + ' ' + (b.getAttribute('aria-label') || '') + ' ' + (b.title || '')) || /삭제/.test((b.getAttribute('aria-label') || '') + (b.title || '')))
          || btns.find((b) => !norm(b.textContent) && b.querySelector('svg, img, i')) || null;
      }
      // EMR로 불러온 뒤 필수(*) 검사 중 값이나 검사일자가 빈 것 → 어차피 저장이 안 되니 바로 실패로 넘김
      function requiredMissing(skipRe) {
        const miss = [];
        for (const r of labTableRows()) {
          if (!r.required) continue;
          if (skipRe && skipRe.test(r.label)) continue;
          const name = r.label.replace(/\(.*$/, '').replace(/\*/g, '');
          if (!rowHasValue(r.tr)) { miss.push(name); continue; }
          const db = r.tr.querySelector(DATE_SEL);
          if (db && !digits(db.textContent)) miss.push(name + ' 검사일자');
        }
        return Array.from(new Set(miss));
      }
      function failIfRequiredMissing(skipRe) {
        const miss = requiredMissing(skipRe);
        if (miss.length) throw new Error('저장 실패(필수값 없음): EMR 결과 가져오기 후에도 ' + miss.join(', ') + ' 값이 없어 저장하지 않고 다음 건으로 넘어감');
      }
      async function clearOptionalLabs() {
        const removed = [], failed = [];
        const all = labTableRows();
        if (!all.length) return '';
        // '*' 표시를 하나도 못 읽으면(화면 구조가 바뀐 경우) 잘못 지울 수 있어서 아무것도 지우지 않음
        if (!all.some((r) => r.required)) return '필수(*) 표시를 읽지 못해 검사값은 지우지 않음';
        const targets = labTableRows().filter((r) => !r.required && !r.keep && rowHasValue(r.tr)).map((r) => r.label);
        // 지우기 전에 남겨야 할 줄(필수·혈압) 중 값이 있는 줄을 기억해 두고, 지울 때마다 그 값이 그대로인지 확인
        const guarded = all.filter((r) => (r.required || r.keep) && rowHasValue(r.tr)).map((r) => r.label);
        const emptyKeep = all.filter((r) => r.keep && !rowHasValue(r.tr)).map((r) => r.label.replace(/\(.*$/, ''));
        const lostGuarded = () => guarded.filter((l) => { const x = labTableRows().find((y) => y.label === l); return !x || !rowHasValue(x.tr); });
        for (const label of targets) {
          const r = labTableRows().find((x) => x.label === label); // 지울 때마다 화면이 다시 그려질 수 있어 새로 찾음
          if (!r || !rowHasValue(r.tr)) continue;
          const name = label.replace(/\(.*$/, '');
          // 휴지통은 마우스를 올린(선택된) 줄에 붙는 버튼이라, 먼저 그 줄에 마우스를 올린 것처럼 만들고 그 줄의 입력칸을 선택함
          try { r.tr.scrollIntoView({ block: 'center' }); } catch (e) {}
          const hover = (el) => ['pointerover', 'pointerenter', 'mouseover', 'mouseenter', 'mousemove'].forEach((t) => {
            try { el.dispatchEvent(new MouseEvent(t, { bubbles: t !== 'mouseenter' && t !== 'pointerenter', cancelable: true })); } catch (e) {}
          });
          hover(r.tr);
          Array.from(r.tr.children).forEach(hover);
          const inp0 = r.tr.querySelector('input:not([type="hidden"]):not([type="radio"]):not([type="checkbox"])');
          if (inp0) { try { inp0.focus(); } catch (e) {} }
          await sleep(300);
          const tb = trashButton(r.tr);
          if (!tb) { failed.push(name + '(휴지통 버튼 없음)'); continue; }
          // 이 휴지통이 정말 이 줄 안에 있는지, 남겨야 할 줄의 버튼은 아닌지 다시 확인
          const owner = tb.closest('tr');
          const ownerRow = labTableRows().find((x) => x.tr === owner);
          if (owner !== r.tr || (ownerRow && (ownerRow.required || ownerRow.keep))) { failed.push(name + '(휴지통이 다른 줄 것이라 누르지 않음)'); continue; }
          hover(tb);
          realClick(tb);
          await sleep(400);
          const cf = findConfirm(['삭제', '지우', '초기화']);
          if (cf) { cf.btn.click(); await sleep(400); }
          const gone = await waitFor(() => { const x = labTableRows().find((y) => y.label === label); return !x || !rowHasValue(x.tr); }, 2500, 200);
          (gone ? removed : failed).push(name);
          const lost = lostGuarded();
          if (lost.length) throw new Error('"' + name + '" 휴지통을 눌렀더니 남겨야 할 ' + lost.map((l) => l.replace(/\(.*$/, '')).join(', ') + ' 값이 지워져서 저장하지 않고 멈춤 — 이 환자는 직접 확인해 주세요');
        }
        const notes = [];
        if (emptyKeep.length) notes.push('EMR에서 ' + Array.from(new Set(emptyKeep)).join(', ') + ' 값을 불러오지 못함(지운 것 아님)');
        if (removed.length) notes.push('필수 아닌 검사 지움: ' + removed.join(', '));
        if (failed.length) notes.push('못 지운 검사: ' + failed.join(', '));
        return notes.join(' · ');
      }

      async function runPcpItem(it, dry) {
        pstep('1/6 대상자 목록에서 환자 찾는 중…');
        if (openDialogs()[0]) throw new Error('열려 있는 팝업이 있어요. 닫고 다시 눌러 주세요');
        if (!onListPage() && !(await backToList())) throw new Error('대상자 목록 화면에서 시작해 주세요 (목록으로 돌아가지 못함)');
        await clearSearch();
        if (!it.birth6) { // 엑셀에 생년월일이 없던 줄: 이름으로 찾아 한 명일 때만 그 생년월일로 진행
          pstep('1/6 생년월일 찾는 중… (이름으로 대상자 목록 검색)');
          await resolveBirth(it);
          prender();
        }

        const same = (r) => r.name === it.name && r.birth6 === it.birth6;
        const pre = pcpRows();
        let rows = pre.filter(same);
        let scanNote = '';
        if (rows.length === 0) {
          const inp = searchInput();
          if (inp) { await pressEnterSearch(inp, it.name); rows = pcpRows().filter(same); }
        }
        if (rows.length === 0) {
          // 검색으로도 안 나오면 필터를 풀고 다른 쪽을 차례로 확인
          await clearSearch();
          const sc = await scanAllPages(null, () => pcpRows(), same);
          rows = pcpRows().filter(same);
          scanNote = sc.pages ? ' · ' + sc.tried + '쪽까지 확인함' : ' · 쪽 이동 버튼을 찾지 못해 현재 쪽만 확인함';
        }
        if (rows.length === 0) {
          const sameBirth = pre.filter((r) => r.birth6 === it.birth6).map((r) => r.name);
          const sameName = pcpRows().concat(pre).filter((r) => r.name === it.name).map((r) => r.birth6);
          let hint = '';
          if (sameBirth.length) hint += ' · 같은 생년월일의 이름: ' + Array.from(new Set(sameBirth)).join(', ');
          if (sameName.length) hint += ' · 같은 이름의 생년월일: ' + Array.from(new Set(sameName)).join(', ');
          if (!hint) hint = ' · 완료·중단 환자는 목록에서 숨겨져 있을 수 있어요';
          throw new Error('대상자 목록에서 환자를 찾지 못함' + hint + scanNote);
        }
        if (rows.length > 1) throw new Error('이름·생년월일이 같은 환자가 2명 이상');
        const row = rows[0];
        const eb = evalButton(row);
        if (!eb) throw new Error('포괄평가 문진·검사 버튼이 없음 (' + rowSummary(row) + ')');
        if (eb.disabled || eb.getAttribute('aria-disabled') === 'true') throw new Error('문진·검사 버튼이 비활성 상태 (' + rowSummary(row) + ')');
        eb.click();

        pstep('2/6 임상검사값 입력 화면으로 이동 중…');
        const first = await waitFor(() => findDialogButton('임상검사값입력하기') || (onPatientPage() ? 'page' : null), 8000, 200);
        if (!first) {
          const er = findErrorDialog();
          if (er) { const t = er.text; er.ok.click(); throw new Error('포괄평가 입력 화면으로 갈 수 없음: ' + t); }
          throw new Error('문진·검사 버튼을 눌러도 반응이 없음 — 아직 입력할 수 없는 상태일 수 있어요 (' + rowSummary(row) + ')' + screenNote());
        }
        if (first !== 'page') first.click();
        const ready = await waitFor(
          () => { dismissCarePlan(); return onPatientPage() && headerBirth6() === it.birth6 && (document.body.innerText || '').includes(it.name); },
          20000, 300
        );
        if (!ready) throw new Error('환자 포괄평가 화면으로 이동하지 못함(또는 환자 정보가 다름)' + screenNote());
        await sleep(800);
        await clearCarePlan();

        pstep('3/6 계획수립일자 선택 중…');
        const dateBtn = findPlanDateBtn();
        if (!dateBtn) throw new Error('계획 수립 일자 입력칸을 찾지 못함');
        dateBtn.click();
        const grid = await waitFor(() => calReady(), 4000);
        if (!grid) throw new Error('달력이 열리지 않음');
        if (!(await gotoMonth(it.date.slice(0, 7)))) throw new Error('달력에서 해당 달로 이동하지 못함');
        const dayBtn = calDayButton(it.date);
        if (!dayBtn) throw new Error('달력에서 날짜를 찾지 못함');
        if (calDisabled(dayBtn)) {
          throw new Error('계획수립일자로 선택할 수 없는 날짜(달력에서 비활성)' + cycleNote(it.date));
        }
        if (dry) {
          dayBtn.style.outline = '3px solid #e53935';
          const miss = [];
          if (!emrTrigger()) miss.push('EMR 결과 가져오기');
          if (!saveBtn()) miss.push('임시저장(심평원)');
          if (!findByText(['button'], '심평원 바로가기')) miss.push('심평원 바로가기');
          if (miss.length) throw new Error('버튼을 찾지 못함: ' + miss.join(', '));
          return '연습 완료: 환자 화면·날짜 위치·버튼 확인(날짜 클릭·저장 안 함)';
        }
        dayBtn.click();
        const want = it.date.replace(/-/g, '');
        const filled = await waitFor(() => {
          const b = findPlanDateBtn();
          const d = b ? digits(b.textContent) : '';
          return d && (d === want || d === want.slice(2));
        }, 4000);
        if (!filled) throw new Error('계획수립일자가 입력되지 않음');

        pstep('4/6 EMR 검사 결과 일괄 반영 중…');
        const trig = emrTrigger();
        if (!trig) throw new Error('"EMR 결과 가져오기" 선택 상자를 찾지 못함');
        if (!(await selectOption(trig, '검사결과일괄반영'))) throw new Error('"검사 결과 일괄 반영" 항목을 선택하지 못함' + screenNote());
        const r1 = await watchPopups(7000, ['반영']);
        if (r1.kind === 'fail') throw new Error(withLabel('EMR 결과 반영 실패', r1.text));
        if (r1.kind === 'unknown') it.note = 'EMR 반영 뒤 팝업: ' + r1.text;

        await sleep(800); // 일괄 반영 값이 화면에 그려질 때까지
        failIfRequiredMissing(/총콜레스테롤|HDL|LDL/); // 이 셋은 자동계산으로 채울 수 있어서 다음 단계에서 확인
        pstep('4/6 콜레스테롤 자동계산 확인 중…');
        const ldlNote = await fillLipid(it);
        if (ldlNote) it.note = (it.note ? it.note + ' · ' : '') + ldlNote;
        failIfRequiredMissing();
        const clrNote = await clearOptionalLabs();
        if (clrNote) it.note = (it.note ? it.note + ' · ' : '') + clrNote;

        pstep('5/6 임시저장 중…');
        const sb = saveBtn();
        if (!sb) throw new Error('임시저장(심평원) 버튼을 찾지 못함');
        sb.click();
        const cf = await waitFor(() => findConfirm(['저장하시겠습니까']), 6000, 200);
        if (!cf) {
          const er = findErrorDialog();
          if (er) { const t = er.text; er.ok.click(); throw new Error(withLabel('임시저장 실패', t)); }
          let hint = '';
          try {
            const toast = Array.from(document.querySelectorAll('[role="alert"], [class*="toast"], [class*="Toast"], [class*="snackbar"]')).map((e) => (e.textContent || '').trim()).filter(Boolean).join(' / ');
            if (toast) hint += ' · 알림 문구: ' + toast.slice(0, 120);
            const empties = Array.from(document.querySelectorAll('tr')).filter((tr) => {
              const c0 = tr.children[0] ? tr.children[0].textContent : '';
              const c1 = tr.children[1] ? tr.children[1].textContent : '';
              return (c0.indexOf('*') >= 0 || c1.indexOf('*') >= 0) && Array.from(tr.querySelectorAll('input:not([type="checkbox"]):not([type="radio"]):not([type="hidden"])')).some((i) => !i.value);
            }).map((tr) => norm(tr.children[0].textContent).slice(0, 10));
            if (empties.length) hint += ' · 비어 있는 필수(*) 항목: ' + Array.from(new Set(empties)).join(', ');
          } catch (e) {}
          throw new Error('임시저장 확인 팝업이 뜨지 않음' + hint + screenNote());
        }
        cf.btn.click();
        const r2 = await watchPopups(7000, []);
        if (r2.kind === 'fail') throw new Error(withLabel('임시저장 실패', r2.text));
        if (r2.kind === 'unknown') it.note = (it.note ? it.note + ' · ' : '') + '임시저장 뒤 팝업: ' + r2.text;

        if (saveOnly()) return '임시저장 완료(심평원 생략)';
        pstep('6/6 심평원 바로가기…');
        const hb = findByText(['button'], '심평원 바로가기');
        if (!hb) throw new Error('임시저장은 됐지만 심평원 바로가기 버튼을 찾지 못함 (심평원에서 직접 작성완료 필요)');
        pSignal = null;
        hb.click();
        return '임시저장 완료 → 심평원 이동함';
      }

      // ============================================================
      // 점검평가: 포괄평가 문진값(생활습관·체중)을 읽어 점검평가에 입력
      //   규칙: 식습관개선=달성 고정 / 신체활동=포괄평가 '목표값' 그대로 /
      //         흡연·음주는 현재 하고 있으면 '개선되지 않음', 금연·금주(·절주)면 그대로 / 체중=그대로
      //         복부둘레·BMI는 채우지 않음
      // ============================================================
      const cInputs = (root) => root.querySelectorAll('input[type="checkbox"], input[type="radio"]');
      function cLabel(input) { // 선택 항목(라디오·체크박스)의 글자: 그 항목 하나만 담은 가장 가까운 상위 요소의 글자
        let el = input;
        for (let i = 0; i < 5; i++) {
          el = el.parentElement;
          if (!el) break;
          const n = cInputs(el).length;
          if (n === 1) {
            const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
            if (t) return t;
          } else if (n > 1) break;
        }
        return '';
      }
      const textInput = (root) => root.querySelector('input:not([type="checkbox"]):not([type="radio"]):not([type="hidden"])');
      const headerCell = (tr) => Array.from(tr.children).find((c) => c.tagName === 'TH') || null;
      // 항목 이름 칸: 행의 앞 두 칸(th/td 무관) 중 글자가 일치하는 칸 (예: '생활습관'이 앞에 붙은 '흡연' 행)
      const labelCells = (tr) => Array.from(tr.children).slice(0, 2);
      function rowByHeader(hNorm) {
        return Array.from(document.querySelectorAll('tr')).find((tr) => labelCells(tr).some((c) => norm(c.textContent) === hNorm)) || null;
      }
      function weightRow() {
        return Array.from(document.querySelectorAll('tr')).find((tr) =>
          labelCells(tr).some((c) => norm(c.textContent).indexOf('체중') === 0) &&
          tr.querySelector('input:not([type="checkbox"]):not([type="radio"]):not([type="hidden"])')) || null;
      }
      // 문진표가 화면 아래쪽에 있어 스크롤해야 보이는 구조라서, 스크롤 가능한 영역을 조금씩 내려 가며 찾음
      function scrollPage(y) {
        try {
          document.querySelectorAll('*').forEach((e) => {
            if (e.scrollHeight > e.clientHeight + 50 && /(auto|scroll)/.test(getComputedStyle(e).overflowY)) e.scrollTop = y;
          });
          window.scrollTo(0, y);
        } catch (e) {}
      }
      function rowLabelsNote() {
        return ' · 화면의 항목 이름: ' + Array.from(document.querySelectorAll('tr'))
          .filter((tr) => tr.querySelector('input[type="radio"], input[type="checkbox"]'))
          .map((tr) => labelCells(tr).map((c) => norm(c.textContent).slice(0, 8)).join('/')).slice(0, 14).join(' | ');
      }
      const checkedLabels = (row) => Array.from(row.querySelectorAll('input[type="radio"]')).filter((i) => i.checked).map(cLabel);

      function findTab(textNorm) { // 가장 안쪽에 있는 일치 요소 (탭 글자)
        const all = Array.from(document.querySelectorAll('button, a, [role="tab"], li, div, span'))
          .filter((e) => e.children.length <= 3 && norm(e.textContent) === textNorm);
        const inner = all.filter((e) => !all.some((o) => o !== e && e.contains(o)));
        return inner[0] || null;
      }

      function readPcpValues() { // 포괄평가 문진 화면에서 읽기 (값은 바꾸지 않음)
        const sRow = rowByHeader('흡연');
        const aRow = rowByHeader('음주');
        const pRow = rowByHeader('신체활동');
        const wRow = weightRow();
        const miss = [];
        if (!sRow) miss.push('흡연');
        if (!aRow) miss.push('음주');
        if (!pRow) miss.push('신체활동');
        if (!wRow) miss.push('체중');
        if (miss.length) throw new Error('포괄평가 문진 화면에서 ' + miss.join(', ') + ' 항목을 찾지 못함');
        const out = { smoke: checkedLabels(sRow), alcohol: checkedLabels(aRow), walk: null, ex: null };
        const cbs = Array.from(pRow.querySelectorAll('input[type="checkbox"]'));
        const walkCb = cbs.find((i) => norm(cLabel(i)).indexOf('걷기') >= 0);
        const exCb = cbs.find((i) => norm(cLabel(i)).indexOf('운동주') >= 0);
        if (walkCb && walkCb.checked) {
          const td = walkCb.closest('td') || pRow;
          // 같은 칸에 '운동함/안 함' 같은 다른 선택지도 있어서, 걷기 구간(○천보…) 선택지만 봄
          const r = Array.from(td.querySelectorAll('input[type="radio"]')).find((i) => i.checked && /보/.test(cLabel(i)));
          out.walk = r ? cLabel(r) : '';
        }
        if (exCb && exCb.checked) {
          const td = exCb.closest('td') || pRow;
          const ti = textInput(td);
          out.ex = ti ? ti.value.trim() : '';
        }
        const wi = textInput(wRow);
        out.weight = wi ? wi.value.trim() : '';
        return out;
      }

      function buildPlan(v) { // 읽은 값 → 점검평가에 넣을 값 (규칙 밖이면 추측하지 않고 오류)
        const has = (arr, s) => arr.some((l) => norm(l).indexOf(s) >= 0);
        const plan = { smoke: null, alcohol: null, walk: null, ex: null, weight: null };
        if (has(v.smoke, '현재흡연')) plan.smoke = '개선되지 않음';
        else if (has(v.smoke, '금연')) plan.smoke = '금연';
        else if (has(v.smoke, '흡연하지않음')) plan.smoke = '금연'; // 기존 내역 없으면 금연으로 고정
        else throw new Error('포괄평가 문진의 흡연 값이 선택되어 있지 않음');
        if (has(v.alcohol, '현재음주')) plan.alcohol = '개선되지 않음';
        else if (has(v.alcohol, '금주')) plan.alcohol = '금주';
        else if (has(v.alcohol, '절주')) plan.alcohol = '절주';
        else if (has(v.alcohol, '음주하지않음')) plan.alcohol = '금주'; // 기존 내역 없으면 금주로 고정
        else throw new Error('포괄평가 문진의 음주 값이 선택되어 있지 않음');
        if (v.walk !== null) {
          if (!v.walk) throw new Error('포괄평가에서 걷기 목표는 체크돼 있는데 걷기 구간이 선택되어 있지 않음');
          plan.walk = v.walk;
        }
        if (v.ex !== null) {
          if (!/^\d+$/.test(v.ex)) throw new Error('포괄평가의 운동 주 N회 목표 값이 숫자가 아님 (' + (v.ex || '빈칸') + ')');
          plan.ex = v.ex;
        }
        if (v.walk === null && v.ex === null) { plan.walk = '2천보 미만'; plan.ex = '1'; } // 포괄평가 신체활동 목표가 없으면 걷기 2천보 미만 + 운동 주 1회로 고정
        if (!/^\d+(\.\d+)?$/.test(v.weight || '')) throw new Error('포괄평가의 체중 값이 올바르지 않음 (' + (v.weight || '빈칸') + ')');
        plan.weight = v.weight;
        return plan;
      }
      function planText(v, plan) {
        const j = (a) => (a.length ? a.join('/') : '선택 없음');
        return '읽은 값[흡연 ' + j(v.smoke) + ' · 음주 ' + j(v.alcohol) + ' · 걷기목표 ' + (v.walk === null ? '없음' : v.walk) +
          ' · 운동목표 ' + (v.ex === null ? '없음' : v.ex + '회') + ' · 체중 ' + (v.weight || '빈칸') + '] → 넣을 값[흡연 ' + (plan.smoke || '비움') +
          ' · 음주 ' + (plan.alcohol || '비움') + ' · 식습관 달성 · 걷기 ' + (plan.walk || '없음') + ' · 운동 ' + (plan.ex ? '주 ' + plan.ex + '회' : '없음') +
          ' · 체중 ' + plan.weight + ']';
      }

      // ---- 점검평가 화면 도우미 ----
      function assessRounds() { // 1차, 2차… 머리줄과 작성상태
        return Array.from(document.querySelectorAll('label'))
          .filter((l) => /^\d+차$/.test(norm(l.textContent)))
          .map((l) => {
            const hdr = l.closest('[class*="monitoring-assessment_table_header"]') || l.parentElement;
            const wrap = hdr.closest('[class*="monitoring-assessment_table_wrapper"]') || hdr.parentElement;
            const si = Array.from(hdr.querySelectorAll('input')).find((i) => i.type !== 'checkbox');
            return { label: norm(l.textContent), hdr, wrap, status: si ? si.value.trim() : '' };
          });
      }
      // 1차·2차가 가로로 나란히 있으면 표가 모두 한 상자 안에 있음 → 가로 위치로 어느 차수의 칸인지 구분
      // (각 요소는, 머리줄의 왼쪽 위치가 자기보다 왼쪽이면서 가장 가까운 차수에 속함)
      function roundLabelOf(el) {
        const x = el.getBoundingClientRect().left + 80;
        let best = null;
        assessRounds().forEach((r) => {
          const left = r.hdr.getBoundingClientRect().left;
          if (left <= x && (!best || left > best.left)) best = { label: r.label, left };
        });
        return best ? best.label : null;
      }
      const inRound = (els, label) => els.filter((e) => roundLabelOf(e) === label);
      function medGroups(scope, label) {
        return inRound(Array.from(scope.querySelectorAll('[class*="monitoring-assessment_radio_group_flex"]')), label);
      }
      function medSevere(scope, label) {
        return inRound(Array.from(scope.querySelectorAll('input[type="checkbox"]')), label).find((i) => norm(cLabel(i)).indexOf('중증사례발생이력없음') >= 0) || null;
      }
      function tableByHeaders(root, hs, label) {
        let ts = Array.from(root.querySelectorAll('table')).filter((t) => {
          const x = norm(t.textContent);
          return hs.every((h) => x.indexOf(h) >= 0);
        });
        if (label) ts = inRound(ts, label);
        ts.sort((a, b) => a.textContent.length - b.textContent.length);
        return ts[0] || null;
      }
      const resultCells = (table) => {
        const tr = Array.from(table.querySelectorAll('tr')).find((r) => r.children[0] && norm(r.children[0].textContent) === '결과');
        return tr ? Array.from(tr.children).slice(1) : null;
      };
      function setText(el, v) {
        if (el.focus) el.focus();
        setInputValue(el, v);
        el.dispatchEvent(new Event('change', { bubbles: true }));
        try { el.dispatchEvent(new FocusEvent('focusout', { bubbles: true })); } catch (e) {}
        try { el.dispatchEvent(new FocusEvent('blur')); } catch (e) {}
      }
      async function chooseRadio(cell, labelNorm) {
        const r = Array.from(cell.querySelectorAll('input[type="radio"]')).find((i) => norm(cLabel(i)) === labelNorm);
        if (!r) throw new Error('점검평가에서 "' + labelNorm + '" 선택 항목을 찾지 못함');
        if (r.disabled) await waitFor(() => !r.disabled, 1500, 100);
        if (r.disabled) throw new Error('점검평가 "' + labelNorm + '" 선택칸이 비활성 상태예요 (앞 단계 선택이 필요하거나 입력할 수 없는 상태)');
        try { r.scrollIntoView({ block: 'center' }); } catch (e) {}
        const lab = r.closest('label') || r.parentElement;
        const tries = [
          () => r.click(),
          () => { if (lab) realClick(lab); },
          () => { if (lab) lab.click(); },
          () => { ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'].forEach((t) => r.dispatchEvent(new MouseEvent(t, { bubbles: true, cancelable: true }))); },
        ];
        for (const t of tries) {
          if (r.checked) break;
          try { t(); } catch (e) {}
          await waitFor(() => r.checked, 700, 100);
        }
        if (!r.checked) throw new Error('점검평가 "' + labelNorm + '" 선택이 반영되지 않음 (' + (r.isConnected ? '선택칸이 눌리지 않음' : '화면이 다시 그려져 선택칸이 바뀜') + ')');
      }
      async function setChecked(input, want) {
        if (input.checked !== want) { input.click(); await sleep(250); }
        return input.checked === want;
      }
      async function pickDate(dateBtn, date, label) { // 달력에서 날짜를 찾아 그 날짜 버튼을 돌려줌 (누르는 것은 호출한 쪽)
        if (dateBtn.disabled) await waitFor(() => !dateBtn.disabled, 3000, 200);
        if (dateBtn.disabled) throw new Error(label + ' 칸이 잠겨 있어 달력을 열 수 없음 (차수 체크가 반영되지 않았거나 입력할 수 없는 차수)');
        try { dateBtn.scrollIntoView({ block: 'center' }); } catch (e) {}
        // 화면이 늦게 반응할 때가 있어서 여러 방법으로 최대 3번 눌러 봄
        const opens = [() => dateBtn.click(), () => realClick(dateBtn), () => { const i = dateBtn.querySelector('svg, i, span'); realClick(i || dateBtn); }];
        let grid = null;
        for (const o of opens) {
          try { o(); } catch (e) {}
          grid = await waitFor(() => calReady(), 2500);
          if (grid) break;
        }
        if (!grid) throw new Error(label + ' 달력이 열리지 않음 (3번 눌러 봄)');
        if (!(await gotoMonth(date.slice(0, 7)))) throw new Error('달력에서 해당 달로 이동하지 못함');
        const dayBtn = calDayButton(date);
        if (!dayBtn) throw new Error('달력에서 날짜를 찾지 못함');
        if (calDisabled(dayBtn)) {
          throw new Error(label + '로 선택할 수 없는 날짜(달력에서 비활성)' + cycleNote(date));
        }
        return dayBtn;
      }
      function dateIs(btn, date) {
        if (!btn) return false;
        const want = date.replace(/-/g, '');
        const d = digits(btn.textContent);
        return !!d && (d === want || d === want.slice(2));
      }
      async function resolveBirth(it) {
        if (it.birth6) return;
        const births = new Map();
        const collect = (r) => { if (r.name === it.name) births.set(r.birth6, true); return false; };
        const inp = searchInput();
        if (inp) { await pressEnterSearch(inp, it.name); await sleep(500); pcpRows().forEach(collect); }
        if (!births.size) { await clearSearch(); await scanAllPages(null, () => pcpRows(), collect); }
        await clearSearch();
        const list = Array.from(births.keys());
        if (!list.length) throw new Error('대상자 목록에서 "' + it.name + '" 이름을 찾지 못함 (엑셀에 생년월일이 없어 이름으로만 찾음)');
        if (list.length > 1) throw new Error('동명이인 ' + list.length + '명(생년월일 ' + list.join(', ') + ') — 엑셀에 생년월일이 없어 확인 필요');
        it.birth6 = list[0];
        it.note = (it.note ? it.note + ' · ' : '') + '생년월일 ' + list[0] + '(웰체크에서 찾음)';
      }
      async function findPatientRow(it) {
        await resolveBirth(it);
        const same = (r) => r.name === it.name && r.birth6 === it.birth6;
        const pre = pcpRows();
        let rows = pre.filter(same);
        let scanNote = '';
        if (rows.length === 0) {
          const inp = searchInput();
          if (inp) { await pressEnterSearch(inp, it.name); rows = pcpRows().filter(same); }
        }
        if (rows.length === 0) {
          await clearSearch();
          const sc = await scanAllPages(null, () => pcpRows(), same);
          rows = pcpRows().filter(same);
          scanNote = sc.pages ? ' · ' + sc.tried + '쪽까지 확인함' : ' · 쪽 이동 버튼을 찾지 못해 현재 쪽만 확인함';
        }
        if (rows.length === 0) {
          const sameBirth = pre.filter((r) => r.birth6 === it.birth6).map((r) => r.name);
          const sameName = pcpRows().concat(pre).filter((r) => r.name === it.name).map((r) => r.birth6);
          let hint = '';
          if (sameBirth.length) hint += ' · 같은 생년월일의 이름: ' + Array.from(new Set(sameBirth)).join(', ');
          if (sameName.length) hint += ' · 같은 이름의 생년월일: ' + Array.from(new Set(sameName)).join(', ');
          if (!hint) hint = ' · 완료·중단 환자는 목록에서 숨겨져 있을 수 있어요';
          throw new Error('대상자 목록에서 환자를 찾지 못함' + hint + scanNote);
        }
        if (rows.length > 1) throw new Error('이름·생년월일이 같은 환자가 2명 이상');
        return rows[0];
      }
      async function saveAndOpenHira(it, n) { // 임시저장 → (확인) → 심평원 바로가기
        pstep(n + ' 임시저장 중…');
        const sb = saveBtn();
        if (!sb) throw new Error('임시저장(심평원) 버튼을 찾지 못함');
        sb.click();
        const cf = await waitFor(() => findConfirm(['저장하시겠습니까']), 6000, 200);
        if (!cf) {
          const er = findErrorDialog();
          if (er) { const t = er.text; er.ok.click(); throw new Error(withLabel('임시저장 실패', t)); }
          let hint = '';
          try {
            const toast = Array.from(document.querySelectorAll('[role="alert"], [class*="toast"], [class*="Toast"], [class*="snackbar"]')).map((e) => (e.textContent || '').trim()).filter(Boolean).join(' / ');
            if (toast) hint += ' · 알림 문구: ' + toast.slice(0, 120);
            const empties = Array.from(document.querySelectorAll('tr')).filter((tr) => {
              const c0 = tr.children[0] ? tr.children[0].textContent : '';
              const c1 = tr.children[1] ? tr.children[1].textContent : '';
              return (c0.indexOf('*') >= 0 || c1.indexOf('*') >= 0) && Array.from(tr.querySelectorAll('input:not([type="checkbox"]):not([type="radio"]):not([type="hidden"])')).some((i) => !i.value);
            }).map((tr) => norm(tr.children[0].textContent).slice(0, 10));
            if (empties.length) hint += ' · 비어 있는 필수(*) 항목: ' + Array.from(new Set(empties)).join(', ');
          } catch (e) {}
          throw new Error('임시저장 확인 팝업이 뜨지 않음' + hint + screenNote());
        }
        cf.btn.click();
        const r2 = await watchPopups(7000, []);
        if (r2.kind === 'fail') throw new Error(withLabel('임시저장 실패', r2.text));
        if (r2.kind === 'unknown') it.note = (it.note ? it.note + ' · ' : '') + '임시저장 뒤 팝업: ' + r2.text;
        if (saveOnly()) return '임시저장 완료(심평원 생략)';
        pstep('심평원 바로가기…');
        const hb = findByText(['button'], '심평원 바로가기');
        if (!hb) throw new Error('임시저장은 됐지만 심평원 바로가기 버튼을 찾지 못함 (심평원에서 직접 작성완료 필요)');
        pSignal = null;
        hb.click();
        return '임시저장 완료 → 심평원 이동함';
      }

      async function runAssessItem(it, dry) {
        pstep('1/8 대상자 목록에서 환자 찾는 중…');
        if (openDialogs()[0]) throw new Error('열려 있는 팝업이 있어요. 닫고 다시 눌러 주세요');
        if (!onListPage() && !(await backToList())) throw new Error('대상자 목록 화면에서 시작해 주세요 (목록으로 돌아가지 못함)');
        await clearSearch();
        const row = await findPatientRow(it);

        pstep('2/8 환자 화면으로 이동 중…');
        const nc = row.nameCell;
        const patientOpen = () => { dismissCarePlan(); return !!findTab('점검/평가') && headerBirth6() === it.birth6 && (document.body.innerText || '').indexOf(it.name) >= 0; };
        nc.click();
        let opened = await waitFor(patientOpen, 7000, 300);
        if (!opened) {
          const inner = nc.querySelector('a, button, span');
          if (inner) { inner.click(); opened = await waitFor(patientOpen, 7000, 300); }
        }
        if (!opened) throw new Error('이름을 눌러도 환자 화면으로 이동하지 못함(또는 환자 정보가 다름)' + screenNote());
        await sleep(600);
        await clearCarePlan();

        pstep('4/8 점검/평가 화면 확인 중…');
        const t2 = findTab('점검/평가');
        if (!t2) throw new Error('"점검/평가" 탭을 찾지 못함');
        t2.click();
        const aok = await waitFor(() => assessRounds().length > 0, 10000, 300);
        if (!aok) throw new Error('점검/평가 화면에서 차수(1차 등)를 찾지 못함' + screenNote());
        await sleep(1800); // 작성상태(작성완료 등)가 늦게 채워질 수 있어서 충분히 기다린 뒤 차수를 고름
        const rounds = assessRounds();
        // 1차→2차… 순서대로: 아직 비어 있는 첫 차수 (연습 모드는 저장하지 않으므로 같은 환자의 n번째 날짜는 n번째 빈 차수로 확인)
        // 작성완료가 아닌 첫 차수: 빈 차수 + '임시저장' 상태(일부만 입력된 차수는 이어서 빈칸만 채움)
        const workable = (r) => !/완료/.test(r.status);
        const blanks = rounds.filter(workable);
        const pickRound = () => assessRounds().filter(workable)[dry ? (it.roundSkip || 0) : 0];
        let round = blanks[dry ? (it.roundSkip || 0) : 0];
        if (!round) throw new Error('작성할 수 있는 차수가 없음 (' + rounds.map((r) => r.label + ':' + (r.status || '빈칸')).join(', ') + ')');
        // 화면이 다시 그려져도 되도록 요소는 필요할 때마다 새로 찾음
        const locate = () => {
          const rd = assessRounds().find((r) => r.label === round.label) || round;
          const scope = rd.wrap || document;
          const lt = tableByHeaders(scope, ['식습관개선', '신체활동'], round.label) || tableByHeaders(document, ['식습관개선', '신체활동'], round.label);
          const bt = tableByHeaders(scope, ['복부둘레', '체중', 'BMI'], round.label) || tableByHeaders(document, ['복부둘레', '체중', 'BMI'], round.label);
          return {
            rd, scope,
            dateBtn: rd.hdr.querySelector(DATE_SEL),
            lcells: lt ? resultCells(lt) : null,
            bcells: bt ? resultCells(bt) : null,
          };
        };
        let L0 = locate();
        if (!L0.dateBtn) throw new Error(round.label + ' 실시 일자 입력칸을 찾지 못함');
        const miss = [];
        if (!L0.lcells || L0.lcells.length < 4) miss.push('생활습관 결과 칸');
        if (!L0.bcells || L0.bcells.length < 2) miss.push('체중 결과 칸');
        if (!emrTrigger()) miss.push('EMR 결과 가져오기');
        if (!saveBtn()) miss.push('임시저장(심평원)');
        if (!findByText(['button'], '심평원 바로가기')) miss.push('심평원 바로가기');
        if (medGroups(L0.scope, round.label).length < 4) miss.push('약물모니터링 4항목');
        if (!medSevere(L0.scope, round.label)) miss.push('중증사례 발생 이력 없음 체크칸');
        if (miss.length) throw new Error('점검평가 화면에서 못 찾은 항목: ' + miss.join(', '));

        // 환자가 문진(생활습관·체중)을 직접 입력했으면 그 값을 그대로 두고, 비어 있는 항목만 포괄평가 값으로 채움
        const anyRadio = (c) => !!c && Array.from(c.querySelectorAll('input[type="radio"]')).some((i) => i.checked);
        const anyCheck = (c) => !!c && Array.from(c.querySelectorAll('input[type="checkbox"]')).some((i) => i.checked);
        const wiPre = L0.bcells[1] ? textInput(L0.bcells[1]) : null;
        const pre = {
          smoke: anyRadio(L0.lcells[0]),
          alcohol: anyRadio(L0.lcells[1]),
          diet: anyRadio(L0.lcells[2]),
          act: anyCheck(L0.lcells[3]) || anyRadio(L0.lcells[4]),
          weight: !!(wiPre && wiPre.value.trim()),
        };
        const preLabel = round.label;
        const empties = Object.keys(pre).filter((k) => !pre[k]);
        let plan = null;
        if (empties.length === 0) {
          it.planText = '생활습관·체중이 이미 입력되어 있어 포괄평가 값은 가져오지 않음';
        } else {
          pstep('3/8 포괄평가 문진값 읽는 중… (비어 있는 항목: ' + empties.join(', ') + ')');
          const t1 = findTab('포괄평가계획수립');
          if (!t1) throw new Error('"포괄평가 계획수립" 탭을 찾지 못함');
          t1.click();
          let sy = 0;
          const readable = await waitFor(() => { dismissCarePlan(); sy += 500; scrollPage(sy); return rowByHeader('흡연') && weightRow(); }, 12000, 400);
          if (!readable) { scrollPage(0); throw new Error('포괄평가 문진 화면(흡연·체중 항목)을 찾지 못함' + rowLabelsNote() + screenNote()); }
          try { rowByHeader('흡연').scrollIntoView({ block: 'center' }); } catch (e) {}
          await sleep(900);
          // 문진 값은 화면이 열린 뒤 조금 늦게 채워질 수 있어서, 흡연·음주 선택이 반영될 때까지 기다림
          await waitFor(() => {
            const a = rowByHeader('흡연'), b = rowByHeader('음주');
            return a && b && checkedLabels(a).length > 0 && checkedLabels(b).length > 0;
          }, 8000, 400);
          await sleep(400);
          const v = readPcpValues();
          scrollPage(0);
          plan = buildPlan(v);
          it.planText = planText(v, plan) + ' · 이미 입력된 항목은 그대로 둠: ' + (Object.keys(pre).filter((k) => pre[k]).join(', ') || '없음');
          const t2b = findTab('점검/평가');
          if (!t2b) throw new Error('"점검/평가" 탭으로 돌아가지 못함');
          t2b.click();
          if (!(await waitFor(() => assessRounds().length > 0, 10000, 300))) throw new Error('점검/평가 화면으로 돌아오지 못함' + screenNote());
          await sleep(700);
          await sleep(1200);
          round = pickRound() || round; // 탭을 오가면 화면이 다시 그려지므로 작성상태를 보고 차수를 다시 고름
          L0 = locate();
          if (!L0.dateBtn || !L0.lcells || !L0.bcells) throw new Error('점검/평가 화면으로 돌아온 뒤 입력칸을 다시 찾지 못함');
        }

        if (dry) {
          let dnote = '';
          if (!L0.dateBtn.disabled) {
            const dayBtn = await pickDate(L0.dateBtn, it.date, '점검평가 실시일자');
            dayBtn.style.outline = '3px solid #e53935';
          } else {
            dnote = ' (날짜 칸은 차수를 체크해야 열리는 것 같아 연습에서는 확인하지 않음)';
          }
          return '연습 완료: ' + round.label + ' 입력 예정, 화면 항목 확인됨' + dnote;
        }

        pstep('5/8 ' + round.label + ' 선택·실시일자 입력 중…');
        { // 누르기 직전에 한 번 더: 작성상태를 새로 읽어 '작성완료'가 아닌 첫 차수인지 확인
          const rr = pickRound();
          if (!rr) throw new Error('작성할 수 있는 차수가 없음 (' + assessRounds().map((r) => r.label + ':' + (r.status || '빈칸')).join(', ') + ')');
          if (rr.label !== preLabel) throw new Error('작성할 차수가 중간에 ' + preLabel + '에서 ' + rr.label + '로 바뀌었어요 (작성상태가 늦게 표시된 것 같아요). 다시 실행해 주세요');
          round = rr;
        }
        const rcb = round.hdr.querySelector('input[type="checkbox"]');
        if (rcb && rcb.disabled && !rcb.checked) throw new Error(round.label + ' 선택칸이 비활성 상태예요 (앞 차수가 작성되어야 하거나 작성할 수 없는 차수)');
        if (rcb && !rcb.checked) { rcb.click(); await sleep(500); }
        await waitFor(() => { const d = locate().dateBtn; return d && !d.disabled; }, 4000, 200);
        const curDate = locate().dateBtn;
        if (curDate && dateIs(curDate, it.date)) {
          // 이미 요청한 날짜와 같음
        } else {
          // 비어 있거나 다른 날짜(임시저장 때 입력된 날짜 등)면 요청한 날짜로 바꿈
          const had = curDate && !curDate.textContent.includes('--') ? curDate.textContent.trim() : '';
          const dayBtn = await pickDate(locate().dateBtn, it.date, '점검평가 실시일자');
          dayBtn.click();
          if (!(await waitFor(() => dateIs(locate().dateBtn, it.date), 4000))) throw new Error('점검평가 실시일자가 입력되지 않음');
          if (had) it.note = (it.note ? it.note + ' · ' : '') + round.label + ' 실시일자를 ' + had + ' → ' + it.date + '로 변경';
        }

        pstep('6/8 EMR 검사 결과 일괄 반영 중…');
        const trig = emrTrigger();
        if (!(await selectOption(trig, '검사결과일괄반영'))) throw new Error('"검사 결과 일괄 반영" 항목을 선택하지 못함' + screenNote());
        const r1 = await watchPopups(7000, ['반영']);
        if (r1.kind === 'fail') throw new Error(withLabel('EMR 결과 반영 실패', r1.text));
        if (r1.kind === 'unknown') it.note = 'EMR 반영 뒤 팝업: ' + r1.text;
        await sleep(800); // 일괄 반영 값이 화면에 그려질 때까지
        failIfRequiredMissing();
        const clrNoteA = await clearOptionalLabs();
        if (clrNoteA) it.note = (it.note ? it.note + ' · ' : '') + clrNoteA;

        pstep('7/8 생활습관·체중 입력 중…');
        const L = locate();
        if (!L.lcells || L.lcells.length < 4 || !L.bcells || L.bcells.length < 2) throw new Error('점검평가 생활습관·체중 결과 칸을 다시 찾지 못함');
        const lcells = L.lcells;
        const bcells = L.bcells;
        try { lcells[0].scrollIntoView({ block: 'center' }); } catch (e) {}
        await sleep(300);
        if (!pre.smoke && plan && plan.smoke) await chooseRadio(lcells[0], norm(plan.smoke));
        if (!pre.alcohol && plan && plan.alcohol) await chooseRadio(lcells[1], norm(plan.alcohol));
        if (!pre.diet) await chooseRadio(lcells[2], '달성');
        if (!pre.act && plan && (plan.walk || plan.ex)) {
          const cbs = Array.from(lcells[3].querySelectorAll('input[type="checkbox"]'));
          if (plan.walk) {
            const wcb = cbs.find((i) => norm(cLabel(i)).indexOf('걷기') >= 0);
            if (!wcb || !(await setChecked(wcb, true))) throw new Error('점검평가 걷기(일평균) 체크가 반영되지 않음');
            await chooseRadio(lcells[4] || lcells[3], norm(plan.walk)); // 걷기 구간 선택칸은 신체활동 칸 오른쪽(colspan 2) 칸에 있음
          }
          if (plan.ex) {
            const ecb = cbs.find((i) => norm(cLabel(i)).indexOf('운동주') >= 0);
            const eti = textInput(lcells[3]);
            if (!ecb || !eti || !(await setChecked(ecb, true))) throw new Error('점검평가 운동 주 N회 입력칸을 찾지 못함');
            setText(eti, plan.ex);
            await sleep(200);
            if (parseFloat(eti.value) !== parseFloat(plan.ex)) throw new Error('점검평가 운동 주 횟수가 입력되지 않음');
          }
        }
        if (!pre.weight && plan) {
          const wi = textInput(bcells[1]);
          if (!wi) throw new Error('점검평가 체중 결과 입력칸을 찾지 못함');
          setText(wi, plan.weight);
          await sleep(200);
          if (parseFloat(wi.value) !== parseFloat(plan.weight)) throw new Error('점검평가 체중이 입력되지 않음 (입력칸 값: ' + (wi.value || '빈칸') + ')');
        }

        // 약물모니터링·중증사례는 고정값: 이미 선택돼 있어도 항상 예/양호/없음/아니오, '발생 이력 없음' 체크
        // (복용여부가 '아니오'면 뒤 항목이 잠겨 있으므로 앞에서부터 차례로 선택하고, 선택할 때마다 칸을 새로 찾음)
        const medVals = ['예', '양호', '없음', '아니오'];
        for (let gi = 0; gi < 4; gi++) {
          await sleep(300);
          const gs = medGroups(locate().scope, round.label);
          if (gs.length < 4) throw new Error('약물모니터링 4항목 선택칸을 찾지 못함');
          await chooseRadio(gs[gi], medVals[gi]);
        }
        const sc = medSevere(locate().scope, round.label);
        if (!sc) throw new Error('중증사례 "발생 이력 없음" 체크칸을 찾지 못함');
        if (!(await setChecked(sc, true))) throw new Error('중증사례 "발생 이력 없음" 체크가 반영되지 않음');

        return saveAndOpenHira(it, '8/8');
      }

      // ---- 패널 ----
      const pp = document.createElement('div');
      pp.id = 'wc-pcp-panel';
      pp.style.cssText =
        'position:fixed;right:12px;bottom:44px;width:330px;max-height:70vh;overflow:auto;' +
        'z-index:2147483000;background:#fff;color:#222;border:1px solid #888;border-radius:8px;' +
        'box-shadow:0 4px 16px rgba(0,0,0,.25);font:12px/1.5 sans-serif;padding:8px;display:none;';
      pp.innerHTML =
        '<div style="font-weight:700;margin-bottom:0"><span id="wcp-title">🐵 칼퇴몽키 · 포괄평가 v11.3</span> <span id="wcp-count"></span>' +
        '<span id="wcp-min" style="float:right;cursor:pointer;padding:0 6px">−</span></div>' +
        '<div style="color:#777;font-size:11px;margin-bottom:4px">클릭은 몽키가, 퇴근은 내가</div>' +
        '<div id="wcp-body">' + MODE_ROW('wcp') +
        '<div style="display:flex;gap:4px;margin-bottom:4px">' +
        '<button id="wcp-t-pcp" type="button" style="flex:1;font-weight:700;background:#e3f2fd">포괄평가</button>' +
        '<button id="wcp-t-edu" type="button" style="flex:1">교육상담</button>' +
        '<button id="wcp-t-assess" type="button" style="flex:1">점검평가</button></div>' +
        '<textarea id="wcp-text" rows="5" style="width:100%;box-sizing:border-box;font:11px monospace" ' +
        'placeholder="이름,생년월일,계획수립일자 (쉼표·탭·공백 가능)\n홍길동,960824,2026-09-28"></textarea>' +
        '<div style="display:flex;gap:4px;margin:4px 0"><button id="wcp-xl" type="button" style="flex:1">엑셀 불러오기</button>' +
        '<input id="wcp-xl-file" type="file" accept=".xlsx,.xls,.csv" style="display:none"></div>' +
        '<div style="display:flex;gap:4px;margin-bottom:4px">' +
        '<button id="wcp-apply" type="button" style="flex:1">목록 적용</button>' +
        '<button id="wcp-one" type="button" style="flex:1;font-weight:700">다음 건 실행</button>' +
        '<button id="wcp-clear" type="button" style="flex:1">비우기</button></div>' +
        '<div style="display:flex;gap:4px;margin-bottom:4px">' +
        '<button id="wcp-all" type="button" style="flex:1;font-weight:700">연속 실행</button>' +
        '<button id="wcp-stop" type="button" style="flex:1">정지</button></div>' +
        '<label style="display:block;margin:2px 0">심평원 이동 후 대기 <input id="wcp-wait" type="number" value="2" min="0" style="width:52px;border:1px solid #aaa"> 초 (심평원 창은 혼자 마무리해요)</label>' +
        '<div id="wcp-msg" style="color:#c62828;min-height:16px"></div>' +
        '<div id="wcp-report" style="display:none;border:1px solid #e65100;border-radius:4px;padding:4px;margin:4px 0">' +
        '<div id="wcp-report-body" style="white-space:pre-wrap;user-select:text"></div>' +
        '<button id="wcp-copy" type="button" style="margin-top:4px">확인 필요 목록 복사</button></div>' +
        '<div id="wcp-list"></div></div>';
      const oldPp = document.getElementById('wc-pcp-panel');
      if (oldPp) oldPp.remove();
      document.body.appendChild(pp);
      const pq = (sel) => pp.querySelector(sel);
      const pmsg = (t, color) => { const m = pq('#wcp-msg'); m.textContent = t || ''; m.style.color = color || '#c62828'; };
      function pstep(t) { pmsg(t, '#1565c0'); }

      function prender() {
        const pending = pItems.filter((x) => x.status === '대기').length;
        pq('#wcp-count').textContent = pItems.length ? '(' + pItems.length + '건, 대기 ' + pending + ')' : '';
        const list = pq('#wcp-list');
        list.innerHTML = '';
        pItems.forEach((it, idx) => {
          const d = document.createElement('div');
          d.style.cssText = 'border-top:1px solid #eee;padding:2px 0;';
          d.textContent = (idx + 1) + '. ' + it.name + ' ' + it.birth6 + ' ' + it.date + ' · ' + it.status + (it.note ? ' (' + it.note + ')' : '');
          list.appendChild(d);
        });
      }
      function pShowReport(problems) {
        const box = pq('#wcp-report');
        if (!problems.length) { box.style.display = 'none'; pReportText = ''; return; }
        const lines = problems.map((p) => p.it.name + ' ' + p.it.birth6 + ' ' + p.it.date + ' — ' + p.reason);
        pReportText = '확인이 필요한 건 ' + problems.length + '건\n' + lines.join('\n');
        pq('#wcp-report-body').textContent = pReportText;
        box.style.display = 'block';
      }
      function parsePcp(text) {
        let raw = parseQueue(text);
        const seen = new Set();
        if (task === 'assess') {
          // 점검평가: 같은 환자의 여러 날짜는 날짜 순(1차→2차…)으로 정렬, 같은 날짜 중복만 건너뜀
          const first = {};
          raw.forEach((it, i) => { const k = it.name + '|' + it.birth6; if (!(k in first)) first[k] = i; });
          raw = raw.map((it, i) => ({ it, i })).sort((a, b) => {
            const ka = a.it.name + '|' + a.it.birth6, kb = b.it.name + '|' + b.it.birth6;
            if (ka !== kb) return first[ka] - first[kb];
            return (a.it.date || '9') < (b.it.date || '9') ? -1 : (a.it.date || '9') > (b.it.date || '9') ? 1 : a.i - b.i;
          }).map((x) => x.it);
          raw.forEach((it) => {
            if (it.status !== '대기') return;
            const k = it.name + '|' + it.birth6 + '|' + it.date;
            if (seen.has(k)) it.status = '건너뜀(같은 환자·같은 날짜 중복)';
            seen.add(k);
          });
          const cnt = {};
          raw.forEach((it) => { if (it.status !== '대기') return; const k = it.name + '|' + it.birth6; it.roundSkip = cnt[k] || 0; cnt[k] = it.roundSkip + 1; });
          return raw;
        }
        raw.forEach((it) => {
          if (it.status !== '대기') return;
          const k = it.name + '|' + it.birth6;
          if (seen.has(k)) it.status = '건너뜀(같은 환자의 중복 줄)';
          seen.add(k);
        });
        return raw;
      }

      async function runAll(single) {
        if (pRunning && Date.now() - pSince < 90000) { pmsg('이미 실행 중이에요.'); return; }
        if (!pItems.some((x) => x.status === '대기')) { pmsg('대기 중인 건이 없어요. 먼저 "목록 적용"을 눌러 주세요.'); return; }
        pRunning = true; badgeReset();
        pSince = Date.now();
        pStop = false;
        pq('#wcp-report').style.display = 'none';
        const dry = false; // 연습 모드는 없앰
        const queue = dry ? pItems.filter((x) => x.status === '대기') : null;
        let qi = 0;
        let done = 0;
        let failed = 0;
        const problems = [];
        let stopReason = '';
        try {
          while (!pStop) {
            const it = dry ? queue[qi++] : pItems.find((x) => x.status === '대기');
            if (!it) break;
            pSince = Date.now();
            it.note = '';
            it.planText = '';
            it.status = '진행 중…';
            prender();
            let ok = false;
            try {
              const r = await (task === 'assess' ? runAssessItem : runPcpItem)(it, dry);
              if (dry) { it.status = '대기'; it.note = '연습 성공' + (it.planText ? ' · ' + it.planText : ''); done++; }
              else { it.status = r; ok = true; }
            } catch (e) {
              const m = (e && e.message ? e.message : String(e)) + (openDialogs()[0] ? screenNote() : '');
              console.error('[자동화] 포괄평가 실행 중 실패:', e);
              if (dry) { it.status = '대기'; it.note = '연습 실패: ' + m + (it.planText ? ' · ' + it.planText : ''); failed++; }
              else { it.status = '실패: ' + m; problems.push({ it, reason: m }); }
            }
            prender();
            if (!dry && ok && saveOnly()) { done++; prender(); }
            else if (!dry && ok) {
              // 심평원 쪽이 작성완료를 누를 때까지(신호가 오면 바로) 최대 N초 대기
              const pwv = parseInt(pq('#wcp-wait').value, 10);
              const waitMs = Math.max(0, isNaN(pwv) ? 2 : pwv) * 1000;
              const t1 = Date.now();
              while (Date.now() - t1 < waitMs && !pStop && !pSignal) { pSince = Date.now(); await sleep(500); }
              it.status = pSignal ? '심평원 작성완료 클릭됨' : '임시저장 완료(심평원은 자동 진행)';
              done++;
              prender();
            }
            if (!(await cleanup())) { stopReason = '대상자 목록으로 돌아가지 못해 멈췄어요. 화면을 확인해 주세요.'; break; }
            if (single) break;
            await sleep(800);
          }
        } catch (e) {
          console.error('[자동화] 포괄평가 연속 실행 오류:', e);
          stopReason = '오류: ' + (e && e.message ? e.message : e);
        }
        pRunning = false;
        prender();
        pShowReport(problems);
        if (dry) {
          const sm = ' (연습 성공 ' + done + '건, 실패 ' + failed + '건)';
          if (stopReason) pmsg(stopReason + sm);
          else pmsg('✔ 연습이 끝났어요. 저장·제출은 하지 않았어요.' + sm, '#2e7d32');
          return;
        }
        const summary = ' (완료 ' + done + '건, 확인 필요 ' + problems.length + '건)';
        if (stopReason) pmsg(stopReason + summary);
        else if (problems.length) pmsg('끝까지 진행했어요. 아래 ' + problems.length + '건은 확인이 필요해요.' + summary, '#e65100');
        else { pmsg('🙈 모든 건을 처리했어요. 이제 퇴근하세요!' + summary, '#2e7d32'); badgeDone(done + '건 완료, 이제 퇴근하세요'); }
      }

      function setTask(t) {
        if (t === task) return;
        if (pRunning) { pmsg('실행 중에는 업무를 바꿀 수 없어요.'); return; }
        store[task].items = pItems;
        store[task].text = pq('#wcp-text').value;
        task = t;
        pItems = store[task].items;
        pq('#wcp-text').value = store[task].text;
        const isA = t === 'assess';
        pq('#wcp-title').textContent = '🐵 칼퇴몽키 · ' + (isA ? '점검평가' : '포괄평가') + ' v11.3';
        pq('#wcp-text').placeholder = isA
          ? '이름,생년월일,점검평가일자(AX003 처방일) (쉼표·탭·공백 가능)\n홍길동,960824,2026-09-28'
          : '이름,생년월일,계획수립일자 (쉼표·탭·공백 가능)\n홍길동,960824,2026-09-28';
        pq('#wcp-t-pcp').style.background = isA ? '' : '#e3f2fd';
        pq('#wcp-t-pcp').style.fontWeight = isA ? '' : '700';
        pq('#wcp-t-assess').style.background = isA ? '#e3f2fd' : '';
        pq('#wcp-t-assess').style.fontWeight = isA ? '700' : '';
        pq('#wcp-report').style.display = 'none';
        pmsg('');
        prender();
      }
      pq('#wcp-t-pcp').addEventListener('click', () => setTask('pcp'));
      pq('#wcp-t-assess').addEventListener('click', () => setTask('assess'));
      // 교육상담 탭: 교육상담(내원) 목록으로 이동하면 교육상담 패널이 뜸
      pq('#wcp-t-edu').addEventListener('click', () => {
        if (pRunning) { pmsg('실행 중에는 업무를 바꿀 수 없어요.'); return; }
        const eduMenu = findByText(['a', 'button', 'span', 'li', 'div'], '교육상담(내원)');
        if (!eduMenu) { pmsg('"교육상담(내원)" 메뉴를 찾지 못했어요. 메뉴에서 직접 열어 주세요.'); return; }
        eduMenu.click();
      });
      pcpShared.setTask = setTask;
      pq('#wcp-min').addEventListener('click', () => {
        const b = pq('#wcp-body');
        b.style.display = b.style.display === 'none' ? '' : 'none';
      });
      pq('#wcp-m-hira').addEventListener('click', () => { if (pRunning) { pmsg('실행 중에는 모드를 바꿀 수 없어요.'); return; } setMode('hira'); });
      pq('#wcp-m-save').addEventListener('click', () => { if (pRunning) { pmsg('실행 중에는 모드를 바꿀 수 없어요.'); return; } setMode('save'); });
      applyModeUI();
      pq('#wcp-xl').addEventListener('click', () => pq('#wcp-xl-file').click());
      pq('#wcp-xl-file').addEventListener('change', (e) => {
        const f = e.target.files && e.target.files[0];
        e.target.value = '';
        if (f) importExcel(f, pmsg, task === 'assess' ? 'AX003' : 'AX001'); // 지금 선택된 탭(포괄평가/점검평가) 기준
      });
      pcpShared.isRunning = () => pRunning;
      pcpShared.loadLists = (pcpText, assessText) => { // 포괄평가·점검평가 대기열을 엑셀 결과로 채움
        const orig = task;
        if (pcpText !== null) { setTask('pcp'); pq('#wcp-text').value = pcpText; pItems = parsePcp(pcpText); }
        if (assessText !== null) { setTask('assess'); pq('#wcp-text').value = assessText; pItems = parsePcp(assessText); }
        setTask(orig);
        prender();
      };
      pq('#wcp-apply').addEventListener('click', () => {
        pItems = parsePcp(pq('#wcp-text').value);
        pmsg(pItems.length ? '' : '해석할 수 있는 줄이 없어요 (이름,생년월일,계획수립일자)');
        prender();
      });
      pq('#wcp-clear').addEventListener('click', () => {
        pq('#wcp-report').style.display = 'none';
        pItems = [];
        pq('#wcp-text').value = '';
        pmsg('');
        prender();
      });
      pq('#wcp-one').addEventListener('click', () => {
        runAll(true).catch((e) => { pRunning = false; pmsg('오류: ' + (e && e.message ? e.message : e)); });
      });
      pq('#wcp-all').addEventListener('click', () => {
        runAll(false).catch((e) => { pRunning = false; pmsg('오류: ' + (e && e.message ? e.message : e)); });
      });
      pq('#wcp-stop').addEventListener('click', () => {
        pStop = true;
        pmsg('정지 요청: 지금 진행 중인 건이 끝나면 멈춰요.', '#1565c0');
      });
      pq('#wcp-copy').addEventListener('click', async () => {
        try {
          await navigator.clipboard.writeText(pReportText);
          pmsg('확인 필요 목록을 복사했어요.', '#2e7d32');
        } catch (e) {
          pmsg('복사하지 못했어요. 위 글자를 직접 선택해서 복사해 주세요.');
        }
      });

      // 심평원 탭이 작성완료를 눌렀다는 신호
      function onPcpMsg(e) {
        if (e.origin !== 'https://biz.hira.or.kr') return;
        const t = e.data && e.data.wcAuto;
        if (t === 'done' || t === 'clicked') pSignal = t;
      }
      window.addEventListener('message', onPcpMsg);

      // 대상자 목록 또는 환자 포괄평가 화면에서만 패널 표시
      pcpShared.findPatientRow = findPatientRow;
      pcpShared.onListPage = onListPage;
      pcpShared.diag = { pcpRows, evalButton, onPatientPage, emrTrigger, saveBtn, labTableRows, labCells };
      pcpShared.resolveBirth = resolveBirth;
      const pShowId = setInterval(() => {
        pp.style.display = isPcpPatientPage() || isAssessPage() || pcpRows().length > 0 ? 'block' : 'none';
      }, 1000);

      console.log('[자동화] 포괄평가 대기열 패널 시작됨');

      return function pcpTeardown() {
        clearInterval(pShowId);
        window.removeEventListener('message', onPcpMsg);
        pp.remove();
      };
    }

    const pcpTeardown = startPcp();

    // 교육상담 화면에서 패널 표시: 교육상담(내원) 목록 화면 + 환자 화면의 '교육 상담' 탭(…/patients/○○/edu-consult)
    const onEduConsultTab = () => /\/patients\/[^/]+\/edu-consult/.test(location.pathname);
    const showId = setInterval(() => {
      const pe0 = document.getElementById('wc-pcp-panel');
      const pcpShown = pe0 && pe0.style.display === 'block';
      panel.style.display = (eduListRoot() || (onEduConsultTab() && !pcpShown) || running) ? 'block' : 'none';
      // 전체 대상자 목록으로 이동해 있을 때는 포괄평가 패널 옆으로 비켜서 표시
      const pe = document.getElementById('wc-pcp-panel');
      panel.style.right = running && pe && pe.style.display === 'block' ? '354px' : '12px';
    }, 1000);

    console.log('[자동화] 교육상담 대기열 패널 시작됨');

    // ============================================================
    // 관리자용: 칼퇴몽키 점검 + 웰체크 화면 변경 감지
    // ============================================================
    if (ADMIN) (function startDiag() {
      const path = () => location.pathname;
      const isShownEl = (el) => !!el && el.getClientRects().length > 0 && !el.closest('[hidden]');
      const textBtns = (re, root) => Array.from((root || document).querySelectorAll('button, [role="button"]')).filter((b) => re.test(norm(b.textContent)) && isShownEl(b));
      const reflDlg = () => openDialogs().find((x) => norm(x.textContent).includes('교육상담반영내용'));
      const D = () => pcpShared.diag || {};

      // 화면 종류 알아내기
      function pageKind() {
        if (/\/monitoringAssessment/.test(path()) || isAssessPage()) return 'assess';
        if (/\/edu-consult/.test(path())) return 'eduTab';
        if (/\/patientEducation/.test(path()) || eduListRoot()) return 'eduList';
        if (isPcpPatientPage()) return 'pcpPatient';
        if (D().pcpRows && D().pcpRows().length) return 'pcpList';
        return 'other';
      }
      const KIND_NAME = { eduList: '교육상담(내원) 화면', eduTab: '환자 교육 상담 탭', pcpList: '일만사 대상자 목록', pcpPatient: '포괄평가 환자 화면', assess: '점검평가 화면', other: '기타 화면' };

      // 점검 항목: need=true 는 그 화면에 늘 있어야 하는 것(변경 감지에 사용), false 는 상황에 따라 있는 것
      function checksFor(kind) {
        const c = [];
        const add = (name, need, fn, area) => c.push({ name, need, fn, area });
        if (kind === 'eduList') {
          add('환자 목록 표', true, () => !!eduListRoot(), () => document.querySelector('main'));
          add('환자 목록 행', true, () => { const r = eduListRoot(); return !!r && r.querySelectorAll('table tbody tr').length > 0; }, () => eduListRoot());
          add('환자명 검색칸', false, () => !!document.querySelector('input[placeholder*="환자명"]'), () => eduListRoot());
        }
        if (kind === 'eduList' || kind === 'eduTab') {
          add('환자 머리줄 생년월일(○○○○년 ○월 ○일(○세))', kind === 'eduTab', () => /(\d{4})년\s*\d{1,2}월\s*\d{1,2}일\s*\(\d+세\)/.test(document.body.innerText), () => document.querySelector('main'));
          add('교육자료 체크칸', kind === 'eduTab', () => materialCheckboxes(false).length > 0, () => { const t = Array.from(document.querySelectorAll('section, div')).find((e) => norm(e.textContent).startsWith('교육자료발송')); return t || document.querySelector('main'); });
          add('발송 바구니 버튼 묶음', kind === 'eduTab', () => !!document.querySelector('[aria-label="교육자료 발송 버튼"]') || textBtns(/^발송$/).some((b) => !b.closest('[role="dialog"]')), () => document.body);
          add('바구니의 "교육자료 발송" 버튼 (자료를 체크했을 때)', false, () => textBtns(/^교육자료발송$/).length > 0, () => document.querySelector('[role="dialog"][aria-label="발송 바구니"]') || document.body);
          add('반영내용 창의 날짜 칸 (창이 열렸을 때)', false, () => { const d = reflDlg(); return !!d && !!d.querySelector(DATE_SEL); }, () => reflDlg() || document.body);
          add('반영내용 창의 "발송" 버튼 (창이 열렸을 때)', false, () => { const d = reflDlg(); return !!d && textBtns(/^발송$/, d).length > 0; }, () => reflDlg() || document.body);
          add('달력 날짜 버튼 (달력이 열렸을 때)', false, () => calDayButtons().length > 0 && !!calShownYm(), () => document.body);
        }
        if (kind === 'pcpList') {
          add('대상자 목록 행', true, () => D().pcpRows().length > 0, () => document.querySelector('main'));
          add('문진·검사 버튼', true, () => D().pcpRows().some((r) => !!D().evalButton(r)), () => { const r = D().pcpRows()[0]; return r ? r.tr : document.querySelector('main'); });
          add('환자명 검색칸', true, () => !!document.querySelector('input[placeholder*="환자명"]'), () => document.querySelector('main'));
          add('쪽 이동 버튼', false, () => textBtns(/^[2-9]$/).length > 0 || document.querySelectorAll('[aria-label*="다음"], [class*="pagination"] button').length > 0, () => document.querySelector('main'));
        }
        if (kind === 'pcpPatient') {
          add('계획수립일자 칸', true, () => !!planDateButton(false), () => document.querySelector('main'));
          add('"EMR 결과 가져오기"', true, () => !!(D().emrTrigger && D().emrTrigger()), () => document.querySelector('main'));
          add('"임시저장(심평원)" 버튼', true, () => !!(D().saveBtn && D().saveBtn()), () => document.querySelector('main'));
          add('임상검사 표 (검사 줄·* 표시)', true, () => { const r = D().labTableRows ? D().labTableRows() : []; return r.length > 0 && r.some((x) => x.required); }, () => Array.from(document.querySelectorAll('table')).find((t) => /임상검사|검사항목/.test(norm(t.textContent))) || document.querySelector('main'));
          add('콜레스테롤 칸 3개 (총·HDL·LDL)', true, () => ['총콜레스테롤', 'HDL-콜레스테롤', 'LDL-콜레스테롤'].every((k) => D().labCells && D().labCells(k)), () => Array.from(document.querySelectorAll('table')).find((t) => /콜레스테롤/.test(t.textContent)) || document.body);
          add('"자동계산" 버튼', true, () => textBtns(/자동계산/).length > 0, () => Array.from(document.querySelectorAll('table')).find((t) => /콜레스테롤/.test(t.textContent)) || document.body);
          add('검사일자 칸', true, () => { const r = D().labTableRows ? D().labTableRows() : []; return r.some((x) => x.tr.querySelector(DATE_SEL)); }, () => Array.from(document.querySelectorAll('table')).find((t) => /검사항목|임상검사/.test(norm(t.textContent))) || document.body);
          add('휴지통 버튼 (값이 있을 때)', false, () => { const r = D().labTableRows ? D().labTableRows() : []; return r.some((x) => Array.from(x.tr.querySelectorAll('button')).some((b) => !norm(b.textContent) && b.querySelector('svg, img, i') && !b.matches(DATE_SEL))); }, () => document.body);
        }
        if (kind === 'assess') {
          const form = () => document.querySelector('form[class*="monitoring-assessment"]');
          add('점검평가 입력 화면', true, () => !!form(), () => document.querySelector('main'));
          add('차수 체크칸(1차·2차…)', true, () => !!form() && form().querySelectorAll('button[role="checkbox"], input[type="checkbox"]').length > 0, () => form() || document.body);
          add('실시일자 칸', true, () => !!form() && !!form().querySelector(DATE_SEL), () => form() || document.body);
          add('약물모니터링 선택지', true, () => /약물/.test(document.body.innerText) && document.querySelectorAll('input[type="radio"], button[role="radio"]').length > 0, () => form() || document.body);
          add('"EMR 결과 가져오기"', true, () => !!(D().emrTrigger && D().emrTrigger()), () => form() || document.body);
          add('임상검사 표 (검사 줄·* 표시)', true, () => { const r = D().labTableRows ? D().labTableRows() : []; return r.length > 0 && r.some((x) => x.required); }, () => form() || document.body);
          add('"임시저장(심평원)" 버튼', true, () => !!(D().saveBtn && D().saveBtn()), () => form() || document.body);
        }
        return c;
      }

      function runChecks() {
        const kind = pageKind();
        const res = checksFor(kind).map((x) => { let ok = false; try { ok = !!x.fn(); } catch (e) { ok = false; } return Object.assign({}, x, { ok }); });
        return { kind, res };
      }

      // 화면 구조 요약 (개인정보 제외): 태그·클래스·역할·이름표만, 글자는 버튼·머리글·라벨에서만, 숫자는 #로 가림
      function skeleton(root, maxLen) {
        const out = [];
        let len = 0;
        const mask = (t) => t.replace(/\d/g, '#').replace(/\s+/g, ' ').trim().slice(0, 24);
        const LABEL_TAGS = /^(BUTTON|TH|LABEL|LEGEND|OPTION)$/;
        function walk(el, depth) {
          if (len > maxLen || depth > 9 || !el || el.nodeType !== 1) return;
          if (el.id && /^wc-/.test(el.id)) return; // 칼퇴몽키 자신은 제외
          if (/^(SCRIPT|STYLE|SVG|PATH|NOSCRIPT|IMG)$/i.test(el.tagName)) return;
          const cls = (typeof el.className === 'string' ? el.className : '').split(/\s+/).filter((k) => k && !/^(flex|w-|h-|p[xytblr]?-|m[xytblr]?-|gap-|text-|bg-|border|rounded|items-|justify-|min-|max-|shadow|leading|font-|relative|absolute|overflow|z-|top-|left-|right-|bottom-|hover:|group|shrink|grow|inline|block|hidden$|cursor)/.test(k)).slice(0, 3).join('.');
          const attrs = [];
          ['role', 'aria-label', 'type', 'placeholder', 'data-state', 'aria-checked', 'disabled', 'hidden'].forEach((a) => { if (el.hasAttribute(a)) attrs.push(a + (el.getAttribute(a) ? '="' + mask(el.getAttribute(a)) + '"' : '')); });
          let txt = '';
          if (LABEL_TAGS.test(el.tagName) || el.getAttribute('role') === 'button') txt = mask(el.textContent || '');
          const line = '  '.repeat(depth) + el.tagName.toLowerCase() + (cls ? '.' + cls : '') + (attrs.length ? '[' + attrs.join(' ') + ']' : '') + (txt ? ' "' + txt + '"' : '');
          out.push(line); len += line.length + 1;
          if (txt && el.tagName === 'BUTTON') return;
          Array.from(el.children).slice(0, 25).forEach((ch) => walk(ch, depth + 1));
        }
        walk(root, 0);
        return out.join('\n') + (len > maxLen ? '\n… (길어서 생략)' : '');
      }

      function report(r) {
        const lines = [];
        lines.push('[칼퇴몽키 점검] v' + (GM_VERSION || '') + ' · ' + new Date().toLocaleString('ko-KR'));
        lines.push('화면: ' + KIND_NAME[r.kind] + ' · 주소: ' + location.pathname.replace(/[0-9a-f]{16,}/gi, '(환자ID)'));
        lines.push('');
        r.res.forEach((x) => lines.push((x.ok ? '✅ ' : (x.need ? '❌ ' : '⚪ ')) + x.name + (x.ok ? '' : (x.need ? ' — 못 찾음' : ' — 지금은 없음(상황에 따라 정상)'))));
        const bad = r.res.filter((x) => !x.ok && x.need);
        const extra = r.res.filter((x) => !x.ok && !x.need);
        const areas = new Set();
        bad.concat(extra.slice(0, 2)).forEach((x) => {
          let a = null; try { a = x.area && x.area(); } catch (e) {}
          if (!a || areas.has(a)) return;
          areas.add(a);
          lines.push('');
          lines.push('---- 화면 구조: ' + x.name + ' 주변 ----');
          lines.push(skeleton(a, 5000));
        });
        const dl = openDialogs();
        if (dl.length) { lines.push(''); lines.push('---- 열린 창 ----'); dl.slice(0, 2).forEach((d) => lines.push(skeleton(d, 3000))); }
        return lines.join('\n');
      }

      // ---- 점검 버튼과 결과 상자 ----
      ['wc-diag-btn', 'wc-diag-box'].forEach((id) => { const o = document.getElementById(id); if (o) o.remove(); }); // 다시 시작될 때 중복 방지
      const btn = document.createElement('button');
      btn.id = 'wc-diag-btn';
      btn.type = 'button';
      btn.textContent = '🔍 점검';
      btn.style.cssText = 'position:fixed;left:8px;bottom:36px;z-index:2147483646;background:#fff;color:#333;border:1px solid #999;border-radius:14px;font:12px sans-serif;padding:3px 10px;cursor:pointer;box-shadow:0 1px 4px rgba(0,0,0,.2)';
      const box = document.createElement('div');
      box.id = 'wc-diag-box';
      box.style.cssText = 'position:fixed;left:8px;bottom:66px;z-index:2147483646;width:360px;max-height:60vh;overflow:auto;background:#fff;color:#222;border:1px solid #888;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.25);font:12px/1.5 sans-serif;padding:8px;display:none';
      document.body.appendChild(btn);
      document.body.appendChild(box);
      let lastReport = '';
      function copyText(t) {
        try { if (navigator.clipboard) return navigator.clipboard.writeText(t).then(() => true, () => fallback()); } catch (e) {}
        return Promise.resolve(fallback());
        function fallback() { const ta = document.createElement('textarea'); ta.value = t; document.body.appendChild(ta); ta.select(); let ok = false; try { ok = document.execCommand('copy'); } catch (e) {} ta.remove(); return ok; }
      }
      function showBox() {
        const r = runChecks();
        lastReport = report(r);
        const bad = r.res.filter((x) => !x.ok && x.need).length;
        box.innerHTML = '';
        const h = document.createElement('div');
        h.style.cssText = 'font-weight:700;margin-bottom:4px';
        h.textContent = '🔍 칼퇴몽키 점검 · ' + KIND_NAME[r.kind] + (bad ? ' · 문제 ' + bad + '개' : ' · 이상 없음');
        h.style.color = bad ? '#c62828' : '#2e7d32';
        box.appendChild(h);
        r.res.forEach((x) => {
          const d = document.createElement('div');
          d.textContent = (x.ok ? '✅ ' : (x.need ? '❌ ' : '⚪ ')) + x.name;
          if (!x.ok && x.need) d.style.color = '#c62828';
          if (!x.ok && !x.need) d.style.color = '#888';
          box.appendChild(d);
        });
        if (!r.res.length) { const d = document.createElement('div'); d.textContent = '이 화면은 점검할 항목이 없어요. 교육상담·대상자 목록·포괄평가·점검평가 화면에서 눌러 주세요.'; box.appendChild(d); }
        const note = document.createElement('div');
        note.style.cssText = 'color:#777;margin:6px 0 4px';
        note.textContent = '⚪는 창이 열렸을 때만 있는 항목이에요. 해당 단계에서 다시 누르면 확인돼요.';
        box.appendChild(note);
        const row = document.createElement('div');
        row.style.cssText = 'display:flex;gap:6px';
        const cp = document.createElement('button'); cp.type = 'button'; cp.textContent = '점검 결과 복사'; cp.style.cssText = 'flex:1';
        const cl = document.createElement('button'); cl.type = 'button'; cl.textContent = '닫기';
        cp.onclick = () => copyText(lastReport).then((ok) => { cp.textContent = ok === false ? '복사 실패' : '복사됨 · Claude에게 붙여넣기'; });
        cl.onclick = () => { box.style.display = 'none'; };
        row.appendChild(cp); row.appendChild(cl);
        box.appendChild(row);
        box.style.display = 'block';
        setAlert(bad ? r : null);
      }
      btn.addEventListener('click', () => { if (box.style.display === 'block') box.style.display = 'none'; else showBox(); });

      // ---- 화면 변경 자동 감지: 화면에 들어가고 잠시 뒤, 늘 있어야 할 항목을 확인 (두 번 연속 없을 때만 경고) ----
      function setAlert(r) {
        if (r) {
          const miss = r.res.filter((x) => !x.ok && x.need).map((x) => x.name);
          btn.textContent = '⚠️ 점검 필요';
          btn.style.background = '#ffebee'; btn.style.borderColor = '#c62828'; btn.style.color = '#c62828';
          btn.title = '웰체크 화면이 바뀐 것 같아요: ' + miss.join(', ');
        } else {
          btn.textContent = '🔍 점검';
          btn.style.background = '#fff'; btn.style.borderColor = '#999'; btn.style.color = '#333';
          btn.title = '';
        }
      }
      let lastPath = '', settleAt = 0, missStreak = 0, alerted = '';
      setInterval(() => {
        if (location.pathname !== lastPath) { lastPath = location.pathname; settleAt = Date.now() + 5000; missStreak = 0; return; }
        if (Date.now() < settleAt) return;
        if (running || (pcpShared.isRunning && pcpShared.isRunning())) return; // 실행 중에는 화면이 계속 바뀌므로 확인 안 함
        const r = runChecks();
        const miss = r.res.filter((x) => !x.ok && x.need);
        if (miss.length) {
          missStreak++;
          if (missStreak >= 2) {
            setAlert(r);
            const key = r.kind + ':' + miss.map((x) => x.name).join(',');
            if (alerted !== key) {
              alerted = key;
              badge('⚠️ 웰체크 화면이 바뀐 것 같아요 (' + KIND_NAME[r.kind] + ': ' + miss.map((x) => x.name).join(', ') + ') · 왼쪽 아래 점검을 눌러 주세요');
            }
          }
        } else { missStreak = 0; if (alerted && alerted.startsWith(r.kind + ':')) { alerted = ''; setAlert(null); } }
        settleAt = Date.now() + 10000; // 10초마다 다시 확인
      }, 1000);
      console.log('[자동화] 관리자용 점검·화면 변경 감지 시작됨');
    })();

    return function teardown() {
      clearInterval(showId);
      window.removeEventListener('message', onMsg);
      panel.remove();
      pcpTeardown();
    };
  }

  // ============================================================
  // SPA 페이지 이동 감지 로직
  // 웰체크는 SPA라서 페이지 안에서 이동할 땐 새로고침이 안 일어나
  // 스크립트가 재실행되지 않음. 주소를 계속 지켜보다가 바뀌면 재시작함.
  // ============================================================

  let currentKey = null;
  let currentTeardown = null;

  function computeKey() {
    const h = location.hostname;
    const p = location.pathname;
    // [변경] 화면 주소(경로)와 무관하게 웰체크 전체에서 감시 (동작은 화면 요소 기준)
    if (h === 'web.well-check.co.kr') return 'edu:all';
    return null;
  }

  function activate() {
    const key = computeKey();
    if (key === currentKey) return;

    if (currentTeardown) {
      try {
        currentTeardown();
      } catch (e) {}
      currentTeardown = null;
    }

    currentKey = key;

    if (!key) return;

    if (key.startsWith('edu:')) {
      const t1 = startEduConsult();
      const t2 = startQueuePanel();
      currentTeardown = function () { t1(); t2(); };
    }
  }

  setInterval(activate, 500);
  activate();
})();