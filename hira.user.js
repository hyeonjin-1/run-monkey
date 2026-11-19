// ==UserScript==
// @name         🐵 칼퇴몽키 – 심평원용
// @namespace    wellcheck-hira-automation
// @version      1.9
// @description  클릭은 몽키가, 퇴근은 내가 · 웰체크에서 심평원 이동 후 작성완료 자동 클릭 (교육상담: 마지막 행 체크 후 / 포괄평가: 바로)
// @match        https://biz.hira.or.kr/*
// @updateURL    https://raw.githubusercontent.com/hyeonjin-1/run-monkey/main/hira.user.js
// @downloadURL  https://raw.githubusercontent.com/hyeonjin-1/run-monkey/main/hira.user.js
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  console.log('[자동화] 심평원 스크립트 로드됨:', location.href);

  // 페이지 환경에서 직접 실행하므로 창은 window 그대로 사용
  const PAGE = window;

  // ---- 새 버전 알림: GitHub에 올라간 버전이 더 높으면 화면 왼쪽 아래에 알림 카드를 띄움 ----
  // 버전 숫자만 읽어 오고, 화면의 정보는 어디로도 보내지 않음
  (function kaltoeUpdateNotice(cfg) {
    try {
      if (window.top !== window.self) return; // 맨 바깥 화면에서만
      const cur = (typeof GM_info !== 'undefined' && GM_info.script && GM_info.script.version) || cfg.fallback;
      const newer = (a, b) => { const x = String(a).split('.').map(Number), y = String(b).split('.').map(Number); for (let i = 0; i < Math.max(x.length, y.length); i++) { const d = (x[i] || 0) - (y[i] || 0); if (d) return d > 0; } return false; };
      const key = 'kaltoe-upd-' + cfg.id;
      const getText = (url) => new Promise((res, rej) => {
        if (typeof GM_xmlhttpRequest === 'function') {
          GM_xmlhttpRequest({ method: 'GET', url, headers: { 'Cache-Control': 'no-cache' }, timeout: 8000, onload: (r) => (r.status === 200 ? res(r.responseText) : rej(r.status)), onerror: rej, ontimeout: rej });
        } else fetch(url, { cache: 'no-store' }).then((r) => (r.ok ? r.text() : Promise.reject(r.status))).then(res, rej);
      });
      setTimeout(() => {
        getText(cfg.url + '?t=' + Date.now()).then((txt) => {
          const m = /@version\s+([\d.]+)/.exec(txt || '');
          if (!m || !newer(m[1], cur)) return;
          const latest = m[1];
          try { const sn = JSON.parse(localStorage.getItem(key) || 'null'); if (sn && sn.v === latest && Date.now() < sn.until) return; } catch (e) {}
          if (document.getElementById('kaltoe-upd-card-' + cfg.id)) return;
          const c = document.createElement('div');
          c.id = 'kaltoe-upd-card-' + cfg.id;
          c.style.cssText = 'position:fixed;left:8px;bottom:' + (cfg.bottom || 70) + 'px;z-index:2147483646;width:280px;background:#fff;color:#222;border:2px solid #ef6c00;border-radius:12px;padding:12px 14px;box-shadow:0 4px 16px rgba(0,0,0,.2);font:13px/1.5 "Malgun Gothic",sans-serif';
          c.innerHTML = '<div style="font-weight:700;font-size:14px;margin-bottom:4px">🐵 ' + cfg.title + ' 새 버전이 나왔어요</div>' +
            '<div style="color:#666;margin-bottom:10px">지금 v' + cur + ' → 새 버전 v' + latest + '</div>' +
            '<div style="display:flex;gap:6px"><button type="button" data-a="go" style="flex:1;background:#ef6c00;color:#fff;border:0;border-radius:6px;padding:6px;cursor:pointer;font-weight:700">지금 업데이트</button>' +
            '<button type="button" data-a="later" style="background:#fff;border:1px solid #bbb;border-radius:6px;padding:6px 10px;cursor:pointer">나중에</button></div>';
          c.addEventListener('click', (e) => {
            const a = e.target && e.target.getAttribute('data-a');
            if (a === 'go') {
              window.open(cfg.url, '_blank'); // Tampermonkey 덮어쓰기(Overwrite) 화면이 열림
              c.querySelector('div:nth-child(2)').textContent = '새 탭에서 Overwrite(덮어쓰기)를 누른 뒤 이 화면을 새로고침해 주세요.';
            } else if (a === 'later') {
              try { localStorage.setItem(key, JSON.stringify({ v: latest, until: Date.now() + 24 * 3600 * 1000 })); } catch (e2) {}
              c.remove();
            }
          });
          (document.body || document.documentElement).appendChild(c);
        }).catch(() => {}); // 확인 실패(인터넷·보안 설정)는 조용히 넘어감
      }, cfg.delay || 3000);
    } catch (e) {}
  })({ id: 'hira', title: '칼퇴몽키 심평원용', url: 'https://raw.githubusercontent.com/hyeonjin-1/run-monkey/main/hira.user.js', fallback: '1.9', bottom: 40, delay: 1500 });

  function realClick(el) {
    if (!el) return false;
    ['mousedown', 'mouseup', 'click'].forEach((type) => {
      const ev = new MouseEvent(type, { bubbles: true, cancelable: true, view: window });
      el.dispatchEvent(ev);
    });
    return true;
  }

  // 화면 왼쪽 아래에 현재 상태 표시
  function badge(text) {
    // 칼퇴몽키 상태 표시: 진행 중에는 🐵, 작성완료되면 🙈
    const done = /작성완료 (확인됨|클릭함)/.test(text);
    text = (done ? '🙈 칼퇴몽키 · ' : '🐵 칼퇴몽키 · ') + text;
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

  // ============================================================
  // 2) 심평원 (biz.hira.or.kr) 자동화
  // ============================================================
  function startHira() {
    function collectFrameDocs() {
      const docs = [document];
      function walk(win) {
        let frames;
        try {
          frames = Array.from(win.frames);
        } catch (e) {
          return;
        }
        for (const f of frames) {
          try {
            if (f.document) {
              docs.push(f.document);
              walk(f);
            }
          } catch (e) {}
        }
      }
      walk(PAGE);
      if (PAGE !== window) walk(window);
      return Array.from(new Set(docs));
    }

    function queryAllFrames(selector) {
      for (const doc of collectFrameDocs()) {
        try {
          const el = doc.querySelector(selector);
          if (el) return el;
        } catch (e) {}
      }
      return null;
    }

    // [추가] 행의 맨 앞(첫 번째 열) 체크박스 찾기: 기존 방식 → 안 되면 첫 열 요소 중 체크박스로 보이는 것
    function findRowCheckbox(rowIdx) {
      const exact =
        queryAllFrames(`[id$="_gridrow_${rowIdx}_cell_${rowIdx}_0_controlcheckbox"]`) ||
        queryAllFrames(`[id*="_gridrow_${rowIdx}_cell_${rowIdx}_0"]`);
      if (exact) return exact;

      const colRe = new RegExp('_gridrow_' + rowIdx + '_cell_\\d+_0(?:_|$)');
      const cands = [];
      for (const doc of collectFrameDocs()) {
        try {
          doc.querySelectorAll('[id*="_gridrow_' + rowIdx + '_cell_"]').forEach((el) => {
            if (colRe.test(el.id)) cands.push(el);
          });
        } catch (e) {}
      }
      if (!cands.length) return null;
      const cb = cands.find((el) => /checkbox/i.test(el.id));
      if (cb) return cb;
      cands.sort((a, b) => a.id.length - b.id.length);
      return cands[0];
    }

    function findMaxRowIndex() {
      let maxIdx = -1;
      for (const doc of collectFrameDocs()) {
        try {
          Array.from(doc.querySelectorAll('[id*="_gridrow_"]')).forEach((el) => {
            const m = el.id.match(/_gridrow_(\d+)_/);
            if (m) {
              const idx = parseInt(m[1], 10);
              if (idx > maxIdx) maxIdx = idx;
            }
          });
        } catch (e) {}
      }
      return maxIdx >= 0 ? String(maxIdx) : null;
    }

    function overrideConfirmEverywhere() {
      function override(win) {
        try {
          win.confirm = function () {
            return true;
          };
          win.alert = function () {
            return undefined;
          };
        } catch (e) {}
      }
      override(PAGE);
      if (PAGE !== window) override(window);
      function walk(win) {
        let frames;
        try {
          frames = Array.from(win.frames);
        } catch (e) {
          return;
        }
        for (const f of frames) {
          try {
            override(f);
            walk(f);
          } catch (e) {}
        }
      }
      walk(PAGE);
      if (PAGE !== window) walk(window);
    }
    overrideConfirmEverywhere();

    // [추가] 작성완료가 실제로 처리됐는지 행 상태로 확인 → 웰체크 탭에 알리고, 확인되면 이 탭을 닫음
    function rowStatusText(rowIdx) {
      let txt = '';
      let found = false;
      for (const doc of collectFrameDocs()) {
        try {
          doc.querySelectorAll('[id*="_gridrow_' + rowIdx + '_cell_"]').forEach((el) => {
            found = true;
            txt += ' ' + (el.textContent || '');
          });
        } catch (e) {}
      }
      return { found, txt };
    }
    function notifyOpener(type) {
      try {
        if (window.opener) window.opener.postMessage({ wcAuto: type }, 'https://web.well-check.co.kr');
      } catch (e) {}
    }
    function verifyDone(rowIdx) {
      const t0 = Date.now();
      let goneCount = 0;
      const id = setInterval(() => {
        const { found, txt } = rowStatusText(rowIdx);
        goneCount = found ? 0 : goneCount + 1;
        const completed = (found && txt.includes('작성완료') && !txt.includes('임시저장')) || goneCount >= 3;
        const shown = found ? txt.replace(/\s+/g, ' ').trim().slice(-24) : '행 사라짐';
        badge('심평원 · 완료 확인 중 (' + Math.round((Date.now() - t0) / 1000) + '초) · ' + shown);
        if (completed) {
          clearInterval(id);
          badge('심평원 · 작성완료 확인됨');
          notifyOpener('done');
          setTimeout(() => { try { window.close(); } catch (e) {} }, 1500);
        } else if (Date.now() - t0 > 25000) {
          clearInterval(id);
          badge('심평원 · 작성완료를 확인하지 못함 (이 탭을 확인하세요)');
          notifyOpener('unconfirmed');
        }
      }, 700);
    }

    // [추가] 포괄평가 화면: 행 선택 없이 작성완료만 누름 (화면에 '개선목표' 문구가 있으면 포괄평가로 판단)
    let pcpSeenAt = 0;
    let pcpDone = false;
    function pageText() {
      let txt = '';
      for (const doc of collectFrameDocs()) {
        try { txt += ' ' + (doc.body ? doc.body.textContent : ''); } catch (e) {}
      }
      return txt;
    }
    function handlePcp() {
      if (pcpDone) return true;
      const txt = pageText();
      // 포괄평가: '개선목표' 문구가 있는 화면 / 점검평가: 행 없이 '점검' 문구만 있는 화면(약 8초 동안 행이 안 뜰 때)
      let kind = null;
      if (txt.includes('개선목표')) kind = '포괄평가';
      else if (/약물\s*변경\s*필요/.test(txt)) kind = '점검평가'; // 점검평가 화면에만 있는 문구
      else if (tries > 16 && !/교육[·ㆍ・.]?상담/.test(txt) && /점검/.test(txt) && findMaxRowIndex() === null) kind = '점검평가';
      if (!kind) return false;
      const btn = queryAllFrames('[id$="btnS0001"]');
      if (!btn) { badge('심평원 · ' + kind + ' 화면 확인됨 · 작성완료 버튼 기다리는 중'); return true; }
      if (!pcpSeenAt) pcpSeenAt = Date.now();
      if (Date.now() - pcpSeenAt < 3000) { badge('심평원 · ' + kind + ' 화면 로딩 확인 중…'); return true; }
      pcpDone = true;
      clearInterval(intervalId);
      realClick(btn);
      badge('심평원 · ' + kind + ' 작성완료 클릭함');
      setTimeout(() => {
        notifyOpener('clicked');
        badge('심평원 · ' + kind + ' 작성완료 클릭함 (탭이 곧 닫혀요)');
        setTimeout(() => { try { window.close(); } catch (e) {} }, 1500);
      }, 6000);
      return true;
    }

    let tries = 0;

    const intervalId = setInterval(() => {
      tries++;
      overrideConfirmEverywhere();
      if (handlePcp()) {
        if (!pcpDone && tries > 240) { clearInterval(intervalId); badge('심평원 · 작성완료 버튼을 찾지 못해 중단함 (포괄평가·점검평가 화면 확인)'); }
        return;
      }

      const rowIdx = findMaxRowIndex();
      badge('심평원 자동화 · 시도 ' + tries + '회 · 프레임 ' + collectFrameDocs().length + '개 · 행 ' + (rowIdx === null ? '못 찾음' : rowIdx));

      if (rowIdx !== null) {
        const checkboxEl = findRowCheckbox(rowIdx);
        if (!checkboxEl) {
          const any = queryAllFrames('[id*="_gridrow_' + rowIdx + '_"]');
          badge('심평원 · 행 ' + rowIdx + ' 체크박스 못 찾음 · 예시 id: ' + (any ? any.id.slice(-60) : '없음'));
        }
        if (checkboxEl) {
          realClick(checkboxEl);
          console.log('[자동화] 마지막 행(' + rowIdx + ') 체크박스 클릭 완료');
          badge('심평원 · 행 ' + rowIdx + ' 체크함 → 작성완료 시도 중');
          clearInterval(intervalId);

          setTimeout(() => {
            const btn = queryAllFrames('[id$="btnS0001"]');
            if (btn) {
              realClick(btn);
              console.log('[자동화] 작성완료 버튼 클릭 시도 완료');
              badge('심평원 · 작성완료 클릭 시도함');
              verifyDone(rowIdx);
            } else {
              console.log('[자동화] 작성완료 버튼을 찾지 못함');
              badge('심평원 · 작성완료 버튼 못 찾음');
            }
          }, 500);

          return;
        }
      }

      if (tries % 10 === 0) {
        console.log('[자동화] 그리드 행을 아직 못 찾음 (시도 ' + tries + '회, 프레임 ' + collectFrameDocs().length + '개)');
      }

      if (tries > 240) {
        clearInterval(intervalId);
        console.log('[자동화] 그리드 행을 끝내 찾지 못해 중단함 (2분 초과)');
        badge('심평원 · 행을 못 찾아 중단함(2분 초과)');
      }
    }, 500);

    console.log('[자동화] 심평원 자동화 시작됨');

    return function teardown() {
      clearInterval(intervalId);
    };
  }


  // 최상위 창에서만 실행
  if (window.top === window.self) {
    badge('심평원 자동화 v1.9 시작');
    startHira();
  }
})();
