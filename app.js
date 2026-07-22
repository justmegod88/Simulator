(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);

  const state = {
    master: [], content: [], edu: [], perRaw: [], per: [], qm: [],
    sales: [], rec: [], send: [], filtered: [], query: '', gapFilter: '', competitorRows: []
  };

  // Performance indexes/caches
  const metricsCache = new Map();
  const masterIndex = new Map();
  const salesIndex = new Map();
  const perceptionIndex = new Map();
  const educationIndex = new Map();
  const recommendationIndex = new Map();

  const aliases = {
    master: ['01_안경사마스터', '안경사마스터'],
    content: ['02_교육콘텐츠', '교육콘텐츠마스터'],
    edu: ['03_교육참여', '교육참여이력', '교육이력'],
    per: ['04_인식조사', '인식조사', 'Sheet1'],
    qm: ['인식문항마스터', '문항마스터'],
    sales: ['06_피팅판매', '피팅판매'],
    rec: ['AI추천결과', '교육추천결과', '08_교육추천'],
    send: ['발송요청관리', '발송결과로그', '09_발송']
  };

  const LIKERT = {
    '전혀 그렇지 않다': 1, '그렇지 않다': 2, '보통이다': 3, '비슷하다': 3,
    '그렇다': 4, '매우 그렇다': 5
  };

  const PRODUCT_RULES = {
    max: { label: 'MAX 성장률', match: (p) => /MAX|맥스/i.test(p || '') },
    ast: { label: '난시 성장률', match: (p) => (/난시|토릭|TORIC|ASD/i.test(p || '') && !/MAX|맥스/i.test(p || '')) },
    mf: {
      label: '멀티포컬 성장률',
      match: (p) => /멀티포컬|MULTIFOCAL|다초점|노안|\bMF\b/i.test(p || '') ||
        ((/MAX|맥스/i.test(p || '')) && (/난시|토릭|TORIC|ASD/i.test(p || '')))
    }
  };

  const SUMMARY_ALIASES = {
    max: { cur: ['MAX 성장률', 'MAX성장률', 'MAX제품군 성장률', 'MAX제품군', 'MAX YoY'], py: ['전년 MAX 성장률', 'MAX PY 성장률', '작년 MAX 성장률'] },
    ast: { cur: ['난시 성장률', '난시성장률', '난시제품 성장률', '난시제품', '토릭 성장률', 'ASD 성장률'], py: ['전년 난시 성장률', '난시 PY 성장률', '작년 난시 성장률'] },
    mf: { cur: ['멀티포컬 성장률', '멀티포컬성장률', 'MF 성장률', 'MF성장률', '멀티포컬제품 성장률'], py: ['전년 멀티포컬 성장률', '멀티포컬 PY 성장률', '작년 멀티포컬 성장률', '전년 MF 성장률'] }
  };

  const COMPETITOR_RULES = [
    { vendor: '알콘', patterns: ['알콘', 'ALCON', '토탈원', 'TOTAL1', '데일리스', 'DAILIES', '프리시전', 'PRECISION'] },
    { vendor: '바슈롬', patterns: ['바슈롬', 'BAUSCH', 'BAUSCH+LOMB', '울트라', 'ULTRA', '바이오트루', 'BIOTRUE', '레이셀', 'LACELLE'] },
    { vendor: '쿠퍼비전', patterns: ['쿠퍼비전', 'COOPERVISION', '마이데이', 'MYDAY', '클라리티', 'CLARITI', '바이오피니티', 'BIOFINITY'] }
  ];
  const COMPETITOR_KEYWORDS = ['멀티포컬', '난시', '렌즈', '블루라이트', '실리콘하이드로겔', '콘택트렌즈', '피팅', '캠페인', '신제품', '프로모션', '행사', '할인'];

  const clean = (v) => v == null ? '' : String(v).trim();
  const norm = (s) => clean(s).replace(/[\s_\-()\/]/g, '').toLowerCase();
  const esc = (s) => clean(s).replace(/[&<>"']/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[m]));
  const num = (v) => {
    if (v == null || v === '') return null;
    const n = Number(String(v).replace(/,/g, '').replace(/%/g, ''));
    return Number.isFinite(n) ? n : null;
  };
  const avg = (arr, fn = (x) => x) => {
    const v = arr.map(fn).filter((x) => x != null && Number.isFinite(Number(x))).map(Number);
    return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
  };
  const sum = (arr, fn) => arr.reduce((s, x) => s + (Number(fn(x)) || 0), 0);
  const pct = (v) => v == null ? '데이터 없음' : `${Math.round(Number(v) * 100)}%`;
  const rate = (v) => v == null ? '데이터 없음' : `${v >= 0 ? '+' : ''}${Number(v).toFixed(1)}%`;
  const pp = (v) => v == null ? '데이터 없음' : `${v >= 0 ? '+' : ''}${Number(v).toFixed(1)}%p`;

  function get(r, names) {
    if (!r) return '';
    const map = {};
    Object.keys(r).forEach((k) => { map[norm(k)] = r[k]; });
    for (const n of names) {
      const v = map[norm(n)];
      if (v !== undefined && clean(v) !== '') return v;
    }
    return '';
  }

  function sheet(wb, names) {
    const name = wb.SheetNames.find((s) => names.some((a) => norm(s).includes(norm(a))));
    return name ? XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '', raw: true }) : [];
  }

  function infer(q) {
    q = clean(q);
    if (/멀티포컬|다초점|노안/.test(q)) return '멀티포컬';
    if (/난시|토릭|저난시|ASD/.test(q)) return '난시';
    if (/블루라이트|눈건강|기술|편안함/.test(q)) return '눈건강/기술';
    if (/아큐브|추천|원데이|브랜드/.test(q)) return '브랜드/추천';
    return '기타';
  }

  function normMaster(rows) {
    return rows.map((r, i) => ({
      ...r,
      안경사ID: clean(get(r, ['안경사ID', 'ID'])) || `AUTO-${i + 1}`,
      안경사명: clean(get(r, ['안경사명', '이름', '성명'])),
      안경원코드: clean(get(r, ['안경원코드', '매장코드', '거래처코드', 'ShipTo', 'SoldTo'])),
      안경원명: clean(get(r, ['안경원명', '안경원', '매장명', '거래처명'])),
      지역: clean(get(r, ['지역', '시도'])),
      연차: clean(get(r, ['연차'])),
      Tier: clean(get(r, ['Tier', '티어', '등급'])),
      채널: clean(get(r, ['채널', 'Channel'])),
      담당영업사원: clean(get(r, ['담당영업사원', '담당자', '영업사원']))
    })).filter((r) => r.안경사ID || r.안경사명);
  }

  function normQm(rows) {
    return rows.map((r, i) => {
      const q = clean(get(r, ['문항', '문항명', 'Question']));
      return {
        문항ID: clean(get(r, ['문항ID', 'QuestionID'])) || `Q${String(i + 1).padStart(3, '0')}`,
        문항: q,
        영역: clean(get(r, ['인식영역', '영역'])) || infer(q),
        제품군: clean(get(r, ['제품군'])) || infer(q),
        응답척도: clean(get(r, ['응답척도', '응답유형'])),
        긍정방향: clean(get(r, ['긍정방향'])) || (/역코딩/.test(q) ? '낮을수록 긍정' : '높을수록 긍정'),
        목표값: num(get(r, ['목표값'])) ?? 4,
        추천교육ID: clean(get(r, ['추천교육ID', '교육ID'])),
        사용: clean(get(r, ['분석사용여부', '사용여부'])) || 'Y'
      };
    }).filter((x) => x.문항ID || x.문항);
  }

  function qmFor(col) {
    const c = clean(col);
    const bare = c.replace(/^Q\d{2,4}[_\-\s]*/i, '');
    const id = (c.match(/^(Q\d{2,4})/i) || [])[1] || c;
    return state.qm.find((q) => norm(q.문항ID) === norm(id)) ||
      state.qm.find((q) => norm(q.문항) === norm(bare)) ||
      state.qm.find((q) => bare.includes(q.문항) || q.문항.includes(bare));
  }

  function score(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return v;
    return LIKERT[clean(v)] ?? num(v);
  }

  function analyzable(q, val) {
    if (!q || q.사용 === 'N') return false;
    if (/0~100|%|Y\/N|브랜드선택|선택형|N\/A/i.test(q.응답척도)) return false;
    const s = score(val);
    return s != null && s >= 1 && s <= 5;
  }

  function normPer(rows) {
    const out = [];
    const meta = ['안경사ID', '안경사명', '안경원명', '안경원코드', '지역', '연차', 'Tier', 'SEG', 'No', '번호'];
    rows.forEach((r) => {
      const id = clean(get(r, ['안경사ID', 'ID']));
      if (!id) return;
      Object.keys(r).forEach((col) => {
        if (meta.some((m) => norm(m) === norm(col))) return;
        const q = qmFor(col) || {
          문항ID: col, 문항: clean(col), 영역: infer(col), 제품군: infer(col), 응답척도: '1~5점',
          긍정방향: /역코딩/.test(col) ? '낮을수록 긍정' : '높을수록 긍정', 목표값: 4, 추천교육ID: '', 사용: 'Y'
        };
        if (!analyzable(q, r[col])) return;
        const raw = score(r[col]);
        const adj = /낮을수록/.test(q.긍정방향) ? 6 - raw : raw;
        out.push({ 안경사ID: id, 문항ID: q.문항ID, 문항: q.문항, 영역: q.영역, 제품군: q.제품군, 추천교육ID: q.추천교육ID, 원응답: r[col], 보정점수: adj, 목표값: Number(q.목표값 || 4), gap: adj < Number(q.목표값 || 4) });
      });
    });
    return out;
  }

  function rebuildIndexes() {
    metricsCache.clear();
    masterIndex.clear();
    salesIndex.clear();
    perceptionIndex.clear();
    educationIndex.clear();
    recommendationIndex.clear();

    state.master.forEach((r) => {
      if (r.안경사ID) masterIndex.set(r.안경사ID, r);
    });

    state.sales.forEach((r) => {
      const keys = [
        clean(get(r, ['안경사ID', 'ID'])),
        clean(get(r, ['안경원코드', '매장코드', '거래처코드', 'ShipTo', 'SoldTo'])),
        clean(get(r, ['안경원명', '매장명', '거래처명']))
      ].filter(Boolean);
      keys.forEach((k) => {
        if (!salesIndex.has(k)) salesIndex.set(k, []);
        salesIndex.get(k).push(r);
      });
    });

    state.per.forEach((r) => {
      if (!perceptionIndex.has(r.안경사ID)) perceptionIndex.set(r.안경사ID, []);
      perceptionIndex.get(r.안경사ID).push(r);
    });

    state.edu.forEach((r) => {
      const id = clean(get(r, ['안경사ID', 'ID']));
      if (!educationIndex.has(id)) educationIndex.set(id, []);
      educationIndex.get(id).push(r);
    });

    state.rec.forEach((r) => {
      const id = clean(get(r, ['안경사ID', 'ID']));
      if (!recommendationIndex.has(id)) recommendationIndex.set(id, []);
      recommendationIndex.get(id).push(r);
    });
  }

  function eduDoneRow(r) {
    const flag = clean(get(r, ['완료여부', '수료여부', '참여여부', '시청여부'])).toUpperCase();
    if (['Y', 'YES', 'TRUE', '완료', '수료', 'DONE', 'COMPLETED'].includes(flag)) return true;
    if (['N', 'NO', 'FALSE', '미완료', '미수료'].includes(flag)) return false;
    const rateVal = num(get(r, ['완료율', '시청완료율', '진도율', '진행률']));
    return rateVal != null ? rateVal >= 100 : false;
  }

  function salesRowsForPerson(id) {
    const p = masterIndex.get(id);
    if (!p) return [];
    const keys = [p.안경사ID, p.안경원코드, p.안경원명].filter(Boolean);
    const seen = new Set();
    const rows = [];
    keys.forEach((k) => {
      (salesIndex.get(k) || []).forEach((r) => {
        if (!seen.has(r)) { seen.add(r); rows.push(r); }
      });
    });
    return rows;
  }

  function productNameOf(row) { return clean(get(row, ['제품명', '상품명', 'Product', '제품', 'SKU', '품목명', '브랜드제품명'])); }
  function val2025(row) { return num(get(row, ['2025팩수', '25년팩수', '2025', '2025년', '25년', 'PY팩수', '전년팩수', '작년팩수', '2025 판매팩수', '25년 판매팩수'])); }
  function val2026(row) { return num(get(row, ['2026팩수', '26년팩수', '2026', '2026년', '26년', 'CY팩수', '올해팩수', '현재팩수', '2026 판매팩수', '26년 판매팩수'])); }
  function directSummary(rows, key, kind) { return avg(rows, (r) => num(get(r, SUMMARY_ALIASES[key][kind]))); }

  function wideProductSum(rows, key, year) {
    let total = 0;
    let seen = false;
    rows.forEach((r) => {
      Object.keys(r).forEach((col) => {
        const c = clean(col);
        const isYear = year === '2026' ? /(2026|26년|CY|올해|당년)/i.test(c) : /(2025|25년|PY|전년|작년)/i.test(c);
        if (!isYear) return;
        const name = c.replace(/2026|2025|26년|25년|CY|PY|올해|당년|전년|작년|팩수|판매|수량/gi, '');
        if (PRODUCT_RULES[key].match(name)) {
          const v = num(r[col]);
          if (v != null) { total += v; seen = true; }
        }
      });
    });
    return seen ? total : null;
  }

  function productSummaryRows(rows, key) {
    const px = rows.filter((r) => {
      const pn = productNameOf(r);
      return pn && PRODUCT_RULES[key].match(pn);
    });
    if (!px.length) return null;
    const y25 = sum(px, val2025);
    const y26 = sum(px, val2026);
    const has25 = px.some((r) => val2025(r) != null);
    const has26 = px.some((r) => val2026(r) != null);
    if (!has25 && !has26) return null;
    return { cur: y25 === 0 && y26 > 0 ? 100 : y25 ? ((y26 - y25) / y25 * 100) : null, py: null, packs25: y25, packs26: y26 };
  }

  function groupGrowth(rows, key) {
    const productRows = productSummaryRows(rows, key);
    if (productRows) return productRows;
    const w25 = wideProductSum(rows, key, '2025');
    const w26 = wideProductSum(rows, key, '2026');
    if (w25 != null || w26 != null) return { cur: w25 === 0 && w26 > 0 ? 100 : w25 ? ((w26 - w25) / w25 * 100) : null, py: null, packs25: w25, packs26: w26 };
    const cur = directSummary(rows, key, 'cur');
    const py = directSummary(rows, key, 'py');
    return { cur, py, packs25: null, packs26: null };
  }

  function eduRows(id) { return educationIndex.get(id) || []; }
  function recRows(id) { return recommendationIndex.get(id) || []; }

  function calcPriority(m) {
    const text = [get(m.rec, ['AI인사이트', '추천사유', '추천교육명'])].join(' ');
    if (m.gaps.length >= 3 || m.growth < 0 || /긴급|개선 필요|낮음|부족|미완료/.test(text)) return '높음';
    if (m.gaps.length >= 1 || (m.eduRate != null && m.eduRate < 1) || m.growth === 0) return '중간';
    return '낮음';
  }

  function metrics(id) {
    if (metricsCache.has(id)) return metricsCache.get(id);

    const p = masterIndex.get(id);
    if (!p) {
      const empty = { p: null, perc: [], gaps: [], growths: { ast: { cur: null }, mf: { cur: null }, max: { cur: null } }, eduRate: null, rec: null, priority: '낮음' };
      metricsCache.set(id, empty);
      return empty;
    }

    const rows = salesRowsForPerson(id);
    const perc = perceptionIndex.get(id) || [];
    const gaps = perc.filter((x) => x.gap);
    const edu = eduRows(id);
    const eduRate = edu.length ? edu.filter(eduDoneRow).length / edu.length : null;
    const rec = recRows(id)[0] || null;
    const growths = { ast: groupGrowth(rows, 'ast'), mf: groupGrowth(rows, 'mf'), max: groupGrowth(rows, 'max') };
    const priority = calcPriority({ gaps, growth: avg([growths.ast.cur, growths.mf.cur, growths.max.cur], (x) => x), eduRate, rec });
    const result = { p, perc, gaps, growths, eduRate, rec, priority };

    metricsCache.set(id, result);
    return result;
  }

  function follow(m) {
    if (m.eduRate == null || m.eduRate < 1) return '교육 완료 독려';
    if (avg([m.growths.ast.cur, m.growths.mf.cur, m.growths.max.cur], (x) => x) < 0) return '현장 Follow-up/코칭';
    if (m.gaps.length >= 3) return '핵심 Gap 문항 코칭';
    if (m.gaps.length >= 1) return '문항별 설명 보완';
    return '모니터링';
  }

  function badge(v) {
    const c = v === '높음' ? 'high' : v === '중간' ? 'medium' : 'low';
    return `<span class="pill ${c}">${esc(v || '낮음')}</span>`;
  }

  function deltaClass(v) { return v == null ? '' : v < 0 ? 'negative' : 'positive'; }
  function kpi(label, value, note) { return `<div class="kpi-card"><span>${label}</span><strong>${value}</strong><small>${note}</small></div>`; }

  function kpiGrowth(key, rows, peerRows) {
    const cur = avg(rows, (r) => metrics(r.안경사ID).growths[key].cur);
    const group = avg(peerRows, (r) => metrics(r.안경사ID).growths[key].cur);
    const groupDelta = cur != null && group != null ? cur - group : null;
    return kpi(
      PRODUCT_RULES[key].label,
      `전년 대비 ${rate(cur)}`,
      `<span class="delta ${deltaClass(groupDelta)}">전체 대비 ${pp(groupDelta)}</span>`
    );
  }

  function filterByDropdown() {
    let rows = [...state.master];
    const f = { regionFilter: '지역', yearsFilter: '연차', tierFilter: 'Tier', channelFilter: '채널', repFilter: '담당영업사원' };
    Object.entries(f).forEach(([id, field]) => {
      const v = $(id)?.value;
      if (v) rows = rows.filter((r) => clean(r[field]) === v);
    });
    return rows;
  }

  function filtered() {
    let rows = filterByDropdown();
    const q = clean(state.query);

    if (q) {
      const y = (q.match(/(\d+)\s*년차/) || [])[1];
      if (y) rows = rows.filter((r) => clean(r.연차).includes(y));

      const wantGap = /인식|Gap|갭|문항/.test(q);
      const eduIncomplete = /미완료|미수료|교육.*필요/.test(q);
      const salesLow = /성장률.*낮|성장.*낮|판매.*낮|미전환|성장률 음수|역성장/.test(q);
      const high = /고우선|높음|우선순위/.test(q);
      const text = q.replace(/\d+\s*년차|MAX|맥스|멀티포컬|다초점|난시|저난시|ASD|토릭|인식|Gap|갭|문항|미완료|미수료|교육|필요|성장률|성장|판매|낮은|안경사|고우선|높음|우선순위|음수|역성장|중/g, '').trim();
      if (text.length >= 2) rows = rows.filter((r) => [r.안경사ID, r.안경사명, r.안경원명, r.지역, r.Tier, r.채널, r.담당영업사원].join(' ').includes(text));

      rows = rows.filter((r) => {
        const m = metrics(r.안경사ID);
        if (wantGap && !m.gaps.length) return false;
        if (eduIncomplete && !(m.eduRate == null || m.eduRate < 1)) return false;
        if (salesLow && !(avg([m.growths.ast.cur, m.growths.mf.cur, m.growths.max.cur], (x) => x) < 0)) return false;
        if (high && m.priority !== '높음') return false;
        return true;
      });
    }

    if (state.gapFilter) {
      rows = rows.filter((r) => {
        const m = metrics(r.안경사ID);
        if (state.gapFilter === 'educationIncomplete') return m.eduRate == null || m.eduRate < 1;
        if (state.gapFilter === 'perceptionGap') return m.gaps.length > 0;
        if (state.gapFilter === 'astNegative') return m.growths.ast.cur != null && m.growths.ast.cur < 0;
        if (state.gapFilter === 'mfNegative') return m.growths.mf.cur != null && m.growths.mf.cur < 0;
        if (state.gapFilter === 'maxNegative') return m.growths.max.cur != null && m.growths.max.cur < 0;
        return true;
      });
    }

    return rows;
  }

  function render() {
    console.time('render');
    const rows = filtered();
    const peerRows = filterByDropdown();
    state.filtered = rows;
    const ms = rows.map((r) => metrics(r.안경사ID));
    const eduComplete = ms.filter((m) => m.eduRate === 1).length;
    const reached = ms.filter((m) => m.gaps.length === 0 && m.perc.length > 0).length;

    $('kpiGrid').innerHTML = [
      kpi('전체 관리 안경사', rows.length.toLocaleString('ko-KR'), '현재 필터'),
      kpi('교육 완료 안경사', eduComplete.toLocaleString('ko-KR'), `${pct(rows.length ? eduComplete / rows.length : null)} 완료`),
      kpi('인식 목표 도달 안경사', reached.toLocaleString('ko-KR'), `${pct(rows.length ? reached / rows.length : null)} 도달`),
      kpiGrowth('ast', rows, peerRows),
      kpiGrowth('mf', rows, peerRows),
      kpiGrowth('max', rows, peerRows)
    ].join('');

    const eduIncomplete = ms.filter((m) => m.eduRate == null || m.eduRate < 1).length;
    const gapPeople = ms.filter((m) => m.gaps.length).length;
    const astNegative = ms.filter((m) => m.growths.ast.cur != null && m.growths.ast.cur < 0).length;
    const mfNegative = ms.filter((m) => m.growths.mf.cur != null && m.growths.mf.cur < 0).length;
    const maxNegative = ms.filter((m) => m.growths.max.cur != null && m.growths.max.cur < 0).length;

    const gapItems = [
      ['education', '교육 미완료', eduIncomplete, 'educationIncomplete'],
      ['perception', '인식 목표 미달', gapPeople, 'perceptionGap'],
      ['sales ast', '난시 역성장', astNegative, 'astNegative'],
      ['sales mf', '멀티포컬 역성장', mfNegative, 'mfNegative'],
      ['sales max', 'MAX 역성장', maxNegative, 'maxNegative']
    ];

    $('gapCards').innerHTML = gapItems.map((x) => `<div class="gap-card ${x[0]}" data-gap="${x[3]}" role="button" tabindex="0"><span>${x[1]}</span><b>${x[2]}명</b><small>클릭 시 대상 리스트 보기</small></div>`).join('');

    document.querySelectorAll('#gapCards .gap-card').forEach((card) => {
      const run = () => {
        state.gapFilter = card.dataset.gap || '';
        state.query = '';
        $('smartQuery').value = '';
        metricsCache.clear();
        render();
        const label = card.querySelector('span')?.textContent || 'Gap';
        $('queryExplanation').textContent = `${label} 필터 적용 / 결과 ${state.filtered.length}명`;
        view('segment');
      };
      card.onclick = run;
      card.onkeydown = (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          run();
        }
      };
    });

    renderQuestionTop(rows);
    renderTopEdu(ms);
    renderSegment(rows, ms);
    console.timeEnd('render');
  }

  function renderQuestionTop(rows) {
    const ids = new Set(rows.map((r) => r.안경사ID));
    const cnt = {};
    state.per.filter((p) => ids.has(p.안경사ID) && p.gap).forEach((p) => {
      const key = p.문항ID + '|' + p.문항;
      cnt[key] = (cnt[key] || 0) + 1;
    });
    $('questionTop').innerHTML = Object.entries(cnt).sort((a, b) => b[1] - a[1]).slice(0, 7).map(([k, v], i) => `<div class="rank-item"><span class="rank-no">${i + 1}</span><b>${esc(k.split('|')[1])}</b><span>${v}명</span></div>`).join('') || '<div class="empty-state">인식 Gap 문항이 없습니다.</div>';
  }

  function contentName(id) {
    const c = state.content.find((x) => clean(get(x, ['교육ID'])) === clean(id));
    return clean(get(c, ['교육명', '콘텐츠명'])) || clean(id);
  }

  function renderTopEdu(ms) {
    const cnt = {};
    ms.forEach((m) => {
      const name = clean(get(m.rec, ['추천교육명', '교육명'])) || contentName(get(m.rec, ['추천교육ID', '교육ID']));
      if (name) cnt[name] = (cnt[name] || 0) + 1;
    });
    $('topEducation').innerHTML = Object.entries(cnt).sort((a, b) => b[1] - a[1]).slice(0, 8).map((x, i) => `<div class="rank-item"><span class="rank-no">${i + 1}</span><b>${esc(x[0])}</b><span>${x[1]}명</span></div>`).join('') || '<div class="empty-state">추천 교육 데이터가 없습니다.</div>';
  }

  function renderSegment(rows, ms) {
    $('resultCount').textContent = `${rows.length.toLocaleString('ko-KR')}명`;
    $('segmentSummary').innerHTML = `<div class="three-col"><div>${kpi('난시 평균 성장률', rate(avg(ms, (m) => m.growths.ast.cur)), '')}</div><div>${kpi('멀티포컬 평균 성장률', rate(avg(ms, (m) => m.growths.mf.cur)), '')}</div><div>${kpi('MAX 평균 성장률', rate(avg(ms, (m) => m.growths.max.cur)), '')}</div></div>`;
    $('segmentTable').innerHTML = ms.map((m) => {
      const p = m.p || {};
      const eduName = clean(get(m.rec, ['추천교육명', '교육명'])) || contentName(get(m.rec, ['추천교육ID', '교육ID']));
      return `<tr data-id="${esc(p.안경사ID)}"><td><b>${esc(p.안경사명)}</b><small><br>${esc(p.안경사ID)}</small></td><td>${esc(p.안경원명)}<small><br>${esc(p.지역)} · ${esc(p.채널)}</small></td><td>${esc(p.연차)} / ${esc(p.Tier)}</td><td>${m.eduRate == null ? '데이터 없음' : pct(m.eduRate)}</td><td>${m.gaps.length}개</td><td>${rate(m.growths.ast.cur)}</td><td>${rate(m.growths.mf.cur)}</td><td>${rate(m.growths.max.cur)}</td><td>${esc(eduName || '없음')}</td><td>${badge(m.priority)}</td></tr>`;
    }).join('');
    document.querySelectorAll('#segmentTable tr').forEach((tr) => { tr.onclick = () => showProfile(tr.dataset.id); });
  }

  function showProfile(id) {
    const m = metrics(id);
    if (!m.p) return;
    $('profilePanel').hidden = false;
    const cards = [
      ['교육완료', m.eduRate == null ? '데이터 없음' : pct(m.eduRate)],
      ['인식 Gap', m.gaps.length + '개'],
      ['난시 성장률', rate(m.growths.ast.cur)],
      ['멀티포컬 성장률', rate(m.growths.mf.cur)],
      ['MAX 성장률', rate(m.growths.max.cur)],
      ['우선순위', m.priority]
    ].map((x) => `<div class="status-card"><small>${x[0]}</small><h3>${x[1]}</h3></div>`).join('');
    const gaps = (m.gaps.length ? m.gaps : m.perc.slice(0, 5)).slice(0, 10).map((g) => `<div class="question-card"><b>${esc(g.문항)}</b><br><small>${esc(g.영역)} · 응답 ${esc(g.원응답)} · 목표 ${g.목표값}점 이상${g.gap ? ' · Gap' : ''}</small></div>`).join('') || '<div class="empty-state">분석 가능한 문항이 없습니다.</div>';
    $('profileContent').innerHTML = `<h3>${esc(m.p.안경사명)} <small>${esc(m.p.안경사ID)}</small></h3><p>${esc(m.p.안경원명)} · ${esc(m.p.지역)} · ${esc(m.p.연차)} / ${esc(m.p.Tier)}</p><div class="profile-grid">${cards}</div><h3>문항별 Gap</h3>${gaps}`;
    $('profilePanel').scrollIntoView({ behavior: 'smooth' });
    view('segment');
  }

  function buildFilters() {
    const f = { regionFilter: '지역', yearsFilter: '연차', tierFilter: 'Tier', channelFilter: '채널', repFilter: '담당영업사원' };
    Object.entries(f).forEach(([id, field]) => {
      const el = $(id);
      const vals = [...new Set(state.master.map((r) => clean(r[field])).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ko', { numeric: true }));
      el.innerHTML = '<option value="">전체</option>' + vals.map((v) => `<option>${esc(v)}</option>`).join('');
      el.onchange = () => { state.query = ''; state.gapFilter = ''; $('smartQuery').value = ''; metricsCache.clear(); render(); };
    });
  }

  async function upload(file) {
    if (!window.XLSX) throw new Error('XLSX 라이브러리가 로드되지 않았습니다. xlsx.full.min.js 또는 CDN 로딩을 확인하세요.');
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
    console.log('===== 업로드 파일 =====', file.name);
    console.table(wb.SheetNames);

    const masterRows = sheet(wb, aliases.master);
    const contentRows = sheet(wb, aliases.content);
    const eduRowsData = sheet(wb, aliases.edu);
    const qmRows = sheet(wb, aliases.qm);
    const perRows = sheet(wb, aliases.per);
    const salesRows = sheet(wb, aliases.sales);
    const recRowsData = sheet(wb, aliases.rec);
    const sendRowsData = sheet(wb, aliases.send);

    console.table({ master: masterRows.length, content: contentRows.length, education: eduRowsData.length, questionMaster: qmRows.length, perception: perRows.length, sales: salesRows.length, recommendation: recRowsData.length, send: sendRowsData.length });
    if (!masterRows.length) throw new Error('01_안경사마스터 시트를 찾지 못했습니다.');

    state.master = normMaster(masterRows);
    state.content = contentRows;
    state.edu = eduRowsData;
    state.qm = normQm(qmRows);
    state.perRaw = perRows;
    state.per = normPer(state.perRaw);
    state.sales = salesRows;
    state.rec = recRowsData;
    state.send = sendRowsData;

    rebuildIndexes();
    $('uploadStatus').textContent = file.name;
    buildFilters();
    render();
    toast(`업로드 완료: 안경사 ${state.master.length}명, 판매행 ${state.sales.length}건`);
  }

  function resetAll() {
    state.query = '';
    state.gapFilter = '';
    $('smartQuery').value = '';
    ['regionFilter', 'yearsFilter', 'tierFilter', 'channelFilter', 'repFilter'].forEach((id) => { if ($(id)) $(id).value = ''; });
    $('queryExplanation').textContent = '필터를 선택하거나 검색어를 입력하세요.';
    metricsCache.clear();
    render();
  }

  function seed() {
    state.master = normMaster([
      { 안경사ID: 'A001', 안경사명: '이창훈', 안경원명: '으뜸50안경 신풍점', 지역: '서울', 연차: '9년차', Tier: 'Gold', 채널: 'Top50', 담당영업사원: '유아영' },
      { 안경사ID: 'A002', 안경사명: '최용운', 안경원명: '으뜸50안경 청라점', 지역: '인천', 연차: '4년차', Tier: 'Silver', 채널: 'Top50', 담당영업사원: '이자영' }
    ]);
    state.edu = [{ 안경사ID: 'A001', 완료여부: 'N' }, { 안경사ID: 'A002', 완료여부: 'Y' }];
    state.qm = normQm([
      { 문항ID: 'Q001', 문항: '멀티포컬 추천 또는 피팅시 자신있게 응대한다', 인식영역: '멀티포컬', 응답척도: '1~5점', 긍정방향: '높을수록 긍정', 목표값: 4 },
      { 문항ID: 'Q002', 문항: '난시용 렌즈를 긴 상담시간 때문에 권유하지 않는다 ★역코딩★', 인식영역: '난시', 응답척도: '1~5점', 긍정방향: '낮을수록 긍정', 목표값: 4 }
    ]);
    state.perRaw = [
      { 안경사ID: 'A001', 'Q001_멀티포컬 추천 또는 피팅시 자신있게 응대한다': '보통이다', 'Q002_난시용 렌즈를 긴 상담시간 때문에 권유하지 않는다 ★역코딩★': '그렇다' },
      { 안경사ID: 'A002', 'Q001_멀티포컬 추천 또는 피팅시 자신있게 응대한다': '그렇다', 'Q002_난시용 렌즈를 긴 상담시간 때문에 권유하지 않는다 ★역코딩★': '그렇지 않다' }
    ];
    state.per = normPer(state.perRaw);
    state.sales = [
      { 안경사ID: 'A001', 제품명: 'MAX 구면', 2025: 100, 2026: 130 },
      { 안경사ID: 'A001', 제품명: 'MAX 난시', 2025: 20, 2026: 34 },
      { 안경사ID: 'A001', 제품명: '모이스트 난시', 2025: 200, 2026: 230 },
      { 안경사ID: 'A001', 제품명: 'MAX 멀티포컬 난시', 2025: 10, 2026: 16 },
      { 안경사ID: 'A002', 제품명: '오아시스 난시', 2025: 100, 2026: 80 },
      { 안경사ID: 'A002', 제품명: 'MAX 난시', 2025: 50, 2026: 55 },
      { 안경사ID: 'A002', 제품명: 'MAX 멀티포컬', 2025: 20, 2026: 35 }
    ];
    state.rec = [
      { 안경사ID: 'A001', 추천교육명: '멀티포컬 기초 교육', 추천사유: '멀티포컬 인식 Gap 보완 필요' },
      { 안경사ID: 'A002', 추천교육명: 'ASD 난시 시뮬레이터 체험', 추천사유: '난시 인식 및 성장률 개선 필요' }
    ];
    rebuildIndexes();
    buildFilters();
    render();
  }

  function view(id) {
    document.querySelectorAll('.view').forEach((v) => v.classList.remove('active'));
    document.querySelectorAll('.tab').forEach((t) => t.classList.remove('active'));
    $(id)?.classList.add('active');
    document.querySelector(`.tab[data-view="${id}"]`)?.classList.add('active');
  }

  function toast(msg) {
    const t = $('toast');
    t.textContent = msg;
    t.classList.add('show');
    setTimeout(() => t.classList.remove('show'), 2200);
  }

  function download() {
    const rows = state.filtered.map((p) => {
      const m = metrics(p.안경사ID);
      return { 안경사ID: p.안경사ID, 안경사명: p.안경사명, 교육완료율: m.eduRate, 인식Gap문항수: m.gaps.length, 난시성장률: m.growths.ast.cur, 멀티포컬성장률: m.growths.mf.cur, MAX성장률: m.growths.max.cur, 추천교육: get(m.rec, ['추천교육명', '교육명']), 우선순위: m.priority, 후속조치: follow(m) };
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), '대상목록');
    XLSX.writeFile(wb, 'ACUVUE_추천운영_대상목록.xlsx');
  }

  function parseCsv(text) {
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i], nx = text[i + 1];
      if (ch === '"' && q && nx === '"') { cell += '"'; i++; continue; }
      if (ch === '"') { q = !q; continue; }
      if (ch === ',' && !q) { row.push(cell); cell = ''; continue; }
      if ((ch === '\n' || ch === '\r') && !q) {
        if (ch === '\r' && nx === '\n') i++;
        row.push(cell);
        if (row.some((v) => clean(v))) rows.push(row);
        row = []; cell = ''; continue;
      }
      cell += ch;
    }
    row.push(cell);
    if (row.some((v) => clean(v))) rows.push(row);
    if (!rows.length) return [];
    const head = rows.shift().map(clean);
    return rows.map((r) => Object.fromEntries(head.map((h, i) => [h, clean(r[i])])));
  }

  function detectVendor(row) {
    const text = Object.values(row).join(' ').toUpperCase();
    const hits = COMPETITOR_RULES.filter((rule) => rule.patterns.some((p) => text.includes(p.toUpperCase()))).map((rule) => rule.vendor);
    return hits.length ? [...new Set(hits)].join(', ') : '미분류';
  }

  function originalLink(row) {
    return clean(get(row, ['원본링크', '링크', 'URL', 'url', 'link', 'Link', '게시물URL', '게시물 링크', 'source_url', 'source', 'href']));
  }

  function rowTitle(row) {
    return clean(get(row, ['제목', 'title', 'Title', '게시글제목', '게시물제목', 'subject', '본문', '내용'])) || '제목 없음';
  }

  function vendorBadge(vendor) {
    const cls = vendor.includes('알콘') ? 'neutral' : vendor.includes('바슈롬') ? 'medium' : vendor.includes('쿠퍼비전') ? 'low' : 'neutral';
    return `<span class="pill ${cls}">${esc(vendor)}</span>`;
  }

  function renderExternal(rows, source = 'output/Competitor_Activity.csv') {
    const raw = rows.length;
    rows = rows.map((r) => {
      const joined = Object.values(r).join(' ');
      return { ...r, __kw: COMPETITOR_KEYWORDS.filter((k) => joined.includes(k)), __vendor: detectVendor(r), __link: originalLink(r) };
    }).filter((r) => r.__kw.length || r.__vendor !== '미분류');

    state.competitorRows = rows;

    const vendorCount = rows.reduce((acc, r) => {
      acc[r.__vendor] = (acc[r.__vendor] || 0) + 1;
      return acc;
    }, {});

    if (!rows.length) {
      $('externalInsight').innerHTML = `CSV 연결됨. 업체/키워드 매칭 0건 / 전체 ${raw}건`;
      return;
    }

    const summary = Object.entries(vendorCount).sort((a, b) => b[1] - a[1]).map(([k, v]) => `${esc(k)} ${v}건`).join(' · ');

    $('externalInsight').innerHTML = `
      <div class="query-explanation">
        ${esc(source)} · 업체/키워드 매칭 ${rows.length}건 / 전체 ${raw}건<br>
        업체 기준: 알콘, 바슈롬, 쿠퍼비전 · ${summary}<br>
        키워드: ${COMPETITOR_KEYWORDS.join(', ')}
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>업체</th><th>매칭키워드</th><th>제목/내용</th><th>월</th><th>게시일</th><th>지역</th><th>원본 링크</th>
            </tr>
          </thead>
          <tbody>
            ${rows.slice(0, 300).map((r, i) => {
              const link = r.__link;
              return `<tr data-i="${i}">
                <td>${vendorBadge(r.__vendor)}</td>
                <td>${esc(r.__kw.join(', ') || '-')}</td>
                <td>${esc(rowTitle(r)).slice(0, 120)}</td>
                <td>${esc(get(r, ['월', 'month', 'Month']))}</td>
                <td>${esc(get(r, ['게시일', '작성일', 'date', 'Date', 'published_at']))}</td>
                <td>${esc(get(r, ['지역', 'region', 'Region']))}</td>
                <td>${link ? `<a href="${esc(link)}" target="_blank" rel="noopener noreferrer">원본 보기</a>` : '-'}</td>
              </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>`;

    document.querySelectorAll('#externalInsight tr[data-i]').forEach((tr) => {
      tr.onclick = () => {
        const r = rows[Number(tr.dataset.i)];
        $('externalDetailPanel').hidden = false;
        const link = r.__link;
        $('externalDetail').innerHTML = `
          <div class="detail-title">${esc(rowTitle(r))}</div>
          <div>${vendorBadge(r.__vendor)} ${r.__kw.map((k) => `<span class="pill neutral">${esc(k)}</span>`).join(' ')}</div>
          ${link ? `<p><a href="${esc(link)}" target="_blank" rel="noopener noreferrer">원본 링크 열기</a></p>` : ''}
          <div class="detail-grid">
            ${Object.entries(r).filter(([k]) => !k.startsWith('__')).map(([k, v]) => `<div class="detail-item"><small>${esc(k)}</small>${esc(v)}</div>`).join('')}
          </div>`;
        $('externalDetailPanel').scrollIntoView({ behavior: 'smooth' });
      };
    });
  }

  async function loadExternal() {
    try {
      const res = await fetch('output/Competitor_Activity.csv', { cache: 'no-store' });
      if (!res.ok) throw new Error();
      renderExternal(parseCsv(await res.text()));
    } catch (e) {
      $('externalInsight').innerHTML = '자동 연결 실패. 경쟁사 CSV 업로드 버튼으로 파일을 선택하세요.';
    }
  }

  document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('.tab').forEach((t) => { t.onclick = () => view(t.dataset.view); });
    $('workbookInput').onchange = async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try { await upload(file); }
      catch (err) {
        console.error('===== 업로드 실패 =====');
        console.error(err);
        alert('업로드 실패\n\n' + (err.message || err));
        toast(`업로드 실패 : ${err.message || err}`);
      }
    };
    $('runQuery').onclick = () => {
      state.query = $('smartQuery').value;
      state.gapFilter = '';
      metricsCache.clear();
      render();
      $('queryExplanation').textContent = `검색 조건 적용: ${state.query || '없음'} / 결과 ${state.filtered.length}명`;
      view('segment');
    };
    $('smartQuery').onkeydown = (e) => { if (e.key === 'Enter') $('runQuery').click(); };
    $('clearQuery').onclick = resetAll;
    $('resetFilters').onclick = resetAll;
    document.querySelectorAll('.examples button').forEach((b) => {
      b.onclick = () => {
        state.query = b.dataset.query;
        state.gapFilter = '';
        $('smartQuery').value = state.query;
        metricsCache.clear();
        render();
        $('queryExplanation').textContent = `검색 조건 적용: ${state.query} / 결과 ${state.filtered.length}명`;
        view('segment');
      };
    });
    $('downloadResults').onclick = download;
    $('closeProfile').onclick = () => { $('profilePanel').hidden = true; };
    $('competitorInput').onchange = (e) => e.target.files[0] && e.target.files[0].text().then((t) => renderExternal(parseCsv(t), e.target.files[0].name));
    $('closeExternalDetail').onclick = () => { $('externalDetailPanel').hidden = true; };
    seed();
    loadExternal();
  });
})();
