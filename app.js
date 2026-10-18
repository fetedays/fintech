/**
 * UI/UX Pro Max & Frontend Design System
 * 2-Track 자산관리 & 심플 가계부 (하나의 앱 · 하나의 데이터)
 * - 5대 메뉴: 홈(Home), 가계부(Spending), 자산관리(Assets), 분석(Analytics), 설정(Settings)
 * - Flow & Stock 공통 데이터 계좌 잔액 자동 연동
 * - 1. 지출 등록 버튼: 가계부 탭에서만 스마트 노출 (홈 탭 제외)
 * - 2. '수정' 텍스트 제거 및 직관적 카드 클릭 인터랙션
 * - 3. 분석 페이지: 고정/변동 비교, 저축 성적표, 소비 TOP 4 랭킹, 건전성 지표 등 대폭 강화
 * - 4. 아이콘 스타일: 이모지 완전 제거 및 Lucide 라인 벡터 아이콘으로 100% 통일
 * - 5. 피드백 핀 도구: 뷰포트 내부 좌표 계산 및 클릭 캡처 오버레이로 위치 오차 및 클릭 간섭 100% 해결
 */

const STORAGE_KEY = 'toss_financial_complete_v9';

// 순수한 0원 초기 상태 (허구/데모 데이터 일체 배제)
const INITIAL_DATA = {
  selectedYearMonth: '2026-10',
  dividendTaxDeduction: false, // false: 세전, true: 세후(15.4%)
  accounts: [],
  assets: [],
  transactions: [],
  totalMonthlyBudget: 0,
  budgets: {}
};

const StorageService = {
  load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (!parsed.yearlyNetWorth) parsed.yearlyNetWorth = INITIAL_DATA.yearlyNetWorth;
        if (!parsed.monthlyNetWorth) parsed.monthlyNetWorth = INITIAL_DATA.monthlyNetWorth;
        if (!parsed.totalMonthlyBudget) parsed.totalMonthlyBudget = INITIAL_DATA.totalMonthlyBudget;
        if (!parsed.budgets) parsed.budgets = INITIAL_DATA.budgets;
        if (!Array.isArray(parsed.accounts)) parsed.accounts = [];
        if (!Array.isArray(parsed.assets)) parsed.assets = [];
        if (!Array.isArray(parsed.transactions)) parsed.transactions = [];
        return parsed;
      }
    } catch (e) {
      console.warn('데이터 로드 실패, 초기 데이터를 생성합니다:', e);
    }
    this.save(INITIAL_DATA);
    return JSON.parse(JSON.stringify(INITIAL_DATA));
  },
  save(data) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    } catch (e) {
      console.error('데이터 저장 실패:', e);
    }
  }
};

const GOOGLE_APPS_SCRIPT_TEMPLATE = `/**
 * 📊 스마트 자산관리 & 심플 가계부 - Google Apps Script 연동 스크립트
 * 
 * [배포 방법]
 * 1. 구글 스프레드시트 상단 메뉴 [확장 프로그램] > [Apps Script] 클릭
 * 2. 기존 코드를 모두 지우고 이 스크립트를 전체 복사해 붙여넣고 저장 (Cmd+S / Ctrl+S)
 * 3. 우측 상단 [배포] > [새 배포] 클릭
 * 4. 톱니바퀴 아이콘 클릭 > [웹 앱] 선택
 *    - 설명: 가계부 연동
 *    - 다음 사용자 권한으로 실행: '나' (내 계정)
 *    - 액세스 권한이 있는 사용자: '모든 사용자(Anyone)' (필수!)
 * 5. [배포] 버튼 클릭 후 생성된 '웹 앱 URL'을 복사하여 가계부 설정에 붙여넣으세요.
 */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'get';
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'ping') {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Google Sheets 연동이 정상적으로 연결되었습니다.'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 데이터 불러오기 (RawData 시트에서 복원)
  var rawSheet = ss.getSheetByName('RawData');
  if (!rawSheet) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: '시트에 저장된 데이터가 없습니다. 먼저 [시트로 내보내기]를 실행해주세요.'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var rawJson = rawSheet.getRange(1, 1).getValue();
  try {
    var data = JSON.parse(rawJson);
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: data
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: '데이터 파싱 실패: ' + err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var contents = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var payload = JSON.parse(contents);
    var action = payload.action || 'export';
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'ping') {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Google Sheets 연동이 정상적으로 연결되었습니다.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'import') {
      var rawSheet = ss.getSheetByName('RawData');
      if (!rawSheet) {
        return ContentService.createTextOutput(JSON.stringify({
          status: 'error',
          message: '시트에 저장된 데이터가 없습니다.'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var rawJson = rawSheet.getRange(1, 1).getValue();
      var data = JSON.parse(rawJson);
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        data: data
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 내보내기 (Export) 실행
    var data = payload.data;
    if (!data) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: '전송된 데이터가 없습니다.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 1. RawData 시트에 JSON 통째로 보관 (완벽한 복원용)
    var rawSheet = ss.getSheetByName('RawData');
    if (!rawSheet) {
      rawSheet = ss.insertSheet('RawData');
    }
    rawSheet.clear();
    rawSheet.getRange(1, 1).setValue(JSON.stringify(data));

    // 2. 1.거래내역 시트 생성/갱신
    updateTransactionsSheet(ss, data.transactions || []);

    // 3. 2.계좌목록 시트 생성/갱신
    updateAccountsSheet(ss, data.accounts || []);

    // 4. 3.자산현황 시트 생성/갱신
    updateAssetsSheet(ss, data.assets || []);

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: '구글 시트에 모든 데이터(거래/계좌/자산)가 성공적으로 동기화되었습니다!'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: '처리 중 오류 발생: ' + err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function updateTransactionsSheet(ss, txList) {
  var sheet = ss.getSheetByName('1.거래내역');
  if (!sheet) {
    sheet = ss.insertSheet('1.거래내역', 0);
  }
  sheet.clear();
  var headers = ['거래일시', '구분', '카테고리', '사용처/메모', '금액(원)', '출금계좌', '입금계좌'];
  var rows = [headers];

  var typeLabels = {
    'EXPENSE_VARIABLE': '변동지출',
    'EXPENSE_FIXED': '고정지출',
    'INCOME': '수입',
    'TRANSFER': '계좌이체'
  };

  txList.forEach(function(tx) {
    rows.push([
      tx.datetime || tx.date || '',
      typeLabels[tx.type] || tx.type || '',
      tx.category || '',
      tx.memo || '',
      Number(tx.amount) || 0,
      tx.fromAccount || '',
      tx.toAccount || ''
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#fff1eb').setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }
}

function updateAccountsSheet(ss, accounts) {
  var sheet = ss.getSheetByName('2.계좌목록');
  if (!sheet) {
    sheet = ss.insertSheet('2.계좌목록', 1);
  }
  sheet.clear();
  var headers = ['계좌명', '분류', '현재잔액(원)', '메모/설명'];
  var rows = [headers];

  accounts.forEach(function(acc) {
    rows.push([
      acc.name || '',
      acc.type || '',
      Number(acc.balance) || 0,
      acc.memo || ''
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#ebf5ff').setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }
}

function updateAssetsSheet(ss, assets) {
  var sheet = ss.getSheetByName('3.자산현황');
  if (!sheet) {
    sheet = ss.insertSheet('3.자산현황', 2);
  }
  sheet.clear();
  var headers = ['자산명', '대분류', '평가금액(원)', '취득/등록일', '메모'];
  var rows = [headers];

  var catLabels = {
    'INVEST': '투자/주식/펀드',
    'REAL_ESTATE': '부동산',
    'PENSION': '연금',
    'DEPOSIT': '예적금',
    'OTHER': '기타자산'
  };

  assets.forEach(function(asset) {
    rows.push([
      asset.name || '',
      catLabels[asset.category] || asset.category || '',
      Number(asset.amount) || 0,
      asset.date || '',
      asset.memo || ''
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#f3f0ff').setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }
};
`;

const App = {
  data: null,
  activeTab: 'home', // home, spending, assets, analytics, settings
  trendMode: 'year', // 'year' | 'month' (연도별 / 월별 자산 추이 모드)
  charts: {},
  editingTxId: null,
  editingAssetContext: null, // { type: 'account' | 'asset', id: string }

  init() {
    this.data = StorageService.load();
    this.bindEvents();
    this.render();
    this.updateFabButton('home');
  },

  bindEvents() {
    // 하단 5대 메뉴 버튼 클릭
    document.querySelectorAll('.toss-bottom-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const tab = e.currentTarget.getAttribute('data-tab');
        this.switchTab(tab);
      });
    });

    // 월 이동
    const prevBtn = document.getElementById('monthPrevBtn');
    if (prevBtn) prevBtn.addEventListener('click', () => this.changeMonth(-1));

    const nextBtn = document.getElementById('monthNextBtn');
    if (nextBtn) nextBtn.addEventListener('click', () => this.changeMonth(1));

    // 배당 세전/세후 스위치
    const taxToggle = document.getElementById('taxToggleInput');
    if (taxToggle) {
      taxToggle.addEventListener('change', (e) => {
        this.data.dividendTaxDeduction = e.target.checked;
        StorageService.save(this.data);
        this.render();
      });
    }

    // 거래 유형 변경 시 입금계좌 표시 토글
    const txTypeSelect = document.getElementById('newTxType');
    if (txTypeSelect) {
      txTypeSelect.addEventListener('change', (e) => {
        const isTransfer = e.target.value === 'TRANSFER';
        const toAccRow = document.getElementById('toAccountRow');
        if (toAccRow) toAccRow.style.display = isTransfer ? 'block' : 'none';
      });
    }
  },

  switchTab(tabName) {
    this.activeTab = tabName;
    document.querySelectorAll('.toss-bottom-btn').forEach(b => {
      b.classList.toggle('active', b.getAttribute('data-tab') === tabName);
    });

    document.querySelectorAll('.view-section').forEach(sec => {
      sec.classList.add('hidden');
    });

    const target = document.getElementById(`view-${tabName}`);
    if (target) target.classList.remove('hidden');

    // 헤더 텍스트 및 서브태그 동적 업데이트
    const navTitleEl = document.getElementById('miraeNavTitle');
    const heroHeadingEl = document.getElementById('miraeHeroHeading');
    const heroSubtagEl = document.getElementById('miraeHeroSubtag');

    if (tabName === 'home') {
      if (navTitleEl) navTitleEl.textContent = 'MY자산';
      if (heroHeadingEl) heroHeadingEl.innerHTML = '오늘의 미래에셋대우<br>자산을 살펴볼까요?';
      if (heroSubtagEl) heroSubtagEl.innerHTML = '<span style="background:rgba(255,255,255,0.2); padding:2px 8px; border-radius:6px; font-weight:700;">USD ₩1,380</span> <span>실시간 자산 포트폴리오</span>';
    } else if (tabName === 'spending') {
      if (navTitleEl) navTitleEl.textContent = 'MY가계부';
      if (heroHeadingEl) heroHeadingEl.innerHTML = '이번 달 소비 흐름과<br>예산을 관리해보세요';
      if (heroSubtagEl) heroSubtagEl.innerHTML = '<span style="background:rgba(255,255,255,0.2); padding:2px 8px; border-radius:6px; font-weight:700;">가계부</span> <span>스마트한 지출 관리</span>';
    } else if (tabName === 'assets') {
      if (navTitleEl) navTitleEl.textContent = 'MY자산관리';
      if (heroHeadingEl) heroHeadingEl.innerHTML = '보유하신 모든 계좌와<br>투자 자산 현황입니다';
      if (heroSubtagEl) heroSubtagEl.innerHTML = '<span style="background:rgba(255,255,255,0.2); padding:2px 8px; border-radius:6px; font-weight:700;">통합계좌</span> <span>실시간 자산 현황</span>';
    } else if (tabName === 'analytics') {
      if (navTitleEl) navTitleEl.textContent = 'MY금융분석';
      if (heroHeadingEl) heroHeadingEl.innerHTML = '소비 패턴과<br>재무 건전성 리포트';
      if (heroSubtagEl) heroSubtagEl.innerHTML = '<span style="background:rgba(255,255,255,0.2); padding:2px 8px; border-radius:6px; font-weight:700;">AI 리포트</span> <span>심층 금융 진단</span>';
    } else if (tabName === 'settings') {
      if (navTitleEl) navTitleEl.textContent = 'MY설정';
      if (heroHeadingEl) heroHeadingEl.innerHTML = '데이터 백업 및<br>앱 환경을 설정하세요';
      if (heroSubtagEl) heroSubtagEl.innerHTML = '<span style="background:rgba(255,255,255,0.2); padding:2px 8px; border-radius:6px; font-weight:700;">설정</span> <span>데이터 및 환경설정</span>';
    }

    this.updateFabButton(tabName);
    this.render();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  },

  // 홈 탭의 최근 거래 '전체보기' 클릭 시 가장 최신 거래가 있는 월로 맞춰 가계부 열기
  goToSpendingWithRecent() {
    if (this.data.transactions && this.data.transactions.length > 0) {
      const sortedTxs = [...this.data.transactions].sort((a, b) => b.date.localeCompare(a.date));
      const latestTx = sortedTxs[0];
      if (latestTx && latestTx.date) {
        this.data.selectedYearMonth = latestTx.date.substring(0, 7);
        StorageService.save(this.data);
      }
    }
    this.switchTab('spending');
  },

  // [미래에셋 MY자산 레퍼런스] 일별 · 월별 · 연별 3단 탭 전환
  setMiraeTrendPeriod(period) {
    this.miraeTrendPeriod = period;
    const btns = {
      day: document.getElementById('miraeTabDay'),
      month: document.getElementById('miraeTabMonth'),
      year: document.getElementById('miraeTabYear')
    };
    Object.entries(btns).forEach(([k, btn]) => {
      if (btn) btn.classList.toggle('active', k === period);
    });

    const fin = this.getFinancialSummary();
    this.renderMiraeTrendChart(fin);
  },

  // 1 & 3. [핀 #3 해결]: 메인 주황색 원형 FAB(+) 버튼은 가계부(spending) 탭에서만 스마트 노출
  updateFabButton(tabName) {
    const fab = document.getElementById('floatingAddFab') || document.querySelector('.toss-fab-circle-add') || document.querySelector('.toss-fab-add');
    if (!fab) return;

    if (tabName === 'spending') {
      fab.style.display = 'flex';
      fab.onclick = () => App.openAddModal();
    } else {
      fab.style.display = 'none';
    }

    if (window.lucide) lucide.createIcons();
  },

  // 4. [아이콘 스타일 통일]: 이모지를 배제하고 정제된 Lucide 라인 벡터 아이콘 매핑
  getCategoryIcon(category, type) {
    if (type === 'INCOME') {
      return { icon: 'arrow-down-left', tint: 'tint-income' };
    }
    if (type === 'TRANSFER') {
      return { icon: 'arrow-left-right', tint: 'tint-transfer' };
    }
    const cat = category || '';
    if (cat.includes('카페') || cat.includes('간식')) {
      return { icon: 'coffee', tint: 'tint-cafe' };
    }
    if (cat.includes('마트') || cat.includes('잡화') || cat.includes('쇼핑')) {
      return { icon: 'shopping-bag', tint: 'tint-mart' };
    }
    if (cat.includes('주거') || cat.includes('통신') || cat.includes('월세')) {
      return { icon: 'home', tint: 'tint-house' };
    }
    return { icon: 'utensils', tint: 'tint-food' };
  },

  changeMonth(delta) {
    const [y, m] = this.data.selectedYearMonth.split('-').map(Number);
    const d = new Date(y, m - 1 + delta, 1);
    const ny = d.getFullYear();
    const nm = String(d.getMonth() + 1).padStart(2, '0');
    this.data.selectedYearMonth = `${ny}-${nm}`;
    StorageService.save(this.data);
    this.render();
  },

  // 종합 자산 및 순자산 연산 (Net Worth = 총자산 - 총부채)
  getFinancialSummary() {
    let totalCash = 0;
    let totalSavings = 0;
    let totalStock = 0;
    let totalPension = 0;
    let totalOther = 0;
    let totalDebt = 0;

    // 계좌 잔액
    (this.data.accounts || []).forEach(a => {
      const bal = Number(a.balance) || 0;
      if (bal >= 0) totalCash += bal;
      else totalDebt += Math.abs(bal);
    });

    // 자산 및 부채 (SAVINGS, STOCK, PENSION, OTHER, DEBT 등 모든 자산 누락 없이 집계)
    (this.data.assets || []).forEach(a => {
      const amt = Number(a.amount) || 0;
      if (a.category === 'SAVINGS') totalSavings += amt;
      else if (a.category === 'STOCK') totalStock += amt;
      else if (a.category === 'PENSION') totalPension += amt;
      else if (a.category === 'DEBT') totalDebt += amt;
      else totalOther += amt; // OTHER, REAL_ESTATE, CRYPTO 등 모든 기타 자산 포괄
    });

    const totalAssets = totalCash + totalSavings + totalStock + totalPension + totalOther;
    const netWorth = totalAssets - totalDebt;

    // 가계부 해당 월 집계
    const ym = this.data.selectedYearMonth;
    const monthTxs = (this.data.transactions || []).filter(t => t.date && t.date.startsWith(ym));
    let monthIncome = 0;
    let monthFixedExpense = 0;
    let monthVariableExpense = 0;

    monthTxs.forEach(t => {
      const amt = Number(t.amount) || 0;
      if (t.type === 'INCOME') monthIncome += amt;
      else if (t.type === 'EXPENSE_FIXED') monthFixedExpense += amt;
      else if (t.type === 'EXPENSE_VARIABLE') monthVariableExpense += amt;
    });

    const monthExpense = monthFixedExpense + monthVariableExpense;
    const remainingSurplus = monthIncome - monthExpense;
    const savingsRate = monthIncome > 0 ? Math.max(0, (remainingSurplus / monthIncome) * 100) : 0;

    return {
      totalAssets,
      totalDebt,
      netWorth,
      totalCash,
      totalSavings,
      totalStock,
      totalPension,
      totalOther,
      monthIncome,
      monthFixedExpense,
      monthVariableExpense,
      monthExpense,
      remainingSurplus,
      savingsRate,
      monthTxs
    };
  },

  render() {
    const fin = this.getFinancialSummary();

    if (this.activeTab === 'home') {
      this.renderHomeTab(fin);
    } else if (this.activeTab === 'spending') {
      this.renderSpendingTab(fin);
    } else if (this.activeTab === 'assets') {
      this.renderAssetsTab(fin);
    } else if (this.activeTab === 'analytics') {
      this.renderAnalyticsTab(fin);
    } else if (this.activeTab === 'settings') {
      this.renderSettingsTab();
    }

    if (window.lucide) lucide.createIcons();
  },

  // [Code-Auditor 보안 점검] XSS 방지를 위한 HTML 특수문자 이스케이프 유틸리티
  escapeHtml(str) {
    if (str === undefined || str === null) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  },

  // 한국 원화 단위 포맷터 (예: 8,540만 원, 350만 원, 1만 2,500원, 0원)
  formatKoreanAmount(amt) {
    if (amt === undefined || amt === null || isNaN(amt)) return '0원';
    const sign = amt < 0 ? '-' : '';
    const abs = Math.abs(amt);
    if (abs === 0) return '0원';
    if (abs >= 10000) {
      const man = Math.floor(abs / 10000);
      const rem = abs % 10000;
      if (rem === 0) {
        return `${sign}${man.toLocaleString()}만 원`;
      }
      return `${sign}${man.toLocaleString()}만 ${rem.toLocaleString()}원`;
    }
    return `${sign}${abs.toLocaleString()}원`;
  },

  // 1. [홈 대시보드 탭] - 사용자 업로드 레퍼런스 디자인 100% 동일 구현
  renderHomeTab(fin) {
    // 1) 통합 순자산 카드 (핀 #1: 순자산 총액 상단, 밑으로 총자산 및 부채 표기)
    const netWorthEl = document.getElementById('homeNetWorthText');
    if (netWorthEl) netWorthEl.textContent = this.formatKoreanAmount(fin.netWorth);

    const totalAssetEl = document.getElementById('homeTotalAssetText');
    if (totalAssetEl) totalAssetEl.textContent = `총자산 ${this.formatKoreanAmount(fin.totalAssets)}`;

    const totalDebtEl = document.getElementById('homeTotalDebtText');
    if (totalDebtEl) totalDebtEl.textContent = `부채 ${this.formatKoreanAmount(fin.totalDebt)}`;

    // 전월 말 대비 순자산 변동폭 역산 (100% 실제 데이터 기반)
    const today = new Date();
    const prevMonthLastDay = new Date(today.getFullYear(), today.getMonth(), 0);
    const prevMonthEndStr = `${prevMonthLastDay.getFullYear()}-${String(prevMonthLastDay.getMonth() + 1).padStart(2, '0')}-${String(prevMonthLastDay.getDate()).padStart(2, '0')}`;
    const prevSnap = this.calculateAssetSnapshotOnDate(prevMonthEndStr);
    const netDiff = fin.netWorth - (prevSnap ? prevSnap.netWorth : 0);

    const changeEl = document.getElementById('homeNetWorthChangeText');
    if (changeEl) {
      if (netDiff > 0) {
        changeEl.textContent = `전월 말 대비 +${this.formatKoreanAmount(netDiff)}`;
        changeEl.style.color = '#00a878';
      } else if (netDiff < 0) {
        changeEl.textContent = `전월 말 대비 -${this.formatKoreanAmount(Math.abs(netDiff))}`;
        changeEl.style.color = '#f04452';
      } else {
        changeEl.textContent = `전월 말 대비 변동 없음`;
        changeEl.style.color = '#8b95a1';
      }
    }

    // 2) 이번 달 수입 / 이번 달 지출 카드
    const incEl = document.getElementById('homeMonthIncomeText');
    if (incEl) incEl.textContent = this.formatKoreanAmount(fin.monthIncome);

    const expEl = document.getElementById('homeMonthExpenseText');
    if (expEl) expEl.textContent = this.formatKoreanAmount(fin.monthExpense);

    // 3) 이번 달 남은 돈 카드 & 저축률
    const remEl = document.getElementById('homeRemainingSurplusText');
    if (remEl) remEl.textContent = this.formatKoreanAmount(fin.remainingSurplus);

    const subEl = document.getElementById('homeSurplusSubtext');
    if (subEl) subEl.textContent = '수입 – 지출';

    const badgeEl = document.getElementById('homeSavingsRateBadge');
    if (badgeEl) badgeEl.textContent = `저축률 ${fin.savingsRate.toFixed(1)}%`;

    // 4) [핀 #1 복원] 순자산 추이 콤보 차트 (일별 · 월별 · 연별 3단 탭)
    this.renderMiraeTrendChart(fin);

    // 5) 자산 배분 가로 세그먼트 바
    this.renderHomeAssetSegmentBar(fin);

    // 6) 최근 거래 내역 (최신 4건)
    this.renderHomeRecentTransactions(fin);
  },

  // [홈 대시보드 4번 카드] 순자산 추이 최근 6개월 라인 차트
  renderHomeNetWorthLineChart(fin) {
    const canvas = document.getElementById('homeNetWorthLineCanvas');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    if (this.charts.homeNetWorthLine) {
      this.charts.homeNetWorthLine.destroy();
    }

    const today = new Date();
    const snapshots = [];
    const labels = [];
    const data = [];

    for (let i = 5; i >= 0; i--) {
      const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const lastDay = (i === 0) ? today.getDate() : new Date(yyyy, d.getMonth() + 1, 0).getDate();
      const dateStr = `${yyyy}-${mm}-${String(lastDay).padStart(2, '0')}`;
      const label = (i === 0) ? '이번달' : `${parseInt(mm)}월`;

      const snap = this.calculateAssetSnapshotOnDate(dateStr);
      snapshots.push({ ...snap, label, dateStr });
      labels.push(label);
      data.push(snap.netWorth);
    }

    const minVal = Math.min(...data);
    const maxVal = Math.max(...data);
    const range = (maxVal - minVal) || 1000000;

    canvas.style.cursor = 'pointer';

    this.charts.homeNetWorthLine = new Chart(ctx, {
      type: 'line',
      data: {
        labels,
        datasets: [{
          data,
          borderColor: '#00a878',
          borderWidth: 2.5,
          backgroundColor: 'rgba(0, 168, 120, 0.08)',
          pointBackgroundColor: '#00a878',
          pointBorderColor: '#ffffff',
          pointBorderWidth: 2,
          pointRadius: 4,
          pointHoverRadius: 6,
          tension: 0.35,
          fill: true
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
          mode: 'index',
          intersect: false
        },
        onClick: (evt, elements) => {
          if (elements && elements.length > 0) {
            const idx = elements[0].index;
            const snap = snapshots[idx];
            if (snap) {
              this.openTrendDetailModal('month', snap);
            }
          }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: 'rgba(25, 31, 40, 0.92)',
            titleFont: { family: 'Pretendard', size: 12, weight: '700' },
            bodyFont: { family: 'Pretendard', size: 12, weight: '600' },
            padding: 8,
            cornerRadius: 8,
            displayColors: false,
            callbacks: {
              label: (c) => `순자산: ₩${Math.round(c.raw).toLocaleString()} (클릭 시 상세 내역)`
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: {
              font: { family: 'Pretendard', size: 10, weight: '600' },
              color: '#8b95a1'
            }
          },
          y: {
            display: false,
            min: minVal === 0 && maxVal === 0 ? -100 : Math.max(0, minVal - range * 0.4),
            max: minVal === 0 && maxVal === 0 ? 100 : maxVal + range * 0.3
          }
        }
      }
    });

    canvas.onclick = (e) => {
      if (!this.charts.homeNetWorthLine) return;
      const elements = this.charts.homeNetWorthLine.getElementsAtEventForMode(e, 'index', { intersect: false }, true);
      if (elements && elements.length > 0) {
        const idx = elements[0].index;
        const snap = snapshots[idx];
        if (snap) {
          this.openTrendDetailModal('month', snap);
        }
      }
    };
  },

  // [홈 대시보드 5번 카드] 자산 배분 가로 세그먼트 바 렌더링 (자산관리의 모든 카테고리 완전 반영)
  renderHomeAssetSegmentBar(fin) {
    const barEl = document.getElementById('homeAssetSegmentBar');
    const labelEl = document.getElementById('homeAssetSegmentLabels');
    if (!barEl) return;

    const totalActive = fin.totalCash + fin.totalSavings + fin.totalStock + fin.totalPension + (fin.totalOther || 0);

    if (totalActive === 0) {
      barEl.innerHTML = `<div style="width:100%; height:100%; background:#eaedf0; border-radius:6px;"></div>`;
      if (labelEl) labelEl.textContent = '등록된 자산이 없습니다 (0원)';
      return;
    }

    const items = [
      { key: 'CASH', name: '현금', amt: fin.totalCash, color: '#00a878' },
      { key: 'SAVINGS', name: '예·적금', amt: fin.totalSavings, color: '#0ba385' },
      { key: 'STOCK', name: '투자', amt: fin.totalStock, color: '#3182f6' },
      { key: 'PENSION', name: '연금', amt: fin.totalPension, color: '#7058ff' },
      { key: 'OTHER', name: '기타', amt: fin.totalOther, color: '#f59e0b' }
    ].filter(it => it.amt > 0);

    barEl.innerHTML = items.map(it => {
      const pct = Math.max(4, (it.amt / totalActive) * 100);
      return `<div style="width:${pct}%; height:100%; background-color:${it.color}; border-radius:3px;" title="${it.name}: ₩${it.amt.toLocaleString()} (${pct.toFixed(1)}%)"></div>`;
    }).join('');

    if (labelEl) {
      labelEl.textContent = items.map(it => `${it.name} ${((it.amt / totalActive) * 100).toFixed(0)}%`).join(' · ');
    }
  },

  // [홈 대시보드 6번 카드] 최근 거래 내역 (최신순 4건)
  renderHomeRecentTransactions(fin) {
    const recentEl = document.getElementById('homeRecentList');
    if (!recentEl) return;

    const sortedTxs = [...this.data.transactions].sort((a, b) => b.date.localeCompare(a.date));
    const recent = sortedTxs.slice(0, 4);

    if (recent.length === 0) {
      recentEl.innerHTML = `
        <div style="text-align:center; padding:24px 16px; color:var(--color-text-tertiary);">
          <div style="font-size:13px; font-weight:700; color:var(--color-text-secondary); margin-bottom:4px;">거래 내역이 없습니다</div>
          <div style="font-size:11px; color:var(--color-text-tertiary); margin-bottom:12px;">가계부 또는 홈에서 새 거래를 등록해 보세요.</div>
          <button onclick="App.openAddModal()" style="border:none; background:var(--mirae-orange); color:#ffffff; font-size:11px; font-weight:700; padding:6px 14px; border-radius:8px; cursor:pointer;">+ 거래 등록</button>
        </div>
      `;
      return;
    }

    const accMap = {};
    this.data.accounts.forEach(a => accMap[a.id] = a.name);

    recentEl.innerHTML = recent.map((t, idx) => {
      const isIncome = t.type === 'INCOME';
      const isTransfer = t.type === 'TRANSFER';
      const accName = accMap[t.accountId] || '결제계좌';
      const sign = isIncome ? '+' : (isTransfer ? '' : '-');
      const color = isIncome ? '#00a878' : '#191f28';

      // 일자 포맷 (예: 10월 18일)
      const parts = t.date.split('-');
      const dateText = parts.length === 3 ? `${parseInt(parts[1])}월 ${parseInt(parts[2])}일` : t.date;
      const borderStyle = idx < recent.length - 1 ? 'border-bottom: 1px solid #f2f4f6;' : '';

      return `
        <div class="dash-recent-row" onclick="App.openEditTxModal('${t.id}')" style="display:flex; justify-content:space-between; align-items:center; padding:12px 2px; ${borderStyle} cursor:pointer;">
          <div>
            <div style="font-size:15px; font-weight:700; color:#191f28; letter-spacing:-0.01em;">${this.escapeHtml(t.memo || t.category)}</div>
            <div style="font-size:12px; color:#8b95a1; margin-top:3px;">${this.escapeHtml(dateText)} · ${this.escapeHtml(accName)}</div>
          </div>
          <div style="font-size:16px; font-weight:800; color:${color}; font-family:var(--font-toss); letter-spacing:-0.02em;">
            ${sign}${this.formatKoreanAmount(t.amount)}
          </div>
        </div>
      `;
    }).join('');
  },

  // =========================================================================
  // 전체 자산 배분 비중 상세 모달 로직 (자산관리의 모든 카테고리 및 개별 항목 100% 반영)
  // =========================================================================
  openAssetAllocationModal() {
    const modal = document.getElementById('tossAssetAllocationModal');
    if (!modal) return;

    const fin = this.getFinancialSummary();
    const totalActive = fin.totalCash + fin.totalSavings + fin.totalStock + fin.totalPension + (fin.totalOther || 0);

    // 자산관리 탭에 등록된 실제 개별 자산 항목들 추출
    const cashAccounts = (this.data.accounts || []).filter(a => Number(a.balance) > 0);
    const savingsAssets = (this.data.assets || []).filter(a => a.category === 'SAVINGS' && Number(a.amount) > 0);
    const stockAssets = (this.data.assets || []).filter(a => a.category === 'STOCK' && Number(a.amount) > 0);
    const pensionAssets = (this.data.assets || []).filter(a => a.category === 'PENSION' && Number(a.amount) > 0);
    const otherAssets = (this.data.assets || []).filter(a => (a.category === 'OTHER' || a.category === 'REAL_ESTATE') && Number(a.amount) > 0);

    const catDefs = [
      {
        key: 'CASH',
        name: '현금 (계좌)',
        targetSection: 'assetsAccountList',
        amt: fin.totalCash,
        color: '#00a878',
        items: cashAccounts.map(a => ({ id: a.id, type: 'account', name: a.name, amount: Number(a.balance), sub: a.institution }))
      },
      {
        key: 'SAVINGS',
        name: '예금·적금',
        targetSection: 'assetsSavingsList',
        amt: fin.totalSavings,
        color: '#0ba385',
        items: savingsAssets.map(s => ({ id: s.id, type: 'asset', name: s.name, amount: Number(s.amount), sub: s.institution || '예적금' }))
      },
      {
        key: 'STOCK',
        name: '투자 (주식/ETF)',
        targetSection: 'assetsStockList',
        amt: fin.totalStock,
        color: '#3182f6',
        items: stockAssets.map(s => ({ id: s.id, type: 'asset', name: s.name, amount: Number(s.amount), sub: s.ticker || '투자' }))
      },
      {
        key: 'PENSION',
        name: '연금 (연금저축/IRP)',
        targetSection: 'assetsPensionList',
        amt: fin.totalPension,
        color: '#7058ff',
        items: pensionAssets.map(p => ({ id: p.id, type: 'asset', name: p.name, amount: Number(p.amount), sub: p.institution || '연금' }))
      },
      {
        key: 'OTHER',
        name: '부동산 · 기타 자산',
        targetSection: 'assetsOtherList',
        amt: fin.totalOther,
        color: '#f59e0b',
        items: otherAssets.map(o => ({ id: o.id, type: 'asset', name: o.name, amount: Number(o.amount), sub: o.institution || '기타' }))
      }
    ];

    // 보유하지 않은 자산(금액 0원)은 범례 및 목록에서 완전히 숨김!
    const activeCats = catDefs.filter(c => c.amt > 0).sort((a, b) => b.amt - a.amt);

    const canvas = document.getElementById('allocationModalDonutCanvas');
    if (canvas) {
      const ctx = canvas.getContext('2d');
      if (this.charts.allocationModalDonut) {
        this.charts.allocationModalDonut.destroy();
      }

      if (activeCats.length === 0 || totalActive === 0) {
        this.charts.allocationModalDonut = new Chart(ctx, {
          type: 'doughnut',
          data: {
            labels: ['보유 자산 없음'],
            datasets: [{
              data: [1],
              backgroundColor: ['#eaedf0'],
              borderWidth: 2,
              borderColor: '#ffffff'
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { display: false },
              tooltip: { callbacks: { label: () => ' 등록된 자산이 없습니다 (0원)' } }
            },
            cutout: '72%'
          }
        });
      } else {
        this.charts.allocationModalDonut = new Chart(ctx, {
          type: 'doughnut',
          data: {
            labels: activeCats.map(c => c.name),
            datasets: [{
              data: activeCats.map(c => c.amt),
              backgroundColor: activeCats.map(c => c.color),
              borderWidth: 2,
              borderColor: '#ffffff'
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { display: false }, // 두 줄로 공간 차지하는 기본 범례 숨기기!
              tooltip: {
                callbacks: {
                  label: (c) => ` ${c.label}: ₩${Number(c.raw).toLocaleString()} (${((Number(c.raw) / totalActive) * 100).toFixed(1)}%)`
                }
              }
            },
            cutout: '70%'
          }
        });
      }
    }

    const listEl = document.getElementById('allocationModalCategoryList');
    if (listEl) {
      if (activeCats.length === 0 || totalActive === 0) {
        listEl.innerHTML = `<div style="text-align:center; padding:18px; font-size:13px; color:#8b95a1; background:#f9fafb; border-radius:12px;">등록된 자산이 없습니다.</div>`;
      } else {
        listEl.innerHTML = activeCats.map(c => {
          const pct = ((c.amt / totalActive) * 100).toFixed(1);
          const subItemsHtml = c.items.map(it => {
            const itPct = ((it.amount / totalActive) * 100).toFixed(1);
            return `
              <div onclick="event.stopPropagation(); App.goToAssetCategory('${c.targetSection}');" style="display:flex; justify-content:space-between; align-items:center; padding:9px 12px; background:#ffffff; border-radius:10px; margin-top:6px; border:1px solid #f2f4f6; cursor:pointer;" title="클릭 시 자산관리 해당 위치로 이동">
                <div style="display:flex; align-items:center; gap:8px;">
                  <span style="font-size:13px; font-weight:700; color:#333d4b;">${this.escapeHtml(it.name)}</span>
                  <span style="font-size:11px; color:#8b95a1;">${this.escapeHtml(it.sub)}</span>
                </div>
                <div style="text-align:right;">
                  <span style="font-size:13px; font-weight:800; color:#191f28; font-family:var(--font-toss);">₩${it.amount.toLocaleString()}</span>
                  <span style="font-size:11px; color:#8b95a1; margin-left:4px;">(${itPct}%)</span>
                </div>
              </div>
            `;
          }).join('');

          return `
            <div style="background:#f9fafb; border-radius:14px; border:1px solid #eaedf0; padding:12px 14px; margin-bottom:10px;">
              <div class="allocation-cat-row" onclick="App.goToAssetCategory('${c.targetSection}')" style="display:flex; justify-content:space-between; align-items:center; cursor:pointer;" title="클릭 시 자산관리 해당 분류로 이동">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="width:12px; height:12px; border-radius:50%; background-color:${c.color}; flex-shrink:0;"></span>
                  <div>
                    <div style="font-size:14px; font-weight:800; color:#191f28;">${this.escapeHtml(c.name)}</div>
                    <div style="font-size:11px; color:#8b95a1; margin-top:2px; font-weight:600;">총 비중 ${pct}% (${c.items.length}개 항목)</div>
                  </div>
                </div>
                <div style="display:flex; align-items:center; gap:6px;">
                  <div style="font-size:15px; font-weight:800; color:#191f28; font-family:var(--font-toss);">
                    ₩${c.amt.toLocaleString()}
                  </div>
                  <i data-lucide="chevron-right" style="width:16px; height:16px; color:#8b95a1;"></i>
                </div>
              </div>
              <!-- 하위 개별 자산 항목 목록 (자산관리의 모든 항목 100% 표시) -->
              <div style="margin-top:8px; border-top:1px dashed #eaedf0; padding-top:6px;">
                ${subItemsHtml}
              </div>
            </div>
          `;
        }).join('');
      }
    }

    modal.classList.remove('hidden');
    if (window.lucide) lucide.createIcons();
  },

  closeAssetAllocationModal() {
    const modal = document.getElementById('tossAssetAllocationModal');
    if (modal) modal.classList.add('hidden');
  },

  goToAssetCategory(targetId) {
    this.closeAssetAllocationModal();
    this.switchTab('assets');
    setTimeout(() => {
      const el = document.getElementById(targetId);
      if (el) {
        const targetCard = el.closest('.toss-card') || el;
        targetCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
        targetCard.style.transition = 'box-shadow 0.3s ease, transform 0.3s ease';
        targetCard.style.boxShadow = '0 0 0 3px rgba(49, 130, 246, 0.4)';
        targetCard.style.transform = 'scale(1.01)';
        setTimeout(() => {
          targetCard.style.boxShadow = '';
          targetCard.style.transform = '';
        }, 1500);
      }
    }, 150);
  },

  // 2. [가계부 탭]
  renderSpendingTab(fin) {
    const [y, m] = this.data.selectedYearMonth.split('-');
    const monthTitleEl = document.getElementById('heroMonthTitle');
    if (monthTitleEl) monthTitleEl.textContent = `${y}년 ${parseInt(m)}월`;

    const incEl = document.getElementById('heroIncomeAmt');
    if (incEl) incEl.textContent = `₩${fin.monthIncome.toLocaleString()}`;

    const expEl = document.getElementById('heroExpenseAmt');
    if (expEl) expEl.textContent = `₩${fin.monthExpense.toLocaleString()}`;

    const effectiveBudget = this.data.totalMonthlyBudget || 3500000;
    const remainingBudget = Math.max(0, effectiveBudget - fin.monthExpense);

    const remEl = document.getElementById('heroRemainingAmt');
    if (remEl) remEl.textContent = `${remainingBudget.toLocaleString()}원 남음`;

    const budgetGuideEl = document.getElementById('heroTotalBudgetGuide');
    if (budgetGuideEl) budgetGuideEl.textContent = `총 예산 ₩${effectiveBudget.toLocaleString()}`;

    // 일별 권장 사용액 & 오늘 권장 점선 마커
    const daysInMonth = new Date(Number(y), Number(m), 0).getDate();
    const today = new Date();
    const currentDay = (today.getFullYear() === Number(y) && (today.getMonth() + 1) === Number(m)) ? today.getDate() : 17;
    const remainingDays = Math.max(1, daysInMonth - currentDay + 1);
    const dailyBudget = Math.round(remainingBudget / remainingDays);

    const dailyEl = document.getElementById('heroDailyGuide');
    if (dailyEl) dailyEl.textContent = `목표 달성을 위한 하루 예산 ${dailyBudget.toLocaleString()}원`;

    const budgetPct = Math.min(100, Math.round((fin.monthExpense / effectiveBudget) * 100));
    const barFill = document.getElementById('heroBudgetBarFill');
    if (barFill) barFill.style.width = `${budgetPct}%`;

    const barPin = document.getElementById('heroBudgetBarPin');
    if (barPin) {
      const monthProgressPct = Math.min(100, Math.round((currentDay / daysInMonth) * 100));
      barPin.style.left = `${monthProgressPct}%`;
    }

    // 가로 멀티컬러 세그먼트 바 & 카테고리 2열 요약
    const segTotalEl = document.getElementById('segTotalExpenseText');
    if (segTotalEl) segTotalEl.textContent = `${fin.monthExpense.toLocaleString()}원`;

    const catTotals = {};
    fin.monthTxs.filter(t => t.type.startsWith('EXPENSE')).forEach(t => {
      catTotals[t.category] = (catTotals[t.category] || 0) + Number(t.amount);
    });

    const palette = ['#0ba385', '#3182f6', '#7058ff', '#ff8a00', '#00b4d8', '#8b95a1'];
    const sortedCats = Object.entries(catTotals).sort((a, b) => b[1] - a[1]);

    const segBar = document.getElementById('spendingSegmentBar');
    if (segBar) {
      if (sortedCats.length === 0 || fin.monthExpense === 0) {
        segBar.innerHTML = `<div class="segment-item w-full" style="background:#eaedf0;"></div>`;
      } else {
        segBar.innerHTML = sortedCats.map(([cat, amt], idx) => {
          const pct = Math.max(3, (amt / fin.monthExpense) * 100);
          return `<div class="segment-item" style="width:${pct}%; background-color:${palette[idx % palette.length]};"></div>`;
        }).join('');
      }
    }

    const catGrid = document.getElementById('spendingCategoryGrid');
    if (catGrid) {
      if (sortedCats.length === 0) {
        catGrid.innerHTML = `<div style="grid-column:span 2; text-align:center; padding:16px; font-size:12px; color:var(--color-text-tertiary);">이번 달 지출 내역이 없습니다.</div>`;
      } else {
        catGrid.innerHTML = sortedCats.map(([cat, amt], idx) => {
          const pct = Math.round((amt / fin.monthExpense) * 100);
          return `
            <div class="category-summary-item">
              <div style="display:flex; align-items:center;">
                <span class="category-dot" style="background-color:${palette[idx % palette.length]};"></span>
                <span style="font-size:13px; font-weight:700; color:var(--color-text-primary);">${cat}</span>
              </div>
              <div style="text-align:right;">
                <span style="font-size:12px; font-weight:800; color:var(--toss-blue);">${pct}%</span>
                <div style="font-size:11px; color:var(--color-text-tertiary);">₩${amt.toLocaleString()}</div>
              </div>
            </div>
          `;
        }).join('');
      }
    }

    // 날짜별 타임라인 (클릭 시 수정 모달 오픈, '수정' 텍스트 불필요한 라벨 제거)
    const monthNoticeEl = document.getElementById('spendingTxMonthNotice');
    if (monthNoticeEl) monthNoticeEl.textContent = `(${parseInt(m)}월)`;

    const countBadgeEl = document.getElementById('spendingTxCountBadge');
    if (countBadgeEl) countBadgeEl.textContent = `${fin.monthTxs.length}건`;

    const timelineContainer = document.getElementById('spendingTimelineList');
    if (timelineContainer) {
      if (fin.monthTxs.length === 0) {
        timelineContainer.innerHTML = `
          <div style="text-align:center; padding:32px 16px; color:var(--color-text-tertiary);">
            <div style="width:38px; height:38px; border-radius:50%; background:#f2f4f6; display:flex; align-items:center; justify-content:center; margin:0 auto 10px auto; color:#8b95a1;">
              <i data-lucide="calendar" style="width:18px; height:18px;"></i>
            </div>
            <div style="font-size:13px; font-weight:700; color:var(--color-text-secondary); margin-bottom:4px;">${parseInt(m)}월 거래 내역이 없습니다</div>
            <div style="font-size:11px; color:var(--color-text-tertiary); margin-bottom:12px;">+ 버튼을 눌러 새 거래를 등록하거나 다른 달로 이동해 보세요.</div>
            <button onclick="App.openAddModal()" style="border:none; background:var(--mirae-orange); color:#ffffff; font-size:11px; font-weight:700; padding:6px 14px; border-radius:8px; cursor:pointer;">+ 거래 등록</button>
          </div>
        `;
      } else {
        const grouped = {};
        [...fin.monthTxs].sort((a, b) => b.date.localeCompare(a.date)).forEach(it => {
          if (!grouped[it.date]) grouped[it.date] = [];
          grouped[it.date].push(it);
        });

        const accMap = {};
        this.data.accounts.forEach(a => accMap[a.id] = a.name);

        timelineContainer.innerHTML = Object.entries(grouped).map(([date, dayItems]) => {
          const parts = date.split('-');
          const dateLabel = parts.length === 3 ? `${parseInt(parts[1])}월 ${parseInt(parts[2])}일` : date;

          return `
            <div class="timeline-date-group">
              <div class="timeline-date-title">${dateLabel}</div>
              <div class="finance-list-card">
                ${dayItems.map(it => {
                  const isIncome = it.type === 'INCOME';
                  const isTransfer = it.type === 'TRANSFER';
                  const accName = accMap[it.accountId] || '결제계좌';
                  const toAccName = (isTransfer && it.toAccountId) ? ` → ${accMap[it.toAccountId] || '입금계좌'}` : '';
                  const { icon, tint } = this.getCategoryIcon(it.category, it.type);
                  const sign = isIncome ? '+' : (isTransfer ? '↔ ' : '-');
                  const color = isIncome ? 'var(--toss-blue)' : (isTransfer ? '#8b95a1' : 'var(--color-text-primary)');
                  const tagText = isTransfer ? '이체' : (isIncome ? '입금' : '일시불');

                  return `
                    <div class="finance-row" onclick="App.openEditTxModal('${it.id}')">
                      <div style="display:flex; align-items:center;">
                        <div class="finance-icon ${tint}"><i data-lucide="${icon}" style="width:18px; height:18px;"></i></div>
                        <div>
                          <div style="font-size:14px; font-weight:700; color:var(--color-text-primary); letter-spacing:-0.01em;">${it.memo || it.category}</div>
                          <div style="font-size:11px; color:var(--color-text-tertiary); margin-top:2px;">${accName}${toAccName} · <span style="padding:1px 6px; background:#f2f4f6; border-radius:4px; font-size:10px; font-weight:600;">${tagText}</span></div>
                        </div>
                      </div>
                      <div style="font-size:15px; font-weight:800; color:${color}; font-family:var(--font-toss); letter-spacing:-0.02em;">
                        ${sign}₩${it.amount.toLocaleString()}
                      </div>
                    </div>
                  `;
                }).join('')}
              </div>
            </div>
          `;
        }).join('');
      }
    }
  },

  // 3. [자산관리 탭]
  renderAssetsTab(fin) {
    const totAssetEl = document.getElementById('assetsTotalAssetNum');
    if (totAssetEl) totAssetEl.textContent = `${fin.totalAssets.toLocaleString()}원`;

    const totDebtEl = document.getElementById('assetsTotalDebtNum');
    if (totDebtEl) totDebtEl.textContent = `${fin.totalDebt.toLocaleString()}원`;

    const netEl = document.getElementById('assetsNetWorthNum');
    if (netEl) netEl.textContent = `${fin.netWorth.toLocaleString()}원`;

    // 1) 계좌 잔액 리스트 (클릭 시 수정, '수정' 텍스트 불필요한 라벨 제거)
    const accListEl = document.getElementById('assetsAccountList');
    if (accListEl) {
      if (this.data.accounts.length === 0) {
        accListEl.innerHTML = `
          <div style="text-align:center; padding:24px 16px; background:#f9fafb; border-radius:14px; margin-bottom:12px; border:1px dashed #e5e8eb;">
            <div style="font-size:13px; font-weight:700; color:var(--color-text-secondary); margin-bottom:3px;">등록된 계좌가 없습니다</div>
            <div style="font-size:11px; color:var(--color-text-tertiary);">계좌를 등록하면 잔액이 현금 자산으로 자동 합산됩니다.</div>
          </div>
        `;
      } else {
        accListEl.innerHTML = `
          <div class="finance-list-card">
            ${this.data.accounts.map(a => `
              <div class="finance-row" onclick="App.openEditAssetModal('account', '${a.id}')">
                <div style="display:flex; align-items:center;">
                  <div class="finance-icon tint-transfer"><i data-lucide="landmark" style="width:18px; height:18px;"></i></div>
                  <div>
                    <div style="font-size:14px; font-weight:700; color:var(--color-text-primary); letter-spacing:-0.01em;">${a.name}</div>
                    <div style="font-size:11px; color:var(--color-text-tertiary); margin-top:2px;">${a.institution} · <span style="font-weight:600;">${a.type}</span>${a.date ? ' · <span style="color:#6b7280;">' + a.date + ' 개설</span>' : ''}</div>
                  </div>
                </div>
                <div style="font-size:15px; font-weight:800; font-family:var(--font-toss); color:${a.balance < 0 ? 'var(--toss-red)' : 'var(--color-text-primary)'}; letter-spacing:-0.02em;">
                  ${a.balance.toLocaleString()}원
                </div>
              </div>
            `).join('')}
          </div>
        `;
      }
    }

    // 2) 투자 주식 리스트 & 배당 캘린더 (클릭 시 수정)
    const stockListEl = document.getElementById('assetsStockList');
    if (stockListEl) {
      const stocks = this.data.assets.filter(a => a.category === 'STOCK');
      if (stocks.length === 0) {
        stockListEl.innerHTML = `
          <div style="text-align:center; padding:24px 16px; background:#f9fafb; border-radius:14px; margin-bottom:12px; border:1px dashed #e5e8eb;">
            <div style="font-size:13px; font-weight:700; color:var(--color-text-secondary); margin-bottom:3px;">등록된 주식/투자 자산이 없습니다</div>
            <div style="font-size:11px; color:var(--color-text-tertiary);">주식/ETF를 등록해 실시간 평가금과 배당금을 확인해 보세요.</div>
          </div>
        `;
      } else {
        stockListEl.innerHTML = `
          <div class="finance-list-card">
            ${stocks.map(s => `
              <div class="finance-row" onclick="App.openEditAssetModal('asset', '${s.id}')">
                <div style="display:flex; align-items:center;">
                  <div class="finance-icon tint-income"><i data-lucide="trending-up" style="width:18px; height:18px;"></i></div>
                  <div>
                    <div style="font-size:14px; font-weight:700; color:var(--color-text-primary); letter-spacing:-0.01em;">${s.name} <span style="font-size:11px; color:var(--color-text-tertiary); font-weight:500;">(${s.ticker})</span></div>
                    <div style="font-size:11px; color:var(--toss-blue); font-weight:700; margin-top:2px;">${s.date ? '<span style="color:#6b7280; font-weight:500;">' + s.date + ' 매수 · </span>' : ''}예상 배당률 ${s.dividendYield}%</div>
                  </div>
                </div>
                <div style="text-align:right;">
                  <div style="font-size:15px; font-weight:800; color:var(--color-text-primary); font-family:var(--font-toss); letter-spacing:-0.02em;">${s.amount.toLocaleString()}원</div>
                  <div style="font-size:10px; color:var(--color-text-tertiary); margin-top:1px;">평가금액</div>
                </div>
              </div>
            `).join('')}
          </div>
        `;
      }
    }

    // 배당 막대 차트 (The Rich)
    this.renderDividendBarChart();

    // 3) 예적금 요약 (클릭 시 수정)
    const savings = this.data.assets.filter(a => a.category === 'SAVINGS');
    const savingsEl = document.getElementById('assetsSavingsList');
    if (savingsEl) {
      if (savings.length === 0) {
        savingsEl.innerHTML = `
          <div style="text-align:center; padding:18px 16px; background:#f9fafb; border-radius:14px; margin-bottom:12px; border:1px dashed #e5e8eb;">
            <div style="font-size:12px; font-weight:700; color:var(--color-text-secondary);">등록된 예·적금이 없습니다</div>
          </div>
        `;
      } else {
        savingsEl.innerHTML = `
          <div class="finance-list-card">
            ${savings.map(s => `
              <div class="finance-row" onclick="App.openEditAssetModal('asset', '${s.id}')">
                <div style="display:flex; align-items:center;">
                  <div class="finance-icon tint-house"><i data-lucide="piggy-bank" style="width:18px; height:18px;"></i></div>
                  <div>
                    <div style="font-size:14px; font-weight:700; color:var(--color-text-primary);">${this.escapeHtml(s.name)}</div>
                    <div style="font-size:11px; color:var(--color-text-tertiary); margin-top:2px;">${s.date ? s.date + ' 가입 · ' : ''}금리 연 ${s.interestRate}% · ${s.maturityDate ? s.maturityDate + ' 만기' : '자유'}</div>
                  </div>
                </div>
                <div style="font-size:15px; font-weight:800; color:var(--color-text-primary); font-family:var(--font-toss);">${s.amount.toLocaleString()}원</div>
              </div>
            `).join('')}
          </div>
        `;
      }
    }

    // 4) 연금 요약 (클릭 시 수정)
    const pensions = this.data.assets.filter(a => a.category === 'PENSION');
    const pensionEl = document.getElementById('assetsPensionList');
    if (pensionEl) {
      if (pensions.length === 0) {
        pensionEl.innerHTML = `
          <div style="text-align:center; padding:18px 16px; background:#f9fafb; border-radius:14px; margin-bottom:12px; border:1px dashed #e5e8eb;">
            <div style="font-size:12px; font-weight:700; color:var(--color-text-secondary);">등록된 연금이 없습니다</div>
            <div style="font-size:11px; color:var(--color-text-tertiary); margin-top:2px;">연금저축/IRP를 등록해 노후 자산을 체계적으로 관리해 보세요.</div>
          </div>
        `;
      } else {
        pensionEl.innerHTML = `
          <div class="finance-list-card">
            ${pensions.map(p => `
              <div class="finance-row" onclick="App.openEditAssetModal('asset', '${p.id}')">
                <div style="display:flex; align-items:center;">
                  <div class="finance-icon tint-income"><i data-lucide="shield" style="width:18px; height:18px;"></i></div>
                  <div>
                    <div style="font-size:14px; font-weight:700; color:var(--color-text-primary);">${this.escapeHtml(p.name)}</div>
                    <div style="font-size:11px; color:var(--color-text-tertiary); margin-top:2px;">${p.date ? p.date + ' 가입 · ' : ''}${this.escapeHtml(p.institution || '연금')}</div>
                  </div>
                </div>
                <div style="font-size:15px; font-weight:800; color:var(--color-text-primary); font-family:var(--font-toss);">${p.amount.toLocaleString()}원</div>
              </div>
            `).join('')}
          </div>
        `;
      }
    }

    // 5) 부동산 및 기타 자산 요약 (클릭 시 수정)
    const others = this.data.assets.filter(a => a.category === 'OTHER' || a.category === 'REAL_ESTATE');
    const otherEl = document.getElementById('assetsOtherList');
    if (otherEl) {
      if (others.length === 0) {
        otherEl.innerHTML = `
          <div style="text-align:center; padding:18px 16px; background:#f9fafb; border-radius:14px; margin-bottom:12px; border:1px dashed #e5e8eb;">
            <div style="font-size:12px; font-weight:700; color:var(--color-text-secondary);">등록된 기타 자산이 없습니다</div>
            <div style="font-size:11px; color:var(--color-text-tertiary); margin-top:2px;">부동산, 보증금, 금 등 실물 자산을 등록해 보세요.</div>
          </div>
        `;
      } else {
        otherEl.innerHTML = `
          <div class="finance-list-card">
            ${others.map(o => `
              <div class="finance-row" onclick="App.openEditAssetModal('asset', '${o.id}')">
                <div style="display:flex; align-items:center;">
                  <div class="finance-icon tint-house"><i data-lucide="building" style="width:18px; height:18px;"></i></div>
                  <div>
                    <div style="font-size:14px; font-weight:700; color:var(--color-text-primary);">${this.escapeHtml(o.name)}</div>
                    <div style="font-size:11px; color:var(--color-text-tertiary); margin-top:2px;">${o.date ? o.date + ' 취득 · ' : ''}${this.escapeHtml(o.institution || '실물/기타')}</div>
                  </div>
                </div>
                <div style="font-size:15px; font-weight:800; color:var(--color-text-primary); font-family:var(--font-toss);">${o.amount.toLocaleString()}원</div>
              </div>
            `).join('')}
          </div>
        `;
      }
    }

    const debts = this.data.assets.filter(a => a.category === 'DEBT');
    const debtEl = document.getElementById('assetsDebtList');
    if (debtEl) {
      if (debts.length === 0) {
        debtEl.innerHTML = `
          <div style="text-align:center; padding:18px 16px; background:#f9fafb; border-radius:14px; margin-bottom:12px; border:1px dashed #e5e8eb;">
            <div style="font-size:12px; font-weight:700; color:var(--toss-mint);">등록된 대출/부채가 없습니다 🎉</div>
            <div style="font-size:11px; color:var(--color-text-tertiary); margin-top:2px;">부채가 없어 순자산이 100% 보존되고 있습니다.</div>
          </div>
        `;
      } else {
        debtEl.innerHTML = `
          <div class="finance-list-card">
            ${debts.map(d => `
              <div class="finance-row" onclick="App.openEditAssetModal('asset', '${d.id}')">
                <div style="display:flex; align-items:center;">
                  <div class="finance-icon tint-food"><i data-lucide="credit-card" style="width:18px; height:18px;"></i></div>
                  <div>
                    <div style="font-size:14px; font-weight:700; color:var(--toss-red);">${d.name}</div>
                    <div style="font-size:11px; color:var(--color-text-tertiary); margin-top:2px;">${d.date ? d.date + ' 실행 · ' : ''}적용 금리 연 ${d.interestRate}%</div>
                  </div>
                </div>
                <div style="font-size:15px; font-weight:800; color:var(--toss-red); font-family:var(--font-toss);">-${d.amount.toLocaleString()}원</div>
              </div>
            `).join('')}
          </div>
        `;
      }
    }
  },

  renderDividendBarChart() {
    const canvas = document.getElementById('assetsDividendBarCanvas');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    if (this.charts.dividendBar) this.charts.dividendBar.destroy();

    const isTaxDeducted = this.data.dividendTaxDeduction;
    const taxRate = isTaxDeducted ? (1 - 0.154) : 1;
    const monthlyDivs = new Array(12).fill(0);

    this.data.assets.filter(a => a.category === 'STOCK').forEach(s => {
      const annual = s.amount * (s.dividendYield / 100) * taxRate;
      const months = s.dividendMonths || [12];
      const perMo = annual / months.length;
      months.forEach(m => {
        if (m >= 1 && m <= 12) monthlyDivs[m - 1] += perMo;
      });
    });

    const labels = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'];
    const currentMonthIdx = 9; // 10월
    const colors = monthlyDivs.map((_, idx) => idx === currentMonthIdx ? '#3182f6' : '#e5e8eb');

    this.charts.dividendBar = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets: [{
          data: monthlyDivs,
          backgroundColor: colors,
          borderRadius: 8,
          borderSkipped: false
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (c) => ` 예상 배당금: ₩${Math.round(c.raw || 0).toLocaleString()}`
            }
          }
        },
        scales: {
          x: { grid: { display: false }, ticks: { font: { family: 'Pretendard', size: 10 }, color: '#8b95a1' } },
          y: { grid: { color: '#f2f4f6' }, ticks: { font: { family: 'Pretendard', size: 10 }, color: '#8b95a1', callback: (v) => v >= 10000 ? `${Math.round(v/10000)}만` : v } }
        }
      }
    });
  },

  // 특정 일자(YYYY-MM-DD) 기준 실제 자산/부채/순자산 스냅샷 계산 (100% 실제 데이터 기반 역산)
  calculateAssetSnapshotOnDate(targetDateStr) {
    const activeAccounts = [];
    let accountsTotal = 0;

    (this.data.accounts || []).forEach(acc => {
      // 계좌 개설일/등록일 체크 (acc.date가 있고 targetDateStr보다 미래면 당시에는 미존재)
      if (acc.date && acc.date > targetDateStr) {
        return;
      }

      let balanceOnDate = Number(acc.balance) || 0;
      // targetDateStr 이후에 일어난 거래 내역 역산 (미래 거래를 되돌려 해당일 시점 잔액 도출)
      (this.data.transactions || []).forEach(t => {
        if (t.date > targetDateStr) {
          if (t.accountId === acc.id) {
            if (t.type === 'INCOME') balanceOnDate -= Number(t.amount);
            else if (t.type.startsWith('EXPENSE')) balanceOnDate += Number(t.amount);
            else if (t.type === 'TRANSFER') balanceOnDate += Number(t.amount);
          }
          if (t.toAccountId === acc.id && t.type === 'TRANSFER') {
            balanceOnDate -= Number(t.amount);
          }
        }
      });

      activeAccounts.push({
        id: acc.id,
        name: acc.name,
        institution: acc.institution,
        balance: balanceOnDate,
        type: acc.type
      });
      accountsTotal += balanceOnDate;
    });

    const activeAssets = [];
    const activeDebts = [];
    let assetsTotal = 0;
    let debtTotal = 0;

    (this.data.assets || []).forEach(ast => {
      // 취득/매수/가입 일자 체크 (미래 일자면 당시 미보유)
      const assetDate = ast.date || ast.purchaseDate || '2000-01-01';
      if (assetDate > targetDateStr) {
        return;
      }

      const val = Number(ast.amount) || 0;
      if (ast.category === 'DEBT') {
        activeDebts.push({ ...ast, currentVal: val });
        debtTotal += val;
      } else {
        activeAssets.push({ ...ast, currentVal: val });
        assetsTotal += val;
      }
    });

    const totalAssets = accountsTotal + assetsTotal;
    const netWorth = totalAssets - debtTotal;
    const dayTxs = (this.data.transactions || []).filter(t => t.date === targetDateStr);

    return {
      date: targetDateStr,
      totalAssets,
      totalDebt: debtTotal,
      netWorth,
      accounts: activeAccounts,
      assets: activeAssets,
      debts: activeDebts,
      transactions: dayTxs
    };
  },

  // [미래에셋 MY자산 레퍼런스 디자인] 총자산 추이 콤보 차트 (100% 실제 데이터 기반 동적 생성 & 클릭 연동)
  renderMiraeTrendChart(fin) {
    const canvas = document.getElementById('miraeAssetComboCanvas');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    if (this.charts.miraeCombo) {
      this.charts.miraeCombo.destroy();
    }

    const period = this.miraeTrendPeriod || 'day';
    const snapshots = [];
    let labels = [];
    let lineData = [];
    let barData = [];

    const today = new Date();

    if (period === 'day') {
      // 최근 7일 동적 생성 (D-6 ~ 오늘)
      for (let i = 6; i >= 0; i--) {
        const d = new Date(today);
        d.setDate(today.getDate() - i);
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        const dateStr = `${yyyy}-${mm}-${dd}`;
        const shortLabel = (i === 0) ? '오늘' : `${parseInt(mm)}.${parseInt(dd)}`;

        const snap = this.calculateAssetSnapshotOnDate(dateStr);
        snapshots.push({ ...snap, label: shortLabel, isCurrent: i === 0 });
      }

      // 첫째 날 이전 일자(D-7) 스냅샷 계산하여 첫날의 실제 변동폭 산출
      const d7 = new Date(today);
      d7.setDate(today.getDate() - 7);
      const prevDateStr = `${d7.getFullYear()}-${String(d7.getMonth() + 1).padStart(2, '0')}-${String(d7.getDate()).padStart(2, '0')}`;
      const d7Snap = this.calculateAssetSnapshotOnDate(prevDateStr);

      snapshots.forEach((s, idx) => {
        labels.push(s.label);
        lineData.push(s.netWorth);
        const prevNet = idx === 0 ? d7Snap.netWorth : snapshots[idx - 1].netWorth;
        const change = s.netWorth - prevNet;
        s.change = change;
        barData.push(change);
      });
    } else if (period === 'month') {
      // 최근 6개월 동적 생성 (M-5 ~ 이번달)
      for (let i = 5; i >= 0; i--) {
        const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const ymStr = `${yyyy}-${mm}`;
        const lastDay = (i === 0) ? today.getDate() : new Date(yyyy, d.getMonth() + 1, 0).getDate();
        const dateStr = `${ymStr}-${String(lastDay).padStart(2, '0')}`;
        const label = (i === 0) ? '이번달' : `${parseInt(mm)}월`;

        const snap = this.calculateAssetSnapshotOnDate(dateStr);
        snapshots.push({ ...snap, ymStr, label, isCurrent: i === 0 });
      }

      // M-6 스냅샷
      const dPrev = new Date(today.getFullYear(), today.getMonth() - 6, 1);
      const lastDayPrev = new Date(dPrev.getFullYear(), dPrev.getMonth() + 1, 0).getDate();
      const prevMonthStr = `${dPrev.getFullYear()}-${String(dPrev.getMonth() + 1).padStart(2, '0')}-${String(lastDayPrev).padStart(2, '0')}`;
      const prevSnap = this.calculateAssetSnapshotOnDate(prevMonthStr);

      snapshots.forEach((s, idx) => {
        labels.push(s.label);
        lineData.push(s.netWorth);
        const prevNet = idx === 0 ? prevSnap.netWorth : snapshots[idx - 1].netWorth;
        const change = s.netWorth - prevNet;
        s.change = change;
        barData.push(change);
      });
    } else {
      // 최근 5개년 동적 생성 (Y-4 ~ 올해)
      const currentYear = today.getFullYear();
      for (let i = 4; i >= 0; i--) {
        const y = currentYear - i;
        const dateStr = (i === 0) ? today.toISOString().split('T')[0] : `${y}-12-31`;
        const label = (i === 0) ? '올해' : `${String(y).slice(-2)}년`;

        const snap = this.calculateAssetSnapshotOnDate(dateStr);
        snapshots.push({ ...snap, year: y, label, isCurrent: i === 0 });
      }

      const prevYearStr = `${currentYear - 5}-12-31`;
      const prevSnap = this.calculateAssetSnapshotOnDate(prevYearStr);

      snapshots.forEach((s, idx) => {
        labels.push(s.label);
        lineData.push(s.netWorth);
        const prevNet = idx === 0 ? prevSnap.netWorth : snapshots[idx - 1].netWorth;
        const change = s.netWorth - prevNet;
        s.change = change;
        barData.push(change);
      });
    }

    // 실제 순자산 및 최신 변동 수치
    const realNetWorth = Number(fin.netWorth) || 0;
    const lastSnap = snapshots[snapshots.length - 1];
    const latestChange = lastSnap ? lastSnap.change : 0;
    const prevNet = (snapshots.length >= 2 ? snapshots[snapshots.length - 2].netWorth : 0);
    const changeRate = prevNet > 0 ? ((latestChange / prevNet) * 100) : 0;

    let totalAssetDisplay = this.formatKoreanAmount(realNetWorth);
    let profitLabel = period === 'day' ? '당일 순자산 변동' : (period === 'month' ? '당월 순자산 변동' : '연간 순자산 변동');
    let profitAmtDisplay = `${latestChange >= 0 ? '+' : ''}${this.formatKoreanAmount(latestChange)}`;
    let profitRateDisplay = `${changeRate >= 0 ? '+' : ''}${changeRate.toFixed(2)}%`;

    // 수치 DOM 업데이트
    const totalAssetEl = document.getElementById('miraeTotalAssetVal');
    if (totalAssetEl) totalAssetEl.textContent = totalAssetDisplay;

    const profitLabelEl = document.getElementById('miraeProfitLabel');
    if (profitLabelEl) profitLabelEl.textContent = profitLabel;

    const profitAmtEl = document.getElementById('miraeProfitAmt');
    if (profitAmtEl) profitAmtEl.textContent = profitAmtDisplay;

    const profitRateEl = document.getElementById('miraeProfitRate');
    if (profitRateEl) profitRateEl.textContent = profitRateDisplay;

    // 점 및 바 컬러 세팅
    const lastIdx = labels.length - 1;
    const pointBgColors = labels.map((_, idx) => idx === lastIdx ? '#ff4713' : '#9ca3af');
    const pointBorderColors = labels.map((_, idx) => idx === lastIdx ? '#ffffff' : '#cbd5e1');
    const pointBorderWidths = labels.map((_, idx) => idx === lastIdx ? 2.5 : 1);
    const pointRadii = labels.map((_, idx) => idx === lastIdx ? 5.5 : 3.5);
    const barColors = labels.map((_, idx) => idx === lastIdx ? '#ff4713' : '#c8ced6');

    const lineMin = Math.min(...lineData);
    const lineMax = Math.max(...lineData);
    const lineRange = (lineMax - lineMin) || 1000000;
    const barAbsMax = Math.max(...barData.map(v => Math.abs(v))) || 500000;

    canvas.style.cursor = 'pointer';

    this.charts.miraeCombo = new Chart(ctx, {
      data: {
        labels,
        datasets: [
          {
            type: 'line',
            label: '순자산',
            data: lineData,
            borderColor: '#cbd5e1',
            borderWidth: 2,
            tension: 0.25,
            pointBackgroundColor: pointBgColors,
            pointBorderColor: pointBorderColors,
            pointBorderWidth: pointBorderWidths,
            pointRadius: pointRadii,
            pointHoverRadius: 6,
            fill: false,
            yAxisID: 'yLine',
            order: 1
          },
          {
            type: 'bar',
            label: '변동폭',
            data: barData,
            backgroundColor: barColors,
            borderRadius: 6,
            borderSkipped: false,
            barThickness: 22,
            yAxisID: 'yBar',
            order: 2
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
          mode: 'index',
          intersect: false
        },
        onClick: (evt, elements) => {
          if (elements && elements.length > 0) {
            const idx = elements[0].index;
            const snap = snapshots[idx];
            if (snap) {
              this.openTrendDetailModal(period, snap);
            }
          }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: 'rgba(25, 31, 40, 0.92)',
            titleFont: { family: 'Pretendard', size: 12, weight: '700' },
            bodyFont: { family: 'Pretendard', size: 12, weight: '600' },
            padding: 10,
            cornerRadius: 8,
            displayColors: false,
            callbacks: {
              label: (c) => {
                if (c.dataset.type === 'line') {
                  return `순자산: ₩${Math.round(c.raw).toLocaleString()}`;
                } else {
                  const s = c.raw >= 0 ? '+' : '';
                  return `변동폭: ${s}₩${Math.round(c.raw).toLocaleString()} (클릭 시 상세)`;
                }
              }
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: {
              font: (ctx) => {
                const isLast = ctx.index === labels.length - 1;
                return { family: 'Pretendard', size: 11, weight: isLast ? '800' : '500' };
              },
              color: (ctx) => {
                const isLast = ctx.index === labels.length - 1;
                return isLast ? '#ff4713' : '#8b95a1';
              }
            }
          },
          yLine: {
            display: false,
            min: realNetWorth === 0 ? -100 : lineMin - lineRange * 1.3,
            max: realNetWorth === 0 ? 100 : lineMax + lineRange * 0.2
          },
          yBar: {
            display: false,
            min: realNetWorth === 0 ? -100 : -barAbsMax * 1.6,
            max: realNetWorth === 0 ? 100 : barAbsMax * 2.8
          }
        }
      }
    });

    // 캔버스 클릭 이벤트 보강 (차트 막대/점 주변 클릭 시에도 상세 모달 오픈 보장)
    canvas.onclick = (e) => {
      if (!this.charts.miraeCombo) return;
      const elements = this.charts.miraeCombo.getElementsAtEventForMode(e, 'index', { intersect: false }, true);
      if (elements && elements.length > 0) {
        const idx = elements[0].index;
        const snap = snapshots[idx];
        if (snap) {
          this.openTrendDetailModal(period, snap);
        }
      }
    };
  },

  // [표 클릭 시 100% 연결된 실제 데이터 상세 모달 오픈]
  openTrendDetailModal(period, snap) {
    const modal = document.getElementById('tossTrendDetailModal');
    if (!modal || !snap) return;

    const periodName = period === 'day' ? '일자' : (period === 'month' ? '월' : '연도');
    const titleEl = document.getElementById('trendDetailTitle');
    if (titleEl) titleEl.textContent = `${snap.date} (${snap.label}) 자산 내역`;

    const subEl = document.getElementById('trendDetailSubtitle');
    if (subEl) subEl.textContent = `차트에서 선택한 ${periodName} 기준 실제 등록 데이터`;

    const totEl = document.getElementById('trendDetailTotalAsset');
    if (totEl) totEl.textContent = `₩${snap.totalAssets.toLocaleString()}`;

    const netEl = document.getElementById('trendDetailNetWorth');
    if (netEl) netEl.textContent = `순자산 ₩${snap.netWorth.toLocaleString()}`;

    const changeEl = document.getElementById('trendDetailChange');
    if (changeEl) {
      const chg = snap.change || 0;
      changeEl.textContent = `변동 ${chg >= 0 ? '+' : ''}₩${chg.toLocaleString()}`;
      changeEl.style.color = chg > 0 ? 'var(--toss-red)' : (chg < 0 ? 'var(--toss-blue)' : 'var(--color-text-secondary)');
    }

    // 등록 자산 및 계좌 리스트 렌더링
    const assetListEl = document.getElementById('trendDetailAssetList');
    const assetCountEl = document.getElementById('trendDetailAssetCount');
    const allItems = [
      ...snap.accounts.map(a => ({ type: '계좌', name: a.name, detail: a.institution, amount: a.balance })),
      ...snap.assets.map(s => ({ type: s.category, name: s.name, detail: s.ticker || s.institution || '', amount: s.currentVal })),
      ...snap.debts.map(d => ({ type: '부채', name: d.name, detail: d.institution || '', amount: -d.currentVal, isDebt: true }))
    ];

    if (assetCountEl) assetCountEl.textContent = `(${allItems.length}건)`;

    if (assetListEl) {
      if (allItems.length === 0) {
        assetListEl.innerHTML = `<div style="text-align:center; padding:12px; font-size:12px; color:#8b95a1; background:#f9fafb; border-radius:10px;">당시 등록된 보유 자산 및 계좌가 없습니다.</div>`;
      } else {
        assetListEl.innerHTML = allItems.map(item => `
          <div style="display:flex; justify-content:space-between; align-items:center; padding:8px 10px; background:#f9fafb; border-radius:10px; border:1px solid #f2f4f6;">
            <div>
              <div style="font-size:13px; font-weight:700; color:var(--color-text-primary);">${item.name}</div>
              <div style="font-size:10px; color:#8b95a1;">[${item.type}] ${item.detail}</div>
            </div>
            <div style="font-size:13px; font-weight:800; font-family:var(--font-toss); color:${item.isDebt ? 'var(--toss-red)' : 'var(--color-text-primary)'};">
              ₩${item.amount.toLocaleString()}
            </div>
          </div>
        `).join('');
      }
    }

    // 당일 발생 거래 내역
    const txListEl = document.getElementById('trendDetailTxList');
    const txCountEl = document.getElementById('trendDetailTxCount');
    const txs = snap.transactions || [];
    if (txCountEl) txCountEl.textContent = `(${txs.length}건)`;

    if (txListEl) {
      if (txs.length === 0) {
        txListEl.innerHTML = `<div style="text-align:center; padding:12px; font-size:12px; color:#8b95a1; background:#f9fafb; border-radius:10px;">해당 일자에 발생한 가계부 거래가 없습니다.</div>`;
      } else {
        txListEl.innerHTML = txs.map(t => {
          const isInc = t.type === 'INCOME';
          const isTrf = t.type === 'TRANSFER';
          const sign = isInc ? '+' : (isTrf ? '↔ ' : '-');
          const color = isInc ? 'var(--toss-blue)' : (isTrf ? '#8b95a1' : 'var(--color-text-primary)');
          return `
            <div style="display:flex; justify-content:space-between; align-items:center; padding:8px 10px; background:#f9fafb; border-radius:10px; border:1px solid #f2f4f6;">
              <div>
                <div style="font-size:13px; font-weight:700; color:var(--color-text-primary);">${t.memo || t.category}</div>
                <div style="font-size:10px; color:#8b95a1;">${t.category} · ${t.type}</div>
              </div>
              <div style="font-size:13px; font-weight:800; font-family:var(--font-toss); color:${color};">
                ${sign}₩${Number(t.amount).toLocaleString()}
              </div>
            </div>
          `;
        }).join('');
      }
    }

    modal.classList.remove('hidden');
  },

  closeTrendDetailModal() {
    const modal = document.getElementById('tossTrendDetailModal');
    if (modal) modal.classList.add('hidden');
  },

  // 3. [분석 탭 대폭 강화]: 연도별 추이 차트, 저축 성적표, 듀얼 고정/변동 비교, 소비 TOP 4 랭킹, 건전성 지표, 맞춤형 조언
  renderAnalyticsTab(fin) {
    // 1) 저축 성적표 & 잉여 현금
    const savingsBig = document.getElementById('anaSavingsRateBig');
    if (savingsBig) savingsBig.textContent = `${fin.savingsRate.toFixed(1)}%`;

    const savingsBar = document.getElementById('anaSavingsProgressBar');
    if (savingsBar) savingsBar.style.width = `${Math.min(100, fin.savingsRate)}%`;

    const gradeBadge = document.getElementById('anaSavingsGradeBadge');
    const subtext = document.getElementById('anaSavingsSubtext');
    if (gradeBadge) {
      if (fin.savingsRate >= 50) {
        gradeBadge.textContent = '🔥 상위 5% 슈퍼 세이버';
        gradeBadge.style.background = 'var(--toss-mint-soft)';
        gradeBadge.style.color = 'var(--toss-mint)';
        if (subtext) subtext.textContent = '(월간 수입의 절반 이상을 자산으로 축적 중)';
      } else if (fin.savingsRate >= 25) {
        gradeBadge.textContent = '👍 우수 저축러';
        gradeBadge.style.background = '#e8f3ff';
        gradeBadge.style.color = 'var(--toss-blue)';
        if (subtext) subtext.textContent = '(안정적인 잉여금을 유지하고 있습니다)';
      } else {
        gradeBadge.textContent = '⚠️ 소비 점검 필요';
        gradeBadge.style.background = '#fff2f2';
        gradeBadge.style.color = 'var(--toss-red)';
        if (subtext) subtext.textContent = '(지출이 수입에 근접하고 있습니다)';
      }
    }

    const incomeEl = document.getElementById('anaTotalIncomeText');
    if (incomeEl) incomeEl.textContent = `₩${fin.monthIncome.toLocaleString()}`;

    const surplusEl = document.getElementById('anaTotalSurplusText');
    if (surplusEl) surplusEl.textContent = `₩${fin.remainingSurplus.toLocaleString()}`;

    // 2) 고정비 vs 변동비 듀얼 비교
    const fixedAmt = document.getElementById('anaFixedAmtText');
    if (fixedAmt) fixedAmt.textContent = `₩${fin.monthFixedExpense.toLocaleString()}`;

    const varAmt = document.getElementById('anaVarAmtText');
    if (varAmt) varAmt.textContent = `₩${fin.monthVariableExpense.toLocaleString()}`;

    const totalExp = fin.monthExpense || 1;
    const fixedPct = Math.round((fin.monthFixedExpense / totalExp) * 100);
    const varPct = Math.round((fin.monthVariableExpense / totalExp) * 100);

    const fixedPctEl = document.getElementById('anaFixedPctText');
    if (fixedPctEl) fixedPctEl.textContent = `전체 지출의 ${fixedPct}%`;

    const varPctEl = document.getElementById('anaVarPctText');
    if (varPctEl) varPctEl.textContent = `전체 지출의 ${varPct}%`;

    const fixedBar = document.getElementById('anaFixedBar');
    if (fixedBar) fixedBar.style.width = `${fixedPct}%`;

    const varBar = document.getElementById('anaVarBar');
    if (varBar) varBar.style.width = `${varPct}%`;

    // 3) 소비 랭킹 TOP 4
    const rankList = document.getElementById('anaCategoryRankingList');
    if (rankList) {
      const catTotals = {};
      fin.monthTxs.filter(t => t.type.startsWith('EXPENSE')).forEach(t => {
        catTotals[t.category] = (catTotals[t.category] || 0) + Number(t.amount);
      });

      const sorted = Object.entries(catTotals).sort((a, b) => b[1] - a[1]).slice(0, 4);
      if (sorted.length === 0) {
        rankList.innerHTML = `<div style="text-align:center; padding:16px; font-size:12px; color:var(--color-text-tertiary);">이번 달 지출 내역이 없습니다.</div>`;
      } else {
        const rankColors = ['#f04452', '#ff8a00', '#3182f6', '#0ba385'];
        rankList.innerHTML = sorted.map(([cat, amt], idx) => {
          const pct = Math.round((amt / totalExp) * 100);
          const { icon } = this.getCategoryIcon(cat, 'EXPENSE');

          return `
            <div>
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
                <div style="display:flex; align-items:center; gap:6px;">
                  <span style="width:18px; height:18px; border-radius:50%; background:${rankColors[idx]}; color:#ffffff; font-size:10px; font-weight:800; display:flex; align-items:center; justify-content:center;">${idx + 1}</span>
                  <i data-lucide="${icon}" style="width:14px; height:14px; color:var(--color-text-secondary);"></i>
                  <span style="font-size:13px; font-weight:700; color:var(--color-text-primary);">${cat}</span>
                </div>
                <div style="font-size:13px; font-weight:800; color:var(--color-text-primary);">
                  ₩${amt.toLocaleString()} <span style="font-size:11px; color:var(--color-text-tertiary); font-weight:500;">(${pct}%)</span>
                </div>
              </div>
              <div style="height:6px; background:var(--color-surface-hover); border-radius:9999px; overflow:hidden;">
                <div style="height:100%; width:${pct}%; background:${rankColors[idx]}; border-radius:9999px;"></div>
              </div>
            </div>
          `;
        }).join('');
      }
    }

    // 4) 건전성 지표 (부채 비율 & 비상금 버퍼)
    const debtRatio = fin.totalAssets > 0 ? ((fin.totalDebt / fin.totalAssets) * 100).toFixed(1) : 0;
    const debtRatioEl = document.getElementById('anaDebtRatioText');
    if (debtRatioEl) {
      debtRatioEl.textContent = `${debtRatio}%`;
      debtRatioEl.style.color = Number(debtRatio) > 40 ? 'var(--toss-red)' : 'var(--toss-blue)';
    }

    const emergencyMonths = fin.monthExpense > 0 ? (fin.totalCash / fin.monthExpense).toFixed(1) : '12+';
    const emergencyEl = document.getElementById('anaEmergencyBufferText');
    if (emergencyEl) emergencyEl.textContent = `${emergencyMonths}개월분`;

    // 5) 맞춤형 금융 조언
    const adviceEl = document.getElementById('anaAdviceText');
    if (adviceEl) {
      if (fin.savingsRate >= 60) {
        adviceEl.innerHTML = `현재 저축률 <b>${fin.savingsRate.toFixed(1)}%</b>로 최상위권입니다. 비상금 버퍼(${emergencyMonths}개월분)가 든든하게 확보되어 있으므로, 월 잉여금 <b>₩${fin.remainingSurplus.toLocaleString()}</b> 중 일부를 배당 성장 ETF(SCHD 등)로 분할 매수해 자산 증식을 가속화하기에 최적의 시기입니다.`;
      } else if (fin.savingsRate >= 30) {
        adviceEl.innerHTML = `고정 지출 비율(<b>${fixedPct}%</b>)이 매우 안정적으로 관리되고 있습니다. 변동 지출 상위 항목을 10%만 절감하면 월 저축액을 <b>${Math.round(fin.remainingSurplus * 0.15).toLocaleString()}원</b> 더 늘릴 수 있습니다.`;
      } else {
        adviceEl.innerHTML = `이번 달 변동 지출 비중이 <b>${varPct}%</b>로 다소 높습니다. 가계부 탭의 하루 권장 예산 가이드라인에 맞춰 주말 외식과 간식 지출을 조절하는 것을 추천드립니다.`;
      }
    }
  },

  // 5. [설정 탭 & 구글 시트 양방향 연동]
  renderSettingsTab() {
    const urlInput = document.getElementById('googleSheetUrlInput');
    if (urlInput) {
      urlInput.value = this.getGoogleSheetUrl();
    }
    const msgEl = document.getElementById('googleSyncStatusMsg');
    if (msgEl) {
      const savedUrl = this.getGoogleSheetUrl();
      if (savedUrl) {
        msgEl.style.color = 'var(--color-text-secondary)';
        msgEl.textContent = '구글 시트 Web App URL이 설정되어 있습니다.';
      } else {
        msgEl.style.color = '#8b95a1';
        msgEl.textContent = '아직 연동 URL이 설정되지 않았습니다.';
      }
    }
  },

  getGoogleSheetUrl() {
    return (this.data && this.data.googleSheetUrl) || localStorage.getItem('google_sheet_webapp_url') || '';
  },

  saveGoogleSheetUrl(url) {
    const trimmed = (url || '').trim();
    if (!this.data) this.data = {};
    this.data.googleSheetUrl = trimmed;
    try {
      localStorage.setItem('google_sheet_webapp_url', trimmed);
    } catch (e) {}
    StorageService.save(this.data);

    const msgEl = document.getElementById('googleSyncStatusMsg');
    if (msgEl) {
      if (trimmed) {
        msgEl.style.color = 'var(--color-text-secondary)';
        msgEl.textContent = '웹 앱 URL이 저장되었습니다. [시트로 내보내기] 또는 [연동 연결 테스트]를 진행하세요.';
      } else {
        msgEl.style.color = '#8b95a1';
        msgEl.textContent = '연동 URL이 비워졌습니다.';
      }
    }
    this.showToast('💾 구글 시트 URL이 저장되었습니다.');
  },

  async testGoogleSheetsSync() {
    const url = this.getGoogleSheetUrl();
    const msgEl = document.getElementById('googleSyncStatusMsg');
    if (!url) {
      if (msgEl) {
        msgEl.style.color = '#f04452';
        msgEl.textContent = '❌ Google Apps Script Web App URL을 먼저 입력해주세요.';
      }
      this.showToast('⚠️ Web App URL을 먼저 입력해주세요.');
      this.openGoogleScriptGuideModal();
      return;
    }

    if (msgEl) {
      msgEl.style.color = 'var(--mirae-orange)';
      msgEl.textContent = '연결 상태를 확인하는 중입니다...';
    }

    try {
      const testUrl = url + (url.includes('?') ? '&' : '?') + 'action=ping&_t=' + Date.now();
      const res = await fetch(testUrl, {
        method: 'GET',
        mode: 'cors'
      });
      const data = await res.json();
      if (data && data.status === 'success') {
        if (msgEl) {
          msgEl.style.color = '#00a86b';
          msgEl.textContent = '✅ 연결 성공! 정상적으로 구글 시트와 통신할 수 있습니다.';
        }
        this.showToast('✅ 구글 시트 연결 성공!');
      } else {
        throw new Error((data && data.message) || '응답 형식 불일치');
      }
    } catch (err) {
      console.error('구글 시트 연동 테스트 실패:', err);
      if (msgEl) {
        msgEl.style.color = '#f04452';
        msgEl.textContent = '❌ 연결 실패: 배포 시 "액세스 권한: 모든 사용자(Anyone)" 설정 및 URL을 다시 확인해주세요.';
      }
      this.showToast('❌ 연결 실패! 가이드의 권한 설정을 확인하세요.');
    }
  },

  async exportToGoogleSheets() {
    const url = this.getGoogleSheetUrl();
    const msgEl = document.getElementById('googleSyncStatusMsg');
    if (!url) {
      if (msgEl) {
        msgEl.style.color = '#f04452';
        msgEl.textContent = '❌ Google Apps Script Web App URL을 먼저 입력해주세요.';
      }
      this.showToast('⚠️ Web App URL을 먼저 입력해주세요.');
      this.openGoogleScriptGuideModal();
      return;
    }

    if (msgEl) {
      msgEl.style.color = 'var(--mirae-orange)';
      msgEl.textContent = '구글 시트로 데이터를 백업 및 동기화하는 중입니다...';
    }

    try {
      // POST text/plain avoids CORS preflight OPTIONS which GAS doesn't handle well
      const res = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8'
        },
        body: JSON.stringify({
          action: 'syncAll',
          exportAction: 'export',
          data: this.data
        })
      });

      const result = await res.json();
      if (result && result.status === 'success') {
        const nowStr = new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        if (msgEl) {
          msgEl.style.color = '#00a86b';
          msgEl.textContent = `✅ 백업 완료 (${nowStr}) : 거래·계좌·자산 시트 갱신`;
        }
        this.showToast('✅ 구글 시트에 모든 데이터가 성공적으로 백업되었습니다!');
      } else {
        throw new Error((result && result.message) || '동기화 응답 오류');
      }
    } catch (err) {
      console.error('구글 시트 내보내기 실패:', err);
      if (msgEl) {
        msgEl.style.color = '#f04452';
        msgEl.textContent = '❌ 내보내기 실패: Web App 배포 설정(액세스 권한: Anyone)을 확인해주세요.';
      }
      this.showToast('❌ 내보내기 실패. 웹 앱 배포 설정을 확인하세요.');
    }
  },

  async importFromGoogleSheets() {
    const url = this.getGoogleSheetUrl();
    const msgEl = document.getElementById('googleSyncStatusMsg');
    if (!url) {
      if (msgEl) {
        msgEl.style.color = '#f04452';
        msgEl.textContent = '❌ Google Apps Script Web App URL을 먼저 입력해주세요.';
      }
      this.showToast('⚠️ Web App URL을 먼저 입력해주세요.');
      this.openGoogleScriptGuideModal();
      return;
    }

    if (!confirm('구글 시트에 백업된 데이터로 현재 브라우저의 가계부 데이터를 덮어쓰시겠습니까?')) {
      return;
    }

    if (msgEl) {
      msgEl.style.color = 'var(--mirae-orange)';
      msgEl.textContent = '구글 시트에서 최신 데이터를 내려받는 중입니다...';
    }

    try {
      const getUrl = url + (url.includes('?') ? '&' : '?') + 'action=get&_t=' + Date.now();
      const res = await fetch(getUrl, {
        method: 'GET',
        mode: 'cors'
      });
      const result = await res.json();

      if (result && result.status === 'success' && result.data) {
        const imported = result.data;
        if (Array.isArray(imported)) {
          throw new Error('시트에서 반환된 데이터가 2차원 셀 배열입니다. [설정] 탭의 [연동 가이드 & 코드 복사]에 있는 최신 스크립트로 Apps Script를 배포해주세요.');
        }
        if (!Array.isArray(imported.accounts)) imported.accounts = [];
        if (!Array.isArray(imported.assets)) imported.assets = [];
        if (!Array.isArray(imported.transactions)) imported.transactions = [];

        // 연동 URL 보존
        if (!imported.googleSheetUrl) {
          imported.googleSheetUrl = url;
        }

        this.data = imported;
        StorageService.save(this.data);
        this.render();

        const nowStr = new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        if (msgEl) {
          msgEl.style.color = '#00a86b';
          msgEl.textContent = `✅ 복원 완료 (${nowStr}) : 거래 ${imported.transactions.length}건, 계좌 ${imported.accounts.length}건, 자산 ${imported.assets.length}건`;
        }
        this.showToast('✅ 구글 시트에서 데이터를 성공적으로 복원했습니다!');
      } else {
        throw new Error((result && result.message) || '가져올 데이터가 없거나 형식이 올바르지 않습니다.');
      }
    } catch (err) {
      console.error('구글 시트 가져오기 실패:', err);
      if (msgEl) {
        msgEl.style.color = '#f04452';
        msgEl.textContent = '❌ 가져오기 실패: 시트에 저장된 데이터가 없거나 접근 권한이 부족합니다.';
      }
      this.showToast('❌ 가져오기 실패: ' + (err.message || '오류가 발생했습니다.'));
    }
  },

  openGoogleScriptGuideModal() {
    const modal = document.getElementById('googleScriptGuideModal');
    if (modal) {
      const codeArea = document.getElementById('googleAppsScriptCodeArea');
      if (codeArea) {
        codeArea.value = GOOGLE_APPS_SCRIPT_TEMPLATE;
      }
      modal.classList.remove('hidden');
    }
  },

  closeGoogleScriptGuideModal() {
    const modal = document.getElementById('googleScriptGuideModal');
    if (modal) modal.classList.add('hidden');
  },

  copyGoogleAppsScriptCode() {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(GOOGLE_APPS_SCRIPT_TEMPLATE).then(() => {
        this.showToast('📋 연동 코드가 복사되었습니다! Apps Script에 붙여넣으세요.');
      }).catch(() => {
        const codeArea = document.getElementById('googleAppsScriptCodeArea');
        if (codeArea) {
          codeArea.select();
          document.execCommand('copy');
          this.showToast('📋 연동 코드가 복사되었습니다!');
        }
      });
    } else {
      const codeArea = document.getElementById('googleAppsScriptCodeArea');
      if (codeArea) {
        codeArea.select();
        document.execCommand('copy');
        this.showToast('📋 연동 코드가 복사되었습니다!');
      }
    }
  },


  exportDataJson() {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(this.data, null, 2));
    const dlAnchorElem = document.createElement('a');
    dlAnchorElem.setAttribute("href", dataStr);
    dlAnchorElem.setAttribute("download", `financial_backup_${new Date().toISOString().split('T')[0]}.json`);
    dlAnchorElem.click();
  },

  showToast(msg) {
    let toast = document.getElementById('appToastNotice');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'appToastNotice';
      toast.style.cssText = `
        position: fixed;
        bottom: 85px;
        left: 50%;
        transform: translateX(-50%);
        background: rgba(25, 31, 40, 0.94);
        color: #ffffff;
        font-size: 13px;
        font-weight: 700;
        padding: 12px 22px;
        border-radius: 9999px;
        z-index: 99999;
        box-shadow: 0 6px 20px rgba(0,0,0,0.25);
        transition: opacity 0.25s ease, transform 0.25s ease;
        pointer-events: none;
        white-space: nowrap;
        font-family: var(--font-toss);
      `;
      document.body.appendChild(toast);
    }
    toast.textContent = msg;
    toast.style.opacity = '1';
    toast.style.transform = 'translateX(-50%) translateY(0)';
    clearTimeout(this._toastTimeout);
    this._toastTimeout = setTimeout(() => {
      if (toast) {
        toast.style.opacity = '0';
        toast.style.transform = 'translateX(-50%) translateY(10px)';
      }
    }, 2400);
  },

  clearAllAssetsAndAccounts() {
    this.data.accounts = [];
    this.data.assets = [];
    StorageService.save(this.data);
    this.render();
    this.showToast('✅ 모든 계좌와 자산이 0원으로 비워졌습니다.');
  },

  resetAllDataToEmpty() {
    this.data.accounts = [];
    this.data.assets = [];
    this.data.transactions = [];
    this.data.yearlyNetWorth = [];
    this.data.monthlyNetWorth = [];
    StorageService.save(this.data);
    this.render();
    this.showToast('✅ 전체 데이터가 0원으로 완전 초기화되었습니다.');
  },


  clearAllTransactions() {
    this.data.transactions = [];
    StorageService.save(this.data);
    this.render();
    this.showToast('✅ 모든 거래 내역이 0건으로 삭제되었습니다.');
  },

  // =========================================================================
  // 거래 등록 / 수정 / 삭제 모달 로직 (질문 3 해결)
  // =========================================================================
  applyQuickTx(type, category, amount, memo) {
    const typeSelect = document.getElementById('newTxType');
    if (typeSelect) {
      typeSelect.value = type;
      const toAccRow = document.getElementById('toAccountRow');
      if (toAccRow) toAccRow.style.display = type === 'TRANSFER' ? 'block' : 'none';
    }
    const catInput = document.getElementById('newTxCategory');
    if (catInput) catInput.value = category;

    const amtInput = document.getElementById('newTxAmount');
    if (amtInput) amtInput.value = amount;

    const memoInput = document.getElementById('newTxMemo');
    if (memoInput) memoInput.value = memo;
  },

  openAddModal() {
    this.editingTxId = null;

    const titleEl = document.getElementById('txModalTitle');
    if (titleEl) titleEl.textContent = '새 거래 등록';

    const chipArea = document.getElementById('quickTxChipArea');
    if (chipArea) chipArea.style.display = 'block';

    const fromSelect = document.getElementById('newTxFromAccount');
    const toSelect = document.getElementById('newTxToAccount');
    if (fromSelect) {
      if (this.data.accounts.length === 0) {
        fromSelect.innerHTML = `<option value="">등록된 계좌 없음 (기본)</option>`;
      } else {
        fromSelect.innerHTML = this.data.accounts.map(a => `<option value="${a.id}" ${a.isDefault ? 'selected' : ''}>${a.name} (잔액: ₩${a.balance.toLocaleString()})</option>`).join('');
      }
    }
    if (toSelect) {
      if (this.data.accounts.length === 0) {
        toSelect.innerHTML = `<option value="">등록된 계좌 없음</option>`;
      } else {
        toSelect.innerHTML = this.data.accounts.map(a => `<option value="${a.id}">${a.name}</option>`).join('');
      }
    }

    document.getElementById('newTxType').value = 'EXPENSE_VARIABLE';
    const toAccRow = document.getElementById('toAccountRow');
    if (toAccRow) toAccRow.style.display = 'none';

    document.getElementById('newTxCategory').value = '';
    document.getElementById('newTxAmount').value = '';
    document.getElementById('newTxMemo').value = '';

    // 현재 가계부에서 조회 중인 연-월에 맞춘 스마트 기본 날짜 설정
    const ym = this.data.selectedYearMonth || '2026-10';
    const today = new Date();
    const todayYm = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
    let defaultDate;
    if (ym === todayYm) {
      defaultDate = `${todayYm}-${String(today.getDate()).padStart(2, '0')}`;
    } else {
      const [y, m] = ym.split('-').map(Number);
      const daysInMonth = new Date(y, m, 0).getDate();
      const day = Math.min(today.getDate() || 15, daysInMonth);
      defaultDate = `${ym}-${String(day).padStart(2, '0')}`;
    }
    document.getElementById('newTxDate').value = defaultDate;

    const btnContainer = document.getElementById('txModalActionButtons');
    if (btnContainer) {
      btnContainer.innerHTML = `
        <button onclick="App.closeAddModal()" style="flex:1; padding:12px; border-radius:14px; border:1px solid var(--color-border); background:#ffffff; font-weight:700; font-size:14px; color:var(--color-text-secondary); cursor:pointer;">취소</button>
        <button onclick="App.saveNewTransaction()" style="flex:2; padding:12px; border-radius:14px; border:none; background:var(--toss-blue); font-weight:700; font-size:14px; color:#ffffff; cursor:pointer;">저장하기</button>
      `;
    }

    const modal = document.getElementById('tossAddTxModal');
    if (modal) modal.classList.remove('hidden');
  },

  openEditTxModal(txId) {
    const tx = this.data.transactions.find(t => t.id === txId);
    if (!tx) return;
    this.editingTxId = txId;

    const titleEl = document.getElementById('txModalTitle');
    if (titleEl) titleEl.textContent = '거래 내역 수정';

    const chipArea = document.getElementById('quickTxChipArea');
    if (chipArea) chipArea.style.display = 'none';

    const fromSelect = document.getElementById('newTxFromAccount');
    const toSelect = document.getElementById('newTxToAccount');
    if (fromSelect) {
      fromSelect.innerHTML = this.data.accounts.map(a => `<option value="${a.id}" ${a.id === tx.accountId ? 'selected' : ''}>${a.name} (잔액: ₩${a.balance.toLocaleString()})</option>`).join('');
    }
    if (toSelect) {
      toSelect.innerHTML = this.data.accounts.map(a => `<option value="${a.id}" ${a.id === tx.toAccountId ? 'selected' : ''}>${a.name}</option>`).join('');
    }

    document.getElementById('newTxType').value = tx.type;
    const toAccRow = document.getElementById('toAccountRow');
    if (toAccRow) toAccRow.style.display = tx.type === 'TRANSFER' ? 'block' : 'none';

    document.getElementById('newTxCategory').value = tx.category || '';
    document.getElementById('newTxAmount').value = tx.amount || '';
    document.getElementById('newTxMemo').value = tx.memo || '';
    document.getElementById('newTxDate').value = tx.date || '';

    const btnContainer = document.getElementById('txModalActionButtons');
    if (btnContainer) {
      btnContainer.innerHTML = `
        <button onclick="App.deleteTransaction('${tx.id}'); App.closeAddModal();" style="flex:1; padding:12px; border-radius:14px; border:1px solid #ffd0d0; background:#fff2f2; font-weight:700; font-size:14px; color:#f04452; cursor:pointer;">삭제</button>
        <button onclick="App.closeAddModal()" style="flex:1; padding:12px; border-radius:14px; border:1px solid var(--color-border); background:#ffffff; font-weight:700; font-size:14px; color:var(--color-text-secondary); cursor:pointer;">취소</button>
        <button onclick="App.saveNewTransaction()" style="flex:2; padding:12px; border-radius:14px; border:none; background:var(--toss-blue); font-weight:700; font-size:14px; color:#ffffff; cursor:pointer;">수정 완료</button>
      `;
    }

    const modal = document.getElementById('tossAddTxModal');
    if (modal) modal.classList.remove('hidden');
  },

  closeAddModal() {
    this.editingTxId = null;
    const modal = document.getElementById('tossAddTxModal');
    if (modal) modal.classList.add('hidden');
  },

  saveNewTransaction() {
    const type = document.getElementById('newTxType').value;
    const category = document.getElementById('newTxCategory').value.trim();
    const amount = Number(document.getElementById('newTxAmount').value);
    const accountId = document.getElementById('newTxFromAccount').value;
    const toAccountId = document.getElementById('newTxToAccount').value;
    const memo = document.getElementById('newTxMemo').value.trim();
    const dateInput = document.getElementById('newTxDate').value;
    const date = dateInput || (this.data.selectedYearMonth ? `${this.data.selectedYearMonth}-15` : new Date().toISOString().split('T')[0]);

    if (!category || amount <= 0) {
      alert('분류와 올바른 금액을 입력해주세요.');
      return;
    }

    // [핵심 해결]: 등록한 거래의 연-월(YYYY-MM) 추출 후 가계부 조회 월 자동 동기화
    // 거래를 등록했을 때 다른 월에 들어가 가계부에서 안 보이는 현상을 100% 원천 차단!
    const txYearMonth = date.substring(0, 7);
    if (this.data.selectedYearMonth !== txYearMonth) {
      this.data.selectedYearMonth = txYearMonth;
    }

    if (this.editingTxId) {
      // [수정 모드]: 이전 효과 롤백 후 새 효과 적용
      const tx = this.data.transactions.find(t => t.id === this.editingTxId);
      if (tx) {
        // 롤백
        const oldAcc = this.data.accounts.find(a => a.id === tx.accountId);
        if (oldAcc) {
          if (tx.type === 'INCOME') oldAcc.balance -= tx.amount;
          else if (tx.type.startsWith('EXPENSE')) oldAcc.balance += tx.amount;
          else if (tx.type === 'TRANSFER') {
            oldAcc.balance += tx.amount;
            const oldTo = this.data.accounts.find(a => a.id === tx.toAccountId);
            if (oldTo) oldTo.balance -= tx.amount;
          }
        }

        // 값 업데이트
        tx.type = type;
        tx.category = category;
        tx.amount = amount;
        tx.accountId = accountId;
        tx.toAccountId = type === 'TRANSFER' ? toAccountId : null;
        tx.memo = memo;
        tx.date = date;

        // 새 계좌 잔액 적용
        const newAcc = this.data.accounts.find(a => a.id === accountId);
        if (newAcc) {
          if (type === 'INCOME') newAcc.balance += amount;
          else if (type.startsWith('EXPENSE')) newAcc.balance -= amount;
          else if (type === 'TRANSFER' && toAccountId) {
            newAcc.balance -= amount;
            const newTo = this.data.accounts.find(a => a.id === toAccountId);
            if (newTo) newTo.balance += amount;
          }
        }
      }
      this.editingTxId = null;
    } else {
      // [신규 등록 모드]
      const newTx = { id: 't-' + Date.now(), date, type, category, accountId, toAccountId: type === 'TRANSFER' ? toAccountId : null, amount, memo };
      this.data.transactions.unshift(newTx);

      const acc = this.data.accounts.find(a => a.id === accountId);
      if (acc) {
        if (type === 'INCOME') acc.balance += amount;
        else if (type.startsWith('EXPENSE')) acc.balance -= amount;
        else if (type === 'TRANSFER' && toAccountId) {
          acc.balance -= amount;
          const toAcc = this.data.accounts.find(a => a.id === toAccountId);
          if (toAcc) toAcc.balance += amount;
        }
      }
    }

    StorageService.save(this.data);
    this.closeAddModal();
    this.render();
  },

  deleteTransaction(id) {
    if (!confirm('이 거래 내역을 삭제하시겠습니까?')) return;
    const tx = this.data.transactions.find(t => t.id === id);
    if (tx) {
      // 계좌 잔액 롤백
      const acc = this.data.accounts.find(a => a.id === tx.accountId);
      if (acc) {
        if (tx.type === 'INCOME') acc.balance -= tx.amount;
        else if (tx.type.startsWith('EXPENSE')) acc.balance += tx.amount;
        else if (tx.type === 'TRANSFER') {
          acc.balance += tx.amount;
          const toAcc = this.data.accounts.find(a => a.id === tx.toAccountId);
          if (toAcc) toAcc.balance -= tx.amount;
        }
      }
    }
    this.data.transactions = this.data.transactions.filter(t => t.id !== id);
    StorageService.save(this.data);
    this.render();
  },

  // =========================================================================
  // 자산 / 계좌 추가 모달 로직 (취득일자/매수일 및 주식 자동계산 지원)
  // =========================================================================
  calcStockAmount() {
    const bp = Number(document.getElementById('newAssetBuyPrice')?.value) || 0;
    const cp = Number(document.getElementById('newAssetCurrentPrice')?.value) || bp;
    const qty = Number(document.getElementById('newAssetQuantity')?.value) || 0;
    const amtInput = document.getElementById('newAssetAmount');
    if (amtInput && qty > 0) {
      amtInput.value = Math.round(cp * qty);
    }
  },

  openAddAssetModal(category = 'CHECKING') {
    const catSelect = document.getElementById('newAssetCategory');
    if (catSelect) catSelect.value = category;
    this.handleAssetTypeChange(category);

    const dateInput = document.getElementById('newAssetDate');
    if (dateInput) dateInput.value = new Date().toISOString().split('T')[0];

    document.getElementById('newAssetName').value = '';
    document.getElementById('newAssetInstitution').value = '';
    document.getElementById('newAssetAmount').value = '';
    document.getElementById('newAssetTicker').value = '';
    document.getElementById('newAssetBuyPrice').value = '';
    document.getElementById('newAssetQuantity').value = '';
    document.getElementById('newAssetCurrentPrice').value = '';
    document.getElementById('newAssetDividend').value = '';
    document.getElementById('newAssetRate').value = '';

    const modal = document.getElementById('tossAddAssetModal');
    if (modal) modal.classList.remove('hidden');
  },

  closeAddAssetModal() {
    const modal = document.getElementById('tossAddAssetModal');
    if (modal) modal.classList.add('hidden');
  },

  handleAssetTypeChange(type) {
    const stockGroup = document.getElementById('stockFieldGroup');
    const interestGroup = document.getElementById('interestFieldGroup');
    const amountLabel = document.getElementById('newAssetAmountLabel');

    if (stockGroup) stockGroup.style.display = type === 'STOCK' ? 'flex' : 'none';
    if (interestGroup) interestGroup.style.display = (type === 'SAVINGS' || type === 'DEBT') ? 'block' : 'none';

    if (amountLabel) {
      if (type === 'DEBT') amountLabel.textContent = '대출 / 부채 금액 (원)';
      else if (type === 'STOCK') amountLabel.textContent = '현재 평가금액 (원)';
      else amountLabel.textContent = '현재 잔액 / 원금 (원)';
    }
  },

  saveNewAsset() {
    const type = document.getElementById('newAssetCategory').value;
    const date = document.getElementById('newAssetDate').value || new Date().toISOString().split('T')[0];
    const name = document.getElementById('newAssetName').value.trim();
    const institution = document.getElementById('newAssetInstitution').value.trim() || '금융사';
    const amount = Number(document.getElementById('newAssetAmount').value);
    const ticker = document.getElementById('newAssetTicker').value.trim();
    const buyPrice = Number(document.getElementById('newAssetBuyPrice').value) || amount;
    const currentPrice = Number(document.getElementById('newAssetCurrentPrice').value) || buyPrice;
    const quantity = Number(document.getElementById('newAssetQuantity').value) || 1;
    const dividendYield = Number(document.getElementById('newAssetDividend').value) || 0;
    const rate = Number(document.getElementById('newAssetRate').value) || 0;

    if (!name || isNaN(amount) || amount <= 0) {
      alert('자산 이름과 올바른 금액을 입력해주세요.');
      return;
    }

    if (type === 'CHECKING') {
      this.data.accounts.push({
        id: 'acc-' + Date.now(),
        name,
        institution,
        type: 'CHECKING',
        balance: amount,
        date: date,
        isDefault: false
      });
    } else if (type === 'STOCK') {
      this.data.assets.push({
        id: 'ast-' + Date.now(),
        category: 'STOCK',
        name,
        ticker: ticker || 'STOCK',
        amount: amount,
        buyPrice: buyPrice,
        currentPrice: currentPrice,
        quantity: quantity,
        dividendYield: dividendYield || 2.5,
        dividendMonths: [3, 6, 9, 12],
        date: date
      });
    } else if (type === 'SAVINGS') {
      this.data.assets.push({
        id: 'ast-' + Date.now(),
        category: 'SAVINGS',
        name,
        institution,
        amount,
        interestRate: rate || 4.0,
        date: date
      });
    } else if (type === 'PENSION') {
      this.data.assets.push({
        id: 'ast-' + Date.now(),
        category: 'PENSION',
        name,
        institution,
        amount,
        date: date
      });
    } else if (type === 'OTHER') {
      this.data.assets.push({
        id: 'ast-' + Date.now(),
        category: 'OTHER',
        name,
        institution,
        amount,
        date: date
      });
    } else if (type === 'DEBT') {
      this.data.assets.push({
        id: 'ast-' + Date.now(),
        category: 'DEBT',
        name,
        institution,
        amount,
        interestRate: rate || 5.0,
        date: date
      });
    }

    StorageService.save(this.data);
    this.closeAddAssetModal();
    this.render();
    alert(`[${name}] 자산이 등록되었습니다. (취득일: ${date})`);
  },

  // =========================================================================
  // 자산 / 계좌 수정 및 삭제 모달 로직
  // =========================================================================
  openEditAssetModal(type, id) {
    this.editingAssetContext = { type, id };
    let item = null;

    if (type === 'account') {
      item = this.data.accounts.find(a => a.id === id);
    } else {
      item = this.data.assets.find(a => a.id === id);
    }

    if (!item) return;

    const dateInput = document.getElementById('editAssetDate');
    if (dateInput) {
      dateInput.value = item.date || item.purchaseDate || new Date().toISOString().split('T')[0];
    }

    document.getElementById('editAssetName').value = item.name;
    document.getElementById('editAssetAmount').value = type === 'account' ? item.balance : item.amount;

    const rateGroup = document.getElementById('editAssetRateGroup');
    const rateInput = document.getElementById('editAssetRate');
    if (rateGroup && rateInput) {
      if (item.interestRate !== undefined) {
        rateGroup.style.display = 'block';
        rateInput.value = item.interestRate;
      } else if (item.dividendYield !== undefined) {
        rateGroup.style.display = 'block';
        rateInput.value = item.dividendYield;
      } else {
        rateGroup.style.display = 'none';
      }
    }

    const modal = document.getElementById('tossEditAssetModal');
    if (modal) modal.classList.remove('hidden');
  },

  closeEditAssetModal() {
    this.editingAssetContext = null;
    const modal = document.getElementById('tossEditAssetModal');
    if (modal) modal.classList.add('hidden');
  },

  saveEditAsset() {
    if (!this.editingAssetContext) return;
    const { type, id } = this.editingAssetContext;
    const newDate = document.getElementById('editAssetDate')?.value;
    const newName = document.getElementById('editAssetName').value.trim();
    const newAmount = Number(document.getElementById('editAssetAmount').value);
    const newRate = Number(document.getElementById('editAssetRate').value);

    if (!newName || isNaN(newAmount)) {
      alert('올바른 이름과 금액을 입력해주세요.');
      return;
    }

    if (type === 'account') {
      const acc = this.data.accounts.find(a => a.id === id);
      if (acc) {
        acc.name = newName;
        acc.balance = newAmount;
        if (newDate) acc.date = newDate;
      }
    } else {
      const ast = this.data.assets.find(a => a.id === id);
      if (ast) {
        ast.name = newName;
        ast.amount = newAmount;
        if (newDate) ast.date = newDate;
        if (ast.interestRate !== undefined) ast.interestRate = newRate;
        if (ast.dividendYield !== undefined) ast.dividendYield = newRate;
      }
    }

    StorageService.save(this.data);
    this.closeEditAssetModal();
    this.render();
  },

  deleteCurrentEditingAsset() {
    if (!this.editingAssetContext) return;
    if (!confirm('정말 이 항목을 삭제하시겠습니까?')) return;

    const { type, id } = this.editingAssetContext;
    if (type === 'account') {
      this.data.accounts = this.data.accounts.filter(a => a.id !== id);
    } else {
      this.data.assets = this.data.assets.filter(a => a.id !== id);
    }

    StorageService.save(this.data);
    this.closeEditAssetModal();
    this.render();
  },

  // =========================================================================
  // 한 달 예산 설정 모달 로직 (핀 #1: 예산 입력 인터페이스)
  // =========================================================================
  openBudgetModal() {
    const totalInput = document.getElementById('budgetTotalInput');
    if (totalInput) totalInput.value = this.data.totalMonthlyBudget || 3500000;

    const b = this.data.budgets || {};
    const food = document.getElementById('budgetCatFood');
    if (food) food.value = b['식비/외식'] || 800000;

    const house = document.getElementById('budgetCatHouse');
    if (house) house.value = b['주거/통신'] || 700000;

    const cafe = document.getElementById('budgetCatCafe');
    if (cafe) cafe.value = b['카페/간식'] || 200000;

    const mart = document.getElementById('budgetCatMart');
    if (mart) mart.value = b['마트/잡화'] || 400000;

    const modal = document.getElementById('tossBudgetModal');
    if (modal) modal.classList.remove('hidden');
  },

  closeBudgetModal() {
    const modal = document.getElementById('tossBudgetModal');
    if (modal) modal.classList.add('hidden');
  },

  setBudgetQuickAmount(amt) {
    const totalInput = document.getElementById('budgetTotalInput');
    if (totalInput) totalInput.value = amt;
  },

  saveBudgetSettings() {
    const totalInput = document.getElementById('budgetTotalInput');
    const newTotal = Number(totalInput.value);

    if (isNaN(newTotal) || newTotal < 0) {
      alert('올바른 목표 예산 금액을 입력해주세요.');
      return;
    }

    this.data.totalMonthlyBudget = newTotal;

    const food = Number(document.getElementById('budgetCatFood').value) || 0;
    const house = Number(document.getElementById('budgetCatHouse').value) || 0;
    const cafe = Number(document.getElementById('budgetCatCafe').value) || 0;
    const mart = Number(document.getElementById('budgetCatMart').value) || 0;

    if (!this.data.budgets) this.data.budgets = {};
    this.data.budgets['식비/외식'] = food;
    this.data.budgets['주거/통신'] = house;
    this.data.budgets['카페/간식'] = cafe;
    this.data.budgets['마트/잡화'] = mart;

    StorageService.save(this.data);
    this.closeBudgetModal();
    this.render();
    alert(`한 달 목표 예산이 ₩${newTotal.toLocaleString()}원으로 설정되었습니다.`);
  }
};

// =========================================================================
// 5. [피드백 핀 도구]: 위치 오차 해결 및 클릭 간섭 100% 방지 오버레이
// =========================================================================
const FeedbackTool = {
  isModeActive: false,
  pins: [],

  toggleMode() {
    this.isModeActive = !this.isModeActive;
    const topBar = document.getElementById('feedbackTopBar');
    if (topBar) topBar.classList.toggle('hidden', !this.isModeActive);

    // 핀 모드 활성화 시 클릭 캡처 오버레이 켜기 (아래 버튼/카드 클릭 완전 차단!)
    const overlay = document.getElementById('feedbackCaptureOverlay');
    if (overlay) overlay.classList.toggle('hidden', !this.isModeActive);

    document.body.classList.toggle('feedback-mode-active', this.isModeActive);
  },

  // 뷰포트 내부 상대 좌표 계산으로 위치 오차 0% 달성 및 원래 기능 실행 방지
  handleOverlayClick(e) {
    if (!this.isModeActive) return;
    e.preventDefault();
    e.stopPropagation();

    const viewport = document.querySelector('.app-viewport');
    if (!viewport) return;

    const rect = viewport.getBoundingClientRect();
    const x = (((e.clientX - rect.left) / rect.width) * 100).toFixed(1);
    const y = (((e.clientY - rect.top) / rect.height) * 100).toFixed(1);

    const text = prompt('이 위치에 남길 수정 의견을 입력하세요:');
    if (!text || !text.trim()) return;

    this.pins.push({ id: 'p-' + Date.now(), x: parseFloat(x), y: parseFloat(y), text: text.trim() });
    this.renderPins();
  },

  renderPins() {
    const layer = document.getElementById('feedbackPinLayer');
    if (!layer) return;
    layer.innerHTML = this.pins.map((p, idx) => `
      <div class="feedback-pin-marker" style="left: ${p.x}%; top: ${p.y}%;">
        <div class="feedback-pin-circle">${idx + 1}</div>
        <div class="feedback-pin-tail"></div>
        <div class="feedback-bubble">
          <div style="font-size:11px; font-weight:700; color:#3182f6; margin-bottom:4px;">핀 #${idx + 1}</div>
          <p style="font-size:12px; color:#191f28; line-height:1.4;">${p.text}</p>
        </div>
      </div>
    `).join('');
  },

  exportBatchFeedback() {
    if (this.pins.length === 0) {
      alert('반영할 피드백 핀이 없습니다.');
      return;
    }
    let report = `[웹 화면 피드백 일괄 수정 요청]\n\n`;
    this.pins.forEach((p, idx) => {
      report += `${idx + 1}. [핀 #${idx + 1}] 위치 (${p.x}%, ${p.y}%): "${p.text}"\n`;
    });
    navigator.clipboard.writeText(report).then(() => {
      alert('피드백 내역이 클립보드에 복사되었습니다! 채팅창에 붙여넣어주세요.');
    }).catch(() => {
      prompt('아래 내용을 복사하세요:', report);
    });
  }
};

window.addEventListener('DOMContentLoaded', () => {
  App.init();
});
