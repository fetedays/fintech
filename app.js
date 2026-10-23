/**
 * 2-Track 자산관리 & 심플 가계부 (완성본 app.js)
 * - 5대 메뉴: 홈(Home), 가계부(Spending), 자산관리(Assets), 분석(Analytics), 설정(Settings)
 * - 재무 계산 정합성 및 0원 예외 처리 완벽 반영
 */

const STORAGE_KEY = 'toss_financial_complete_v9';

const getTodayStr = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

const getCurrentYearMonthStr = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
};

const INITIAL_DATA = {
  updatedAt: 0,
  selectedYearMonth: getCurrentYearMonthStr(),
  dividendTaxDeduction: false,
  usdExchangeRate: 1380,
  accounts: [],
  assets: [],
  transactions: [],
  recurringExpenses: [],
  totalMonthlyBudget: 0,
  budgets: {},
  quickTransactions: [
    { id: 'qt-1', type: 'EXPENSE_VARIABLE', category: '카페/간식', amount: 4500, memo: '스타벅스 커피' },
    { id: 'qt-2', type: 'EXPENSE_VARIABLE', category: '식비/외식', amount: 11000, memo: '점심 식사' },
    { id: 'qt-3', type: 'EXPENSE_VARIABLE', category: '마트/잡화', amount: 3500, memo: '편의점 간식' },
    { id: 'qt-4', type: 'EXPENSE_FIXED', category: '주거/통신', amount: 650000, memo: '월세 이체' }
  ]
};

const StorageService = {
  // 저장소/구글 시트에서 읽은 데이터에 누락된 기본값을 채움
  normalize(parsed) {
    if (!parsed || typeof parsed !== 'object') parsed = {};
    parsed.updatedAt = Number(parsed.updatedAt) || 0;
    if (!parsed.selectedYearMonth || !/^\d{4}-\d{2}$/.test(parsed.selectedYearMonth)) parsed.selectedYearMonth = getCurrentYearMonthStr();
    if (parsed.totalMonthlyBudget === undefined) parsed.totalMonthlyBudget = 0;
    if (!parsed.budgets) parsed.budgets = {};
    if (!Array.isArray(parsed.accounts)) parsed.accounts = [];
    if (!Array.isArray(parsed.assets)) parsed.assets = [];
    if (!Array.isArray(parsed.transactions)) parsed.transactions = [];
    if (!Array.isArray(parsed.recurringExpenses)) parsed.recurringExpenses = [];
    if (!Array.isArray(parsed.quickTransactions)) {
      parsed.quickTransactions = [
        { id: 'qt-1', type: 'EXPENSE_VARIABLE', category: '카페/간식', amount: 4500, memo: '스타벅스 커피' },
        { id: 'qt-2', type: 'EXPENSE_VARIABLE', category: '식비/외식', amount: 11000, memo: '점심 식사' },
        { id: 'qt-3', type: 'EXPENSE_VARIABLE', category: '마트/잡화', amount: 3500, memo: '편의점 간식' },
        { id: 'qt-4', type: 'EXPENSE_FIXED', category: '주거/통신', amount: 650000, memo: '월세 이체' }
      ];
    }
    if (!parsed.usdExchangeRate) parsed.usdExchangeRate = 1380;
    return parsed;
  },
  load() {
    let raw = null;
    try {
      raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        return this.normalize(JSON.parse(raw));
      }
    } catch (e) {
      console.error('데이터 로드 실패, 손상 대비 안전 백업을 생성합니다:', e);
      if (raw) {
        try {
          const backupKey = `toss_financial_corrupt_backup_${Date.now()}`;
          localStorage.setItem(backupKey, raw);
          alert(`⚠️ 기존 저장 데이터 파싱 중 문제가 발생하여 복구용 백업(${backupKey})을 안전하게 보관했습니다.`);
        } catch (saveErr) {}
      }
    }
    const fresh = JSON.parse(JSON.stringify(INITIAL_DATA));
    this.save(fresh);
    return fresh;
  },
  save(data, markUpdated = false) {
    try {
      if (markUpdated) {
        data.updatedAt = Date.now();
      }
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
      if (markUpdated && typeof App !== 'undefined' && App && typeof App.triggerAutoSync === 'function') {
        App.triggerAutoSync();
      }
    } catch (e) {
      console.error('데이터 저장 실패:', e);
      if (typeof App !== 'undefined' && App && typeof App.showToast === 'function') {
        App.showToast('⚠️ 브라우저 저장 공간이 가득 차 데이터 저장에 실패했습니다.');
      }
    }
  }
};

const App = {
  data: null,
  activeTab: 'home',
  assetsTrendPeriod: 'day',
  charts: {},
  editingTxId: null,
  editingRecurringId: null,
  editingAssetContext: null,
  autoSyncEnabled: true,
  _autoSyncTimer: null,
  _isSyncing: false,
  _lastSyncAttempt: 0,
  _toastTimeout: null,

  init() {
    this.initTheme();
    this.data = StorageService.load();
    this.reconcileTransactionsAndBalances();
    const savedAutoSync = localStorage.getItem('google_sheet_autosync');
    this.autoSyncEnabled = savedAutoSync !== null ? (savedAutoSync === 'true') : true;

    if (this.getGoogleSheetUrl()) {
      if (this.autoSyncEnabled) {
        this.syncWithGoogleSheets(true);
      } else {
        this.updateSyncBadge('idle', '동기화 꺼짐');
      }
    } else {
      this.updateSyncBadge('idle', 'URL 미등록');
    }

    this.bindEvents();
    this.render();
    this.updateFabButton('home');
    this.maybeAutoRefreshStocks();
  },

  initTheme() {
    document.documentElement.removeAttribute('data-theme');
    try { localStorage.removeItem('wealth_theme'); } catch (e) {}
  },

  scheduleIcons() {
    if (this._iconScheduled) return;
    this._iconScheduled = true;
    requestAnimationFrame(() => {
      this._iconScheduled = false;
      if (window.lucide && typeof window.lucide.createIcons === 'function') {
        window.lucide.createIcons();
      }
    });
  },

  bindEvents() {
    document.querySelectorAll('.toss-bottom-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const tab = e.currentTarget.getAttribute('data-tab');
        this.switchTab(tab);
      });
    });

    const prevBtn = document.getElementById('monthPrevBtn');
    if (prevBtn) prevBtn.addEventListener('click', () => this.changeMonth(-1));

    const nextBtn = document.getElementById('monthNextBtn');
    if (nextBtn) nextBtn.addEventListener('click', () => this.changeMonth(1));

    const taxToggle = document.getElementById('taxToggleInput');
    if (taxToggle) {
      taxToggle.checked = !!this.data.dividendTaxDeduction;
      taxToggle.addEventListener('change', (e) => {
        this.data.dividendTaxDeduction = e.target.checked;
        StorageService.save(this.data);
        this.render();
      });
    }

    const txTypeSelect = document.getElementById('newTxType');
    if (txTypeSelect) {
      txTypeSelect.addEventListener('change', () => this.onTxTypeChange());
    }

    // 글로벌 ESC 키 바인딩 (활성화된 최상단 모달 닫기)
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' || e.key === 'Esc') {
        const openModals = document.querySelectorAll('.toss-modal-overlay:not(.hidden)');
        if (openModals.length > 0) {
          const topModal = openModals[openModals.length - 1];
          topModal.classList.add('hidden');
        }
      }
    });

    // 다중 기기(스마트폰/태블릿) 동기화: 화면 복귀 시 동기화, 화면 나갈 시 즉시 전송
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'hidden') {
        this.flushAutoSync();
      } else if (document.visibilityState === 'visible') {
        this.syncWithGoogleSheets(true);
      }
    });
    window.addEventListener('pagehide', () => {
      this.flushAutoSync();
    });
    window.addEventListener('focus', () => {
      this.syncWithGoogleSheets(true);
    });
  },

  // 모바일 화면 닫힘/백그라운드 전환 시 대기 중인 동기화 즉시 전송
  flushAutoSync() {
    if (this._autoSyncTimer) {
      clearTimeout(this._autoSyncTimer);
      this._autoSyncTimer = null;
      this.exportToGoogleSheets(false).catch(() => {});
    }
  },

  updateHeaderCopyright(tabName) {
    const navTitleEl = document.getElementById('miraeNavTitle');
    const navSubEl = document.getElementById('headerSubCopy');
    const settingsBtn = document.getElementById('headerSettingsToggleBtn');

    const copyMap = {
      home: { main: '나의 자산 이야기', sub: '모든 지출과 자산을 한눈에' },
      spending: { main: '나의 소비 기록', sub: '이번 달 수입과 지출' },
      assets: { main: '나의 자산 모으기', sub: '늘어나는 나의 자산' },
      analytics: { main: '나의 자산 변화 살펴보기', sub: '소비와 자산의 변화' },
      settings: { main: '나에게 맞게 설정', sub: '설정 변경' }
    };

    const cur = copyMap[tabName] || copyMap.home;
    if (navTitleEl) navTitleEl.textContent = cur.main;
    if (navSubEl) navSubEl.textContent = cur.sub;

    if (settingsBtn) {
      if (tabName === 'settings') {
        settingsBtn.innerHTML = '<i data-lucide="x" style="width:20px; height:20px;"></i>';
        settingsBtn.title = '설정 닫기';
        settingsBtn.setAttribute('aria-label', '설정 닫기');
        settingsBtn.style.background = 'var(--color-surface-hover)';
        settingsBtn.style.color = 'var(--brand-primary)';
      } else {
        settingsBtn.innerHTML = '<i data-lucide="menu" style="width:20px; height:20px;"></i>';
        settingsBtn.title = '설정 메뉴';
        settingsBtn.setAttribute('aria-label', '설정 메뉴');
        settingsBtn.style.background = 'var(--color-surface-subtle)';
        settingsBtn.style.color = 'var(--color-text-primary)';
      }
      this.scheduleIcons();
    }
  },

  switchTab(tabName) {
    if (this.activeTab && this.activeTab !== 'settings') {
      this.previousTab = this.activeTab;
    }
    this.activeTab = tabName;
    document.querySelectorAll('.toss-bottom-btn').forEach(b => {
      b.classList.toggle('active', b.getAttribute('data-tab') === tabName);
    });

    document.querySelectorAll('.view-section').forEach(sec => {
      sec.classList.add('hidden');
    });

    const target = document.getElementById(`view-${tabName}`);
    if (target) target.classList.remove('hidden');

    this.updateHeaderCopyright(tabName);
    this.updateFabButton(tabName);
    this.render();
    window.scrollTo({ top: 0, behavior: 'smooth' });
    if (tabName === 'assets') this.maybeAutoRefreshStocks();
  },

  toggleSettings() {
    if (this.activeTab === 'settings') {
      this.switchTab(this.previousTab || 'home');
    } else {
      this.switchTab('settings');
    }
  },

  goToSpendingWithRecent() {
    if (this.data.transactions && this.data.transactions.length > 0) {
      const sortedTxs = [...this.data.transactions].sort((a, b) => (b.date || '').localeCompare(a.date || ''));
      const latestTx = sortedTxs[0];
      if (latestTx && latestTx.date) {
        this.data.selectedYearMonth = latestTx.date.substring(0, 7);
        StorageService.save(this.data);
      }
    }
    this.switchTab('spending');
  },

  setAssetsTrendPeriod(period) {
    this.assetsTrendPeriod = period;
    const btns = {
      day: document.getElementById('assetsTabDay'),
      month: document.getElementById('assetsTabMonth'),
      year: document.getElementById('assetsTabYear')
    };
    Object.entries(btns).forEach(([k, btn]) => {
      if (btn) btn.classList.toggle('active', k === period);
    });

    const fin = this.getFinancialSummary();
    this.renderAssetsTrendChart(fin);
  },

  updateFabButton(tabName) {
    const fab = document.getElementById('floatingAddFab');
    if (fab) {
      if (tabName === 'spending') {
        fab.style.display = 'flex';
        fab.onclick = () => App.openAddModal();
      } else {
        fab.style.display = 'none';
      }
    }
    this.scheduleIcons();
  },

  getCategoryIcon(category, type) {
    if (type === 'INCOME') return { icon: 'arrow-down-left', tint: 'tint-income' };
    if (type === 'TRANSFER') return { icon: 'arrow-left-right', tint: 'tint-transfer' };
    const cat = category || '';
    if (cat.includes('카페') || cat.includes('간식') || cat.includes('커피') || cat.includes('음료')) return { icon: 'coffee', tint: 'tint-cafe' };
    if (cat.includes('마트') || cat.includes('잡화') || cat.includes('쇼핑') || cat.includes('편의점')) return { icon: 'shopping-bag', tint: 'tint-mart' };
    if (cat.includes('주거') || cat.includes('통신') || cat.includes('월세') || cat.includes('관리비') || cat.includes('공과금')) return { icon: 'home', tint: 'tint-house' };
    if (cat.includes('교통') || cat.includes('주유') || cat.includes('택시') || cat.includes('지하철') || cat.includes('버스') || cat.includes('차량')) return { icon: 'car', tint: 'tint-transport' };
    if (cat.includes('의료') || cat.includes('건강') || cat.includes('병원') || cat.includes('약국')) return { icon: 'heart-pulse', tint: 'tint-medical' };
    if (cat.includes('문화') || cat.includes('구독') || cat.includes('영화') || cat.includes('여가') || cat.includes('도서') || cat.includes('OTT')) return { icon: 'film', tint: 'tint-culture' };
    if (cat.includes('교육') || cat.includes('학원') || cat.includes('강의') || cat.includes('교재')) return { icon: 'graduation-cap', tint: 'tint-edu' };
    if (cat.includes('금융') || cat.includes('보험') || cat.includes('세금')) return { icon: 'shield', tint: 'tint-finance' };
    if (type === 'EXPENSE_FIXED' || cat.includes('고정')) return { icon: 'calendar-clock', tint: 'tint-fixed' };
    if (cat.includes('식비') || cat.includes('외식') || cat.includes('식사') || cat.includes('점심') || cat.includes('저녁') || cat.includes('배달')) return { icon: 'utensils', tint: 'tint-food' };
    return { icon: 'receipt-text', tint: 'tint-default' };
  },

  changeMonth(delta) {
    const [y, m] = this.data.selectedYearMonth.split('-').map(Number);
    const d = new Date(y, m - 1 + delta, 1);
    const ny = d.getFullYear();
    const nm = String(d.getMonth() + 1).padStart(2, '0');
    this.data.selectedYearMonth = `${ny}-${nm}`;
    this.spendingTimelineLimit = 15;
    StorageService.save(this.data);
    this.render();
  },

  getFinancialSummary() {
    let totalCash = 0;
    let totalSavings = 0;
    let totalStock = 0;
    let totalPension = 0;
    let totalOther = 0;
    let totalDebt = 0;

    const usdRate = Number(this.data.usdExchangeRate) || 1380;

    (this.data.accounts || []).forEach(a => {
      const bal = Number(a.balance) || 0;
      if (bal >= 0) totalCash += bal;
      else totalDebt += Math.abs(bal);
    });

    (this.data.assets || []).forEach(a => {
      let amt = Number(a.amount) || 0;
      if (a.category === 'STOCK') {
        const hasQty = a.quantity !== undefined && a.quantity !== null && a.quantity !== '';
        const qty = hasQty ? Number(a.quantity) : 1;
        const hasCurPrice = a.currentPrice !== undefined && a.currentPrice !== null && a.currentPrice !== '';
        const curPrice = hasCurPrice ? Number(a.currentPrice) : (Number(a.buyPrice) || 0);
        const rate = (a.currency === 'USD') ? usdRate : 1;
        if (qty > 0 && curPrice > 0) {
          amt = Math.round(qty * curPrice * rate);
          a.amount = amt;
        }
        totalStock += amt;
      } else if (a.category === 'SAVINGS') {
        totalSavings += amt;
      } else if (a.category === 'PENSION') {
        totalPension += amt;
      } else if (a.category === 'DEBT') {
        totalDebt += amt;
      } else {
        totalOther += amt;
      }
    });

    const totalAssets = totalCash + totalSavings + totalStock + totalPension + totalOther;
    const netWorth = totalAssets - totalDebt;

    const ym = this.data.selectedYearMonth || getCurrentYearMonthStr();
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

    const actualCurrentYM = getCurrentYearMonthStr();
    let actualMonthTxs, actualIncome, actualFixedExp, actualVarExp, actualExpense, actualSurplus, actualSavingsRate;

    if (ym === actualCurrentYM) {
      actualMonthTxs = monthTxs;
      actualIncome = monthIncome;
      actualFixedExp = monthFixedExpense;
      actualVarExp = monthVariableExpense;
      actualExpense = monthExpense;
      actualSurplus = remainingSurplus;
      actualSavingsRate = savingsRate;
    } else {
      actualMonthTxs = (this.data.transactions || []).filter(t => t.date && t.date.startsWith(actualCurrentYM));
      actualIncome = 0;
      actualFixedExp = 0;
      actualVarExp = 0;
      actualMonthTxs.forEach(t => {
        const amt = Number(t.amount) || 0;
        if (t.type === 'INCOME') actualIncome += amt;
        else if (t.type === 'EXPENSE_FIXED') actualFixedExp += amt;
        else if (t.type === 'EXPENSE_VARIABLE') actualVarExp += amt;
      });
      actualExpense = actualFixedExp + actualVarExp;
      actualSurplus = actualIncome - actualExpense;
      actualSavingsRate = actualIncome > 0 ? Math.max(0, (actualSurplus / actualIncome) * 100) : 0;
    }

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
      monthTxs,
      actualCurrentYM,
      actualMonthIncome: actualIncome,
      actualMonthFixedExpense: actualFixedExp,
      actualMonthVariableExpense: actualVarExp,
      actualMonthExpense: actualExpense,
      actualRemainingSurplus: actualSurplus,
      actualSavingsRate,
      actualMonthTxs
    };
  },

  calculateFinancials() {
    return this.getFinancialSummary();
  },

  render() {
    const fin = this.getFinancialSummary();
    this.updateHeaderCopyright(this.activeTab || 'home');

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

    this.scheduleIcons();
  },

  getFinancialBrandBadge(name, institution, defaultIcon = 'landmark') {
    const text = `${institution || ''} ${name || ''}`.toLowerCase();

    if (text.includes('토스') || text.includes('toss')) {
      return `<div class="asset-circle-avatar" style="background:#0064FF; border-color:#0064FF; color:#ffffff; font-weight:900; font-size:12px; letter-spacing:-0.03em;">toss</div>`;
    }
    if (text.includes('케이뱅크') || text.includes('k뱅크') || text.includes('kbank')) {
      return `<div class="asset-circle-avatar" style="background:#001C44; border-color:#001C44; color:#00E5D2; font-weight:900; font-size:16px;">K</div>`;
    }
    if (text.includes('카카오') || text.includes('kakao')) {
      return `<div class="asset-circle-avatar" style="background:#FEE500; border-color:#FEE500; color:#191919; font-weight:900; font-size:14px;">B</div>`;
    }
    if (text.includes('국민') || text.includes('kb')) {
      return `<div class="asset-circle-avatar" style="background:#FFBC00; border-color:#FFBC00; color:#4B4842; font-weight:900; font-size:12px;">KB</div>`;
    }
    if (text.includes('신한') || text.includes('shinhan')) {
      return `<div class="asset-circle-avatar" style="background:#0046FF; border-color:#0046FF; color:#ffffff; font-weight:900; font-size:12px;">신한</div>`;
    }
    if (text.includes('우리') || text.includes('woori')) {
      return `<div class="asset-circle-avatar" style="background:#007BC3; border-color:#007BC3; color:#ffffff; font-weight:900; font-size:12px;">우리</div>`;
    }
    if (text.includes('하나') || text.includes('hana')) {
      return `<div class="asset-circle-avatar" style="background:#008485; border-color:#008485; color:#ffffff; font-weight:900; font-size:12px;">하나</div>`;
    }
    if (text.includes('농협') || text.includes('nh')) {
      return `<div class="asset-circle-avatar" style="background:#00A040; border-color:#00A040; color:#ffffff; font-weight:900; font-size:12px;">NH</div>`;
    }
    if (text.includes('기업') || text.includes('ibk')) {
      return `<div class="asset-circle-avatar" style="background:#1858A8; border-color:#1858A8; color:#ffffff; font-weight:900; font-size:12px;">IBK</div>`;
    }
    if (text.includes('제일') || text.includes('sc')) {
      return `<div class="asset-circle-avatar" style="background:#008752; border-color:#008752; color:#ffffff; font-weight:900; font-size:12px;">SC</div>`;
    }
    if (text.includes('미래에셋') || text.includes('mirae')) {
      return `<div class="asset-circle-avatar" style="background:#F37321; border-color:#F37321; color:#ffffff; font-weight:900; font-size:12px;">미래</div>`;
    }
    if (text.includes('삼성') || text.includes('samsung')) {
      return `<div class="asset-circle-avatar" style="background:#0C4DA2; border-color:#0C4DA2; color:#ffffff; font-weight:900; font-size:12px;">삼성</div>`;
    }
    if (text.includes('한국투자') || text.includes('한투')) {
      return `<div class="asset-circle-avatar" style="background:#003478; border-color:#003478; color:#ffffff; font-weight:900; font-size:12px;">한투</div>`;
    }
    if (text.includes('새마을')) {
      return `<div class="asset-circle-avatar" style="background:#005BAC; border-color:#005BAC; color:#ffffff; font-weight:900; font-size:11px;">새마을</div>`;
    }
    if (text.includes('우체국')) {
      return `<div class="asset-circle-avatar" style="background:#E61E28; border-color:#E61E28; color:#ffffff; font-weight:900; font-size:11px;">우체국</div>`;
    }

    const label = (institution || name || '').trim().substring(0, 2);
    if (label && label.length > 0) {
      return `<div class="asset-circle-avatar" style="background:var(--pill-bg); border-color:var(--color-border); color:var(--color-text-primary); font-weight:800; font-size:12px;">${this.escapeHtml(label)}</div>`;
    }

    return `<div class="asset-circle-avatar"><i data-lucide="${defaultIcon}" style="width:18px; height:18px;"></i></div>`;
  },

  escapeHtml(str) {
    if (str === undefined || str === null) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  },

  getCategoryKoreanLabel(cat) {
    const map = {
      CHECKING: '입출금',
      STOCK: '투자(주식/ETF)',
      INVEST: '투자(주식/ETF)',
      SAVINGS: '예·적금',
      DEPOSIT: '예·적금',
      PENSION: '연금',
      OTHER: '부동산·기타',
      REAL_ESTATE: '부동산',
      DEBT: '대출·부채',
      INCOME: '수입',
      EXPENSE_VARIABLE: '변동지출',
      EXPENSE_FIXED: '고정지출',
      TRANSFER: '계좌이체'
    };
    return map[cat] || cat || '';
  },

  // 순수 한글 수사 독음 변환 (예: 112,109,300 -> 일억 천이백십만 구천삼백원)
  formatHangulSpelledOut(amt) {
    if (amt === undefined || amt === null || isNaN(amt)) return '영원';
    const sign = amt < 0 ? '마이너스 ' : '';
    let n = Math.round(Math.abs(amt));
    if (n === 0) return '영원';

    const digits = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'];
    const smallUnits = ['', '십', '백', '천'];
    const bigUnits = ['', '만', '억', '조', '경'];

    const blocks = [];
    while (n > 0) {
      blocks.push(n % 10000);
      n = Math.floor(n / 10000);
    }

    const resultParts = [];
    for (let i = blocks.length - 1; i >= 0; i--) {
      const block = blocks[i];
      if (block === 0) continue;

      let blockStr = '';
      let b = block;
      for (let j = 0; j < 4; j++) {
        const d = b % 10;
        if (d > 0) {
          // 십, 백, 천 앞의 1은 '일십' 대신 '십', '일백' 대신 '백' 등으로 부름
          const digitStr = (d === 1 && j > 0) ? '' : digits[d];
          blockStr = digitStr + smallUnits[j] + blockStr;
        }
        b = Math.floor(b / 10);
      }

      if (blockStr) {
        // 1억인 경우 '억' 대신 '일억'
        if (i >= 2 && block === 1) {
          blockStr = '일' + bigUnits[i];
        } else if (i > 0) {
          blockStr += bigUnits[i];
        }
        resultParts.push(blockStr);
      }
    }

    return sign + resultParts.join(' ') + '원';
  },

  formatKoreanAmount(amt) {
    if (amt === undefined || amt === null || isNaN(amt)) return '0원';
    const sign = amt < 0 ? '-' : '';
    const abs = Math.round(Math.abs(amt));
    if (abs === 0) return '0원';

    const eok = Math.floor(abs / 100000000);
    const man = Math.floor((abs % 100000000) / 10000);
    const won = abs % 10000;

    if (eok > 0) {
      if (man > 0) {
        if (won > 0) return `${sign}${eok.toLocaleString()}억 ${man.toLocaleString()}만 ${won.toLocaleString()}원`;
        return `${sign}${eok.toLocaleString()}억 ${man.toLocaleString()}만 원`;
      }
      if (won > 0) return `${sign}${eok.toLocaleString()}억 ${won.toLocaleString()}원`;
      return `${sign}${eok.toLocaleString()}억 원`;
    }
    if (man > 0) {
      if (won > 0) return `${sign}${man.toLocaleString()}만 ${won.toLocaleString()}원`;
      return `${sign}${man.toLocaleString()}만 원`;
    }
    return `${sign}${won.toLocaleString()}원`;
  },

  // 1. [홈 대시보드 탭]
  renderHomeTab(fin) {
    const netWorthEl = document.getElementById('homeNetWorthText');
    if (netWorthEl) {
      const netWorthVal = fin.netWorth || 0;
      const numStr = `${netWorthVal.toLocaleString()}원`;
      const korAmount = this.formatKoreanAmount(netWorthVal);
      netEl = netWorthEl;
      netWorthEl.innerHTML = `
        <div style="font-size:38px; font-weight:900; color:#ffffff; letter-spacing:-0.035em; line-height:1.15; word-break:keep-all;">
          ${numStr}
        </div>
        <div style="font-size:14px; font-weight:600; color:rgba(255, 255, 255, 0.75); margin-top:4px; letter-spacing:-0.01em;">
          약 ${korAmount}
        </div>
      `;
    }

    const totalAssetEl = document.getElementById('homeTotalAssetText');
    if (totalAssetEl) totalAssetEl.textContent = `총자산 ${this.formatKoreanAmount(fin.totalAssets)}`;

    const totalDebtEl = document.getElementById('homeTotalDebtText');
    if (totalDebtEl) totalDebtEl.textContent = `부채 ${this.formatKoreanAmount(fin.totalDebt)}`;

    const today = new Date();
    const prevMonthLastDay = new Date(today.getFullYear(), today.getMonth(), 0);
    const prevMonthEndStr = `${prevMonthLastDay.getFullYear()}-${String(prevMonthLastDay.getMonth() + 1).padStart(2, '0')}-${String(prevMonthLastDay.getDate()).padStart(2, '0')}`;
    const prevSnap = this.calculateAssetSnapshotOnDate(prevMonthEndStr);

    const changeEl = document.getElementById('homeNetWorthChangeText');
    if (changeEl) {
      if (!prevSnap || !prevSnap.hasHistory) {
        if (fin.totalAssets === 0 && fin.totalDebt === 0) {
          changeEl.textContent = '등록된 자산이 없습니다';
          changeEl.style.color = 'rgba(255, 255, 255, 0.75)';
        } else {
          changeEl.textContent = '이번 달 신규 등록 자산';
          changeEl.style.color = 'rgba(255, 255, 255, 0.85)';
        }
      } else {
        const netDiff = fin.netWorth - prevSnap.netWorth;
        if (netDiff > 0) {
          changeEl.textContent = `전월 말 대비 +${this.formatKoreanAmount(netDiff)}`;
          changeEl.style.color = '#fca5a5';
        } else if (netDiff < 0) {
          changeEl.textContent = `전월 말 대비 -${this.formatKoreanAmount(Math.abs(netDiff))}`;
          changeEl.style.color = '#38bdf8';
        } else {
          changeEl.textContent = `전월 말 대비 변동 없음`;
          changeEl.style.color = 'rgba(255, 255, 255, 0.75)';
        }
      }
    }

    const incEl = document.getElementById('homeMonthIncomeText');
    if (incEl) incEl.textContent = this.formatKoreanAmount(fin.actualMonthIncome);

    const expEl = document.getElementById('homeMonthExpenseText');
    if (expEl) expEl.textContent = this.formatKoreanAmount(fin.actualMonthExpense);

    const remEl = document.getElementById('homeRemainingSurplusText');
    if (remEl) remEl.textContent = this.formatKoreanAmount(fin.actualRemainingSurplus);

    const subEl = document.getElementById('homeSurplusSubtext');
    if (subEl) subEl.textContent = '수입 – 지출';

    const badgeEl = document.getElementById('homeSavingsRateBadge');
    if (badgeEl) badgeEl.textContent = `저축률 ${fin.actualSavingsRate.toFixed(1)}%`;

    this.renderHomeAssetGrowthChart(fin);
    this.renderHomeAssetSegmentBar(fin);
    this.renderHomeRecentTransactions(fin);
  },

  renderHomeAssetGrowthChart(fin) {
    const canvas = document.getElementById('homeAssetGrowthCanvas');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    if (this.charts.homeAssetGrowth) {
      this.charts.homeAssetGrowth.destroy();
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

    const firstSnap = snapshots[0];
    const firstVal = data[0] || 0;
    const lastVal = data[data.length - 1] || 0;
    const diff = lastVal - firstVal;
    const badgeEl = document.getElementById('homeGrowthBadge');

    if (badgeEl) {
      const sign = diff > 0 ? '+' : '';
      if (!firstSnap.hasHistory || firstVal === 0) {
        if (lastVal === 0) {
          badgeEl.textContent = '변동 없음 (0원)';
          badgeEl.style.color = '#64748b';
          badgeEl.style.background = 'rgba(100, 116, 139, 0.1)';
        } else {
          badgeEl.textContent = `${sign}${this.formatKoreanAmount(diff)} (신규)`;
          badgeEl.style.color = diff > 0 ? '#f04452' : (diff < 0 ? '#3182f6' : '#64748b');
          badgeEl.style.background = diff > 0 ? 'rgba(240, 68, 82, 0.1)' : (diff < 0 ? 'rgba(49, 130, 246, 0.1)' : 'rgba(100, 116, 139, 0.1)');
        }
      } else {
        const rate = (diff / Math.abs(firstVal)) * 100;
        const rateSign = rate > 0 ? '+' : '';
        badgeEl.textContent = `${sign}${this.formatKoreanAmount(diff)} (${rateSign}${rate.toFixed(1)}%)`;
        badgeEl.style.color = diff > 0 ? '#f04452' : (diff < 0 ? '#3182f6' : '#64748b');
        badgeEl.style.background = diff > 0 ? 'rgba(240, 68, 82, 0.1)' : (diff < 0 ? 'rgba(49, 130, 246, 0.1)' : 'rgba(100, 116, 139, 0.1)');
      }
    }

    const minVal = Math.min(...data);
    const maxVal = Math.max(...data);
    const range = (maxVal - minVal) || 1000000;

    canvas.style.cursor = 'pointer';

    const lastIdx = data.length - 1;
    const pointRadii = data.map((_, idx) => idx === lastIdx ? 5.5 : 3);
    const pointBg = data.map((_, idx) => idx === lastIdx ? '#003049' : '#669bbc');

    this.charts.homeAssetGrowth = new Chart(ctx, {
      type: 'line',
      data: {
        labels,
        datasets: [{
          data,
          borderColor: '#003049',
          borderWidth: 2.5,
          backgroundColor: 'rgba(0, 48, 73, 0.08)',
          pointBackgroundColor: pointBg,
          pointBorderColor: '#ffffff',
          pointBorderWidth: 2,
          pointRadius: pointRadii,
          pointHoverRadius: 6,
          tension: 0.45,
          fill: true
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        onClick: (evt, elements) => {
          if (elements && elements.length > 0) {
            const idx = elements[0].index;
            const snap = snapshots[idx];
            if (snap) this.openTrendDetailModal('month', snap);
          } else {
            this.switchTab('assets');
          }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: 'rgba(15, 23, 42, 0.95)',
            titleColor: '#38bdf8',
            bodyColor: '#ffffff',
            titleFont: { size: 12, weight: '700' },
            bodyFont: { size: 12, weight: '600' },
            padding: 8,
            cornerRadius: 8,
            displayColors: false,
            callbacks: {
              title: (items) => `${labels[items[0].dataIndex]} 순자산 현황`,
              label: (c) => ` ₩${Math.round(c.raw).toLocaleString()} (클릭 시 상세 내역)`
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { font: { size: 11, weight: '600' }, color: '#64748b' }
          },
          y: {
            display: false,
            min: minVal === 0 && maxVal === 0 ? -100 : Math.min(0, minVal) - range * 0.25,
            max: minVal === 0 && maxVal === 0 ? 100 : maxVal + range * 0.3
          }
        }
      }
    });
  },

  // [홈 대시보드 5번 카드] 자산 배분 가로 세그먼트 바 (부채 제외 순수 총자산 100% 기준)
  renderHomeAssetSegmentBar(fin) {
    const barEl = document.getElementById('homeAssetSegmentBar');
    const labelEl = document.getElementById('homeAssetSegmentLabels');
    if (!barEl) return;

    const totalAssetsOnly = fin.totalCash + fin.totalSavings + fin.totalStock + fin.totalPension + (fin.totalOther || 0);

    if (totalAssetsOnly === 0) {
      barEl.innerHTML = `<div style="width:100%; height:100%; background:var(--color-divider); border-radius:6px;"></div>`;
      if (labelEl) {
        labelEl.textContent = fin.totalDebt > 0
          ? `보유 자산 0원 (부채 ${this.formatKoreanAmount(fin.totalDebt)})`
          : '등록된 자산이 없습니다 (0원)';
      }
      return;
    }

    const items = [
      { key: 'CASH', name: '현금', amt: fin.totalCash, color: '#003049' },
      { key: 'SAVINGS', name: '예·적금', amt: fin.totalSavings, color: '#0284c7' },
      { key: 'STOCK', name: '투자', amt: fin.totalStock, color: '#2563eb' },
      { key: 'PENSION', name: '연금', amt: fin.totalPension, color: '#0d9488' },
      { key: 'OTHER', name: '기타', amt: fin.totalOther, color: '#64748b' }
    ].filter(it => it.amt > 0);

    barEl.innerHTML = items.map(it => {
      const pct = Math.max(4, (it.amt / totalAssetsOnly) * 100);
      return `<div style="width:${pct}%; height:100%; background-color:${it.color}; border-radius:3px;" title="${it.name}: ₩${it.amt.toLocaleString()} (${((it.amt / totalAssetsOnly) * 100).toFixed(1)}%)"></div>`;
    }).join('');

    if (labelEl) {
      const assetSummary = items.map(it => `${it.name} ${Math.round((it.amt / totalAssetsOnly) * 100)}%`).join(' · ');
      const debtNote = fin.totalDebt > 0 ? ` (부채 별도 ${this.formatKoreanAmount(fin.totalDebt)})` : '';
      labelEl.textContent = assetSummary + debtNote;
    }
  },

  // [홈 대시보드 6번 카드] 최근 거래 내역 (최신순 4건)
  renderHomeRecentTransactions(fin) {
    const recentEl = document.getElementById('homeRecentList');
    if (!recentEl) return;

    const sortedTxs = [...(this.data.transactions || [])].sort((a, b) => (b.date || '').localeCompare(a.date || ''));
    const recent = sortedTxs.slice(0, 4);

    if (recent.length === 0) {
      recentEl.innerHTML = `
        <div style="text-align:center; padding:24px 16px; color:var(--color-text-tertiary);">
          <div style="font-size:13px; font-weight:700; color:var(--color-text-secondary); margin-bottom:4px;">거래 내역이 없습니다</div>
          <div style="font-size:11px; color:var(--color-text-tertiary); margin-bottom:12px;">가계부 또는 홈에서 새 거래를 등록해 보세요.</div>
          <button onclick="App.openAddModal()" style="border:none; background:var(--brand-primary); color:#ffffff; font-size:11px; font-weight:700; padding:6px 14px; border-radius:8px; cursor:pointer;">+ 거래 등록</button>
        </div>
      `;
      return;
    }

    const accMap = {};
    (this.data.accounts || []).forEach(a => accMap[a.id] = a.name);
    (this.data.assets || []).forEach(a => accMap[a.id] = a.name);

    recentEl.innerHTML = recent.map((t, idx) => {
      const isIncome = t.type === 'INCOME';
      const isTransfer = t.type === 'TRANSFER';
      const accName = accMap[t.accountId] || '결제계좌';
      const sign = isIncome ? '+' : (isTransfer ? '↔ ' : '-');
      const color = isIncome ? 'var(--color-gain)' : (isTransfer ? '#64748b' : 'var(--color-text-primary)');
      const { icon, tint } = this.getCategoryIcon(t.category, t.type);

      const parts = (t.date || '').split('-');
      const dateText = parts.length === 3 ? `${parseInt(parts[1])}월 ${parseInt(parts[2])}일` : t.date;
      const borderStyle = idx < recent.length - 1 ? 'border-bottom: 1px solid var(--color-divider);' : '';

      return `
        <div class="dash-recent-row" onclick="App.openEditTxModal('${t.id}')" style="${borderStyle}">
          <div style="display:flex; align-items:center; gap:10px; min-width:0;">
            <div class="finance-icon ${tint}" style="width:34px; height:34px; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
              <i data-lucide="${icon}" style="width:16px; height:16px;"></i>
            </div>
            <div class="dash-recent-info">
              <div class="dash-recent-title">${this.escapeHtml(t.memo || t.category)}</div>
              <div class="dash-recent-sub">${this.escapeHtml(dateText)} · ${this.escapeHtml(accName)}</div>
            </div>
          </div>
          <div class="dash-recent-amount" style="color:${color};">
            ${sign}${this.formatKoreanAmount(t.amount)}
          </div>
        </div>
      `;
    }).join('');
  },

  // 전체 자산 배분 비중 상세 모달 (순수 총자산 100% 기준 도넛 + 부채 별도 표시)
  openAssetAllocationModal() {
    const modal = document.getElementById('tossAssetAllocationModal');
    if (!modal) return;

    const fin = this.getFinancialSummary();

    const totAssetEl = document.getElementById('allocSummaryTotalAssets');
    if (totAssetEl) totAssetEl.textContent = this.formatKoreanAmount(fin.totalAssets);
    const totDebtEl = document.getElementById('allocSummaryTotalDebt');
    if (totDebtEl) totDebtEl.textContent = this.formatKoreanAmount(fin.totalDebt);
    const netWorthEl = document.getElementById('allocSummaryNetWorth');
    if (netWorthEl) netWorthEl.textContent = this.formatKoreanAmount(fin.netWorth);

    const cashAccounts = (this.data.accounts || []).filter(a => Number(a.balance) > 0);
    const savingsAssets = (this.data.assets || []).filter(a => a.category === 'SAVINGS' && Number(a.amount) > 0);
    const stockAssets = (this.data.assets || []).filter(a => a.category === 'STOCK' && Number(a.amount) > 0);
    const pensionAssets = (this.data.assets || []).filter(a => a.category === 'PENSION' && Number(a.amount) > 0);
    const otherAssets = (this.data.assets || []).filter(a => (a.category === 'OTHER' || a.category === 'REAL_ESTATE') && Number(a.amount) > 0);
    const debtAssets = (this.data.assets || []).filter(a => a.category === 'DEBT' && Number(a.amount) > 0);

    const assetCatDefs = [
      {
        key: 'CASH',
        name: '현금 (입출금 계좌)',
        targetSection: 'card-asset-accounts',
        amt: fin.totalCash,
        color: '#003049',
        items: cashAccounts.map(a => ({ id: a.id, name: a.name, amount: Number(a.balance), sub: a.institution || '은행' }))
      },
      {
        key: 'SAVINGS',
        name: '예금 · 적금',
        targetSection: 'card-asset-savings',
        amt: fin.totalSavings,
        color: '#0284c7',
        items: savingsAssets.map(s => ({ id: s.id, name: s.name, amount: Number(s.amount), sub: s.institution || '예적금' }))
      },
      {
        key: 'STOCK',
        name: '투자 (주식 / ETF)',
        targetSection: 'card-asset-stocks',
        amt: fin.totalStock,
        color: '#2563eb',
        items: stockAssets.map(s => ({ id: s.id, name: s.name, amount: Number(s.amount), sub: s.ticker || '투자' }))
      },
      {
        key: 'PENSION',
        name: '연금 (연금저축 / IRP)',
        targetSection: 'card-asset-pension',
        amt: fin.totalPension,
        color: '#0d9488',
        items: pensionAssets.map(p => ({ id: p.id, name: p.name, amount: Number(p.amount), sub: p.institution || '연금' }))
      },
      {
        key: 'OTHER',
        name: '부동산 · 실물 자산',
        targetSection: 'card-asset-other',
        amt: fin.totalOther,
        color: '#64748b',
        items: otherAssets.map(o => ({ id: o.id, name: o.name, amount: Number(o.amount), sub: o.institution || '기타' }))
      }
    ];

    const activeAssetCats = assetCatDefs.filter(c => c.amt > 0).sort((a, b) => b.amt - a.amt);
    const totalPureAssets = activeAssetCats.reduce((acc, c) => acc + c.amt, 0);

    const canvas = document.getElementById('allocationModalDonutCanvas');
    if (canvas) {
      const ctx = canvas.getContext('2d');
      if (this.charts.allocationModalDonut) {
        this.charts.allocationModalDonut.destroy();
      }

      if (activeAssetCats.length === 0 || totalPureAssets === 0) {
        this.charts.allocationModalDonut = new Chart(ctx, {
          type: 'doughnut',
          data: {
            labels: ['보유 자산 없음'],
            datasets: [{ data: [1], backgroundColor: ['#e2e8f0'], borderWidth: 2, borderColor: '#ffffff' }]
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
            labels: activeAssetCats.map(c => c.name),
            datasets: [{
              data: activeAssetCats.map(c => c.amt),
              backgroundColor: activeAssetCats.map(c => c.color),
              borderWidth: 2,
              borderColor: '#ffffff'
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { display: false },
              tooltip: {
                callbacks: {
                  label: (c) => ` ${c.label}: ₩${Number(c.raw).toLocaleString()} (${((Number(c.raw) / totalPureAssets) * 100).toFixed(1)}%)`
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
      if (activeAssetCats.length === 0 && fin.totalDebt === 0) {
        listEl.innerHTML = `<div style="text-align:center; padding:18px; font-size:13px; color:#8b95a1; background:#f9fafb; border-radius:12px;">등록된 자산이 없습니다.</div>`;
      } else {
        const assetHtml = activeAssetCats.map(c => {
          const pct = totalPureAssets > 0 ? ((c.amt / totalPureAssets) * 100).toFixed(1) : '0.0';
          const subItemsHtml = c.items.map(it => {
            const itPct = totalPureAssets > 0 ? ((it.amount / totalPureAssets) * 100).toFixed(1) : '0.0';
            return `
              <div onclick="event.stopPropagation(); App.goToAssetCategory('${c.targetSection}');" style="display:flex; justify-content:space-between; align-items:center; padding:9px 12px; background:#ffffff; border-radius:10px; margin-top:6px; border:1px solid #f2f4f6; cursor:pointer;">
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
            <div style="background:#f9fafb; border-radius:14px; border:1px solid #eaedf0; padding:12px 14px; margin-bottom:8px;">
              <div onclick="App.goToAssetCategory('${c.targetSection}')" style="display:flex; justify-content:space-between; align-items:center; cursor:pointer;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="width:12px; height:12px; border-radius:50%; background-color:${c.color}; flex-shrink:0;"></span>
                  <div>
                    <div style="font-size:14px; font-weight:800; color:#191f28;">${this.escapeHtml(c.name)}</div>
                    <div style="font-size:11px; color:#8b95a1; margin-top:2px; font-weight:600;">자산 내 비중 ${pct}% (${c.items.length}개 항목)</div>
                  </div>
                </div>
                <div style="display:flex; align-items:center; gap:6px;">
                  <div style="font-size:15px; font-weight:800; color:#191f28; font-family:var(--font-toss);">₩${c.amt.toLocaleString()}</div>
                  <i data-lucide="chevron-right" style="width:16px; height:16px; color:#8b95a1;"></i>
                </div>
              </div>
              <div style="margin-top:8px; border-top:1px dashed #eaedf0; padding-top:6px;">
                ${subItemsHtml}
              </div>
            </div>
          `;
        }).join('');

        // 부채가 있는 경우 하단에 별도(차감 항목) 섹션으로 분리 표기
        let debtHtml = '';
        if (fin.totalDebt > 0) {
          const debtRatio = totalPureAssets > 0 ? ((fin.totalDebt / totalPureAssets) * 100).toFixed(1) : '100.0';
          const debtSubHtml = debtAssets.map(d => `
            <div onclick="event.stopPropagation(); App.goToAssetCategory('card-asset-debt');" style="display:flex; justify-content:space-between; align-items:center; padding:9px 12px; background:#ffffff; border-radius:10px; margin-top:6px; border:1px solid #fee2e2; cursor:pointer;">
              <div style="display:flex; align-items:center; gap:8px;">
                <span style="font-size:13px; font-weight:700; color:#dc2626;">${this.escapeHtml(d.name)}</span>
                <span style="font-size:11px; color:#8b95a1;">${this.escapeHtml(d.institution || '대출/부채')}</span>
              </div>
              <div style="text-align:right;">
                <span style="font-size:13px; font-weight:800; color:#dc2626; font-family:var(--font-toss);">-₩${Number(d.amount).toLocaleString()}</span>
              </div>
            </div>
          `).join('');

          debtHtml = `
            <div style="background:rgba(220, 38, 38, 0.04); border-radius:14px; border:1px solid rgba(220, 38, 38, 0.2); padding:12px 14px; margin-top:4px;">
              <div onclick="App.goToAssetCategory('card-asset-debt')" style="display:flex; justify-content:space-between; align-items:center; cursor:pointer;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="width:12px; height:12px; border-radius:50%; background-color:#dc2626; flex-shrink:0;"></span>
                  <div>
                    <div style="font-size:14px; font-weight:800; color:#dc2626;">대출 · 부채 (차감 항목)</div>
                    <div style="font-size:11px; color:#ef4444; margin-top:2px; font-weight:600;">총자산 대비 부채비율 ${debtRatio}%</div>
                  </div>
                </div>
                <div style="display:flex; align-items:center; gap:6px;">
                  <div style="font-size:15px; font-weight:800; color:#dc2626; font-family:var(--font-toss);">-₩${fin.totalDebt.toLocaleString()}</div>
                  <i data-lucide="chevron-right" style="width:16px; height:16px; color:#dc2626;"></i>
                </div>
              </div>
              ${debtSubHtml ? `<div style="margin-top:8px; border-top:1px dashed rgba(220, 38, 38, 0.2); padding-top:6px;">${debtSubHtml}</div>` : ''}
            </div>
          `;
        }

        listEl.innerHTML = assetHtml + debtHtml;
      }
    }

    modal.classList.remove('hidden');
    this.scheduleIcons();
  },

  closeAssetAllocationModal() {
    const modal = document.getElementById('tossAssetAllocationModal');
    if (modal) modal.classList.add('hidden');
  },

  openStockPortfolioModal() {
    const modal = document.getElementById('stockPortfolioDetailModal');
    if (!modal) return;
    modal.classList.remove('hidden');
    this.renderAssetsTab(this.getFinancialSummary());
    this.scheduleIcons();
  },

  closeStockPortfolioModal() {
    const modal = document.getElementById('stockPortfolioDetailModal');
    if (modal) modal.classList.add('hidden');
  },

  openAccountsModal() {
    const modal = document.getElementById('accountsDetailModal');
    if (!modal) return;
    modal.classList.remove('hidden');
    this.renderAssetsTab(this.getFinancialSummary());
    this.scheduleIcons();
  },

  closeAccountsModal() {
    const modal = document.getElementById('accountsDetailModal');
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
        targetCard.style.boxShadow = '0 0 0 3px rgba(2, 132, 199, 0.4)';
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
    if (incEl) incEl.textContent = `₩${(Number(fin.monthIncome) || 0).toLocaleString()}`;

    const expEl = document.getElementById('heroExpenseAmt');
    if (expEl) expEl.textContent = `₩${(Number(fin.monthExpense) || 0).toLocaleString()}`;

    const hasBudget = Number(this.data.totalMonthlyBudget) > 0;
    const effectiveBudget = Number(this.data.totalMonthlyBudget) || 0;
    const remEl = document.getElementById('heroRemainingAmt');
    const budgetGuideEl = document.getElementById('heroTotalBudgetGuide');
    const dailyEl = document.getElementById('heroDailyGuide');
    const barFill = document.getElementById('heroBudgetBarFill');
    const barPin = document.getElementById('heroBudgetBarPin');

    const daysInMonth = new Date(Number(y), Number(m), 0).getDate();
    const today = new Date();
    const currentYM = getCurrentYearMonthStr();
    const isThisMonth = (this.data.selectedYearMonth === currentYM);
    const isPastMonth = (this.data.selectedYearMonth < currentYM);
    const isFutureMonth = (this.data.selectedYearMonth > currentYM);

    const currentDay = isThisMonth ? today.getDate() : (isPastMonth ? daysInMonth : 0);
    const remainingDays = isThisMonth ? Math.max(1, daysInMonth - currentDay + 1) : (isPastMonth ? 0 : daysInMonth);

    if (!hasBudget) {
      if (remEl) {
        remEl.textContent = '예산 미설정';
        remEl.style.color = 'rgba(255, 255, 255, 0.85)';
      }
      if (budgetGuideEl) budgetGuideEl.textContent = '목표 예산 설정하기 >';
      if (dailyEl) dailyEl.textContent = '한 달 목표 예산을 설정해 보세요';
      if (barFill) barFill.style.width = '0%';
      if (barPin) barPin.style.display = 'none';
    } else {
      const diff = effectiveBudget - fin.monthExpense;
      if (remEl) {
        if (diff >= 0) {
          remEl.textContent = `${diff.toLocaleString()}원 남음`;
          remEl.style.color = '#ffffff';
        } else {
          remEl.textContent = `-${Math.abs(diff).toLocaleString()}원 초과`;
          remEl.style.color = '#fca5a5';
        }
      }
      if (budgetGuideEl) budgetGuideEl.textContent = `총 예산 ₩${effectiveBudget.toLocaleString()}`;

      if (dailyEl) {
        if (isPastMonth) {
          dailyEl.textContent = '해당 월 지출 마감';
        } else {
          const dailyBudget = remainingDays > 0 ? Math.round(Math.max(0, diff) / remainingDays) : 0;
          dailyEl.textContent = `목표 달성을 위한 하루 권장 예산 ${dailyBudget.toLocaleString()}원`;
        }
      }

      const budgetPct = Math.min(100, Math.round((fin.monthExpense / effectiveBudget) * 100));
      if (barFill) {
        barFill.style.width = `${budgetPct}%`;
        barFill.style.backgroundColor = diff < 0 ? '#fca5a5' : '#ffffff';
      }
      if (barPin) {
        if (isFutureMonth) {
          barPin.style.display = 'none';
        } else {
          barPin.style.display = 'block';
          const monthProgressPct = Math.round((currentDay / daysInMonth) * 100);
          const clampedPct = Math.max(6, Math.min(94, monthProgressPct));
          barPin.style.left = `${clampedPct}%`;
        }
      }
    }

    const segTotalEl = document.getElementById('segTotalExpenseText');
    if (segTotalEl) segTotalEl.textContent = `${(Number(fin.monthExpense) || 0).toLocaleString()}원`;

    const catTotals = {};
    fin.monthTxs.filter(t => t.type && t.type.startsWith('EXPENSE')).forEach(t => {
      catTotals[t.category] = (catTotals[t.category] || 0) + Number(t.amount);
    });

    const palette = ['#003049', '#0284c7', '#2563eb', '#f97316', '#0d9488', '#64748b'];
    const sortedCats = Object.entries(catTotals).sort((a, b) => b[1] - a[1]);

    const segBar = document.getElementById('spendingSegmentBar');
    if (segBar) {
      if (sortedCats.length === 0 || fin.monthExpense === 0) {
        segBar.innerHTML = `<div class="segment-item" style="width:100%; background:#eaedf0;"></div>`;
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
          const catBudget = (this.data.budgets && this.data.budgets[cat]) ? Number(this.data.budgets[cat]) : 0;

          return `
            <div class="category-summary-item">
              <div class="category-summary-top">
                <div style="display:flex; align-items:center; min-width:0;">
                  <span class="category-dot" style="background-color:${palette[idx % palette.length]};"></span>
                  <span style="font-size:13px; font-weight:700; color:var(--color-text-primary); overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${this.escapeHtml(cat)}</span>
                </div>
                <span style="font-size:12px; font-weight:800; color:var(--brand-primary); flex-shrink:0;">${pct}%</span>
              </div>
              <div class="category-summary-bottom">
                <span style="font-size:13px; font-weight:800; color:var(--color-text-primary); font-family:var(--font-toss);">₩${amt.toLocaleString()}</span>
                ${catBudget > 0 ? `<span style="font-size:10px; font-weight:700; padding:1px 6px; border-radius:4px; ${catBudget - amt >= 0 ? 'background:#e8f5e9; color:#2e7d32;' : 'background:#fff2f2; color:#dc2626;'} flex-shrink:0;">${catBudget - amt >= 0 ? '남음' : '초과'}</span>` : ''}
              </div>
              ${catBudget > 0 ? `
                <div style="height:3px; background:#eaedf0; border-radius:9999px; overflow:hidden; margin-top:2px;">
                  <div style="height:100%; width:${Math.min(100, Math.round((amt / catBudget) * 100))}%; background:${catBudget - amt >= 0 ? 'var(--brand-primary)' : '#dc2626'}; border-radius:9999px;"></div>
                </div>
                <div style="font-size:10px; color:${catBudget - amt >= 0 ? 'var(--color-text-tertiary)' : '#dc2626'}; margin-top:1px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                  예산 ${this.formatKoreanAmount(catBudget)} · ${catBudget - amt >= 0 ? `${this.formatKoreanAmount(catBudget - amt)} 남음` : `${this.formatKoreanAmount(Math.abs(catBudget - amt))} 초과`}
                </div>
              ` : ''}
            </div>
          `;
        }).join('');
      }
    }

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
            <button onclick="App.openAddModal()" style="border:none; background:var(--brand-primary); color:#ffffff; font-size:11px; font-weight:700; padding:6px 14px; border-radius:8px; cursor:pointer;">+ 거래 등록</button>
          </div>
        `;
      } else {
        const allTxs = [...fin.monthTxs].sort((a, b) => (b.date || '').localeCompare(a.date || ''));
        const totalCount = allTxs.length;
        const limit = this.spendingTimelineLimit || 15;
        const displayTxs = allTxs.slice(0, limit);
        const hasMore = totalCount > limit;
        const remainingCount = totalCount - limit;

        const grouped = {};
        displayTxs.forEach(it => {
          if (!grouped[it.date]) grouped[it.date] = [];
          grouped[it.date].push(it);
        });

        const accMap = {};
        (this.data.accounts || []).forEach(a => accMap[a.id] = a.name);
        (this.data.assets || []).forEach(a => accMap[a.id] = a.name);

        let html = Object.entries(grouped).map(([date, dayItems]) => {
          const parts = (date || '').split('-');
          const dateLabel = parts.length === 3 ? `${parseInt(parts[1])}월 ${parseInt(parts[2])}일` : date;

          return `
            <div class="timeline-date-group">
              <div class="timeline-date-title">${this.escapeHtml(dateLabel)}</div>
              <div class="finance-list-card">
                ${dayItems.map(it => {
                  const isIncome = it.type === 'INCOME';
                  const isTransfer = it.type === 'TRANSFER';
                  const accName = accMap[it.accountId] || '결제계좌';
                  const toAccName = (isTransfer && it.toAccountId) ? ` → ${accMap[it.toAccountId] || '입금계좌'}` : '';
                  const { icon, tint } = this.getCategoryIcon(it.category, it.type);
                  const sign = isIncome ? '+' : (isTransfer ? '↔ ' : '-');
                  const color = isIncome ? 'var(--color-gain)' : (isTransfer ? '#64748b' : 'var(--color-text-primary)');
                  const isFixed = it.type === 'EXPENSE_FIXED';
                  const tagText = isTransfer ? '이체' : (isIncome ? '입금' : (isFixed ? '고정지출' : '변동지출'));
                  const tagStyle = isFixed
                    ? 'padding:1px 6px; background:#ede7f6; color:#512da8; border-radius:4px; font-size:10px; font-weight:700;'
                    : (isTransfer ? 'padding:1px 5px; background:#f2f4f6; color:#4e5968; border-radius:4px; font-size:10px; font-weight:600;'
                    : (isIncome ? 'padding:1px 5px; background:rgba(2,132,199,0.12); color:#0284c7; border-radius:4px; font-size:10px; font-weight:600;'
                    : 'padding:1px 5px; background:#f2f4f6; color:#64748b; border-radius:4px; font-size:10px; font-weight:600;'));

                  return `
                    <div class="finance-row" onclick="App.openEditTxModal('${it.id}')">
                      <div class="finance-row-left">
                        <div class="finance-icon ${tint}"><i data-lucide="${icon}" style="width:18px; height:18px;"></i></div>
                        <div class="finance-row-info">
                          <div class="finance-row-title">${this.escapeHtml(it.memo || it.category)}</div>
                          <div class="finance-row-sub">${this.escapeHtml(accName)}${this.escapeHtml(toAccName)} · <span style="${tagStyle}">${tagText}</span></div>
                        </div>
                      </div>
                      <div class="finance-row-right">
                        <div class="finance-row-amount" style="color:${color};">
                          ${sign}₩${(Number(it.amount) || 0).toLocaleString()}
                        </div>
                      </div>
                    </div>
                  `;
                }).join('')}
              </div>
            </div>
          `;
        }).join('');

        if (hasMore) {
          html += `
            <div style="text-align:center; margin-top:12px; margin-bottom:16px;">
              <button type="button" onclick="App.loadMoreSpendingTimeline()" style="width:100%; padding:12px; border-radius:14px; border:1px solid var(--color-border); background:var(--color-surface); color:var(--brand-primary); font-size:13px; font-weight:700; cursor:pointer; display:flex; align-items:center; justify-content:center; gap:6px; box-shadow:0 1px 3px rgba(0,0,0,0.04); transition:all 0.15s ease;">
                <i data-lucide="chevron-down" style="width:16px; height:16px;"></i>
                <span>이전 거래 내역 더보기 (${remainingCount}건 남음)</span>
              </button>
            </div>
          `;
        } else if (limit > 15 && totalCount > 15) {
          html += `
            <div style="text-align:center; margin-top:12px; margin-bottom:16px;">
              <button type="button" onclick="App.collapseSpendingTimeline()" style="padding:8px 16px; border-radius:10px; border:1px solid var(--color-border); background:var(--color-surface-subtle); color:var(--color-text-tertiary); font-size:12px; font-weight:600; cursor:pointer; display:inline-flex; align-items:center; gap:4px;">
                <i data-lucide="chevron-up" style="width:14px; height:14px;"></i>
                <span>접기 (최근 15건만 보기)</span>
              </button>
            </div>
          `;
        }

        timelineContainer.innerHTML = html;
      }
    }

    this.renderRecurringExpenses(fin);
    this.updateSpendingViewModeUI(fin);
  },

  renderRecurringExpenses(fin) {
    const container = document.getElementById('recurringExpenseListContainer');
    const badge = document.getElementById('recurringCountBadge');
    if (!container) return;

    const list = this.data.recurringExpenses || [];
    const ym = this.data.selectedYearMonth || getCurrentYearMonthStr();

    const activeList = list.filter(r => {
      const start = r.startDate || '2000-01';
      const end = r.endDate;
      if (start > ym) return false;
      if (end && end < ym) return false;
      return true;
    });

    const totalAmt = activeList.reduce((sum, r) => sum + (Number(r.amount) || 0), 0);
    if (badge) badge.textContent = activeList.length > 0 ? `${activeList.length}건 · ₩${this.formatCompactAmount(totalAmt)}` : '0건';

    if (activeList.length === 0) {
      container.innerHTML = `
        <div style="text-align:center; padding:18px 12px; background:#f9fafb; border-radius:12px; border:1px dashed #e5e8eb;">
          <div style="font-size:12px; font-weight:700; color:var(--color-text-secondary); margin-bottom:2px;">이번 달 등록된 정기 고정 지출이 없습니다</div>
          <div style="font-size:11px; color:var(--color-text-tertiary);">우측 상단 [+ 추가] 버튼으로 월세, 통신비, 구독료 등을 등록해 보세요.</div>
        </div>
      `;
      return;
    }

    const accMap = {};
    (this.data.accounts || []).forEach(a => accMap[a.id] = a.name);
    (this.data.assets || []).forEach(a => accMap[a.id] = a.name);

    const monthTxs = (this.data.transactions || []).filter(t => t.date && t.date.startsWith(ym));
    let unappliedCount = 0;

    const rowsHtml = activeList.map(r => {
      const isApplied = monthTxs.some(t => t.recurringId === r.id || (t.type === 'EXPENSE_FIXED' && t.category === r.category && Number(t.amount) === Number(r.amount) && t.memo === r.name));
      if (!isApplied) unappliedCount++;

      const accName = accMap[r.accountId] || '출금계좌';
      const periodLabel = r.endDate ? `${r.endDate}까지` : '무제한 반복';
      const statusBadge = isApplied
        ? `<span style="font-size:10px; font-weight:700; padding:2px 7px; border-radius:9999px; background:#e8f5e9; color:#2e7d32; display:inline-flex; align-items:center; gap:3px; flex-shrink:0;"><i data-lucide="check" style="width:10px; height:10px;"></i>반영완료</span>`
        : `<span style="font-size:10px; font-weight:700; padding:2px 7px; border-radius:9999px; background:#fff8e6; color:#d97706; display:inline-flex; align-items:center; gap:3px; flex-shrink:0;"><i data-lucide="clock" style="width:10px; height:10px;"></i>미반영</span>`;

      return `
        <div class="finance-row" onclick="App.openRecurringModal('${r.id}')" style="padding:10px 12px; border-radius:10px; margin-bottom:6px; background:var(--color-surface-subtle); border:1px solid var(--color-border);">
          <div class="finance-row-left">
            <div style="width:32px; height:32px; border-radius:8px; background:#ede7f6; color:#512da8; display:flex; align-items:center; justify-content:center; font-size:13px; font-weight:800; flex-shrink:0;">
              ${r.dayOfMonth || 25}
            </div>
            <div class="finance-row-info">
              <div class="finance-row-title">${this.escapeHtml(r.name)}</div>
              <div class="finance-row-sub">매월 ${r.dayOfMonth || 25}일 · ${this.escapeHtml(accName)} · ${this.escapeHtml(r.category || '고정지출')}</div>
              <div class="finance-row-sub">${this.escapeHtml(periodLabel)}</div>
            </div>
          </div>
          <div class="finance-row-right">
            <div class="finance-row-amount" style="font-size:14px; color:var(--color-text-primary);">
              ₩${(Number(r.amount) || 0).toLocaleString()}
            </div>
            <div style="margin-top:2px;">${statusBadge}</div>
          </div>
        </div>
      `;
    }).join('');

    const actionHeader = unappliedCount > 0 ? `
      <div style="display:flex; justify-content:space-between; align-items:center; padding:8px 12px; background:var(--color-surface-subtle); border:1px solid var(--color-border); border-radius:10px; margin-bottom:8px;">
        <span style="font-size:11px; font-weight:700; color:var(--brand-primary);">이번 달 미반영 고정지출 <b>${unappliedCount}건</b></span>
        <button onclick="App.applyRecurringExpensesToMonth()" style="border:none; background:var(--brand-primary); color:var(--btn-primary-text); font-size:11px; font-weight:700; padding:4px 10px; border-radius:6px; cursor:pointer;">
          이번 달 일괄 반영하기
        </button>
      </div>
    ` : '';

    container.innerHTML = actionHeader + `<div style="display:flex; flex-direction:column;">${rowsHtml}</div>`;
  },

  loadMoreSpendingTimeline() {
    this.spendingTimelineLimit = (this.spendingTimelineLimit || 15) + 15;
    this.renderSpendingTab(this.getFinancialSummary());
    this.scheduleIcons();
  },

  toggleRecurringAccordion() {
    const container = document.getElementById('recurringExpenseListContainer');
    const icon = document.getElementById('recurringAccordionIcon');
    if (!container) return;
    const isHidden = container.classList.contains('hidden');
    if (isHidden) {
      container.classList.remove('hidden');
      if (icon) icon.style.transform = 'rotate(180deg)';
    } else {
      container.classList.add('hidden');
      if (icon) icon.style.transform = 'rotate(0deg)';
    }
    this.scheduleIcons();
  },

  toggleDividendAccordion() {
    const container = document.getElementById('dividendFoldableContainer');
    const icon = document.getElementById('dividendToggleIcon');
    const label = document.getElementById('dividendToggleBtnLabel');
    if (!container) return;
    const isHidden = container.classList.contains('hidden');
    if (isHidden) {
      container.classList.remove('hidden');
      if (icon) icon.style.transform = 'rotate(180deg)';
      if (label) label.textContent = '배당 닫기';
      this.renderDividendBarChart();
    } else {
      container.classList.add('hidden');
      if (icon) icon.style.transform = 'rotate(0deg)';
      if (label) label.textContent = '배당 분석';
    }
    this.scheduleIcons();
  },

  collapseSpendingTimeline() {
    this.spendingTimelineLimit = 15;
    this.renderSpendingTab(this.getFinancialSummary());
    this.scheduleIcons();
  },

  setSpendingViewMode(mode) {
    this.spendingViewMode = mode;
    this.updateSpendingViewModeUI(this.getFinancialSummary());
    this.scheduleIcons();
  },

  updateSpendingViewModeUI(fin) {
    const mode = this.spendingViewMode || 'list';
    const listBtn = document.getElementById('spendingViewModeListBtn');
    const calBtn = document.getElementById('spendingViewModeCalendarBtn');
    const listCont = document.getElementById('spendingTimelineList');
    const calCont = document.getElementById('spendingCalendarContainer');

    if (mode === 'calendar') {
      if (listBtn) listBtn.classList.remove('active');
      if (calBtn) calBtn.classList.add('active');
      if (listCont) listCont.classList.add('hidden');
      if (calCont) calCont.classList.remove('hidden');
      this.renderSpendingCalendar(fin);
    } else {
      if (listBtn) listBtn.classList.add('active');
      if (calBtn) calBtn.classList.remove('active');
      if (listCont) listCont.classList.remove('hidden');
      if (calCont) calCont.classList.add('hidden');
    }
    this.scheduleIcons();
  },

  formatCompactAmount(amt) {
    if (!amt || isNaN(amt)) return '0';
    const abs = Math.abs(Number(amt));
    if (abs >= 100000000) {
      return `${(abs / 100000000).toFixed(1)}억`;
    }
    if (abs >= 10000) {
      const man = Math.round(abs / 10000);
      return `${man}만`;
    }
    return `${Math.round(abs).toLocaleString()}`;
  },

  selectSpendingCalendarDate(dateStr) {
    this.selectedCalendarDate = dateStr;
    this.renderSpendingCalendar(this.getFinancialSummary());
    this.scheduleIcons();
  },

  renderSpendingCalendar(fin) {
    const gridEl = document.getElementById('spendingCalendarGrid');
    const detailEl = document.getElementById('spendingCalendarDayDetail');
    if (!gridEl) return;

    const ym = this.data.selectedYearMonth || new Date().toISOString().slice(0, 7);
    const [yStr, mStr] = ym.split('-');
    const year = parseInt(yStr, 10);
    const month = parseInt(mStr, 10);

    const firstDayOfWeek = new Date(year, month - 1, 1).getDay(); // 0(일) ~ 6(토)
    const daysInMonth = new Date(year, month, 0).getDate();

    const todayObj = new Date();
    const todayStr = `${todayObj.getFullYear()}-${String(todayObj.getMonth() + 1).padStart(2, '0')}-${String(todayObj.getDate()).padStart(2, '0')}`;

    if (!this.selectedCalendarDate || !this.selectedCalendarDate.startsWith(ym)) {
      this.selectedCalendarDate = todayStr.startsWith(ym) ? todayStr : `${ym}-01`;
    }

    const dailyMap = {};
    for (let d = 1; d <= daysInMonth; d++) {
      const dStr = `${year}-${String(month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      dailyMap[dStr] = { expense: 0, income: 0, txs: [] };
    }

    (fin.monthTxs || []).forEach(tx => {
      if (dailyMap[tx.date]) {
        if (tx.type === 'INCOME') {
          dailyMap[tx.date].income += Number(tx.amount) || 0;
        } else if (tx.type.startsWith('EXPENSE')) {
          dailyMap[tx.date].expense += Number(tx.amount) || 0;
        }
        dailyMap[tx.date].txs.push(tx);
      }
    });

    // 캘린더 요일 헤더 및 날짜 그리드 생성
    const daysHeaderHtml = `
      <div class="calendar-header-day sun">일</div>
      <div class="calendar-header-day">월</div>
      <div class="calendar-header-day">화</div>
      <div class="calendar-header-day">수</div>
      <div class="calendar-header-day">목</div>
      <div class="calendar-header-day">금</div>
      <div class="calendar-header-day sat">토</div>
    `;

    let cellsHtml = '';
    for (let i = 0; i < firstDayOfWeek; i++) {
      cellsHtml += `<div class="calendar-cell is-empty"></div>`;
    }

    for (let d = 1; d <= daysInMonth; d++) {
      const dStr = `${year}-${String(month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const dayData = dailyMap[dStr] || { expense: 0, income: 0, txs: [] };
      const dayOfWeek = (firstDayOfWeek + d - 1) % 7;
      const dowClass = dayOfWeek === 0 ? 'is-sun' : (dayOfWeek === 6 ? 'is-sat' : '');
      const isToday = dStr === todayStr ? 'is-today' : '';
      const isSelected = dStr === this.selectedCalendarDate ? 'is-selected' : '';

      const expenseHtml = dayData.expense > 0 ? `<div class="day-expense">-${this.formatCompactAmount(dayData.expense)}</div>` : '';
      const incomeHtml = dayData.income > 0 ? `<div class="day-income">+${this.formatCompactAmount(dayData.income)}</div>` : '';

      cellsHtml += `
        <div class="calendar-cell ${dowClass} ${isToday} ${isSelected}" onclick="App.selectSpendingCalendarDate('${dStr}')" title="${dStr}">
          <div class="day-num">${d}</div>
          ${expenseHtml}
          ${incomeHtml}
        </div>
      `;
    }

    gridEl.innerHTML = `<div class="calendar-grid">${daysHeaderHtml}${cellsHtml}</div>`;

    // 선택 일자 상세 거래 리스트 렌더링
    if (detailEl) {
      const selectedData = dailyMap[this.selectedCalendarDate] || { expense: 0, income: 0, txs: [] };
      const [selY, selM, selD] = (this.selectedCalendarDate || '').split('-');
      const selDateObj = new Date(parseInt(selY, 10), parseInt(selM, 10) - 1, parseInt(selD, 10));
      const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
      const dayName = dayNames[selDateObj.getDay()] || '';
      const dateTitle = `${parseInt(selM, 10)}월 ${parseInt(selD, 10)}일 (${dayName})`;

      const accMap = {};
      (this.data.accounts || []).forEach(a => accMap[a.id] = a.name);
      (this.data.assets || []).forEach(a => accMap[a.id] = a.name);

      let detailContent = '';
      if (selectedData.txs.length === 0) {
        detailContent = `
          <div style="text-align:center; padding:22px 14px; background:var(--color-surface); border-radius:16px; border:1px solid var(--color-border); color:var(--color-text-tertiary);">
            <div style="font-size:13px; font-weight:700; color:var(--color-text-secondary); margin-bottom:4px;">거래 내역이 없습니다</div>
            <div style="font-size:11px; color:var(--color-text-tertiary);">${this.selectedCalendarDate <= todayStr ? '지출 없이 절약한 날입니다 🌱' : '미래 날짜입니다.'}</div>
          </div>
        `;
      } else {
        const sortedTxs = [...selectedData.txs].sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || ''));
        detailContent = `
          <div class="finance-list-card">
            ${sortedTxs.map(it => {
              const isIncome = it.type === 'INCOME';
              const isTransfer = it.type === 'TRANSFER';
              const accName = accMap[it.accountId] || '결제계좌';
              const toAccName = (isTransfer && it.toAccountId) ? ` → ${accMap[it.toAccountId] || '입금계좌'}` : '';
              const { icon, tint } = this.getCategoryIcon(it.category, it.type);
              const sign = isIncome ? '+' : (isTransfer ? '↔ ' : '-');
              const color = isIncome ? 'var(--color-gain)' : (isTransfer ? '#64748b' : 'var(--color-text-primary)');
              const isFixed = it.type === 'EXPENSE_FIXED';
              const tagText = isTransfer ? '이체' : (isIncome ? '입금' : (isFixed ? '고정지출' : '변동지출'));
              const tagStyle = isFixed
                ? 'padding:1px 6px; background:#ede7f6; color:#512da8; border-radius:4px; font-size:10px; font-weight:700;'
                : (isTransfer ? 'padding:1px 5px; background:#f2f4f6; color:#4e5968; border-radius:4px; font-size:10px; font-weight:600;'
                : (isIncome ? 'padding:1px 5px; background:rgba(2,132,199,0.12); color:#0284c7; border-radius:4px; font-size:10px; font-weight:600;'
                : 'padding:1px 5px; background:#f2f4f6; color:#64748b; border-radius:4px; font-size:10px; font-weight:600;'));

              return `
                <div class="finance-row" onclick="App.openEditTxModal('${it.id}')">
                  <div class="finance-row-left">
                    <div class="finance-icon ${tint}"><i data-lucide="${icon}" style="width:18px; height:18px;"></i></div>
                    <div class="finance-row-info">
                      <div class="finance-row-title">${this.escapeHtml(it.memo || it.category)}</div>
                      <div class="finance-row-sub">${this.escapeHtml(accName)}${this.escapeHtml(toAccName)} · <span style="${tagStyle}">${tagText}</span></div>
                    </div>
                  </div>
                  <div class="finance-row-right">
                    <div class="finance-row-amount" style="color:${color};">
                      ${sign}₩${(Number(it.amount) || 0).toLocaleString()}
                    </div>
                  </div>
                </div>
              `;
            }).join('')}
          </div>
        `;
      }

      detailEl.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; padding:0 4px;">
          <div style="font-size:13px; font-weight:800; color:var(--color-text-primary); display:flex; align-items:center; gap:6px;">
            <i data-lucide="calendar" style="width:14px; height:14px; color:var(--brand-primary);"></i>
            <span>${dateTitle}</span>
            <span style="font-size:11px; font-weight:600; color:var(--color-text-tertiary);">(${selectedData.txs.length}건)</span>
          </div>
          <div style="font-size:11px; font-weight:700;">
            ${selectedData.expense > 0 ? `<span style="color:#ef4444; margin-right:6px;">지출 -₩${selectedData.expense.toLocaleString()}</span>` : ''}
            ${selectedData.income > 0 ? `<span style="color:#2563eb;">수입 +₩${selectedData.income.toLocaleString()}</span>` : ''}
          </div>
        </div>
        ${detailContent}
      `;
    }

    this.scheduleIcons();
  },

  openRecurringModal(recId = null) {
    this.editingRecurringId = recId;
    const titleEl = document.getElementById('recurringModalTitle');
    const sources = this.getPayableSourceList();
    const accSelect = document.getElementById('recExpenseAccount');
    if (accSelect) {
      accSelect.innerHTML = sources.map(s => `<option value="${this.escapeHtml(s.id)}">${this.escapeHtml(s.name)}</option>`).join('');
    }

    const currentYm = this.data.selectedYearMonth || getCurrentYearMonthStr();

    if (recId) {
      const rec = (this.data.recurringExpenses || []).find(r => r.id === recId);
      if (!rec) return;
      if (titleEl) titleEl.textContent = '정기 고정 지출 수정';
      document.getElementById('recExpenseName').value = rec.name || '';
      document.getElementById('recExpenseAmount').value = rec.amount || '';
      document.getElementById('recExpenseDay').value = rec.dayOfMonth || 25;
      document.getElementById('recExpenseCategory').value = rec.category || '주거/통신';
      if (accSelect) accSelect.value = rec.accountId || (sources[0] && sources[0].id);
      document.getElementById('recExpenseStart').value = rec.startDate || currentYm;
      const endInput = document.getElementById('recExpenseEnd');
      const noEndCheck = document.getElementById('recExpenseNoEnd');
      if (rec.endDate) {
        endInput.value = rec.endDate;
        endInput.disabled = false;
        noEndCheck.checked = false;
      } else {
        endInput.value = '';
        endInput.disabled = true;
        noEndCheck.checked = true;
      }

      const btnBox = document.getElementById('recurringModalActionButtons');
      if (btnBox) {
        btnBox.innerHTML = `
          <button onclick="App.deleteRecurringExpense('${rec.id}')" style="flex:1; padding:12px; border-radius:14px; border:1px solid var(--color-loss-soft); background:var(--color-loss-soft); font-weight:700; font-size:14px; color:var(--color-loss); cursor:pointer;">삭제</button>
          <button onclick="App.closeRecurringModal()" style="flex:1; padding:12px; border-radius:14px; border:1px solid var(--color-border); background:var(--color-surface); font-weight:700; font-size:14px; color:var(--color-text-secondary); cursor:pointer;">취소</button>
          <button onclick="App.saveRecurringExpense()" style="flex:2; padding:12px; border-radius:14px; border:none; background:var(--brand-primary); font-weight:700; font-size:14px; color:var(--btn-primary-text); cursor:pointer;">수정 완료</button>
        `;
      }
    } else {
      if (titleEl) titleEl.textContent = '정기 고정 지출 등록';
      document.getElementById('recExpenseName').value = '';
      document.getElementById('recExpenseAmount').value = '';
      document.getElementById('recExpenseDay').value = 25;
      document.getElementById('recExpenseCategory').value = '주거/통신';
      document.getElementById('recExpenseStart').value = currentYm;
      const endInput = document.getElementById('recExpenseEnd');
      endInput.value = '';
      endInput.disabled = true;
      document.getElementById('recExpenseNoEnd').checked = true;

      const btnBox = document.getElementById('recurringModalActionButtons');
      if (btnBox) {
        btnBox.innerHTML = `
          <button onclick="App.closeRecurringModal()" style="flex:1; padding:12px; border-radius:14px; border:1px solid var(--color-border); background:var(--color-surface); font-weight:700; font-size:14px; color:var(--color-text-secondary); cursor:pointer;">취소</button>
          <button onclick="App.saveRecurringExpense()" style="flex:2; padding:12px; border-radius:14px; border:none; background:var(--brand-primary); font-weight:700; font-size:14px; color:var(--btn-primary-text); cursor:pointer;">저장하기</button>
        `;
      }
    }

    const modal = document.getElementById('tossRecurringModal');
    if (modal) modal.classList.remove('hidden');
    this.scheduleIcons();
  },

  closeRecurringModal() {
    this.editingRecurringId = null;
    const modal = document.getElementById('tossRecurringModal');
    if (modal) modal.classList.add('hidden');
  },

  saveRecurringExpense() {
    const name = document.getElementById('recExpenseName')?.value.trim();
    const amount = Number(document.getElementById('recExpenseAmount')?.value);
    const dayOfMonth = Math.max(1, Math.min(31, Number(document.getElementById('recExpenseDay')?.value) || 25));
    const category = document.getElementById('recExpenseCategory')?.value || '주거/통신';
    const accountId = document.getElementById('recExpenseAccount')?.value;
    const startDate = document.getElementById('recExpenseStart')?.value || (this.data.selectedYearMonth || getCurrentYearMonthStr());
    const noEnd = document.getElementById('recExpenseNoEnd')?.checked;
    const endDate = noEnd ? null : (document.getElementById('recExpenseEnd')?.value || null);

    if (!name || isNaN(amount) || amount <= 0) {
      alert('고정 지출명과 올바른 금액을 입력해주세요.');
      return;
    }

    if (!Array.isArray(this.data.recurringExpenses)) this.data.recurringExpenses = [];

    if (this.editingRecurringId) {
      const rec = this.data.recurringExpenses.find(r => r.id === this.editingRecurringId);
      if (rec) {
        rec.name = name;
        rec.amount = amount;
        rec.dayOfMonth = dayOfMonth;
        rec.category = category;
        rec.accountId = accountId;
        rec.startDate = startDate;
        rec.endDate = endDate;
      }
      this.editingRecurringId = null;
    } else {
      this.data.recurringExpenses.push({
        id: 'rec-' + Date.now(),
        name,
        amount,
        dayOfMonth,
        category,
        accountId,
        startDate,
        endDate
      });
    }

    StorageService.save(this.data, true);
    this.closeRecurringModal();
    this.render();
    this.showToast(`✅ [${name}] 매월 ${dayOfMonth}일 정기 고정 지출이 저장되었습니다.`);
  },

  deleteRecurringExpense(id) {
    if (!confirm('이 정기 고정 지출 항목을 삭제하시겠습니까?')) return;
    this.data.recurringExpenses = (this.data.recurringExpenses || []).filter(r => r.id !== id);
    StorageService.save(this.data, true);
    this.closeRecurringModal();
    this.render();
    this.showToast('🗑️ 정기 고정 지출 항목이 삭제되었습니다.');
  },

  applyRecurringExpensesToMonth() {
    const ym = this.data.selectedYearMonth || getCurrentYearMonthStr();
    const activeList = (this.data.recurringExpenses || []).filter(r => {
      const start = r.startDate || '2000-01';
      const end = r.endDate;
      if (start > ym) return false;
      if (end && end < ym) return false;
      return true;
    });

    // 미래 달은 아직 결제 전이므로 반영하지 않음 (잔액이 미리 차감되어 순자산이 어긋나는 것 방지)
    if (ym > getCurrentYearMonthStr()) {
      this.showToast('ℹ️ 아직 오지 않은 달은 일괄 반영할 수 없습니다.');
      return;
    }

    const monthTxs = (this.data.transactions || []).filter(t => t.date && t.date.startsWith(ym));
    const [yStr, mStr] = ym.split('-');
    const daysInMonth = new Date(Number(yStr), Number(mStr), 0).getDate();
    const todayStr = getTodayStr();

    let createdCount = 0;
    let notDueCount = 0;
    activeList.forEach(r => {
      const isApplied = monthTxs.some(t => t.recurringId === r.id || (t.type === 'EXPENSE_FIXED' && t.category === r.category && Number(t.amount) === Number(r.amount) && t.memo === r.name));
      if (!isApplied) {
        const day = Math.min(daysInMonth, r.dayOfMonth || 25);
        const txDate = `${ym}-${String(day).padStart(2, '0')}`;
        if (txDate > todayStr) { notDueCount++; return; }  // 이번 달이라도 결제일 전이면 건너뜀
        const newTx = {
          id: 't-rec-' + Date.now() + '-' + Math.random().toString(36).substr(2, 4),
          recurringId: r.id,
          date: txDate,
          type: 'EXPENSE_FIXED',
          category: r.category,
          accountId: r.accountId,
          toAccountId: null,
          amount: Number(r.amount),
          memo: r.name
        };
        this.data.transactions.unshift(newTx);
        this.applyAccountOrAssetDiff(r.accountId, -Number(r.amount));
        createdCount++;
      }
    });

    const dueNote = notDueCount > 0 ? ` (결제일 전 ${notDueCount}건은 제외)` : '';
    if (createdCount > 0) {
      StorageService.save(this.data, true);
      this.render();
      this.showToast(`✅ ${createdCount}건의 고정 지출이 반영되었습니다.${dueNote}`);
    } else if (notDueCount > 0) {
      this.showToast(`ℹ️ 결제일이 아직 오지 않은 ${notDueCount}건은 결제일 이후에 반영할 수 있습니다.`);
    } else {
      this.showToast('ℹ️ 이미 이번 달 고정 지출이 모두 반영되어 있습니다.');
    }
  },

  // [자산관리 탭] 순자산 추이 콤보 차트
  renderAssetsTrendChart(fin) {
    const canvas = document.getElementById('assetsTrendComboCanvas');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    if (this.charts.assetsCombo) {
      this.charts.assetsCombo.destroy();
    }

    const period = this.assetsTrendPeriod || 'day';
    const snapshots = [];
    const labels = [];
    const lineData = [];
    const barData = [];
    const today = new Date();

    if (period === 'day') {
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

      const d7 = new Date(today);
      d7.setDate(today.getDate() - 7);
      const prevDateStr = `${d7.getFullYear()}-${String(d7.getMonth() + 1).padStart(2, '0')}-${String(d7.getDate()).padStart(2, '0')}`;
      const d7Snap = this.calculateAssetSnapshotOnDate(prevDateStr);

      snapshots.forEach((s, idx) => {
        labels.push(s.label);
        lineData.push(s.netWorth);
        let change = 0;
        if (idx === 0) {
          change = (d7Snap && d7Snap.hasHistory && s.hasHistory) ? (s.netWorth - d7Snap.netWorth) : 0;
        } else {
          const prev = snapshots[idx - 1];
          change = (prev && prev.hasHistory && s.hasHistory) ? (s.netWorth - prev.netWorth) : 0;
        }
        s.change = change;
        barData.push(change);
      });
    } else if (period === 'month') {
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

      const dPrev = new Date(today.getFullYear(), today.getMonth() - 6, 1);
      const lastDayPrev = new Date(dPrev.getFullYear(), dPrev.getMonth() + 1, 0).getDate();
      const prevMonthStr = `${dPrev.getFullYear()}-${String(dPrev.getMonth() + 1).padStart(2, '0')}-${String(lastDayPrev).padStart(2, '0')}`;
      const prevSnap = this.calculateAssetSnapshotOnDate(prevMonthStr);

      snapshots.forEach((s, idx) => {
        labels.push(s.label);
        lineData.push(s.netWorth);
        let change = 0;
        if (idx === 0) {
          change = (prevSnap && prevSnap.hasHistory && s.hasHistory) ? (s.netWorth - prevSnap.netWorth) : 0;
        } else {
          const prev = snapshots[idx - 1];
          change = (prev && prev.hasHistory && s.hasHistory) ? (s.netWorth - prev.netWorth) : 0;
        }
        s.change = change;
        barData.push(change);
      });
    } else {
      const currentYear = today.getFullYear();
      for (let i = 4; i >= 0; i--) {
        const y = currentYear - i;
        const dateStr = (i === 0) ? getTodayStr() : `${y}-12-31`;
        const label = (i === 0) ? '올해' : `${String(y).slice(-2)}년`;

        const snap = this.calculateAssetSnapshotOnDate(dateStr);
        snapshots.push({ ...snap, year: y, label, isCurrent: i === 0 });
      }

      const prevYearStr = `${currentYear - 5}-12-31`;
      const prevSnap = this.calculateAssetSnapshotOnDate(prevYearStr);

      snapshots.forEach((s, idx) => {
        labels.push(s.label);
        lineData.push(s.netWorth);
        let change = 0;
        if (idx === 0) {
          change = (prevSnap && prevSnap.hasHistory && s.hasHistory) ? (s.netWorth - prevSnap.netWorth) : 0;
        } else {
          const prev = snapshots[idx - 1];
          change = (prev && prev.hasHistory && s.hasHistory) ? (s.netWorth - prev.netWorth) : 0;
        }
        s.change = change;
        barData.push(change);
      });
    }

    const lastSnap = snapshots[snapshots.length - 1];
    const prevSnapItem = snapshots.length >= 2 ? snapshots[snapshots.length - 2] : null;
    const latestChange = lastSnap ? lastSnap.change : 0;
    const prevNet = prevSnapItem ? prevSnapItem.netWorth : 0;

    const changeEl = document.getElementById('assetsNetWorthChangeNum');
    if (changeEl) {
      const periodLabel = period === 'day' ? '전일 대비' : (period === 'month' ? '전월 대비' : '전년 대비');
      if (!prevSnapItem || !prevSnapItem.hasHistory || prevNet === 0) {
        if (latestChange === 0) {
          changeEl.textContent = `${periodLabel} 변동 없음`;
          changeEl.style.color = '#64748b';
        } else {
          const sign = latestChange > 0 ? '+' : '';
          changeEl.textContent = `${periodLabel} ${sign}${this.formatKoreanAmount(latestChange)}`;
          changeEl.style.color = latestChange > 0 ? '#f04452' : '#3182f6';
        }
      } else {
        const changeRate = (latestChange / Math.abs(prevNet)) * 100;
        const sign = latestChange > 0 ? '+' : '';
        changeEl.textContent = `${periodLabel} ${sign}${this.formatKoreanAmount(latestChange)} (${sign}${changeRate.toFixed(1)}%)`;
        changeEl.style.color = latestChange > 0 ? '#f04452' : (latestChange < 0 ? '#3182f6' : '#64748b');
      }
    }

    const lastIdx = labels.length - 1;
    const pointBgColors = labels.map((_, idx) => idx === lastIdx ? '#003049' : '#669bbc');
    const pointBorderColors = labels.map((_, idx) => idx === lastIdx ? '#ffffff' : '#e2e8f0');
    const pointBorderWidths = labels.map((_, idx) => idx === lastIdx ? 2.5 : 1);
    const pointRadii = labels.map((_, idx) => idx === lastIdx ? 5 : 3);
    const barColors = barData.map((val, idx) => {
      if (idx === lastIdx) {
        return val > 0 ? '#f04452' : (val < 0 ? '#3182f6' : '#cbd5e1');
      }
      return val > 0 ? 'rgba(240, 68, 82, 0.55)' : (val < 0 ? 'rgba(49, 130, 246, 0.55)' : 'rgba(203, 213, 225, 0.6)');
    });

    canvas.style.cursor = 'pointer';

    this.charts.assetsCombo = new Chart(ctx, {
      data: {
        labels,
        datasets: [
          {
            type: 'line',
            label: '순자산',
            data: lineData,
            borderColor: '#003049',
            borderWidth: 2.2,
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
            borderRadius: 5,
            borderSkipped: false,
            barThickness: 20,
            yAxisID: 'yBar',
            order: 2
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        onClick: (evt, elements) => {
          if (elements && elements.length > 0) {
            const idx = elements[0].index;
            const snap = snapshots[idx];
            if (snap) this.openTrendDetailModal(period, snap);
          }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: '#0f172a',
            padding: 10,
            titleFont: { size: 12, weight: '700' },
            bodyFont: { size: 12 },
            cornerRadius: 8,
            callbacks: {
              title: (items) => `${labels[items[0].dataIndex]} 순자산 현황`,
              label: (c) => {
                if (c.dataset.type === 'line') {
                  return ` 순자산: ₩${Number(c.raw).toLocaleString()}`;
                } else {
                  const val = Number(c.raw);
                  const s = val > 0 ? '+' : '';
                  return ` 변동폭: ${s}₩${val.toLocaleString()}`;
                }
              }
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { color: '#64748b', font: { size: 11, weight: '600' } }
          },
          yLine: { position: 'left', display: false, grace: '10%' },
          yBar: { position: 'right', display: false, grace: '20%' }
        }
      }
    });
  },

  // 3. [자산관리 탭]
  renderAssetsTab(fin) {
    const totAssetEl = document.getElementById('assetsTotalAssetNum');
    if (totAssetEl) totAssetEl.textContent = this.formatKoreanAmount(fin.totalAssets);

    const totDebtEl = document.getElementById('assetsTotalDebtNum');
    if (totDebtEl) totDebtEl.textContent = this.formatKoreanAmount(fin.totalDebt);

    const netEl = document.getElementById('assetsNetWorthNum');
    if (netEl) {
      const netWorthVal = fin.netWorth || 0;
      const numStr = `${netWorthVal.toLocaleString()}원`;
      const korAmount = this.formatKoreanAmount(netWorthVal);
      netEl.innerHTML = `
        <div style="font-size:32px; font-weight:900; color:var(--color-text-primary); letter-spacing:-0.03em; line-height:1.2; word-break:keep-all;">
          ${numStr}
        </div>
        <div style="font-size:13px; font-weight:600; color:var(--color-text-secondary); margin-top:2px; margin-bottom:6px; letter-spacing:-0.01em;">
          약 ${korAmount}
        </div>
      `;
    }

    this.renderAssetsTrendChart(fin);
    this.renderStockRefreshStatus();

    const accounts = this.data.accounts || [];
    const stocks = (this.data.assets || []).filter(a => a.category === 'STOCK');
    const savings = (this.data.assets || []).filter(a => a.category === 'SAVINGS');
    const pensions = (this.data.assets || []).filter(a => a.category === 'PENSION');
    const others = (this.data.assets || []).filter(a => a.category === 'OTHER' || a.category === 'REAL_ESTATE');
    const debts = (this.data.assets || []).filter(a => a.category === 'DEBT');

    const accCard = document.getElementById('card-asset-accounts');
    if (accCard) accCard.style.display = accounts.length > 0 ? 'block' : 'none';

    const stockCard = document.getElementById('card-asset-stocks');
    if (stockCard) stockCard.style.display = stocks.length > 0 ? 'block' : 'none';

    const savingsCard = document.getElementById('card-asset-savings');
    if (savingsCard) savingsCard.style.display = savings.length > 0 ? 'block' : 'none';

    const pensionCard = document.getElementById('card-asset-pension');
    if (pensionCard) pensionCard.style.display = pensions.length > 0 ? 'block' : 'none';

    const otherCard = document.getElementById('card-asset-other');
    if (otherCard) otherCard.style.display = others.length > 0 ? 'block' : 'none';

    const debtCard = document.getElementById('card-asset-debt');
    if (debtCard) debtCard.style.display = debts.length > 0 ? 'block' : 'none';

    const totalCount = accounts.length + stocks.length + savings.length + pensions.length + others.length + debts.length;
    const emptyCard = document.getElementById('assetsEmptyOnboardingCard');
    if (emptyCard) emptyCard.classList.toggle('hidden', totalCount > 0);

    // 1) 계좌 잔액 요약 카드 및 상세 리스트 렌더링
    const totalAccountsBal = accounts.reduce((sum, a) => sum + (Number(a.balance) || 0), 0);
    const accSumBalEl = document.getElementById('accountSummaryTotalBalance');
    if (accSumBalEl) accSumBalEl.textContent = `${totalAccountsBal.toLocaleString()}원`;

    const accSumBadgeEl = document.getElementById('accountSummaryBadge');
    if (accSumBadgeEl) accSumBadgeEl.textContent = `계좌 ${accounts.length}개`;

    const modalAccBalEl = document.getElementById('modalAccountsTotalBalance');
    if (modalAccBalEl) modalAccBalEl.textContent = `${totalAccountsBal.toLocaleString()}원`;

    const modalAccBadgeEl = document.getElementById('modalAccountsCountBadge');
    if (modalAccBadgeEl) modalAccBadgeEl.textContent = `${accounts.length}개`;

    const accSubTextEl = document.getElementById('accountSummarySubText');
    if (accSubTextEl) {
      if (accounts.length === 0) {
        accSubTextEl.textContent = '등록된 계좌 0건';
      } else if (accounts.length === 1) {
        accSubTextEl.textContent = `${accounts[0].name} ${(Number(accounts[0].balance) || 0).toLocaleString()}원`;
      } else {
        accSubTextEl.textContent = `${accounts[0].name} 외 ${accounts.length - 1}개 계좌`;
      }
    }

    const accListEl = document.getElementById('assetsAccountList');
    if (accListEl && accounts.length > 0) {
      accListEl.innerHTML = `
        <div class="finance-list-card">
          ${accounts.map(a => {
            const accNumHtml = a.accountNumber ? ` · <span style="font-family:var(--font-toss); font-size:11px; color:var(--color-text-tertiary);">${this.escapeHtml(a.accountNumber)}</span>` : '';
            return `
              <div class="finance-row" onclick="App.openAccountHistoryModal('account', '${a.id}')">
                <div class="finance-row-left">
                  ${this.getFinancialBrandBadge(a.name, a.institution, 'landmark')}
                  <div class="finance-row-info">
                    <div class="finance-row-title">${this.escapeHtml(a.name)}</div>
                    <div class="finance-row-sub">${this.escapeHtml(a.institution || '은행')}${accNumHtml} · <span style="font-weight:600;">${this.escapeHtml(this.getCategoryKoreanLabel(a.type || 'CHECKING'))}</span></div>
                  </div>
                </div>
                <div class="finance-row-right">
                  <div class="finance-row-amount" style="color:${Number(a.balance) < 0 ? 'var(--color-loss)' : 'var(--color-text-primary)'};">
                    ${(Number(a.balance) || 0).toLocaleString()}원
                  </div>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      `;
    }

    // 2) 투자 주식 리스트
    const stockListEl = document.getElementById('assetsStockList');
    if (stockListEl) {
      const activeFilter = this.currentStockAccountFilter || 'ALL';
      const filteredStocks = (stocks || []).filter(s => {
        if (activeFilter === 'ALL') return true;
        const acct = s.accountType || 'GENERAL';
        return acct === activeFilter;
      });

      const usdRate = Number(this.data.usdExchangeRate) || 1380;
      let totalStockCostKrw = 0;
      let totalStockEvalKrw = 0;

      filteredStocks.forEach(s => {
        const qty = Number(s.quantity) || 1;
        const curPrice = Number(s.currentPrice) || Number(s.buyPrice) || 0;
        const buyPrice = Number(s.buyPrice) || curPrice;
        const isUsd = s.currency === 'USD';
        const rate = isUsd ? usdRate : 1;
        const evalKrw = Math.round(qty * curPrice * rate);
        const costKrw = Math.round(qty * buyPrice * rate);
        totalStockCostKrw += costKrw;
        totalStockEvalKrw += evalKrw;
      });

      const totalStockProfitKrw = totalStockEvalKrw - totalStockCostKrw;
      const totalStockProfitRate = totalStockCostKrw > 0
        ? ((totalStockProfitKrw / totalStockCostKrw) * 100).toFixed(2)
        : '0.00';
      const isProfit = totalStockProfitKrw > 0;
      const isLoss = totalStockProfitKrw < 0;
      const profitColor = isProfit ? 'var(--color-gain)' : (isLoss ? 'var(--color-loss)' : 'var(--color-text-tertiary)');
      const badgeBg = isProfit ? 'rgba(240, 68, 82, 0.08)' : (isLoss ? 'rgba(37, 99, 235, 0.08)' : 'var(--color-surface-hover)');
      const caret = isProfit ? '▲' : (isLoss ? '▼' : '');
      const sign = isProfit ? '+' : '';

      const evalEl = document.getElementById('stockSummaryEvalKrw');
      if (evalEl) evalEl.textContent = `${totalStockEvalKrw.toLocaleString()}원`;

      const rateBadgeEl = document.getElementById('stockSummaryProfitRateBadge');
      if (rateBadgeEl) {
        rateBadgeEl.style.color = profitColor;
        rateBadgeEl.style.background = badgeBg;
        rateBadgeEl.textContent = `${caret} ${sign}${totalStockProfitRate}%`;
      }

      const profitWonEl = document.getElementById('stockSummaryProfitWonText');
      if (profitWonEl) {
        profitWonEl.style.color = profitColor;
        profitWonEl.textContent = `${sign}${totalStockProfitKrw.toLocaleString()}원`;
      }

      const costTextEl = document.getElementById('stockSummaryCostKrwText');
      if (costTextEl) {
        costTextEl.textContent = `원금 ${totalStockCostKrw.toLocaleString()}원`;
      }

      const countTextEl = document.getElementById('stockSummaryCountText');
      if (countTextEl) {
        countTextEl.textContent = `종목 ${filteredStocks.length}개`;
      }

      // 모달 내부 요약 박스 동기화
      const modalEvalEl = document.getElementById('modalStockEvalKrw');
      if (modalEvalEl) modalEvalEl.textContent = `${totalStockEvalKrw.toLocaleString()}원`;

      const modalCostEl = document.getElementById('modalStockCostKrw');
      if (modalCostEl) modalCostEl.textContent = `${totalStockCostKrw.toLocaleString()}원`;

      const modalBadgeEl = document.getElementById('modalStockProfitBadge');
      if (modalBadgeEl) {
        modalBadgeEl.style.color = profitColor;
        modalBadgeEl.innerHTML = `<span>${caret} ${sign}${totalStockProfitKrw.toLocaleString()}원 (${sign}${totalStockProfitRate}%)</span>`;
      }

      if (filteredStocks.length === 0) {
        stockListEl.innerHTML = `
          <div style="padding:22px 16px; text-align:center; color:var(--color-text-tertiary); font-size:13px; background:var(--color-surface-subtle); border-radius:14px; border:1px dashed var(--color-border); margin-top:8px;">
            ${activeFilter === 'ALL' ? '등록된 주식 자산이 없습니다.' : `선택한 [${this.getAccountTypeName(activeFilter)}] 계좌에 등록된 주식이 없습니다.`}
          </div>
        `;
      } else {
        const usdRate = Number(this.data.usdExchangeRate) || 1380;
        stockListEl.innerHTML = `
          <div class="finance-list-card">
            ${filteredStocks.map(s => {
              const qty = Number(s.quantity) || 1;
              const curPrice = Number(s.currentPrice) || Number(s.buyPrice) || 0;
              const buyPrice = Number(s.buyPrice) || curPrice;
              const isUsd = s.currency === 'USD';
              const rate = isUsd ? usdRate : 1;
              const evalKrw = Math.round(qty * curPrice * rate);
              const costKrw = Math.round(qty * buyPrice * rate);
              const profitKrw = evalKrw - costKrw;
              const profitRate = costKrw > 0 ? ((profitKrw / costKrw) * 100).toFixed(2) : '0.00';
              const isProfit = profitKrw > 0;
              const isLoss = profitKrw < 0;
              const profitColor = isProfit ? 'var(--color-gain)' : (isLoss ? 'var(--color-loss)' : 'var(--color-text-tertiary)');
              const caret = isProfit ? '▲' : (isLoss ? '▼' : '-');
              const acctBadge = this.getAccountTypeBadge(s.accountType);

              return `
                <div class="finance-row" onclick="App.openAccountHistoryModal('asset', '${s.id}')">
                  <div class="finance-row-left">
                    ${this.getFinancialBrandBadge(s.name, s.ticker, 'trending-up')}
                    <div class="finance-row-info">
                      <div class="finance-row-title" style="display:flex; align-items:center; gap:5px; flex-wrap:wrap;">
                        <span>${this.escapeHtml(s.name)}</span>
                        ${acctBadge}
                      </div>
                      <div class="finance-row-sub">${this.escapeHtml(s.ticker || 'STOCK')} · ${qty}주 · 배당 ${s.dividendYield || 0}%</div>
                    </div>
                  </div>
                  <div class="finance-row-right">
                    <div class="finance-row-amount">${evalKrw.toLocaleString()}원</div>
                    <div class="finance-row-badge" style="color:${profitColor}; font-weight:700;">
                      ${caret} ${Math.abs(profitRate)}% (${this.formatKoreanAmount(profitKrw)})
                    </div>
                  </div>
                </div>
              `;
            }).join('')}
          </div>
        `;
      }
      this.renderDividendBarChart();
    }

    // 3) 예적금
    const savingsEl = document.getElementById('assetsSavingsList');
    if (savingsEl && savings.length > 0) {
      savingsEl.innerHTML = `
        <div class="finance-list-card">
          ${savings.map(s => {
            let subText = '';
            if (s.date) subText += `${this.escapeHtml(s.date)} 가입`;
            if (s.maturityDate) {
              const diffDays = Math.ceil((new Date(s.maturityDate) - new Date()) / (1000 * 60 * 60 * 24));
              const dDayStr = diffDays > 0 ? `D-${diffDays}` : (diffDays === 0 ? 'D-Day' : `만기+${Math.abs(diffDays)}일`);
              subText += `${subText ? ' · ' : ''}${this.escapeHtml(s.maturityDate)} 만기 (${dDayStr})`;
            }
            if (s.interestRate) {
              subText += `${subText ? ' · ' : ''}금리 연 ${s.interestRate}%`;
            }
            return `
              <div class="finance-row" onclick="App.openAccountHistoryModal('asset', '${s.id}')">
                <div class="finance-row-left">
                  ${this.getFinancialBrandBadge(s.name, s.institution, 'piggy-bank')}
                  <div class="finance-row-info">
                    <div class="finance-row-title">${this.escapeHtml(s.name)}</div>
                    <div class="finance-row-sub">${subText || '예적금'}</div>
                  </div>
                </div>
                <div class="finance-row-right">
                  <div class="finance-row-amount">${(Number(s.amount) || 0).toLocaleString()}원</div>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      `;
    }

    // 4) 연금
    const pensionEl = document.getElementById('assetsPensionList');
    if (pensionEl && pensions.length > 0) {
      pensionEl.innerHTML = `
        <div class="finance-list-card">
          ${pensions.map(p => `
            <div class="finance-row" onclick="App.openAccountHistoryModal('asset', '${p.id}')">
              <div class="finance-row-left">
                ${this.getFinancialBrandBadge(p.name, p.institution, 'shield')}
                <div class="finance-row-info">
                  <div class="finance-row-title">${this.escapeHtml(p.name)}</div>
                  <div class="finance-row-sub">${p.date ? this.escapeHtml(p.date) + ' 가입 · ' : ''}${this.escapeHtml(p.institution || '연금')}</div>
                </div>
              </div>
              <div class="finance-row-right">
                <div class="finance-row-amount">${(Number(p.amount) || 0).toLocaleString()}원</div>
              </div>
            </div>
          `).join('')}
        </div>
      `;
    }

    // 5) 부동산·기타
    const otherEl = document.getElementById('assetsOtherList');
    if (otherEl && others.length > 0) {
      otherEl.innerHTML = `
        <div class="finance-list-card">
          ${others.map(o => `
            <div class="finance-row" onclick="App.openAccountHistoryModal('asset', '${o.id}')">
              <div class="finance-row-left">
                ${this.getFinancialBrandBadge(o.name, o.institution, 'building-2')}
                <div class="finance-row-info">
                  <div class="finance-row-title">${this.escapeHtml(o.name)}</div>
                  <div class="finance-row-sub">${o.date ? this.escapeHtml(o.date) + ' 취득 · ' : ''}${this.escapeHtml(o.institution || '실물/기타')}</div>
                </div>
              </div>
              <div class="finance-row-right">
                <div class="finance-row-amount">${(Number(o.amount) || 0).toLocaleString()}원</div>
              </div>
            </div>
          `).join('')}
        </div>
      `;
    }

    // 6) 대출·부채
    const debtEl = document.getElementById('assetsDebtList');
    if (debtEl && debts.length > 0) {
      debtEl.innerHTML = `
        <div class="finance-list-card">
          ${debts.map(d => `
            <div class="finance-row" onclick="App.openAccountHistoryModal('asset', '${d.id}')">
              <div class="finance-row-left">
                ${this.getFinancialBrandBadge(d.name, d.institution, 'credit-card')}
                <div class="finance-row-info">
                  <div class="finance-row-title" style="color:var(--color-loss);">${this.escapeHtml(d.name)}</div>
                  <div class="finance-row-sub">${d.date ? this.escapeHtml(d.date) + ' 실행 · ' : ''}금리 연 ${d.interestRate || 0}%</div>
                </div>
              </div>
              <div class="finance-row-right">
                <div class="finance-row-amount" style="color:var(--color-loss);">-${(Number(d.amount) || 0).toLocaleString()}원</div>
              </div>
            </div>
          `).join('')}
        </div>
      `;
    }
  },

  getAccountTypeName(type) {
    switch (type) {
      case 'PENSION': return '연금저축';
      case 'IRP': return 'IRP';
      case 'ISA': return 'ISA';
      case 'GENERAL': default: return '일반';
    }
  },

  getAccountTypeBadge(type) {
    switch (type) {
      case 'PENSION':
        return `<span style="font-size:10px; font-weight:800; color:#0284c7; background:rgba(2,132,199,0.12); padding:2px 6px; border-radius:6px;">연금저축</span>`;
      case 'IRP':
        return `<span style="font-size:10px; font-weight:800; color:#7c3aed; background:rgba(124,58,237,0.12); padding:2px 6px; border-radius:6px;">IRP</span>`;
      case 'ISA':
        return `<span style="font-size:10px; font-weight:800; color:#059669; background:rgba(5,150,105,0.12); padding:2px 6px; border-radius:6px;">ISA</span>`;
      default:
        return '';
    }
  },

  filterStockAccountType(type) {
    this.currentStockAccountFilter = type;
    const filterIds = ['ALL', 'GENERAL', 'PENSION', 'IRP', 'ISA'];
    filterIds.forEach(id => {
      const btn = document.getElementById(`stockFilterBtn_${id}`);
      if (btn) {
        if (id === type) btn.classList.add('active');
        else btn.classList.remove('active');
      }
    });
    this.renderAssetsTab(this.getFinancialSummary());
  },

  renderDividendBarChart() {
    const canvas = document.getElementById('assetsDividendBarCanvas');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    if (this.charts.dividendBar) this.charts.dividendBar.destroy();

    const isTaxDeducted = this.data.dividendTaxDeduction;
    const taxRate = isTaxDeducted ? (1 - 0.154) : 1;
    const monthlyDivs = new Array(12).fill(0);

    (this.data.assets || []).filter(a => a.category === 'STOCK').forEach(s => {
      const yieldPct = Number(s.dividendYield) || 0;
      if (yieldPct <= 0) return;
      const annual = (Number(s.amount) || 0) * (yieldPct / 100) * taxRate;
      const months = (s.dividendMonths && s.dividendMonths.length > 0) ? s.dividendMonths : [3, 6, 9, 12];
      const perMo = annual / months.length;
      months.forEach(m => {
        if (m >= 1 && m <= 12) monthlyDivs[m - 1] += perMo;
      });
    });

    // 배당금 3대 핵심 지표 업데이트 (연간 총 예상 배당금, 월평균, 이번 달 수령 예정)
    const currentMonthIdx = new Date().getMonth();
    const annualTotal = monthlyDivs.reduce((a, b) => a + b, 0);
    const monthlyAvg = Math.round(annualTotal / 12);
    const currentMonthDiv = Math.round(monthlyDivs[currentMonthIdx] || 0);

    const annualEl = document.getElementById('dividendAnnualTotal');
    if (annualEl) annualEl.textContent = `₩${Math.round(annualTotal).toLocaleString()}`;
    const avgEl = document.getElementById('dividendMonthlyAvg');
    if (avgEl) avgEl.textContent = `₩${monthlyAvg.toLocaleString()}`;
    const curEl = document.getElementById('dividendCurrentMonth');
    if (curEl) curEl.textContent = `₩${currentMonthDiv.toLocaleString()}`;

    const labels = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'];
    const colors = monthlyDivs.map((_, idx) => idx === currentMonthIdx ? '#003049' : '#cbd5e1');

    this.charts.dividendBar = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets: [{ data: monthlyDivs, backgroundColor: colors, borderRadius: 8, borderSkipped: false }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: '#0f172a',
            callbacks: { label: (c) => ` 예상 배당금: ₩${Math.round(c.raw || 0).toLocaleString()}` }
          }
        },
        scales: {
          x: { grid: { display: false }, ticks: { font: { size: 10 }, color: '#64748b' } },
          y: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 10 }, color: '#64748b', callback: (v) => v >= 10000 ? `${Math.round(v / 10000)}만` : v } }
        }
      }
    });
  },

  calculateAssetSnapshotOnDate(targetDateStr) {
    const activeAccounts = [];
    let accountsTotal = 0;

    (this.data.accounts || []).forEach(acc => {
      // 해당 계좌와 연결된 가장 이른 거래일 확인 (계좌 등록일보다 과거 거래가 있는 경우 보정)
      const relatedTxs = (this.data.transactions || []).filter(t => t.accountId === acc.id || t.toAccountId === acc.id);
      let earliestDate = acc.date || '2000-01-01';
      relatedTxs.forEach(t => {
        if (t.date && t.date < earliestDate) earliestDate = t.date;
      });

      if (earliestDate > targetDateStr) return;

      let balanceOnDate = Number(acc.balance) || 0;
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
      const assetDate = ast.date || ast.purchaseDate || '2000-01-01';
      if (assetDate > targetDateStr) return;

      let val = 0;
      if (ast.category === 'STOCK' && (ast.quantity || ast.currentPrice)) {
        const rate = (ast.currency === 'USD') ? (Number(this.data.usdExchangeRate) || 1380) : 1;
        const cp = Number(ast.currentPrice) || Number(ast.buyPrice) || 0;
        const qty = Number(ast.quantity) || 0;
        val = qty > 0 && cp > 0 ? Math.round(qty * cp * rate) : (Number(ast.amount) || 0);
      } else {
        val = Number(ast.amount) || 0;
      }

      (this.data.transactions || []).forEach(t => {
        if (t.date > targetDateStr) {
          if (t.accountId === ast.id) {
            if (t.type === 'INCOME') val -= Number(t.amount);
            else if (t.type.startsWith('EXPENSE')) val += Number(t.amount);
            else if (t.type === 'TRANSFER') val += Number(t.amount);
          }
          if (t.toAccountId === ast.id && t.type === 'TRANSFER') {
            if (ast.category === 'DEBT') val += Number(t.amount);
            else val -= Number(t.amount);
          }
        }
      });

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
    const hasHistory = (activeAccounts.length > 0 || activeAssets.length > 0 || activeDebts.length > 0);

    return {
      date: targetDateStr,
      hasHistory,
      totalAssets,
      totalDebt: debtTotal,
      netWorth,
      accounts: activeAccounts,
      assets: activeAssets,
      debts: activeDebts,
      transactions: dayTxs
    };
  },

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
      changeEl.style.color = chg > 0 ? 'var(--color-gain)' : (chg < 0 ? 'var(--color-loss)' : 'var(--color-text-secondary)');
    }

    const assetListEl = document.getElementById('trendDetailAssetList');
    const assetCountEl = document.getElementById('trendDetailAssetCount');
    const allItems = [
      ...snap.accounts.map(a => ({ type: '입출금', name: a.name, detail: a.institution || '은행', amount: a.balance })),
      ...snap.assets.map(s => ({ type: this.getCategoryKoreanLabel(s.category), name: s.name, detail: s.ticker || s.institution || '', amount: s.currentVal })),
      ...snap.debts.map(d => ({ type: '대출·부채', name: d.name, detail: d.institution || '', amount: -d.currentVal, isDebt: true }))
    ];

    if (assetCountEl) assetCountEl.textContent = `(${allItems.length}건)`;

    if (assetListEl) {
      if (allItems.length === 0) {
        assetListEl.innerHTML = `<div style="text-align:center; padding:12px; font-size:12px; color:#8b95a1; background:#f9fafb; border-radius:10px;">당시 등록된 보유 자산 및 계좌가 없습니다.</div>`;
      } else {
        assetListEl.innerHTML = allItems.map(item => `
          <div style="display:flex; justify-content:space-between; align-items:center; padding:8px 10px; background:#f9fafb; border-radius:10px; border:1px solid #f2f4f6;">
            <div>
              <div style="font-size:13px; font-weight:700; color:var(--color-text-primary);">${this.escapeHtml(item.name)}</div>
              <div style="font-size:10px; color:#8b95a1;">[${this.escapeHtml(item.type)}] ${this.escapeHtml(item.detail)}</div>
            </div>
            <div style="font-size:13px; font-weight:800; font-family:var(--font-toss); color:${item.isDebt ? 'var(--color-loss)' : 'var(--color-text-primary)'};">
              ₩${item.amount.toLocaleString()}
            </div>
          </div>
        `).join('');
      }
    }

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
          const color = isInc ? 'var(--color-gain)' : (isTrf ? '#64748b' : 'var(--color-text-primary)');
          return `
            <div style="display:flex; justify-content:space-between; align-items:center; padding:8px 10px; background:#f9fafb; border-radius:10px; border:1px solid #f2f4f6;">
              <div>
                <div style="font-size:13px; font-weight:700; color:var(--color-text-primary);">${this.escapeHtml(t.memo || t.category)}</div>
                <div style="font-size:10px; color:#8b95a1;">${this.escapeHtml(t.category)} · ${this.escapeHtml(this.getCategoryKoreanLabel(t.type))}</div>
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
    this.scheduleIcons();
  },

  closeTrendDetailModal() {
    const modal = document.getElementById('tossTrendDetailModal');
    if (modal) modal.classList.add('hidden');
  },

  // 4. [분석 탭] (선택월 라벨 동기화 및 0원 예외 처리 완벽 지원)
  renderAnalyticsTab(fin) {
    const ym = this.data.selectedYearMonth || getCurrentYearMonthStr();
    const [yStr, mStr] = ym.split('-');
    const monthLabel = `${parseInt(mStr)}월`;

    const targetMonthBadge = document.getElementById('anaTargetMonthBadge');
    if (targetMonthBadge) targetMonthBadge.textContent = `${yStr}년 ${monthLabel} 기준`;

    const savingsCardTitle = document.getElementById('anaSavingsCardTitle');
    if (savingsCardTitle) savingsCardTitle.textContent = `${monthLabel} 저축 & 잉여 현금 성적표`;

    const rankingTitle = document.getElementById('anaRankingTitle');
    if (rankingTitle) rankingTitle.textContent = `${monthLabel} 카테고리별 소비 순위 TOP 4`;

    const savingsBig = document.getElementById('anaSavingsRateBig');
    if (savingsBig) savingsBig.textContent = `${fin.savingsRate.toFixed(1)}%`;

    const savingsBar = document.getElementById('anaSavingsProgressBar');
    if (savingsBar) savingsBar.style.width = `${Math.min(100, fin.savingsRate)}%`;

    const gradeBadge = document.getElementById('anaSavingsGradeBadge');
    const subtext = document.getElementById('anaSavingsSubtext');
    const noActivity = (fin.monthIncome === 0 && fin.monthExpense === 0);

    if (gradeBadge) {
      if (noActivity) {
        gradeBadge.innerHTML = `<span style="display:inline-flex; align-items:center; gap:4px;"><i data-lucide="minus-circle" style="width:12px; height:12px;"></i>거래 기록 없음</span>`;
        gradeBadge.style.background = 'var(--color-surface-hover)';
        gradeBadge.style.color = 'var(--color-text-secondary)';
        if (subtext) subtext.textContent = `(${monthLabel}에 등록된 수입·지출 내역이 없습니다)`;
      } else if (fin.savingsRate >= 50) {
        gradeBadge.innerHTML = '<span style="display:inline-flex; align-items:center; gap:4px;"><i data-lucide="sparkles" style="width:12px; height:12px;"></i>탁월한 저축 성과</span>';
        gradeBadge.style.background = 'var(--brand-secondary-soft)';
        gradeBadge.style.color = 'var(--brand-primary)';
        if (subtext) subtext.textContent = '(월간 수입의 절반 이상을 자산으로 건전하게 축적 중)';
      } else if (fin.savingsRate >= 25) {
        gradeBadge.innerHTML = '<span style="display:inline-flex; align-items:center; gap:4px;"><i data-lucide="thumbs-up" style="width:12px; height:12px;"></i>우수한 저축 습관</span>';
        gradeBadge.style.background = 'var(--color-surface-subtle)';
        gradeBadge.style.color = 'var(--brand-primary)';
        if (subtext) subtext.textContent = '(안정적인 잉여 현금 흐름을 유지하고 있습니다)';
      } else {
        gradeBadge.innerHTML = '<span style="display:inline-flex; align-items:center; gap:4px;"><i data-lucide="alert-triangle" style="width:12px; height:12px;"></i>소비 점검 권장</span>';
        gradeBadge.style.background = 'var(--color-loss-soft)';
        gradeBadge.style.color = 'var(--color-loss)';
        if (subtext) subtext.textContent = '(지출이 수입에 근접하거나 초과하여 관리가 필요합니다)';
      }
    }

    const incomeEl = document.getElementById('anaTotalIncomeText');
    if (incomeEl) incomeEl.textContent = `₩${fin.monthIncome.toLocaleString()}`;

    const surplusEl = document.getElementById('anaTotalSurplusText');
    if (surplusEl) surplusEl.textContent = `₩${fin.remainingSurplus.toLocaleString()}`;

    const fixedAmt = document.getElementById('anaFixedAmtText');
    if (fixedAmt) fixedAmt.textContent = `₩${fin.monthFixedExpense.toLocaleString()}`;

    const varAmt = document.getElementById('anaVarAmtText');
    if (varAmt) varAmt.textContent = `₩${fin.monthVariableExpense.toLocaleString()}`;

    const hasExp = fin.monthExpense > 0;
    const fixedPct = hasExp ? Math.round((fin.monthFixedExpense / fin.monthExpense) * 100) : 0;
    const varPct = hasExp ? Math.round((fin.monthVariableExpense / fin.monthExpense) * 100) : 0;

    const fixedPctEl = document.getElementById('anaFixedPctText');
    if (fixedPctEl) fixedPctEl.textContent = `전체 지출의 ${fixedPct}%`;

    const varPctEl = document.getElementById('anaVarPctText');
    if (varPctEl) varPctEl.textContent = `전체 지출의 ${varPct}%`;

    const fixedBar = document.getElementById('anaFixedBar');
    if (fixedBar) fixedBar.style.width = `${fixedPct}%`;

    const varBar = document.getElementById('anaVarBar');
    if (varBar) varBar.style.width = `${varPct}%`;

    const rankList = document.getElementById('anaCategoryRankingList');
    if (rankList) {
      const catTotals = {};
      fin.monthTxs.filter(t => t.type && t.type.startsWith('EXPENSE')).forEach(t => {
        catTotals[t.category] = (catTotals[t.category] || 0) + Number(t.amount);
      });

      const sorted = Object.entries(catTotals).sort((a, b) => b[1] - a[1]).slice(0, 4);
      if (sorted.length === 0) {
        rankList.innerHTML = `<div style="text-align:center; padding:16px; font-size:12px; color:var(--color-text-tertiary);">${monthLabel} 지출 내역이 없습니다.</div>`;
      } else {
        const rankColors = ['#780000', '#c1121f', '#003049', '#0284c7'];
        rankList.innerHTML = sorted.map(([cat, amt], idx) => {
          const pct = Math.round((amt / (fin.monthExpense || 1)) * 100);
          const { icon } = this.getCategoryIcon(cat, 'EXPENSE');

          return `
            <div>
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
                <div style="display:flex; align-items:center; gap:6px;">
                  <span style="width:18px; height:18px; border-radius:50%; background:${rankColors[idx]}; color:#ffffff; font-size:10px; font-weight:800; display:flex; align-items:center; justify-content:center;">${idx + 1}</span>
                  <i data-lucide="${icon}" style="width:14px; height:14px; color:var(--color-text-secondary);"></i>
                  <span style="font-size:13px; font-weight:700; color:var(--color-text-primary);">${this.escapeHtml(cat)}</span>
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

    const debtRatio = fin.totalAssets > 0 ? ((fin.totalDebt / fin.totalAssets) * 100).toFixed(1) : '0.0';
    const debtRatioEl = document.getElementById('anaDebtRatioText');
    if (debtRatioEl) {
      debtRatioEl.textContent = `${debtRatio}%`;
      debtRatioEl.style.color = Number(debtRatio) > 30 ? 'var(--color-loss)' : 'var(--color-text-primary)';
    }

    // 최근 3개월 평균 지출 산출 (소액 지출 왜곡 방지를 위해 목표 예산 및 정기 고정지출 하한선 반영)
    const today = new Date();
    const curYear = today.getFullYear();
    const curMonth = today.getMonth() + 1;
    const past3Months = [];
    for (let i = 0; i < 3; i++) {
      let y = curYear;
      let m = curMonth - i;
      if (m <= 0) { m += 12; y -= 1; }
      past3Months.push(`${y}-${String(m).padStart(2, '0')}`);
    }

    const expensesPerMonth = past3Months.map(pYm => {
      return (this.data.transactions || [])
        .filter(t => t.date && t.date.startsWith(pYm) && t.type && t.type.startsWith('EXPENSE'))
        .reduce((sum, t) => sum + (Number(t.amount) || 0), 0);
    }).filter(exp => exp > 0);

    const rawAvgExpense = expensesPerMonth.length > 0
      ? (expensesPerMonth.reduce((a, b) => a + b, 0) / expensesPerMonth.length)
      : (fin.actualMonthExpense || fin.monthExpense || 0);

    const recurringMonthlySum = (this.data.recurringExpenses || []).reduce((s, r) => s + (Number(r.amount) || 0), 0);
    const baselineExpense = Math.max(rawAvgExpense, recurringMonthlySum);

    const emergencyMonths = baselineExpense > 0
      ? (fin.totalCash / baselineExpense).toFixed(1)
      : (fin.totalCash > 0 ? '지출 미기록' : '0.0');

    const emergencyEl = document.getElementById('anaEmergencyBufferText');
    if (emergencyEl) {
      emergencyEl.textContent = baselineExpense > 0 ? `${emergencyMonths}개월분` : emergencyMonths;
    }

    const adviceEl = document.getElementById('anaAdviceText');
    if (adviceEl) {
      if (noActivity) {
        adviceEl.innerHTML = `${yStr}년 <b>${monthLabel}</b>에는 아직 기록된 수입·지출 내역이 없습니다. 가계부 탭에서 거래를 등록하거나 거래가 있는 달로 이동하시면 맞춤형 소비 및 저축 코칭을 확인하실 수 있습니다.`;
      } else if (fin.savingsRate >= 50) {
        adviceEl.innerHTML = `${monthLabel} 저축률 <b>${fin.savingsRate.toFixed(1)}%</b>로 매우 건강한 현금 흐름을 유지하고 있습니다. 월 잉여금 <b>₩${fin.remainingSurplus.toLocaleString()}</b> 중 일부를 적립식 투자나 예·적금, 연금 저축 등 자산 포트폴리오 다각화에 활용해 보세요.`;
      } else if (fin.savingsRate >= 25) {
        adviceEl.innerHTML = `고정 지출 비중(<b>${fixedPct}%</b>)이 안정적으로 관리되고 있습니다. 변동 지출 상위 항목을 10%만 절감하면 월 저축액을 <b>₩${Math.round(fin.monthVariableExpense * 0.1).toLocaleString()}</b> 더 늘릴 수 있습니다.`;
      } else {
        adviceEl.innerHTML = `${monthLabel} 지출이 수입 대비 높은 편입니다(고정비 <b>${fixedPct}%</b> · 변동비 <b>${varPct}%</b>). 가계부 탭의 한 달 목표 예산을 설정하고 선택적 소비 항목을 점검해 보세요.`;
      }
    }
  },

  // 5. [설정 탭]
  renderSettingsTab() {
    const urlInput = document.getElementById('googleSheetUrlInput');
    if (urlInput) urlInput.value = this.getGoogleSheetUrl();

    const autoSyncEl = document.getElementById('autoSyncToggle');
    if (autoSyncEl) autoSyncEl.checked = !!this.autoSyncEnabled;

    const savedUrl = this.getGoogleSheetUrl();
    const msgEl = document.getElementById('googleSyncStatusMsg');
    if (msgEl) {
      if (savedUrl) {
        msgEl.style.color = 'var(--color-text-secondary)';
        msgEl.textContent = this.autoSyncEnabled
          ? '구글 시트와 실시간 자동 연동 중입니다. (데이터 수정 시 자동 저장)'
          : '구글 시트 URL이 등록되어 있으나 자동 동기화가 꺼져 있습니다.';
      } else {
        msgEl.style.color = '#8b95a1';
        msgEl.textContent = 'Web App URL을 등록하면 실시간 자동 동기화가 활성화됩니다.';
      }
    }

    if (!savedUrl) {
      this.updateSyncBadge('idle', '미연결');
    } else if (this.autoSyncEnabled) {
      this.updateSyncBadge('synced', '자동 동기화 대기');
    } else {
      this.updateSyncBadge('idle', '동기화 꺼짐');
    }
  },

  getGoogleSheetUrl() {
    return (this.data && this.data.googleSheetUrl) || localStorage.getItem('google_sheet_webapp_url') || '';
  },

  saveGoogleSheetUrl(url) {
    const trimmed = (url || '').trim();
    if (!this.data) this.data = {};
    this.data.googleSheetUrl = trimmed;
    try { localStorage.setItem('google_sheet_webapp_url', trimmed); } catch (e) {}
    StorageService.save(this.data);
    this.renderSettingsTab();
    this.showToast('💾 구글 시트 URL이 저장되었습니다.');
  },

  toggleAutoSync(enabled) {
    this.autoSyncEnabled = !!enabled;
    try { localStorage.setItem('google_sheet_autosync', String(this.autoSyncEnabled)); } catch (e) {}
    this.renderSettingsTab();
    this.showToast(this.autoSyncEnabled ? '🟢 실시간 자동 동기화가 켜졌습니다.' : '⚪ 자동 동기화가 꺼졌습니다.');
  },

  triggerAutoSync() {
    if (!this.autoSyncEnabled) return;
    const url = this.getGoogleSheetUrl();
    if (!url) return;

    this.updateSyncBadge('syncing', '동기화 중...');
    clearTimeout(this._autoSyncTimer);
    this._autoSyncTimer = setTimeout(async () => {
      try {
        await this.exportToGoogleSheets(false);
      } catch (err) {
        console.warn('자동 동기화 전송 실패:', err);
      }
    }, 1200);
  },

  animateSyncIcon(spin) {
    const headerIcon = document.getElementById('headerSyncIcon');
    if (headerIcon) headerIcon.classList.toggle('lucide-spin', !!spin);
  },

  updateSyncBadge(status, text) {
    const badge = document.getElementById('googleSyncBadge');
    if (badge) {
      if (status === 'syncing') {
        badge.style.background = '#fff8e6';
        badge.style.color = '#d97706';
        badge.innerHTML = `<span style="display:inline-flex; align-items:center; gap:4px;"><i data-lucide="loader-2" style="width:12px; height:12px;" class="lucide-spin"></i><span>${text}</span></span>`;
      } else if (status === 'synced') {
        badge.style.background = '#e8f5e9';
        badge.style.color = '#2e7d32';
        badge.innerHTML = `<span style="display:inline-flex; align-items:center; gap:4px;"><i data-lucide="check-circle-2" style="width:12px; height:12px;"></i><span>${text}</span></span>`;
      } else if (status === 'error') {
        badge.style.background = '#fff2f2';
        badge.style.color = '#dc2626';
        badge.innerHTML = `<span style="display:inline-flex; align-items:center; gap:4px;"><i data-lucide="alert-circle" style="width:12px; height:12px;"></i><span>${text}</span></span>`;
      } else {
        badge.style.background = '#f2f4f6';
        badge.style.color = '#8b95a1';
        badge.innerHTML = `<span style="display:inline-flex; align-items:center; gap:4px;"><i data-lucide="circle" style="width:10px; height:10px;"></i><span>${text}</span></span>`;
      }
    }

    this.scheduleIcons();
  },

  async testGoogleSheetsSync() {
    const url = this.getGoogleSheetUrl();
    const msgEl = document.getElementById('googleSyncStatusMsg');
    if (!url) {
      this.showToast('⚠️ Web App URL을 먼저 입력해주세요.');
      return;
    }
    if (msgEl) msgEl.textContent = '연결 상태를 확인하는 중입니다...';

    try {
      const testUrl = url + (url.includes('?') ? '&' : '?') + 'action=ping&_t=' + Date.now();
      const res = await fetch(testUrl, { method: 'GET', mode: 'cors' });
      const data = await res.json();
      if (data && data.status === 'success') {
        if (msgEl) {
          msgEl.style.color = '#2e7d32';
          msgEl.textContent = '✅ 연결 성공! 정상적으로 구글 시트와 통신할 수 있습니다.';
        }
        this.updateSyncBadge('synced', '연결 정상');
        this.showToast('✅ 구글 시트 연결 성공!');
      } else {
        throw new Error((data && data.message) || '응답 오류');
      }
    } catch (err) {
      if (msgEl) {
        msgEl.style.color = '#dc2626';
        msgEl.textContent = '❌ 연결 실패: 배포 시 "액세스 권한: 모든 사용자(Anyone)" 설정 및 URL을 확인해주세요.';
      }
      this.updateSyncBadge('error', '연결 실패');
      this.showToast('❌ 구글 시트 연결 실패');
    }
  },

  // 다중 기기(스마트폰/태블릿) 거래 내역 및 자산 스마트 병합 (합집합 보존)
  mergeCloudData(cloudData) {
    if (!cloudData) return this.data;
    const local = this.data || {};
    const localUpdatedAt = Number(local.updatedAt) || 0;
    const cloudUpdatedAt = Number(cloudData.updatedAt) || 0;

    // 1) 거래 내역 (transactions) 병합: 고유 ID 기준 누락 없이 모두 합침
    const txMap = new Map();
    (cloudData.transactions || []).forEach(t => {
      if (t && t.id) txMap.set(t.id, t);
    });
    (local.transactions || []).forEach(t => {
      if (t && t.id) {
        // 클라우드에 아직 없거나, 로컬이 더 최신이면 로컬 거래 우선
        if (!txMap.has(t.id) || localUpdatedAt > cloudUpdatedAt) {
          txMap.set(t.id, t);
        }
      }
    });
    const mergedTransactions = Array.from(txMap.values());
    mergedTransactions.sort((a, b) => (b.date || '').localeCompare(a.date || '') || String(b.id).localeCompare(String(a.id)));

    // 2) 계좌 및 자산: 최신 쪽을 기본으로 하되 누락된 항목 병합
    const isCloudPrimary = cloudUpdatedAt >= localUpdatedAt;
    const primary = isCloudPrimary ? cloudData : local;
    const secondary = isCloudPrimary ? local : cloudData;

    const accMap = new Map();
    (secondary.accounts || []).forEach(a => { if (a && a.id) accMap.set(a.id, a); });
    (primary.accounts || []).forEach(a => { if (a && a.id) accMap.set(a.id, a); });

    const astMap = new Map();
    (secondary.assets || []).forEach(a => { if (a && a.id) astMap.set(a.id, a); });
    (primary.assets || []).forEach(a => { if (a && a.id) astMap.set(a.id, a); });

    // 3) 고정지출 병합
    const recMap = new Map();
    (secondary.recurringExpenses || []).forEach(r => { if (r && r.id) recMap.set(r.id, r); });
    (primary.recurringExpenses || []).forEach(r => { if (r && r.id) recMap.set(r.id, r); });

    const merged = {
      ...primary,
      accounts: Array.from(accMap.values()),
      assets: Array.from(astMap.values()),
      recurringExpenses: Array.from(recMap.values()),
      transactions: mergedTransactions,
      updatedAt: Math.max(cloudUpdatedAt, localUpdatedAt, Date.now())
    };

    return StorageService.normalize(merged);
  },

  // 다중 기기(스마트폰/태블릿) 안전 양방향 스마트 동기화
  async syncWithGoogleSheets(silent = true) {
    if (this._isSyncing) return;
    const url = this.getGoogleSheetUrl();
    if (!url) {
      if (!silent) this.showToast('⚠️ [설정] 탭에서 구글 시트 Web App URL을 먼저 등록해주세요.');
      this.updateSyncBadge('idle', 'URL 미등록');
      return;
    }

    const now = Date.now();
    if (silent && this._lastSyncAttempt && (now - this._lastSyncAttempt < 2500)) return;
    this._lastSyncAttempt = now;

    this._isSyncing = true;
    this.animateSyncIcon(true);
    this.updateSyncBadge('syncing', '동기화 중...');

    try {
      const getUrl = url + (url.includes('?') ? '&' : '?') + 'action=get&_t=' + now;
      const res = await fetch(getUrl, { method: 'GET', mode: 'cors' });
      const result = await res.json();

      if (result && result.status === 'success' && result.data) {
        const cloudData = StorageService.normalize(result.data);
        const localTxCount = (this.data && this.data.transactions ? this.data.transactions.length : 0);
        const cloudTxCount = (cloudData.transactions ? cloudData.transactions.length : 0);

        // 스마트 병합 수행 (양쪽 거래가 모두 살아있도록 합집합 구성)
        const merged = this.mergeCloudData(cloudData);
        const mergedTxCount = merged.transactions.length;

        clearTimeout(this._autoSyncTimer);
        if (!merged.googleSheetUrl) merged.googleSheetUrl = url;
        this.data = merged;
        StorageService.save(this.data, false); // 로컬 저장 (markUpdated = false: 무한 루프 방지)
        this.render();

        const syncTime = new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' });
        this.updateSyncBadge('synced', `동기화 완료 (${syncTime})`);

        if (!silent) {
          this.showToast(`✅ 구글 시트와 동기화 완료! (거래 내역 총 ${mergedTxCount}건)`);
        } else if (mergedTxCount > localTxCount) {
          const diff = mergedTxCount - localTxCount;
          this.showToast(`☁️ 다른 기기에서 입력한 새 거래 ${diff}건이 동기화되었습니다.`);
        }

        // 만약 로컬에만 있던 거래가 병합에 포함되어 클라우드보다 더 많아졌다면 클라우드로도 반영
        if (mergedTxCount > cloudTxCount) {
          await this.exportToGoogleSheets(false);
        }
      } else if (result && result.status === 'error' && (result.message || '').includes('비어')) {
        // 시트에 아직 데이터가 전혀 없는 경우 -> 로컬 데이터가 있으면 업로드
        const hasLocal = (this.data.accounts && this.data.accounts.length > 0) ||
                         (this.data.transactions && this.data.transactions.length > 0) ||
                         (this.data.assets && this.data.assets.length > 0);
        if (hasLocal) {
          await this.exportToGoogleSheets(false);
        }
      } else {
        throw new Error((result && result.message) || '데이터 형식 오류');
      }
    } catch (err) {
      console.warn('구글 시트 자동 동기화 실패 (오프라인):', err);
      this.updateSyncBadge('idle', '동기화 대기');
      if (!silent) {
        this.showToast('⚠️ 구글 시트와 연결할 수 없습니다. URL 및 배포 상태를 확인해주세요.');
      }
    } finally {
      this._isSyncing = false;
      this.animateSyncIcon(false);
    }
  },

  async exportToGoogleSheets(isManual = false) {
    const url = this.getGoogleSheetUrl();
    const msgEl = document.getElementById('googleSyncStatusMsg');
    if (!url) {
      if (isManual) this.showToast('⚠️ Web App URL을 먼저 입력해주세요.');
      return;
    }

    const hasLocalData = (this.data.accounts && this.data.accounts.length > 0) ||
                         (this.data.transactions && this.data.transactions.length > 0) ||
                         (this.data.assets && this.data.assets.length > 0);
    if (!isManual && !hasLocalData) {
      this.updateSyncBadge('idle', '데이터 없음(동기화 보류)');
      return;
    }

    try {
      this.animateSyncIcon(true);
      this.updateSyncBadge('syncing', '동기화 중...');
      if (!this.data.updatedAt) {
        this.data.updatedAt = Date.now();
      }
      const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'syncAll', exportAction: 'export', data: this.data })
      });

      const result = await res.json();
      if (result && result.status === 'success') {
        const nowStr = new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        if (msgEl) {
          msgEl.style.color = '#2e7d32';
          msgEl.textContent = `✅ 최근 동기화: ${nowStr}`;
        }
        this.updateSyncBadge('synced', `동기화 완료 (${nowStr})`);
        if (isManual) this.showToast('✅ 구글 시트에 안전하게 백업되었습니다!');
      } else {
        throw new Error((result && result.message) || '동기화 응답 오류');
      }
    } catch (err) {
      this.updateSyncBadge('error', '동기화 오류');
      if (isManual) this.showToast('❌ 동기화 실패. 배포 설정을 확인하세요.');
      throw err;
    } finally {
      this.animateSyncIcon(false);
    }
  },

  async importFromGoogleSheets(silent = false) {
    const url = this.getGoogleSheetUrl();
    if (!url) {
      if (!silent) this.showToast('⚠️ Web App URL을 먼저 입력해주세요.');
      return;
    }
    if (!silent && !confirm('구글 시트에 백업된 데이터로 현재 데이터를 덮어쓰시겠습니까?')) return;

    try {
      this.animateSyncIcon(true);
      this.updateSyncBadge('syncing', '불러오는 중...');
      const getUrl = url + (url.includes('?') ? '&' : '?') + 'action=get&_t=' + Date.now();
      const res = await fetch(getUrl, { method: 'GET', mode: 'cors' });
      const result = await res.json();

      if (result && result.status === 'success' && result.data) {
        const imported = StorageService.normalize(result.data);
        if (!imported.googleSheetUrl) imported.googleSheetUrl = url;
        if (!imported.updatedAt) imported.updatedAt = Date.now();

        clearTimeout(this._autoSyncTimer);
        this.data = imported;
        StorageService.save(this.data, false);
        this.render();
        const nowStr = new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' });
        this.updateSyncBadge('synced', `동기화 완료 (${nowStr})`);
        if (!silent) this.showToast('✅ 구글 시트에서 데이터를 복원했습니다!');
      } else {
        throw new Error('데이터 형식 오류');
      }
    } catch (err) {
      this.updateSyncBadge('error', '가져오기 실패');
      if (!silent) this.showToast('❌ 가져오기 실패');
    } finally {
      this.animateSyncIcon(false);
    }
  },

  exportDataJson() {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(this.data, null, 2));
    const dlAnchorElem = document.createElement('a');
    dlAnchorElem.setAttribute("href", dataStr);
    dlAnchorElem.setAttribute("download", `financial_backup_${getTodayStr()}.json`);
    dlAnchorElem.click();
  },

  importDataJson(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const parsed = JSON.parse(e.target.result);
        if (parsed && (Array.isArray(parsed.transactions) || Array.isArray(parsed.accounts) || Array.isArray(parsed.assets))) {
          this.data = parsed;
          StorageService.save(this.data);
          this.render();
          this.showToast('✅ 백업 데이터를 성공적으로 복원했습니다.');
        } else {
          alert('유효하지 않은 백업 데이터 파일 형식입니다.');
        }
      } catch (err) {
        alert('백업 파일을 읽는 중 오류가 발생했습니다: ' + err.message);
      }
      event.target.value = '';
    };
    reader.readAsText(file);
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
        background: rgba(15, 23, 42, 0.94);
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

  // 거래 등록 / 수정 / 삭제 모달 로직
  applyQuickTx(type, category, amount, memo) {
    const typeSelect = document.getElementById('newTxType');
    if (typeSelect) {
      typeSelect.value = type;
      this.onTxTypeChange();
    }
    const catInput = document.getElementById('newTxCategory');
    if (catInput) catInput.value = category || '';
    const amtInput = document.getElementById('newTxAmount');
    if (amtInput) amtInput.value = (Number(amount) || 0).toLocaleString();
    const memoInput = document.getElementById('newTxMemo');
    if (memoInput) memoInput.value = memo || '';
    this.showToast(`⚡ [${memo || category}] 간편 입력 완료`);
  },

  // 금액 입력 시 실시간 자리수(천 단위 콤마) 포맷터
  onAmountInput(el) {
    if (!el) return;
    const raw = el.value.replace(/[^0-9]/g, '');
    if (!raw) {
      el.value = '';
      return;
    }
    el.value = Number(raw).toLocaleString();
  },

  // 자주 쓰는 거래 간편 입력 칩 렌더링
  renderQuickTxChips() {
    const container = document.getElementById('quickTxChipList');
    if (!container) return;

    if (!Array.isArray(this.data.quickTransactions)) {
      this.data.quickTransactions = [];
    }

    const items = this.data.quickTransactions;
    let html = '';
    if (items.length === 0) {
      html += `
        <div style="font-size:11px; color:var(--color-text-tertiary); padding:4px 0;">
          등록된 자주 쓰는 거래가 없습니다. '+ 항목 추가'로 등록해보세요.
        </div>
      `;
    } else {
      html += items.map(qt => {
        const amtNum = Number(qt.amount) || 0;
        const amtStr = amtNum.toLocaleString() + '원';
        const iconName = qt.type === 'INCOME' ? 'arrow-down-left' : (qt.type === 'TRANSFER' ? 'arrow-left-right' : 'receipt-text');
        return `
          <button type="button" class="quick-chip" onclick="App.applyQuickTx('${this.escapeHtml(qt.type || 'EXPENSE_VARIABLE')}', '${this.escapeHtml(qt.category || '')}', ${amtNum}, '${this.escapeHtml(qt.memo || '')}')" title="클릭 시 간편 입력">
            <i data-lucide="${iconName}" style="width:12px; height:12px; color:var(--brand-primary);"></i>
            <span>${this.escapeHtml(qt.memo || qt.category || '거래')} ${amtStr}</span>
            <span class="quick-chip-del" onclick="event.stopPropagation(); App.deleteQuickTx('${qt.id}')" title="이 항목 삭제">
              <i data-lucide="x" style="width:11px; height:11px;"></i>
            </span>
          </button>
        `;
      }).join('');
    }

    // 새 항목 직접 추가 버튼
    html += `
      <button type="button" class="quick-chip" onclick="App.addQuickTxPrompt()" style="border-style:dashed; color:var(--brand-primary); background:var(--brand-primary-soft);" title="새 자주 쓰는 거래 직접 등록">
        <i data-lucide="plus" style="width:12px; height:12px; color:var(--brand-primary);"></i>
        <span>+ 항목 추가</span>
      </button>
    `;

    container.innerHTML = html;
    this.scheduleIcons();
  },

  deleteQuickTx(id) {
    if (!id) return;
    if (!confirm('이 자주 쓰는 거래 항목을 삭제하시겠습니까?')) return;
    this.data.quickTransactions = (this.data.quickTransactions || []).filter(q => q.id !== id);
    StorageService.save(this.data, true);
    this.renderQuickTxChips();
    this.showToast('🗑️ 자주 쓰는 거래 항목이 삭제되었습니다.');
  },

  saveCurrentAsQuickTx() {
    const type = document.getElementById('newTxType')?.value || 'EXPENSE_VARIABLE';
    const category = document.getElementById('newTxCategory')?.value.trim();
    const rawAmt = String(document.getElementById('newTxAmount')?.value || '').replace(/,/g, '').trim();
    const amount = Number(rawAmt);
    const memo = document.getElementById('newTxMemo')?.value.trim() || category || '거래';

    if (!category || isNaN(amount) || amount <= 0) {
      alert('자주 쓰는 거래로 저장하려면 먼저 분류와 금액을 입력해주세요.');
      return;
    }

    if (!Array.isArray(this.data.quickTransactions)) {
      this.data.quickTransactions = [];
    }

    const newItem = {
      id: 'qt-' + Date.now(),
      type,
      category,
      amount,
      memo
    };

    this.data.quickTransactions.push(newItem);
    StorageService.save(this.data, true);
    this.renderQuickTxChips();
    this.showToast(`✅ [${memo} ${amount.toLocaleString()}원] 자주 쓰는 거래로 등록되었습니다.`);
  },

  addQuickTxPrompt() {
    const memo = prompt('자주 쓰는 거래 이름(사용처/메모)을 입력하세요:\n(예: 스타벅스 아메리카노, 구내식당 점심)');
    if (!memo || !memo.trim()) return;

    const amtStr = prompt(`[${memo.trim()}]의 기본 금액(원)을 입력하세요:\n(예: 4500, 10000)`);
    if (!amtStr) return;
    const amount = Number(amtStr.replace(/[^0-9]/g, ''));
    if (isNaN(amount) || amount <= 0) {
      alert('올바른 금액 숫자를 입력해주세요.');
      return;
    }

    const category = prompt(`[${memo.trim()}]의 분류(카테고리)를 입력하세요:\n(예: 카페/간식, 식비/외식, 마트/잡화)`, '카페/간식') || '지출';

    if (!Array.isArray(this.data.quickTransactions)) {
      this.data.quickTransactions = [];
    }

    this.data.quickTransactions.push({
      id: 'qt-' + Date.now(),
      type: 'EXPENSE_VARIABLE',
      category: category.trim(),
      amount,
      memo: memo.trim()
    });

    StorageService.save(this.data, true);
    this.renderQuickTxChips();
    this.showToast(`✅ [${memo.trim()} ${amount.toLocaleString()}원] 추가 완료!`);
  },

  getPayableSourceList() {
    const list = [];
    (this.data.accounts || []).forEach(acc => {
      list.push({
        id: acc.id,
        name: `[계좌] ${acc.name} (잔액: ₩${(Number(acc.balance) || 0).toLocaleString()})`,
        rawName: acc.name,
        type: 'account',
        balance: Number(acc.balance) || 0
      });
    });

    (this.data.assets || []).forEach(ast => {
      if (ast.category === 'SAVINGS' || ast.category === 'OTHER') {
        const catLabel = ast.category === 'SAVINGS' ? '예적금' : '기타자산';
        list.push({
          id: ast.id,
          name: `[${catLabel}] ${ast.name} (₩${(Number(ast.amount) || 0).toLocaleString()})`,
          rawName: ast.name,
          type: 'asset',
          balance: Number(ast.amount) || 0
        });
      }
    });

    if (list.length === 0) {
      if (!Array.isArray(this.data.accounts)) this.data.accounts = [];
      const defaultAcc = {
        id: 'acc-default',
        name: '기본 현금 계좌',
        institution: '현금',
        type: 'CHECKING',
        balance: 0,
        date: '2026-01-01',
        isDefault: true
      };
      this.data.accounts.push(defaultAcc);
      list.push({
        id: defaultAcc.id,
        name: `[계좌] ${defaultAcc.name} (잔액: ₩0)`,
        rawName: defaultAcc.name,
        type: 'account',
        balance: 0
      });
    }

    return list;
  },

  // 이체 입금처 목록 (수량×단가로 덮어써지는 개별 주식 STOCK은 제외하여 이체 금액 증발 방지)
  getTransferDestinationList() {
    const list = [];
    (this.data.accounts || []).forEach(acc => {
      list.push({
        id: acc.id,
        name: `[계좌] ${acc.name} (잔액: ₩${(Number(acc.balance) || 0).toLocaleString()})`,
        rawName: acc.name,
        type: 'account',
        category: 'CHECKING'
      });
    });

    const catLabels = {
      SAVINGS: '예적금',
      PENSION: '연금',
      OTHER: '기타자산',
      DEBT: '대출/부채(상환)'
    };

    (this.data.assets || []).forEach(ast => {
      if (ast.category === 'STOCK') return; // 주식 종목은 매수/수량 수정으로 관리
      const catLabel = catLabels[ast.category] || '자산';
      list.push({
        id: ast.id,
        name: `[${catLabel}] ${ast.name} (₩${(Number(ast.amount) || 0).toLocaleString()})`,
        rawName: ast.name,
        type: 'asset',
        category: ast.category
      });
    });

    return list;
  },

  reconcileTransactionsAndBalances() {
    if (!Array.isArray(this.data.transactions) || this.data.transactions.length === 0) return;
    const sources = this.getPayableSourceList();
    if (sources.length === 0) return;
    const defaultSourceId = sources[0].id;

    let changed = false;
    this.data.transactions.forEach(t => {
      const accExists = (this.data.accounts || []).some(a => a.id === t.accountId);
      const astExists = (this.data.assets || []).some(a => a.id === t.accountId);
      if (!t.accountId || (!accExists && !astExists)) {
        t.accountId = defaultSourceId;
        changed = true;
      }
    });

    if (changed) StorageService.save(this.data);
  },

  onTxTypeChange() {
    const typeSelect = document.getElementById('newTxType');
    if (!typeSelect) return;
    const type = typeSelect.value;
    const toAccRow = document.getElementById('toAccountRow');
    const fixedArea = document.getElementById('fixedExpenseConfigArea');
    const fromAccLabel = document.getElementById('newTxFromAccountLabel');

    if (fromAccLabel) {
      if (type === 'INCOME') {
        fromAccLabel.textContent = '입금 계좌';
      } else if (type === 'TRANSFER') {
        fromAccLabel.textContent = '출금 계좌';
      } else {
        fromAccLabel.textContent = '출금 / 결제 계좌';
      }
    }

    if (toAccRow) toAccRow.style.display = type === 'TRANSFER' ? 'block' : 'none';
    if (fixedArea) {
      if (type === 'EXPENSE_FIXED') {
        fixedArea.style.display = 'flex';
        const daySelect = document.getElementById('newTxRecurringDay');
        if (daySelect && daySelect.options.length === 0) {
          daySelect.innerHTML = Array.from({ length: 31 }, (_, i) => `<option value="${i + 1}">${i + 1}일</option>`).join('');
        }
        const dateInput = document.getElementById('newTxDate');
        if (dateInput && dateInput.value && daySelect && !daySelect.value) {
          const d = parseInt(dateInput.value.split('-')[2], 10);
          if (d) daySelect.value = String(d);
        }
      } else {
        fixedArea.style.display = 'none';
      }
    }
  },

  toggleTxRecurringNoEnd() {
    const noEndCb = document.getElementById('newTxRecurringNoEnd');
    const endInput = document.getElementById('newTxRecurringEnd');
    if (!noEndCb || !endInput) return;
    if (noEndCb.checked) {
      endInput.disabled = true;
      endInput.value = '';
    } else {
      endInput.disabled = false;
    }
  },

  openAddModal(initialType) {
    this.editingTxId = null;

    const titleEl = document.getElementById('txModalTitle');
    if (titleEl) titleEl.textContent = '새 거래 등록';

    const chipArea = document.getElementById('quickTxChipArea');
    if (chipArea) chipArea.style.display = 'block';

    const sources = this.getPayableSourceList();
    const destinations = this.getTransferDestinationList();
    const fromSelect = document.getElementById('newTxFromAccount');
    const toSelect = document.getElementById('newTxToAccount');
    if (fromSelect) {
      fromSelect.innerHTML = sources.map((s, idx) => `<option value="${this.escapeHtml(s.id)}" ${idx === 0 ? 'selected' : ''}>${this.escapeHtml(s.name)}</option>`).join('');
    }
    if (toSelect) {
      toSelect.innerHTML = destinations.map((d, idx) => `<option value="${this.escapeHtml(d.id)}" ${idx === 1 || (idx === 0 && destinations.length === 1) ? 'selected' : ''}>${this.escapeHtml(d.name)}</option>`).join('');
    }

    const typeToSet = initialType || 'EXPENSE_VARIABLE';
    document.getElementById('newTxType').value = typeToSet;
    document.getElementById('newTxCategory').value = '';
    document.getElementById('newTxAmount').value = '';
    document.getElementById('newTxMemo').value = '';
    this.renderQuickTxChips();

    const ym = this.data.selectedYearMonth || getCurrentYearMonthStr();
    const today = new Date();
    const todayYm = getCurrentYearMonthStr();
    let defaultDate;
    if (ym === todayYm) {
      defaultDate = getTodayStr();
    } else {
      const [y, m] = ym.split('-').map(Number);
      const daysInMonth = new Date(y, m, 0).getDate();
      const day = Math.min(today.getDate() || 15, daysInMonth);
      defaultDate = `${ym}-${String(day).padStart(2, '0')}`;
    }
    document.getElementById('newTxDate').value = defaultDate;

    const daySelect = document.getElementById('newTxRecurringDay');
    if (daySelect) {
      daySelect.innerHTML = Array.from({ length: 31 }, (_, i) => `<option value="${i + 1}">${i + 1}일</option>`).join('');
      const defaultDay = parseInt(defaultDate.split('-')[2], 10) || 25;
      daySelect.value = String(defaultDay);
    }
    const startInput = document.getElementById('newTxRecurringStart');
    if (startInput) startInput.value = ym;
    const noEndCb = document.getElementById('newTxRecurringNoEnd');
    if (noEndCb) noEndCb.checked = true;
    const endInput = document.getElementById('newTxRecurringEnd');
    if (endInput) {
      endInput.value = '';
      endInput.disabled = true;
    }

    this.onTxTypeChange();

    const btnContainer = document.getElementById('txModalActionButtons');
    if (btnContainer) {
      btnContainer.innerHTML = `
        <button onclick="App.closeAddModal()" style="flex:1; padding:12px; border-radius:14px; border:1px solid var(--color-border); background:#ffffff; font-weight:700; font-size:14px; color:var(--color-text-secondary); cursor:pointer;">취소</button>
        <button onclick="App.saveNewTransaction()" style="flex:2; padding:12px; border-radius:14px; border:none; background:var(--toss-blue); font-weight:700; font-size:14px; color:#ffffff; cursor:pointer;">저장하기</button>
      `;
    }

    const modal = document.getElementById('tossAddTxModal');
    if (modal) modal.classList.remove('hidden');
    this.scheduleIcons();
  },

  openEditTxModal(txId) {
    const tx = (this.data.transactions || []).find(t => t.id === txId);
    if (!tx) return;
    this.editingTxId = txId;

    const titleEl = document.getElementById('txModalTitle');
    if (titleEl) titleEl.textContent = '거래 내역 수정';

    const chipArea = document.getElementById('quickTxChipArea');
    if (chipArea) chipArea.style.display = 'none';

    const sources = this.getPayableSourceList();
    const destinations = this.getTransferDestinationList();
    const fromSelect = document.getElementById('newTxFromAccount');
    const toSelect = document.getElementById('newTxToAccount');
    if (fromSelect) {
      fromSelect.innerHTML = sources.map(s => `<option value="${this.escapeHtml(s.id)}" ${s.id === tx.accountId ? 'selected' : ''}>${this.escapeHtml(s.name)}</option>`).join('');
    }
    if (toSelect) {
      toSelect.innerHTML = destinations.map(d => `<option value="${this.escapeHtml(d.id)}" ${d.id === tx.toAccountId ? 'selected' : ''}>${this.escapeHtml(d.name)}</option>`).join('');
    }

    document.getElementById('newTxType').value = tx.type;

    const daySelect = document.getElementById('newTxRecurringDay');
    if (daySelect) {
      daySelect.innerHTML = Array.from({ length: 31 }, (_, i) => `<option value="${i + 1}">${i + 1}일</option>`).join('');
    }
    const linkedRec = (this.data.recurringExpenses || []).find(r => r.id === tx.recurringId || (r.name === tx.memo && r.category === tx.category && Number(r.amount) === Number(tx.amount)));
    if (linkedRec) {
      if (daySelect) daySelect.value = String(linkedRec.dayOfMonth || 25);
      const startInput = document.getElementById('newTxRecurringStart');
      if (startInput) startInput.value = linkedRec.startDate || (tx.date ? tx.date.substring(0, 7) : getCurrentYearMonthStr());
      const noEndCb = document.getElementById('newTxRecurringNoEnd');
      const endInput = document.getElementById('newTxRecurringEnd');
      if (noEndCb && endInput) {
        if (linkedRec.endDate) {
          noEndCb.checked = false;
          endInput.disabled = false;
          endInput.value = linkedRec.endDate;
        } else {
          noEndCb.checked = true;
          endInput.disabled = true;
          endInput.value = '';
        }
      }
    } else {
      const txDay = tx.date ? parseInt(tx.date.split('-')[2], 10) : 25;
      if (daySelect) daySelect.value = String(txDay || 25);
      const startInput = document.getElementById('newTxRecurringStart');
      if (startInput) startInput.value = tx.date ? tx.date.substring(0, 7) : getCurrentYearMonthStr();
      const noEndCb = document.getElementById('newTxRecurringNoEnd');
      const endInput = document.getElementById('newTxRecurringEnd');
      if (noEndCb && endInput) {
        noEndCb.checked = true;
        endInput.disabled = true;
        endInput.value = '';
      }
    }

    this.onTxTypeChange();

    document.getElementById('newTxCategory').value = tx.category || '';
    document.getElementById('newTxAmount').value = tx.amount ? Number(tx.amount).toLocaleString() : '';
    document.getElementById('newTxMemo').value = tx.memo || '';
    document.getElementById('newTxDate').value = tx.date || '';

    const btnContainer = document.getElementById('txModalActionButtons');
    if (btnContainer) {
      btnContainer.innerHTML = `
        <button onclick="App.deleteTransaction('${tx.id}'); App.closeAddModal();" style="flex:1; padding:12px; border-radius:14px; border:1px solid var(--color-loss-soft); background:var(--color-loss-soft); font-weight:700; font-size:14px; color:var(--color-loss); cursor:pointer;">삭제</button>
        <button onclick="App.closeAddModal()" style="flex:1; padding:12px; border-radius:14px; border:1px solid var(--color-border); background:var(--color-surface); font-weight:700; font-size:14px; color:var(--color-text-secondary); cursor:pointer;">취소</button>
        <button onclick="App.saveNewTransaction()" style="flex:2; padding:12px; border-radius:14px; border:none; background:var(--brand-primary); font-weight:700; font-size:14px; color:var(--btn-primary-text); cursor:pointer;">수정 완료</button>
      `;
    }

    const modal = document.getElementById('tossAddTxModal');
    if (modal) modal.classList.remove('hidden');
    this.scheduleIcons();
  },

  closeAddModal() {
    this.editingTxId = null;
    const modal = document.getElementById('tossAddTxModal');
    if (modal) modal.classList.add('hidden');
  },

  applyAccountOrAssetDiff(targetId, diff) {
    if (!targetId) return;
    const acc = (this.data.accounts || []).find(a => a.id === targetId);
    if (acc) {
      acc.balance = (Number(acc.balance) || 0) + diff;
      return;
    }
    const ast = (this.data.assets || []).find(a => a.id === targetId);
    if (ast) {
      if (ast.category === 'DEBT') {
        ast.amount = Math.max(0, (Number(ast.amount) || 0) - diff);
      } else {
        ast.amount = (Number(ast.amount) || 0) + diff;
      }
    }
  },

  saveNewTransaction() {
    const type = document.getElementById('newTxType').value;
    const category = document.getElementById('newTxCategory').value.trim();
    const rawAmt = String(document.getElementById('newTxAmount').value || '').replace(/,/g, '').trim();
    const amount = Number(rawAmt);
    let accountId = document.getElementById('newTxFromAccount').value;
    const toAccountId = document.getElementById('newTxToAccount').value;
    const memo = document.getElementById('newTxMemo').value.trim();
    const dateInput = document.getElementById('newTxDate').value;
    const date = dateInput || getTodayStr();

    if (!category || isNaN(amount) || amount <= 0) {
      alert('분류와 올바른 금액을 입력해주세요.');
      return;
    }

    if (type === 'TRANSFER' && accountId === toAccountId) {
      alert('출금 계좌와 입금 계좌가 동일합니다. 서로 다른 계좌를 선택해주세요.');
      return;
    }

    const txYearMonth = date.substring(0, 7);
    if (this.data.selectedYearMonth !== txYearMonth) {
      this.data.selectedYearMonth = txYearMonth;
    }

    const sources = this.getPayableSourceList();
    if (!accountId && sources.length > 0) {
      accountId = sources[0].id;
    }

    if (this.editingTxId) {
      const tx = this.data.transactions.find(t => t.id === this.editingTxId);
      if (tx) {
        if (tx.type === 'INCOME') this.applyAccountOrAssetDiff(tx.accountId, -Number(tx.amount));
        else if (tx.type.startsWith('EXPENSE')) this.applyAccountOrAssetDiff(tx.accountId, +Number(tx.amount));
        else if (tx.type === 'TRANSFER') {
          this.applyAccountOrAssetDiff(tx.accountId, +Number(tx.amount));
          if (tx.toAccountId) this.applyAccountOrAssetDiff(tx.toAccountId, -Number(tx.amount));
        }

        tx.type = type;
        tx.category = category;
        tx.amount = amount;
        tx.accountId = accountId;
        tx.toAccountId = type === 'TRANSFER' ? toAccountId : null;
        tx.memo = memo;
        tx.date = date;

        if (type === 'EXPENSE_FIXED') {
          const recDay = Number(document.getElementById('newTxRecurringDay')?.value) || parseInt(date.split('-')[2], 10) || 25;
          const recStart = document.getElementById('newTxRecurringStart')?.value || txYearMonth;
          const noEnd = document.getElementById('newTxRecurringNoEnd')?.checked ?? true;
          const recEnd = noEnd ? null : (document.getElementById('newTxRecurringEnd')?.value || null);

          if (!Array.isArray(this.data.recurringExpenses)) this.data.recurringExpenses = [];
          const existingRec = this.data.recurringExpenses.find(r => r.id === tx.recurringId || (r.name === memo && r.category === category));
          if (existingRec) {
            existingRec.name = memo || category;
            existingRec.category = category;
            existingRec.amount = amount;
            existingRec.accountId = accountId;
            existingRec.dayOfMonth = recDay;
            existingRec.startDate = recStart;
            existingRec.endDate = recEnd;
            tx.recurringId = existingRec.id;
          } else {
            const newRecId = 'rec-' + Date.now();
            this.data.recurringExpenses.push({
              id: newRecId,
              name: memo || category,
              category,
              amount,
              accountId,
              dayOfMonth: recDay,
              startDate: recStart,
              endDate: recEnd
            });
            tx.recurringId = newRecId;
          }
        }

        if (type === 'INCOME') this.applyAccountOrAssetDiff(accountId, +amount);
        else if (type.startsWith('EXPENSE')) this.applyAccountOrAssetDiff(accountId, -amount);
        else if (type === 'TRANSFER') {
          this.applyAccountOrAssetDiff(accountId, -amount);
          if (toAccountId) this.applyAccountOrAssetDiff(toAccountId, +amount);
        }
      }
      this.editingTxId = null;
    } else {
      let recurringId = null;
      if (type === 'EXPENSE_FIXED') {
        const recDay = Number(document.getElementById('newTxRecurringDay')?.value) || parseInt(date.split('-')[2], 10) || 25;
        const recStart = document.getElementById('newTxRecurringStart')?.value || txYearMonth;
        const noEnd = document.getElementById('newTxRecurringNoEnd')?.checked ?? true;
        const recEnd = noEnd ? null : (document.getElementById('newTxRecurringEnd')?.value || null);

        if (!Array.isArray(this.data.recurringExpenses)) this.data.recurringExpenses = [];
        recurringId = 'rec-' + Date.now();
        this.data.recurringExpenses.push({
          id: recurringId,
          name: memo || category,
          category,
          amount,
          accountId,
          dayOfMonth: recDay,
          startDate: recStart,
          endDate: recEnd
        });
      }

      const newTx = {
        id: 't-' + Date.now(),
        date,
        type,
        category,
        accountId,
        toAccountId: type === 'TRANSFER' ? toAccountId : null,
        amount,
        memo,
        recurringId
      };
      this.data.transactions.unshift(newTx);

      if (type === 'INCOME') this.applyAccountOrAssetDiff(accountId, +amount);
      else if (type.startsWith('EXPENSE')) this.applyAccountOrAssetDiff(accountId, -amount);
      else if (type === 'TRANSFER') {
        this.applyAccountOrAssetDiff(accountId, -amount);
        if (toAccountId) this.applyAccountOrAssetDiff(toAccountId, +amount);
      }
    }

    StorageService.save(this.data, true);
    this.closeAddModal();
    this.render();
    this.showToast('✅ 거래 내역이 저장되었습니다.');
  },

  deleteTransaction(id) {
    if (!confirm('이 거래 내역을 삭제하시겠습니까?')) return;
    const tx = (this.data.transactions || []).find(t => t.id === id);
    if (tx) {
      if (tx.type === 'INCOME') this.applyAccountOrAssetDiff(tx.accountId, -Number(tx.amount));
      else if (tx.type.startsWith('EXPENSE')) this.applyAccountOrAssetDiff(tx.accountId, +Number(tx.amount));
      else if (tx.type === 'TRANSFER') {
        this.applyAccountOrAssetDiff(tx.accountId, +Number(tx.amount));
        if (tx.toAccountId) this.applyAccountOrAssetDiff(tx.toAccountId, -Number(tx.amount));
      }

      this.data.transactions = this.data.transactions.filter(t => t.id !== id);
      StorageService.save(this.data, true);
      this.render();
      this.showToast('🗑️ 거래 내역이 삭제되었습니다.');
    }
  },

  // 자산 / 계좌 추가 및 수정 모달 로직
  calcStockAmount() {
    const bp = Number(document.getElementById('newAssetBuyPrice')?.value) || 0;
    const cpInput = document.getElementById('newAssetCurrentPrice')?.value;
    const cp = cpInput !== '' && !isNaN(Number(cpInput)) ? Number(cpInput) : bp;
    const qty = Number(document.getElementById('newAssetQuantity')?.value) || 0;
    const curr = document.getElementById('newAssetCurrency')?.value || 'KRW';
    const rate = curr === 'USD' ? (Number(this.data.usdExchangeRate) || 1380) : 1;
    const amtInput = document.getElementById('newAssetAmount');
    if (amtInput && qty > 0 && cp > 0) {
      amtInput.value = Math.round(cp * qty * rate);
    }
    const deductEl = document.getElementById('stockDeductSummary');
    if (deductEl) {
      const totalCostKrw = Math.round(bp * qty * rate);
      deductEl.textContent = `출금 예정: ₩${totalCostKrw.toLocaleString()}`;
    }
  },

  calcEditStockAmount() {
    const bp = Number(document.getElementById('editAssetBuyPrice')?.value) || 0;
    const cpInput = document.getElementById('editAssetCurrentPrice')?.value;
    const cp = cpInput !== '' && !isNaN(Number(cpInput)) ? Number(cpInput) : bp;
    const qty = Number(document.getElementById('editAssetQuantity')?.value) || 0;
    const curr = document.getElementById('editAssetCurrency')?.value || 'KRW';
    const rate = curr === 'USD' ? (Number(this.data.usdExchangeRate) || 1380) : 1;
    const amtInput = document.getElementById('editAssetAmount');
    if (amtInput && qty > 0 && cp > 0) {
      amtInput.value = Math.round(cp * qty * rate);
    }
  },

  openAddAssetModal(category = 'CHECKING') {
    const accModal = document.getElementById('accountsDetailModal');
    const stockModal = document.getElementById('stockPortfolioDetailModal');
    if (accModal && !accModal.classList.contains('hidden')) {
      this.parentModalBeforeAdd = 'accountsDetailModal';
      accModal.classList.add('hidden');
    } else if (stockModal && !stockModal.classList.contains('hidden')) {
      this.parentModalBeforeAdd = 'stockPortfolioDetailModal';
      stockModal.classList.add('hidden');
    } else {
      this.parentModalBeforeAdd = null;
    }

    const catSelect = document.getElementById('newAssetCategory');
    if (catSelect) catSelect.value = category;
    this.handleAssetTypeChange(category);

    const dateInput = document.getElementById('newAssetDate');
    if (dateInput) dateInput.value = getTodayStr();

    document.getElementById('newAssetName').value = '';
    document.getElementById('newAssetInstitution').value = '';
    document.getElementById('newAssetAmount').value = '';
    document.getElementById('newAssetTicker').value = '';
    document.getElementById('newAssetBuyPrice').value = '';
    document.getElementById('newAssetQuantity').value = '';
    document.getElementById('newAssetCurrentPrice').value = '';
    document.getElementById('newAssetDividend').value = '';
    document.getElementById('newAssetRate').value = '';
    const currSelect = document.getElementById('newAssetCurrency');
    if (currSelect) currSelect.value = 'KRW';
    const marketSelect = document.getElementById('newAssetMarket');
    if (marketSelect) marketSelect.value = 'KR';
    const acctTypeSelect = document.getElementById('newAssetAccountType');
    if (acctTypeSelect) acctTypeSelect.value = 'GENERAL';

    // 매수 출금 계좌 목록 바인딩
    const fundingSelect = document.getElementById('newAssetFundingAccount');
    if (fundingSelect) {
      const accounts = this.data.accounts || [];
      if (accounts.length === 0) {
        fundingSelect.innerHTML = `<option value="">등록된 출금 계좌가 없습니다 (자산만 등록)</option>`;
      } else {
        fundingSelect.innerHTML = `<option value="">출금 안 함 (주식 자산만 등록)</option>` +
          accounts.map(acc => `<option value="${this.escapeHtml(acc.id)}">[출금] ${this.escapeHtml(acc.name)} (잔액: ₩${(Number(acc.balance) || 0).toLocaleString()})</option>`).join('');
      }
      fundingSelect.value = '';
    }
    const deductEl = document.getElementById('stockDeductSummary');
    if (deductEl) deductEl.textContent = '출금 예정: ₩0';

    // 예적금 가입 출금 계좌 바인딩
    const savingsFundingSelect = document.getElementById('newAssetSavingsFundingAccount');
    if (savingsFundingSelect) {
      const accounts = this.data.accounts || [];
      if (accounts.length === 0) {
        savingsFundingSelect.innerHTML = `<option value="">등록된 출금 계좌가 없습니다 (자산만 등록)</option>`;
      } else {
        savingsFundingSelect.innerHTML = `<option value="">출금 안 함 (예적금 자산만 등록)</option>` +
          accounts.map(acc => `<option value="${this.escapeHtml(acc.id)}">[출금] ${this.escapeHtml(acc.name)} (잔액: ₩${(Number(acc.balance) || 0).toLocaleString()})</option>`).join('');
      }
      savingsFundingSelect.value = '';
    }
    const savingsDeductEl = document.getElementById('savingsDeductSummary');
    if (savingsDeductEl) savingsDeductEl.textContent = '출금 예정: ₩0';
    const maturityInput = document.getElementById('newAssetMaturityDate');
    if (maturityInput) maturityInput.value = '';

    this.calcSavingsMaturityPreview();

    const modal = document.getElementById('tossAddAssetModal');
    if (modal) modal.classList.remove('hidden');
    this.scheduleIcons();
  },

  closeAddAssetModal() {
    const modal = document.getElementById('tossAddAssetModal');
    if (modal) modal.classList.add('hidden');
    if (this.parentModalBeforeAdd) {
      const parentModal = document.getElementById(this.parentModalBeforeAdd);
      if (parentModal) {
        parentModal.classList.remove('hidden');
        this.renderAssetsTab(this.getFinancialSummary());
        this.scheduleIcons();
      }
      this.parentModalBeforeAdd = null;
    }
  },

  handleAssetTypeChange(type) {
    const stockGroup = document.getElementById('stockFieldGroup');
    const interestGroup = document.getElementById('interestFieldGroup');
    const savingsExtraGroup = document.getElementById('savingsExtraGroup');
    const rateLabel = document.getElementById('newAssetRateLabel');
    const amountLabel = document.getElementById('newAssetAmountLabel');
    const dateLabel = document.getElementById('newAssetDateLabel');

    if (stockGroup) stockGroup.style.display = type === 'STOCK' ? 'flex' : 'none';
    if (interestGroup) interestGroup.style.display = (type === 'SAVINGS' || type === 'DEBT') ? 'flex' : 'none';
    if (savingsExtraGroup) savingsExtraGroup.style.display = type === 'SAVINGS' ? 'flex' : 'none';

    if (rateLabel) {
      rateLabel.textContent = type === 'DEBT' ? '대출 금리 (%)' : '약정 금리 (%)';
    }

    if (amountLabel) {
      if (type === 'DEBT') amountLabel.textContent = '대출 / 부채 금액 (원)';
      else if (type === 'STOCK') amountLabel.textContent = '현재 평가금액 (원)';
      else if (type === 'SAVINGS') amountLabel.textContent = '가입 원금 (원)';
      else amountLabel.textContent = '현재 잔액 / 평가금액 (원)';
    }

    if (dateLabel) {
      if (type === 'CHECKING') dateLabel.textContent = '개설일 / 잔액 기준일';
      else if (type === 'STOCK') dateLabel.textContent = '매수일 / 투자 시작일';
      else if (type === 'SAVINGS' || type === 'PENSION') dateLabel.textContent = '가입일';
      else if (type === 'DEBT') dateLabel.textContent = '대출 실행일';
      else dateLabel.textContent = '취득일자';
    }

    this.calcSavingsMaturityPreview();
  },

  calcSavingsMaturityPreview() {
    const type = document.getElementById('newAssetCategory')?.value;
    const amount = Number(document.getElementById('newAssetAmount')?.value) || 0;
    const deductEl = document.getElementById('savingsDeductSummary');
    if (deductEl) deductEl.textContent = `출금 예정: ₩${amount.toLocaleString()}`;

    if (type !== 'SAVINGS') return;

    const rate = Number(document.getElementById('newAssetRate')?.value) || 0;
    const startDateStr = document.getElementById('newAssetDate')?.value || getTodayStr();
    const maturityDateStr = document.getElementById('newAssetMaturityDate')?.value;
    const previewCard = document.getElementById('savingsMaturityPreviewCard');
    const previewNetInterest = document.getElementById('savingsPreviewNetInterest');
    const previewTotalPayout = document.getElementById('savingsPreviewTotalPayout');

    if (!previewCard || !previewNetInterest || !previewTotalPayout) return;

    if (amount > 0 && rate > 0 && maturityDateStr && startDateStr) {
      const start = new Date(startDateStr);
      const end = new Date(maturityDateStr);
      const diffDays = Math.round((end - start) / (1000 * 60 * 60 * 24));
      if (diffDays > 0) {
        // 단리 세전 이자 = 원금 * (연이율 / 100) * (예치일수 / 365)
        const grossInterest = Math.round(amount * (rate / 100) * (diffDays / 365));
        // 이자소득세 15.4% 공제 (소득세 14% + 지방소득세 1.4%)
        const tax = Math.floor(grossInterest * 0.154);
        const netInterest = Math.max(0, grossInterest - tax);
        const totalPayout = amount + netInterest;

        previewNetInterest.textContent = `+${netInterest.toLocaleString()}원 (${diffDays}일 기준)`;
        previewTotalPayout.textContent = `${totalPayout.toLocaleString()}원`;
        previewCard.style.display = 'block';
        return;
      }
    }
    previewCard.style.display = 'none';
  },

  saveNewAsset() {
    const type = document.getElementById('newAssetCategory').value;
    const date = document.getElementById('newAssetDate').value || getTodayStr();
    const name = document.getElementById('newAssetName').value.trim();
    const institution = document.getElementById('newAssetInstitution').value.trim() || '금융기관';
    const accountNumber = (document.getElementById('newAssetAccountNumber')?.value || '').trim();
    const amount = Number(document.getElementById('newAssetAmount').value);
    const ticker = document.getElementById('newAssetTicker').value.trim();
    const buyPrice = Number(document.getElementById('newAssetBuyPrice').value) || amount;
    const currentPriceInput = document.getElementById('newAssetCurrentPrice').value;
    const currentPrice = currentPriceInput !== '' && !isNaN(Number(currentPriceInput)) ? Number(currentPriceInput) : buyPrice;
    const quantity = Number(document.getElementById('newAssetQuantity').value) || 1;
    const currency = document.getElementById('newAssetCurrency')?.value || 'KRW';
    const market = document.getElementById('newAssetMarket')?.value || 'KR';
    const accountType = document.getElementById('newAssetAccountType')?.value || 'GENERAL';

    const dividendInput = document.getElementById('newAssetDividend').value.trim();
    const dividendYield = dividendInput !== '' && !isNaN(Number(dividendInput)) ? Number(dividendInput) : 0;

    const rateInput = document.getElementById('newAssetRate').value.trim();
    const rate = rateInput !== '' && !isNaN(Number(rateInput)) ? Number(rateInput) : 0;

    if (!name || isNaN(amount) || amount < 0) {
      alert('항목명과 올바른 금액을 입력해주세요.');
      return;
    }

    if (type === 'CHECKING') {
      this.data.accounts.push({
        id: 'acc-' + Date.now(),
        name,
        institution,
        accountNumber,
        type: 'CHECKING',
        balance: amount,
        date,
        isDefault: false
      });
    } else if (type === 'STOCK') {
      const fundingAccId = document.getElementById('newAssetFundingAccount')?.value;
      const rateKrw = currency === 'USD' ? (Number(this.data.usdExchangeRate) || 1380) : 1;
      const totalCostKrw = Math.round(buyPrice * quantity * rateKrw);

      let fundingAccName = '';
      if (fundingAccId) {
        const acc = (this.data.accounts || []).find(a => a.id === fundingAccId);
        if (acc) {
          fundingAccName = acc.name;
          // 선택된 계좌 잔액에서 매수 총액 차감
          acc.balance = (Number(acc.balance) || 0) - totalCostKrw;

          // 가계부 거래 내역에 '투자/이체' 거래 자동 기록
          if (!Array.isArray(this.data.transactions)) this.data.transactions = [];
          this.data.transactions.push({
            id: 'tx-' + Date.now(),
            type: 'TRANSFER',
            category: '투자/주식',
            amount: totalCostKrw,
            accountId: fundingAccId,
            toAccountId: null,
            date: date || getTodayStr(),
            memo: `주식 매수: ${name} (${quantity}주)`
          });
        }
      }

      this.data.assets.push({
        id: 'ast-' + Date.now(),
        category: 'STOCK',
        name,
        institution,
        accountNumber,
        accountType,
        ticker: ticker || 'STOCK',
        amount,
        buyPrice,
        currentPrice,
        quantity,
        currency,
        market,
        dividendYield,
        dividendMonths: [3, 6, 9, 12],
        date
      });

      StorageService.save(this.data, true);
      this.closeAddAssetModal();
      this.render();
      if (fundingAccName) {
        this.showToast(`✅ [${name}] 등록 완료! [${fundingAccName}]에서 ₩${totalCostKrw.toLocaleString()}원 출금 반영`);
      } else {
        this.showToast(`✅ [${name}] 주식이 등록되었습니다.`);
      }
      this.refreshStockPrices(false);
      return;
    } else if (type === 'SAVINGS') {
      const maturityDate = document.getElementById('newAssetMaturityDate')?.value || null;
      const fundingAccId = document.getElementById('newAssetSavingsFundingAccount')?.value;
      let fundingAccName = '';

      if (fundingAccId) {
        const acc = (this.data.accounts || []).find(a => a.id === fundingAccId);
        if (acc) {
          fundingAccName = acc.name;
          acc.balance = (Number(acc.balance) || 0) - amount;

          if (!Array.isArray(this.data.transactions)) this.data.transactions = [];
          this.data.transactions.push({
            id: 'tx-' + Date.now(),
            type: 'TRANSFER',
            category: '저축/적금',
            amount,
            accountId: fundingAccId,
            toAccountId: null,
            date: date || getTodayStr(),
            memo: `예적금 가입: ${name}`
          });
        }
      }

      this.data.assets.push({
        id: 'ast-' + Date.now(),
        category: 'SAVINGS',
        name,
        institution,
        accountNumber,
        amount,
        interestRate: rate,
        maturityDate,
        date
      });

      StorageService.save(this.data, true);
      this.closeAddAssetModal();
      this.render();
      if (fundingAccName) {
        this.showToast(`✅ [${name}] 예적금 가입 완료! [${fundingAccName}]에서 ₩${amount.toLocaleString()}원 출금 반영`);
      } else {
        this.showToast(`✅ [${name}] 예적금이 등록되었습니다.`);
      }
      return;
    } else if (type === 'PENSION') {
      this.data.assets.push({
        id: 'ast-' + Date.now(),
        category: 'PENSION',
        name,
        institution,
        accountNumber,
        amount,
        date
      });
    } else if (type === 'OTHER') {
      this.data.assets.push({
        id: 'ast-' + Date.now(),
        category: 'OTHER',
        name,
        institution,
        accountNumber,
        amount,
        date
      });
    } else if (type === 'DEBT') {
      this.data.assets.push({
        id: 'ast-' + Date.now(),
        category: 'DEBT',
        name,
        institution,
        accountNumber,
        amount,
        interestRate: rate,
        date
      });
    }

    StorageService.save(this.data, true);
    this.closeAddAssetModal();
    this.render();
    this.showToast(`✅ [${name}] 자산이 등록되었습니다.`);
    if (type === 'STOCK') this.refreshStockPrices(false);
  },

  // ===== 3-1. 계좌 상세 및 입출금 거래 내역 모달 =====
  openAccountHistoryModal(type, id) {
    this.currentAccountHistoryContext = { type, id, filter: 'ALL' };
    const item = type === 'account'
      ? (this.data.accounts || []).find(a => a.id === id)
      : (this.data.assets || []).find(a => a.id === id);

    if (!item) return;

    // 0) 이전 열려있던 부모 리스트 모달(계좌 목록 또는 주식 목록) 감지 및 전환
    const accModal = document.getElementById('accountsDetailModal');
    const stockModal = document.getElementById('stockPortfolioDetailModal');
    const backBtn = document.getElementById('accHistBackBtn');

    if (accModal && !accModal.classList.contains('hidden')) {
      this.parentModalBeforeHistory = 'accountsDetailModal';
      accModal.classList.add('hidden');
      if (backBtn) backBtn.classList.remove('hidden');
    } else if (stockModal && !stockModal.classList.contains('hidden')) {
      this.parentModalBeforeHistory = 'stockPortfolioDetailModal';
      stockModal.classList.add('hidden');
      if (backBtn) backBtn.classList.remove('hidden');
    } else {
      this.parentModalBeforeHistory = null;
      if (backBtn) backBtn.classList.add('hidden');
    }

    // 1) 헤더 정보
    const nameEl = document.getElementById('accHistName');
    const subEl = document.getElementById('accHistSub');
    const badgeArea = document.getElementById('accHistBadgeArea');
    const catBadge = document.getElementById('accHistCategoryBadge');

    const rawType = item.type || item.category || 'CHECKING';
    const institution = item.institution || '금융기관';

    if (nameEl) nameEl.textContent = item.name || '계좌 상세';

    // 뱃지 아이콘
    if (badgeArea) {
      let icon = 'landmark';
      if (rawType === 'STOCK') icon = 'trending-up';
      else if (rawType === 'SAVINGS') icon = 'piggy-bank';
      else if (rawType === 'PENSION') icon = 'shield';
      else if (rawType === 'DEBT') icon = 'credit-card';
      else if (rawType === 'OTHER') icon = 'building-2';
      badgeArea.innerHTML = this.getFinancialBrandBadge(item.name, item.institution || item.ticker, icon);
    }

    // 서브 정보 및 카테고리 배지
    let catName = '입출금 통장';
    let catColor = '#0064FF';
    let catBg = 'rgba(0,100,255,0.1)';

    if (rawType === 'STOCK') {
      catName = item.accountType ? this.getAccountTypeName(item.accountType) : '투자 주식';
      catColor = '#D62828';
      catBg = 'rgba(214,40,40,0.1)';
    } else if (rawType === 'SAVINGS') {
      catName = '예금 · 적금';
      catColor = '#059669';
      catBg = 'rgba(5,150,105,0.1)';
    } else if (rawType === 'PENSION') {
      catName = '개인연금';
      catColor = '#0284c7';
      catBg = 'rgba(2,132,199,0.1)';
    } else if (rawType === 'DEBT') {
      catName = '대출 · 부채';
      catColor = '#e11d48';
      catBg = 'rgba(225,29,72,0.1)';
    } else if (rawType === 'OTHER') {
      catName = '실물 · 기타';
      catColor = '#64748b';
      catBg = 'rgba(100,116,139,0.1)';
    }

    if (catBadge) {
      catBadge.textContent = catName;
      catBadge.style.color = catColor;
      catBadge.style.background = catBg;
    }

    let subText = `${institution} · ${catName}`;
    if (item.accountNumber) subText = `${institution} · ${item.accountNumber} · ${catName}`;
    if (item.date) subText += ` · ${item.date} 등록`;
    if (item.maturityDate) subText += ` · ${item.maturityDate} 만기`;
    if (item.interestRate) subText += ` · 금리 연 ${item.interestRate}%`;
    if (subEl) subEl.textContent = subText;

    // 2) 잔액 요약
    const balanceVal = Number(type === 'account' ? item.balance : item.amount) || 0;
    const balanceNum = document.getElementById('accHistBalanceNum');
    const balanceSpelled = document.getElementById('accHistBalanceSpelled');
    const balanceLabel = document.getElementById('accHistBalanceLabel');

    if (balanceLabel) {
      balanceLabel.textContent = rawType === 'DEBT' ? '대출 잔여 원금' : (rawType === 'STOCK' ? '현재 평가금액' : '현재 잔액');
    }

    if (balanceNum) {
      balanceNum.textContent = `${balanceVal.toLocaleString()}원`;
      balanceNum.style.color = balanceVal < 0 ? 'var(--color-loss)' : 'var(--color-text-primary)';
    }

    if (balanceSpelled) {
      balanceSpelled.textContent = `(${this.formatHangulSpelledOut(balanceVal)})`;
    }

    // 3) 이번 달 입금/출금 합계 요약
    const ym = this.data.selectedYearMonth || getCurrentYearMonthStr();
    let monthIncome = 0;
    let monthExpense = 0;

    (this.data.transactions || []).forEach(t => {
      if (!t.date || !t.date.startsWith(ym)) return;
      const amt = Number(t.amount) || 0;

      if (t.type === 'INCOME' && t.accountId === id) {
        monthIncome += amt;
      } else if (t.type === 'TRANSFER' && t.toAccountId === id) {
        monthIncome += amt;
      } else if (t.type.startsWith('EXPENSE') && t.accountId === id) {
        monthExpense += amt;
      } else if (t.type === 'TRANSFER' && t.accountId === id) {
        monthExpense += amt;
      }
    });

    const mIncEl = document.getElementById('accHistMonthIncome');
    const mExpEl = document.getElementById('accHistMonthExpense');
    if (mIncEl) mIncEl.textContent = `+₩${monthIncome.toLocaleString()}`;
    if (mExpEl) mExpEl.textContent = `-₩${monthExpense.toLocaleString()}`;

    // 4) 필터 탭 초기화 및 거래 내역 렌더링
    this.filterAccountHistory('ALL');

    const modal = document.getElementById('tossAccountHistoryModal');
    if (modal) modal.classList.remove('hidden');
    this.scheduleIcons();
  },

  backFromAccountHistory() {
    const parentId = this.parentModalBeforeHistory;
    this.closeAccountHistoryModal();
    if (parentId) {
      const parentModal = document.getElementById(parentId);
      if (parentModal) {
        parentModal.classList.remove('hidden');
        this.renderAssetsTab(this.getFinancialSummary());
        this.scheduleIcons();
      }
    }
  },

  closeAccountHistoryModal() {
    this.currentAccountHistoryContext = null;
    this.parentModalBeforeHistory = null;
    const backBtn = document.getElementById('accHistBackBtn');
    if (backBtn) backBtn.classList.add('hidden');
    const modal = document.getElementById('tossAccountHistoryModal');
    if (modal) modal.classList.add('hidden');
  },

  filterAccountHistory(filterType) {
    if (!this.currentAccountHistoryContext) return;
    this.currentAccountHistoryContext.filter = filterType;

    ['ALL', 'INCOME', 'EXPENSE', 'TRANSFER'].forEach(ft => {
      const btn = document.getElementById(`accHistFilter_${ft}`);
      if (btn) {
        if (ft === filterType) btn.classList.add('active');
        else btn.classList.remove('active');
      }
    });

    this.renderAccountHistoryList();
  },

  renderAccountHistoryList() {
    if (!this.currentAccountHistoryContext) return;
    const { id, filter } = this.currentAccountHistoryContext;
    const listEl = document.getElementById('accHistTxList');
    const countBadge = document.getElementById('accHistTxCountBadge');
    if (!listEl) return;

    // 해당 계좌와 관련된 거래 추출 (출금 계좌이거나 입금 계좌인 경우)
    const relatedTxs = (this.data.transactions || []).filter(t => {
      return t.accountId === id || t.toAccountId === id;
    });

    // 최신순 정렬 (날짜 내림차순, 같은 날짜면 id 내림차순)
    relatedTxs.sort((a, b) => {
      const dCmp = (b.date || '').localeCompare(a.date || '');
      if (dCmp !== 0) return dCmp;
      return String(b.id || '').localeCompare(String(a.id || ''));
    });

    // 필터링 적용
    const filtered = relatedTxs.filter(t => {
      if (filter === 'ALL') return true;
      if (filter === 'INCOME') {
        return (t.type === 'INCOME' && t.accountId === id) || (t.type === 'TRANSFER' && t.toAccountId === id);
      }
      if (filter === 'EXPENSE') {
        return t.type.startsWith('EXPENSE') && t.accountId === id;
      }
      if (filter === 'TRANSFER') {
        return t.type === 'TRANSFER';
      }
      return true;
    });

    if (countBadge) countBadge.textContent = `${filtered.length}건`;

    if (filtered.length === 0) {
      listEl.innerHTML = `
        <div style="text-align:center; padding:32px 16px; color:var(--color-text-tertiary);">
          <div style="font-size:13px; font-weight:700; color:var(--color-text-secondary); margin-bottom:4px;">해당 거래 내역이 없습니다</div>
          <div style="font-size:11px; color:var(--color-text-tertiary); margin-bottom:12px;">가계부에서 이 계좌로 입금 또는 지출을 등록해보세요.</div>
        </div>
      `;
      return;
    }

    const accMap = {};
    (this.data.accounts || []).forEach(a => accMap[a.id] = a.name);
    (this.data.assets || []).forEach(a => accMap[a.id] = a.name);

    listEl.innerHTML = filtered.map(t => {
      const isDeposit = (t.type === 'INCOME' && t.accountId === id) || (t.type === 'TRANSFER' && t.toAccountId === id);
      const isTransfer = t.type === 'TRANSFER';
      const amt = Number(t.amount) || 0;
      const sign = isDeposit ? '+' : '-';
      const color = isDeposit ? 'var(--color-gain)' : 'var(--color-text-primary)';
      const { icon, tint } = this.getCategoryIcon(t.category, t.type);

      // 날짜 라벨
      const parts = (t.date || '').split('-');
      const dateStr = parts.length === 3 ? `${parseInt(parts[1])}월 ${parseInt(parts[2])}일` : t.date;

      let subDesc = `${dateStr} · ${this.escapeHtml(t.category || '기타')}`;
      if (isTransfer) {
        if (isDeposit) {
          subDesc += ` (이체 입금: ${this.escapeHtml(accMap[t.accountId] || '타계좌')} →)`;
        } else {
          subDesc += ` (이체 출금: → ${this.escapeHtml(accMap[t.toAccountId] || '타계좌')})`;
        }
      }

      return `
        <div class="dash-recent-row" onclick="App.openEditTxModalFromHistory('${t.id}')" style="padding:10px 12px; background:var(--color-surface-subtle); border-radius:12px; border:1px solid var(--color-border); cursor:pointer; transition:background 0.15s ease;">
          <div style="display:flex; align-items:center; gap:10px; min-width:0;">
            <div class="finance-icon ${tint}" style="width:32px; height:32px; border-radius:9px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
              <i data-lucide="${icon}" style="width:15px; height:15px;"></i>
            </div>
            <div style="min-width:0;">
              <div style="font-size:13px; font-weight:700; color:var(--color-text-primary); white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                ${this.escapeHtml(t.memo || t.category || '거래')}
              </div>
              <div style="font-size:11px; color:var(--color-text-tertiary); margin-top:1px;">
                ${subDesc}
              </div>
            </div>
          </div>
          <div style="text-align:right; flex-shrink:0; margin-left:8px;">
            <div style="font-size:14px; font-weight:800; color:${color}; font-family:var(--font-toss);">
              ${sign}${amt.toLocaleString()}원
            </div>
          </div>
        </div>
      `;
    }).join('');

    this.scheduleIcons();
  },

  openEditTxModalFromHistory(txId) {
    this.closeAccountHistoryModal();
    this.openEditTxModal(txId);
  },

  openEditFromAccountHistory() {
    if (!this.currentAccountHistoryContext) return;
    const { type, id } = this.currentAccountHistoryContext;
    const prevParent = this.parentModalBeforeHistory;
    this.closeAccountHistoryModal();
    this.parentModalBeforeEdit = prevParent;
    this.openEditAssetModal(type, id);
  },

  openAddModalForAccount() {
    if (!this.currentAccountHistoryContext) return;
    const { id } = this.currentAccountHistoryContext;
    this.closeAccountHistoryModal();
    this.openAddModal('INCOME');
    // 입금/출금 계좌를 현재 선택된 계좌로 세팅
    const fromSelect = document.getElementById('newTxFromAccount');
    if (fromSelect) fromSelect.value = id;
    const toSelect = document.getElementById('newTxToAccount');
    if (toSelect) toSelect.value = id;
    this.onTxTypeChange();
  },

  openEditAssetModal(type, id) {
    this.editingAssetContext = { type, id };
    const item = type === 'account'
      ? (this.data.accounts || []).find(a => a.id === id)
      : (this.data.assets || []).find(a => a.id === id);

    if (!item) return;

    const dateInput = document.getElementById('editAssetDate');
    if (dateInput) dateInput.value = item.date || item.purchaseDate || getTodayStr();

    document.getElementById('editAssetName').value = item.name || '';
    const accNumEl = document.getElementById('editAssetAccountNumber');
    if (accNumEl) accNumEl.value = item.accountNumber || '';
    document.getElementById('editAssetAmount').value = type === 'account' ? item.balance : item.amount;

    const stockGroup = document.getElementById('editAssetStockGroup');
    if (stockGroup) {
      if (item.category === 'STOCK') {
        stockGroup.style.display = 'flex';
        const acctTypeEl = document.getElementById('editAssetAccountType');
        if (acctTypeEl) acctTypeEl.value = item.accountType || 'GENERAL';
        document.getElementById('editAssetTicker').value = item.ticker || '';
        document.getElementById('editAssetBuyPrice').value = item.buyPrice !== undefined ? item.buyPrice : '';
        document.getElementById('editAssetQuantity').value = item.quantity !== undefined ? item.quantity : '';
        document.getElementById('editAssetCurrentPrice').value = item.currentPrice !== undefined ? item.currentPrice : '';
        document.getElementById('editAssetCurrency').value = item.currency || 'KRW';
        const mk = document.getElementById('editAssetMarket');
        if (mk) mk.value = item.market || (item.currency === 'USD' ? 'US' : 'KR');
        const divEl = document.getElementById('editAssetDividend');
        if (divEl) divEl.value = item.dividendYield !== undefined && item.dividendYield !== null ? item.dividendYield : '';

        // 주식 판매(매도) 섹션 설정
        const sellSection = document.getElementById('stockSellActionSection');
        if (sellSection) {
          sellSection.style.display = 'block';
          const sellQtyInput = document.getElementById('stockSellQuantity');
          if (sellQtyInput) sellQtyInput.value = item.quantity || 1;
          const sellPriceInput = document.getElementById('stockSellPrice');
          if (sellPriceInput) sellPriceInput.value = item.currentPrice || item.buyPrice || 0;
          const depositSelect = document.getElementById('stockSellDepositAccount');
          if (depositSelect) {
            const accounts = this.data.accounts || [];
            if (accounts.length === 0) {
              depositSelect.innerHTML = `<option value="">등록된 입금 계좌가 없습니다</option>`;
            } else {
              depositSelect.innerHTML = accounts.map(a => `<option value="${this.escapeHtml(a.id)}">[입금] ${this.escapeHtml(a.name)} (잔액: ₩${(Number(a.balance) || 0).toLocaleString()})</option>`).join('');
            }
          }
          const sellArea = document.getElementById('stockSellFormArea');
          if (sellArea) sellArea.style.display = 'none';
          const toggleBtn = document.getElementById('stockSellToggleBtn');
          if (toggleBtn) toggleBtn.textContent = '판매창 열기 ▼';
          this.calcStockSellAmount();
        }
      } else {
        stockGroup.style.display = 'none';
      }
    }

    const rateGroup = document.getElementById('editAssetRateGroup');
    const rateInput = document.getElementById('editAssetRate');
    const rateLabel = document.getElementById('editAssetRateLabel');
    if (rateGroup && rateInput) {
      if (item.interestRate !== undefined) {
        rateGroup.style.display = 'block';
        if (rateLabel) rateLabel.textContent = item.category === 'DEBT' ? '대출 금리 (%)' : '약정 금리 (%)';
        rateInput.value = item.interestRate;
      } else {
        rateGroup.style.display = 'none';
      }
    }

    // 예적금 전용 만기일 수정 & 만기 해지 카드 설정
    const matDateGroup = document.getElementById('editAssetMaturityDateGroup');
    const matDateInput = document.getElementById('editAssetMaturityDate');
    const savingsMaturitySection = document.getElementById('savingsMaturityActionSection');
    if (item.category === 'SAVINGS') {
      if (matDateGroup) matDateGroup.style.display = 'block';
      if (matDateInput) matDateInput.value = item.maturityDate || '';
      if (savingsMaturitySection) {
        savingsMaturitySection.style.display = 'block';
        const depositSelect = document.getElementById('savingsMaturityDepositAccount');
        if (depositSelect) {
          const accounts = this.data.accounts || [];
          if (accounts.length === 0) {
            depositSelect.innerHTML = `<option value="">등록된 입금 계좌가 없습니다</option>`;
          } else {
            depositSelect.innerHTML = accounts.map(a => `<option value="${this.escapeHtml(a.id)}">[입금] ${this.escapeHtml(a.name)} (잔액: ₩${(Number(a.balance) || 0).toLocaleString()})</option>`).join('');
          }
        }
        const formArea = document.getElementById('savingsMaturityFormArea');
        if (formArea) formArea.style.display = 'none';
        const toggleBtn = document.getElementById('savingsMaturityToggleBtn');
        if (toggleBtn) toggleBtn.textContent = '해지창 열기 ▼';
        this.calcSavingsMaturityActual();
      }
    } else {
      if (matDateGroup) matDateGroup.style.display = 'none';
      if (savingsMaturitySection) savingsMaturitySection.style.display = 'none';
    }

    const modal = document.getElementById('tossEditAssetModal');
    if (modal) modal.classList.remove('hidden');
    this.scheduleIcons();
  },

  closeEditAssetModal() {
    this.editingAssetContext = null;
    const modal = document.getElementById('tossEditAssetModal');
    if (modal) modal.classList.add('hidden');
    if (this.parentModalBeforeEdit) {
      const parentModal = document.getElementById(this.parentModalBeforeEdit);
      if (parentModal) {
        parentModal.classList.remove('hidden');
        this.renderAssetsTab(this.getFinancialSummary());
        this.scheduleIcons();
      }
      this.parentModalBeforeEdit = null;
    }
  },

  saveEditAsset() {
    if (!this.editingAssetContext) return;
    const { type, id } = this.editingAssetContext;
    const newDate = document.getElementById('editAssetDate')?.value;
    const newName = document.getElementById('editAssetName').value.trim();
    const newAccountNumber = (document.getElementById('editAssetAccountNumber')?.value || '').trim();
    const newAmount = Number(document.getElementById('editAssetAmount').value);
    const rateStr = document.getElementById('editAssetRate')?.value.trim();
    const newRate = rateStr !== '' && !isNaN(Number(rateStr)) ? Number(rateStr) : null;

    if (!newName || isNaN(newAmount)) {
      alert('올바른 이름과 금액을 입력해주세요.');
      return;
    }

    if (type === 'account') {
      const acc = (this.data.accounts || []).find(a => a.id === id);
      if (acc) {
        acc.name = newName;
        acc.accountNumber = newAccountNumber;
        acc.balance = newAmount;
        if (newDate) acc.date = newDate;
      }
    } else {
      const ast = (this.data.assets || []).find(a => a.id === id);
      if (ast) {
        ast.name = newName;
        ast.accountNumber = newAccountNumber;
        ast.amount = newAmount;
        if (newDate) ast.date = newDate;
        if (ast.interestRate !== undefined && newRate !== null) ast.interestRate = newRate;

        if (ast.category === 'STOCK') {
          const acctType = document.getElementById('editAssetAccountType')?.value;
          if (acctType) ast.accountType = acctType;
          const divStr = document.getElementById('editAssetDividend')?.value.trim();
          if (divStr !== '' && !isNaN(Number(divStr))) {
            ast.dividendYield = Number(divStr);
          }
          const ticker = document.getElementById('editAssetTicker')?.value.trim();
          if (ticker) ast.ticker = ticker;
          const bp = Number(document.getElementById('editAssetBuyPrice')?.value);
          if (!isNaN(bp) && bp > 0) ast.buyPrice = bp;
          const qty = Number(document.getElementById('editAssetQuantity')?.value);
          if (!isNaN(qty) && qty > 0) ast.quantity = qty;
          const cp = Number(document.getElementById('editAssetCurrentPrice')?.value);
          if (!isNaN(cp) && cp >= 0) ast.currentPrice = cp;
          const curr = document.getElementById('editAssetCurrency')?.value;
          if (curr) ast.currency = curr;
          const mk = document.getElementById('editAssetMarket')?.value;
          if (mk) ast.market = mk;
        } else if (ast.category === 'SAVINGS') {
          const matDate = document.getElementById('editAssetMaturityDate')?.value;
          ast.maturityDate = matDate || null;
        }
      }
    }

    StorageService.save(this.data, true);
    this.closeEditAssetModal();
    this.render();
    this.showToast('✅ 자산 정보가 수정되었습니다.');
  },

  deleteCurrentEditingAsset() {
    if (!this.editingAssetContext) return;
    if (!confirm('정말 이 항목을 삭제하시겠습니까?')) return;

    const { type, id } = this.editingAssetContext;
    if (type === 'account') {
      this.data.accounts = (this.data.accounts || []).filter(a => a.id !== id);
    } else {
      this.data.assets = (this.data.assets || []).filter(a => a.id !== id);
    }

    StorageService.save(this.data, true);
    this.closeEditAssetModal();
    this.render();
    this.showToast('🗑️ 항목이 삭제되었습니다.');
  },

  // ===== 주식 판매(매도) 및 판매대금 입금 연동 =====
  toggleStockSellForm() {
    const area = document.getElementById('stockSellFormArea');
    const btn = document.getElementById('stockSellToggleBtn');
    if (!area) return;
    const isHidden = area.style.display === 'none' || area.style.display === '';
    area.style.display = isHidden ? 'flex' : 'none';
    if (btn) btn.textContent = isHidden ? '판매창 닫기 ▲' : '판매창 열기 ▼';
    if (isHidden) this.calcStockSellAmount();
  },

  calcStockSellAmount() {
    if (!this.editingAssetContext) return;
    const ast = (this.data.assets || []).find(a => a.id === this.editingAssetContext.id);
    if (!ast) return;

    const qty = Number(document.getElementById('stockSellQuantity')?.value) || 0;
    const price = Number(document.getElementById('stockSellPrice')?.value) || 0;
    const isUsd = ast.currency === 'USD';
    const rate = isUsd ? (Number(this.data.usdExchangeRate) || 1380) : 1;
    const totalKrw = Math.round(qty * price * rate);

    const txt = document.getElementById('stockSellTotalText');
    if (txt) {
      txt.textContent = isUsd
        ? `₩${totalKrw.toLocaleString()} ($${(qty * price).toLocaleString()})`
        : `₩${totalKrw.toLocaleString()}`;
    }
  },

  executeStockSell() {
    if (!this.editingAssetContext) return;
    const { id } = this.editingAssetContext;
    const ast = (this.data.assets || []).find(a => a.id === id);
    if (!ast || ast.category !== 'STOCK') return;

    const currentQty = Number(ast.quantity) || 1;
    const sellQty = Number(document.getElementById('stockSellQuantity')?.value);
    const sellPrice = Number(document.getElementById('stockSellPrice')?.value);
    const depositAccId = document.getElementById('stockSellDepositAccount')?.value;

    if (isNaN(sellQty) || sellQty <= 0) {
      alert('올바른 판매 수량을 입력해주세요.');
      return;
    }
    if (sellQty > currentQty) {
      alert(`보유 수량(${currentQty}주)보다 많은 수량을 판매할 수 없습니다.`);
      return;
    }
    if (isNaN(sellPrice) || sellPrice <= 0) {
      alert('올바른 판매 단가를 입력해주세요.');
      return;
    }
    if (!depositAccId) {
      alert('판매대금을 입금받을 계좌를 선택해주세요.');
      return;
    }

    const acc = (this.data.accounts || []).find(a => a.id === depositAccId);
    if (!acc) {
      alert('선택된 입금 계좌를 찾을 수 없습니다.');
      return;
    }

    const isUsd = ast.currency === 'USD';
    const rate = isUsd ? (Number(this.data.usdExchangeRate) || 1380) : 1;
    const totalSellKrw = Math.round(sellQty * sellPrice * rate);

    if (!confirm(`[${ast.name}] ${sellQty}주를 판매하여 [${acc.name}]으로 ₩${totalSellKrw.toLocaleString()}원을 입금 처리하시겠습니까?`)) {
      return;
    }

    // 1) 입금 계좌 잔액에 판매대금 가산
    acc.balance = (Number(acc.balance) || 0) + totalSellKrw;

    // 2) 가계부에 '투자 회수 / 주식 매도' 거래 내역 기록
    if (!Array.isArray(this.data.transactions)) this.data.transactions = [];
    this.data.transactions.unshift({
      id: 'tx-' + Date.now(),
      type: 'INCOME',
      category: '투자/주식',
      amount: totalSellKrw,
      accountId: depositAccId,
      toAccountId: null,
      date: getTodayStr(),
      memo: `주식 매도: ${ast.name} (${sellQty}주 판매)`
    });

    // 3) 주식 수량 및 평가액 차감 또는 전량 매도 시 삭제
    const isAllSold = (sellQty === currentQty);
    if (isAllSold) {
      this.data.assets = this.data.assets.filter(a => a.id !== id);
    } else {
      ast.quantity = currentQty - sellQty;
      const curPrice = Number(ast.currentPrice) || sellPrice;
      ast.amount = Math.round(ast.quantity * curPrice * rate);
    }

    StorageService.save(this.data, true);
    this.closeEditAssetModal();
    this.render();

    this.showToast(isAllSold
      ? `🎉 [${ast.name}] 전량(${sellQty}주) 매도 완료! [${acc.name}]에 ₩${totalSellKrw.toLocaleString()}원 입금 완료`
      : `✅ [${ast.name}] ${sellQty}주 매도 완료 (잔여 ${ast.quantity}주)! [${acc.name}]에 ₩${totalSellKrw.toLocaleString()}원 입금 완료`);
  },

  // ===== 예적금 만기 해지 및 원금+세후이자 입금 연동 =====
  toggleSavingsMaturityForm() {
    const area = document.getElementById('savingsMaturityFormArea');
    const btn = document.getElementById('savingsMaturityToggleBtn');
    if (!area) return;
    const isHidden = area.style.display === 'none' || area.style.display === '';
    area.style.display = isHidden ? 'flex' : 'none';
    if (btn) btn.textContent = isHidden ? '해지창 닫기 ▲' : '해지창 열기 ▼';
    if (isHidden) this.calcSavingsMaturityActual();
  },

  calcSavingsMaturityActual() {
    if (!this.editingAssetContext) return;
    const { type, id } = this.editingAssetContext;
    if (type === 'account') return;

    const ast = (this.data.assets || []).find(a => a.id === id);
    if (!ast || ast.category !== 'SAVINGS') return;

    const principal = Number(document.getElementById('editAssetAmount')?.value) || 0;
    const rate = Number(document.getElementById('editAssetRate')?.value) || 0;
    const startDateStr = document.getElementById('editAssetDate')?.value || ast.date || getTodayStr();
    const maturityDateStr = document.getElementById('editAssetMaturityDate')?.value || ast.maturityDate;

    let diffDays = 365;
    if (startDateStr && maturityDateStr) {
      const d = Math.round((new Date(maturityDateStr) - new Date(startDateStr)) / (1000 * 60 * 60 * 24));
      if (d > 0) diffDays = d;
    }

    const grossInterest = Math.round(principal * (rate / 100) * (diffDays / 365));
    const tax = Math.floor(grossInterest * 0.154);
    const netInterest = Math.max(0, grossInterest - tax);
    const totalPayout = principal + netInterest;

    const pEl = document.getElementById('savingsMaturityPrincipalText');
    const iEl = document.getElementById('savingsMaturityInterestText');
    const tEl = document.getElementById('savingsMaturityTotalPayoutText');

    if (pEl) pEl.textContent = `₩${principal.toLocaleString()}`;
    if (iEl) iEl.textContent = `+₩${netInterest.toLocaleString()} (${diffDays}일치)`;
    if (tEl) tEl.textContent = `₩${totalPayout.toLocaleString()}`;
  },

  executeSavingsMaturity() {
    if (!this.editingAssetContext) return;
    const { type, id } = this.editingAssetContext;
    if (type === 'account') return;

    const ast = (this.data.assets || []).find(a => a.id === id);
    if (!ast || ast.category !== 'SAVINGS') return;

    const depositAccId = document.getElementById('savingsMaturityDepositAccount')?.value;
    if (!depositAccId) {
      alert('해지 원리금을 입금받을 계좌를 선택해주세요.');
      return;
    }

    const acc = (this.data.accounts || []).find(a => a.id === depositAccId);
    if (!acc) {
      alert('입금 대상 계좌를 찾을 수 없습니다.');
      return;
    }

    const principal = Number(document.getElementById('editAssetAmount')?.value) || Number(ast.amount) || 0;
    const rate = Number(document.getElementById('editAssetRate')?.value) || Number(ast.interestRate) || 0;
    const startDateStr = document.getElementById('editAssetDate')?.value || ast.date || getTodayStr();
    const maturityDateStr = document.getElementById('editAssetMaturityDate')?.value || ast.maturityDate || getTodayStr();

    let diffDays = 365;
    if (startDateStr && maturityDateStr) {
      const d = Math.round((new Date(maturityDateStr) - new Date(startDateStr)) / (1000 * 60 * 60 * 24));
      if (d > 0) diffDays = d;
    }

    const grossInterest = Math.round(principal * (rate / 100) * (diffDays / 365));
    const tax = Math.floor(grossInterest * 0.154);
    const netInterest = Math.max(0, grossInterest - tax);
    const totalPayout = principal + netInterest;

    if (!confirm(`[${ast.name}] 만기 해지를 진행하시겠습니까?\n\n• 입금 계좌: ${acc.name}\n• 수령 원금: ₩${principal.toLocaleString()}원\n• 세후 이자: ₩${netInterest.toLocaleString()}원\n• 총 입금액: ₩${totalPayout.toLocaleString()}원`)) {
      return;
    }

    // 1) 대상 계좌 잔액에 원금 + 세후이자 입금
    acc.balance = (Number(acc.balance) || 0) + totalPayout;

    // 2) 가계부 거래 내역 기록
    if (!Array.isArray(this.data.transactions)) this.data.transactions = [];

    // 만기 세후 이자 수입 거래
    if (netInterest > 0) {
      this.data.transactions.unshift({
        id: 'tx-' + Date.now(),
        type: 'INCOME',
        category: '금융/이자',
        amount: netInterest,
        accountId: depositAccId,
        date: getTodayStr(),
        memo: `[만기이자] ${ast.name} (세후이자)`
      });
    }

    // 원금 반환 이체 기록
    this.data.transactions.unshift({
      id: 'tx-' + (Date.now() + 1),
      type: 'TRANSFER',
      category: '저축/투자 회수',
      amount: principal,
      accountId: null,
      toAccountId: depositAccId,
      date: getTodayStr(),
      memo: `[만기해지 원금] ${ast.name}`
    });

    // 3) 만기 완료된 예적금 자산 목록에서 제거
    this.data.assets = (this.data.assets || []).filter(a => a.id !== ast.id);

    // 4) 저장 & 갱신
    StorageService.save(this.data, true);
    this.closeEditAssetModal();
    this.render();
    this.showToast(`🎉 [${ast.name}] 만기 해지 완료! [${acc.name}]에 총 ₩${totalPayout.toLocaleString()}원(세후이자 ₩${netInterest.toLocaleString()} 포함)이 입금되었습니다.`);
  },

  // ===== 주식 현재가 & 배당수익률 자동 갱신 (구글 시트 Apps Script 연동) =====
  onAssetMarketChange(prefix) {
    const market = document.getElementById(`${prefix}AssetMarket`)?.value;
    const curr = document.getElementById(`${prefix}AssetCurrency`);
    if (curr && market) curr.value = (market === 'US') ? 'USD' : 'KRW';
    if (prefix === 'new') this.calcStockAmount(); else this.calcEditStockAmount();
  },

  // 종목 객체 -> 조회용 토큰 생성 (KS, KQ, KR, US)
  getQuoteToken(a) {
    const t = String(a.ticker || '').trim().toUpperCase();
    if (!t || t === 'STOCK') return null;
    return this.buildQuoteToken(t, a.market, a.currency);
  },

  // 티커/코드 문자열과 시장/통화로 표준 토큰 빌드 (자동 감지 지원)
  buildQuoteToken(ticker, market, currency) {
    let t = String(ticker || '').trim().toUpperCase();
    if (!t || t === 'STOCK') return null;

    // 접미사 .KS, .KQ 분리 지원
    if (t.endsWith('.KS')) return 'KS:' + t.replace('.KS', '');
    if (t.endsWith('.KQ')) return 'KQ:' + t.replace('.KQ', '');

    // 숫자 6자리는 한국 주식 자동 판별
    if (/^\d{6}$/.test(t)) {
      if (market === 'KOSPI') return 'KS:' + t;
      if (market === 'KOSDAQ') return 'KQ:' + t;
      return 'KR:' + t;
    }

    // 영문 티커는 미국 주식
    const isUsd = currency === 'USD' || market === 'US' || /^[A-Z]{1,5}$/.test(t);
    if (isUsd) {
      const cleanTicker = t.replace(/\./g, '-');
      return 'US:' + cleanTicker;
    }

    return (market === 'KOSPI' ? 'KS:' : (market === 'KOSDAQ' ? 'KQ:' : 'KR:')) + t;
  },

  // [신규] 티커 입력 시 현재가 & 배당 수익률 & 종목명 즉시 자동 조회
  async lookupTickerInfo(prefix = 'new') {
    const tickerInput = document.getElementById(`${prefix}AssetTicker`);
    if (!tickerInput) return;
    const rawTicker = tickerInput.value.trim();
    if (!rawTicker) return;

    // 1) 시장 및 통화 자동 보정
    const marketSelect = document.getElementById(`${prefix}AssetMarket`);
    const currSelect = document.getElementById(`${prefix}AssetCurrency`);
    const isKrNumeric = /^\d{6}$/.test(rawTicker) || rawTicker.endsWith('.KS') || rawTicker.endsWith('.KQ');

    if (isKrNumeric) {
      if (marketSelect && marketSelect.value === 'US') marketSelect.value = 'KR';
      if (currSelect) currSelect.value = 'KRW';
    } else if (/^[A-Za-z]{1,5}$/.test(rawTicker)) {
      if (marketSelect && marketSelect.value !== 'US') marketSelect.value = 'US';
      if (currSelect) currSelect.value = 'USD';
    }

    const market = marketSelect?.value;
    const currency = currSelect?.value;
    const token = this.buildQuoteToken(rawTicker, market, currency);
    if (!token) return;

    const url = this.getGoogleSheetUrl();
    if (!url) {
      this.showToast('💡 [설정] 탭에 구글 시트 URL을 등록하면 실시간 현재가와 배당률이 자동 조회돼요.');
      return;
    }

    this.showToast(`🔍 [${rawTicker}] 시세 및 배당 정보 조회 중...`);

    try {
      const reqUrl = url + (url.includes('?') ? '&' : '?') + 'action=quotes&symbols=' + encodeURIComponent(token) + '&_t=' + Date.now();
      const res = await fetch(reqUrl, { method: 'GET', mode: 'cors' });
      const result = await res.json();

      if (!result || result.status !== 'success' || !result.quotes || !result.quotes[token]) {
        throw new Error('종목 정보를 찾을 수 없습니다.');
      }

      const q = result.quotes[token];
      let filledMsg = [];

      // 현재가 자동 입력
      const curPriceInput = document.getElementById(`${prefix}AssetCurrentPrice`);
      if (curPriceInput && q.price) {
        curPriceInput.value = q.price;
        filledMsg.push(`현재가 ${q.currency === 'USD' ? '$' + Number(q.price).toLocaleString() : Number(q.price).toLocaleString() + '원'}`);
      }

      // 예상 배당수익률 자동 입력
      const dividendInput = document.getElementById(`${prefix}AssetDividend`);
      if (dividendInput && q.dividendYield !== undefined && q.dividendYield !== null) {
        dividendInput.value = q.dividendYield;
        filledMsg.push(`배당률 ${q.dividendYield}%`);
      }

      // 종목명 비어있거나 기본값이면 자동 기입
      const nameInput = document.getElementById(`${prefix}AssetName`);
      if (nameInput && (!nameInput.value || nameInput.value.trim() === '' || nameInput.value === '주식' || nameInput.value.toUpperCase() === rawTicker.toUpperCase())) {
        if (q.name) {
          nameInput.value = q.name;
          filledMsg.push(`종목명: ${q.name}`);
        }
      }

      // 매수가격 비어있으면 현재가로 제안 기입
      const buyPriceInput = document.getElementById(`${prefix}AssetBuyPrice`);
      if (buyPriceInput && (!buyPriceInput.value || Number(buyPriceInput.value) <= 0) && q.price) {
        buyPriceInput.value = q.price;
      }

      // 금액 재계산
      if (prefix === 'new') this.calcStockAmount(); else this.calcEditStockAmount();

      this.showToast(`✅ [${q.name || rawTicker}] ${filledMsg.join(' · ')} 자동 조회 완료!`);
    } catch (err) {
      console.warn('티커 조회 오류:', err);
      this.showToast(`⚠️ [${rawTicker}] 정보를 불러오지 못했습니다. 종목코드나 시트 배포 상태를 확인해주세요.`);
    }
  },

  renderStockRefreshStatus(state, msg) {
    const el = document.getElementById('stockRefreshStatus');
    const btn = document.getElementById('stockRefreshBtn');
    if (btn) {
      btn.disabled = (state === 'loading');
      btn.style.opacity = (state === 'loading') ? '0.6' : '1';
      const ic = btn.querySelector('svg, i');
      if (ic) ic.classList.toggle('lucide-spin', state === 'loading');
    }
    if (!el) return;
    if (state === 'loading') { el.textContent = '현재가 및 배당 정보를 불러오는 중...'; el.style.color = 'var(--color-text-secondary)'; return; }
    if (msg) { el.textContent = msg; el.style.color = 'var(--color-loss)'; return; }
    el.style.color = 'var(--color-text-secondary)';
    const last = this.data.stockPricesUpdatedAt;
    if (last) {
      const time = new Date(last).toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' });
      const fx = (this.data.stocks_usdAuto) ? ` · 환율 ${Number(this.data.usdExchangeRate).toLocaleString()}원` : '';
      el.textContent = `마지막 갱신 ${time}${fx} (시세·배당률 동기화)`;
    } else if (this.getGoogleSheetUrl()) {
      el.textContent = '아직 갱신 전이에요. 새로고침을 눌러보세요.';
    } else {
      el.textContent = '설정 탭에 구글 시트 URL을 등록하면 현재가·배당률이 자동으로 갱신돼요.';
    }
  },

  maybeAutoRefreshStocks() {
    if (!this.getGoogleSheetUrl()) return;
    const hasStock = (this.data.assets || []).some(a => a.category === 'STOCK');
    if (!hasStock) return;
    const last = Number(this.data.stockPricesUpdatedAt) || 0;
    if (Date.now() - last < 10 * 60 * 1000) return;   // 10분 이내 갱신했으면 생략
    this.refreshStockPrices(false);
  },

  async refreshStockPrices(isManual = false) {
    if (this._quoteBusy) return;
    const stocks = (this.data.assets || []).filter(a => a.category === 'STOCK');
    if (stocks.length === 0) {
      if (isManual) this.showToast('ℹ️ 등록된 주식이 없습니다.');
      return;
    }
    const url = this.getGoogleSheetUrl();
    if (!url) {
      if (isManual) this.showToast('⚠️ 설정 탭에서 구글 시트 URL을 먼저 등록해주세요.');
      this.renderStockRefreshStatus();
      return;
    }

    const items = [];
    let skipped = 0;
    stocks.forEach(a => {
      const tk = this.getQuoteToken(a);
      if (tk) items.push({ a, tk }); else skipped++;
    });
    const needFx = stocks.some(a => a.currency === 'USD');
    const tokens = [...new Set(items.map(i => i.tk))];
    if (needFx) tokens.push('FX:USDKRW');
    if (tokens.length === 0) {
      if (isManual) this.showToast('⚠️ 티커/시장이 올바른 종목이 없어 조회할 수 없습니다.');
      return;
    }

    this._quoteBusy = true;
    this.renderStockRefreshStatus('loading');
    try {
      const reqUrl = url + (url.includes('?') ? '&' : '?') + 'action=quotes&symbols=' + encodeURIComponent(tokens.join(',')) + '&_t=' + Date.now();
      const res = await fetch(reqUrl, { method: 'GET', mode: 'cors' });
      const result = await res.json();
      if (!result || result.status !== 'success' || !result.quotes) {
        throw new Error((result && result.message) || 'SCRIPT_OUTDATED');
      }

      let updated = 0;
      let failed = 0;
      items.forEach(({ a, tk }) => {
        const q = result.quotes[tk];
        const price = q ? Number(q.price) : NaN;
        if (isFinite(price) && price > 0) {
          a.currentPrice = Math.round(price * 10000) / 10000;
          a.priceUpdatedAt = Date.now();
          updated++;
        } else {
          failed++;
        }

        // 배당 수익률도 함께 자동 업데이트
        if (q && q.dividendYield !== undefined && q.dividendYield !== null) {
          const dy = Number(q.dividendYield);
          if (isFinite(dy) && dy >= 0) {
            a.dividendYield = Math.round(dy * 100) / 100;
          }
        }
      });

      this.data.stocks_usdAuto = false;
      const fx = result.quotes['FX:USDKRW'];
      const fxPrice = fx ? Number(fx.price) : NaN;
      if (isFinite(fxPrice) && fxPrice > 500 && fxPrice < 3000) {
        this.data.usdExchangeRate = Math.round(fxPrice * 100) / 100;
        this.data.stocks_usdAuto = true;
      }

      if (updated > 0 || this.data.stocks_usdAuto) this.data.stockPricesUpdatedAt = Date.now();
      StorageService.save(this.data, false);
      this._quoteBusy = false;
      this.render();

      if (isManual || failed > 0 || skipped > 0) {
        const notes = [];
        if (failed > 0) notes.push(`${failed}건 시세 없음`);
        if (skipped > 0) notes.push(`${skipped}건 티커/시장 확인 필요`);
        this.showToast(updated > 0
          ? `✅ ${updated}개 종목 현재가·배당률 갱신 완료${notes.length ? ' (' + notes.join(', ') + ')' : ''}`
          : `⚠️ 갱신된 종목이 없습니다 (${notes.join(', ') || '응답 없음'})`);
      }
    } catch (err) {
      this._quoteBusy = false;
      const outdated = String(err && err.message) === 'SCRIPT_OUTDATED';
      const msg = outdated
        ? '구글 Apps Script가 최신 버전이 아니에요. 새 버전으로 다시 배포해주세요.'
        : '현재가를 불러오지 못했어요. 잠시 후 다시 시도해주세요.';
      this.renderStockRefreshStatus('error', msg);
      if (isManual) this.showToast('❌ ' + msg);
    }
  },

  // 한 달 예산 설정 모달 로직
  openBudgetModal() {
    const totalInput = document.getElementById('budgetTotalInput');
    if (totalInput) totalInput.value = this.data.totalMonthlyBudget ? this.data.totalMonthlyBudget : '';

    const b = this.data.budgets || {};
    const food = document.getElementById('budgetCatFood');
    if (food) food.value = b['식비/외식'] !== undefined && b['식비/외식'] > 0 ? b['식비/외식'] : '';

    const house = document.getElementById('budgetCatHouse');
    if (house) house.value = b['주거/통신'] !== undefined && b['주거/통신'] > 0 ? b['주거/통신'] : '';

    const cafe = document.getElementById('budgetCatCafe');
    if (cafe) cafe.value = b['카페/간식'] !== undefined && b['카페/간식'] > 0 ? b['카페/간식'] : '';

    const mart = document.getElementById('budgetCatMart');
    if (mart) mart.value = b['마트/잡화'] !== undefined && b['마트/잡화'] > 0 ? b['마트/잡화'] : '';

    const modal = document.getElementById('tossBudgetModal');
    if (modal) modal.classList.remove('hidden');
    this.scheduleIcons();
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
    const newTotal = Number(totalInput.value) || 0;

    if (newTotal < 0) {
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

    StorageService.save(this.data, true);
    this.closeBudgetModal();
    this.render();
    this.showToast(`✅ 한 달 목표 예산이 ₩${newTotal.toLocaleString()}원으로 저장되었습니다.`);
  }
};

window.addEventListener('DOMContentLoaded', () => {
  App.init();
});