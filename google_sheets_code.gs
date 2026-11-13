/**
 * =========================================================================
 * 📊 스마트 자산관리 & 심플 가계부 - Google Apps Script 연동 스크립트 v2.1 (현재가 조회 추가)
 * =========================================================================
 * 
 * [배포 방법]
 * 1. 구글 스프레드시트 상단 메뉴 [확장 프로그램] > [Apps Script] 클릭
 * 2. 기존 코드를 모두 지우고 이 스크립트를 전체 복사해 붙여넣고 저장 (Cmd+S / Ctrl+S)
 * 3. 우측 상단 [배포] > [새 배포] 클릭
 * 4. 톱니바퀴 아이콘 클릭 > [웹 앱(Web app)] 선택
 *    - 설명: 가계부 자산관리 연동 API
 *    - 다음 사용자 권한으로 실행: '나' (내 계정)
 *    - 액세스 권한이 있는 사용자: '모든 사용자(Anyone)' (※ 웹앱에서 로그인 없이 동기화하기 위해 필수!)
 * 5. [배포] 버튼 클릭 후 생성된 '웹 앱 URL'을 복사하여 앱 설정 화면에 붙여넣으세요.
 * 
 * [v2.1 업데이트 시 주의]
 * 코드를 교체한 뒤 반드시 [배포] > [배포 관리] > 연필 아이콘 > 버전: '새 버전' > [배포]
 * 를 눌러야 기존 웹 앱 URL에 반영됩니다. (새 배포를 만들면 URL이 바뀝니다)
 * 처음 실행 시 외부 요청(UrlFetchApp) 권한 승인 창이 뜨면 승인해 주세요.
 *
 * [주요 기능]
 * - 주식 현재가/환율 조회 (action=quotes): 야후 파이낸스 우선, 실패 시 GOOGLEFINANCE 보조
 * - Google Sheets 셀당 50,000자 용량 한계 완벽 대응 (RawData 자동 청크 분할)
 * - 자동 덮어쓰기 전 이전 백업 시트(RawData_Backup) 보존
 * - 출금/입금 계좌 실시간 명칭 매핑 및 한글 표준 카테고리 표시
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

  // 주식 현재가 / 환율 조회
  if (action === 'quotes') {
    try {
      var q = getQuotes_(e.parameter.symbols);
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        quotes: q.quotes,
        errors: q.errors,
        fetchedAt: new Date().toISOString()
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: '시세 조회 실패: ' + err.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 데이터 불러오기 (RawData 시트에서 분할 청크 조립 복원)
  var rawSheet = ss.getSheetByName('RawData');
  if (!rawSheet) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: '시트에 저장된 데이터가 없습니다. 먼저 웹앱에서 [시트로 내보내기]를 실행해주세요.'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var rawValues = rawSheet.getDataRange().getValues();
    var rawJson = rawValues.map(function(r) { return r[0] || ''; }).join('');
    if (!rawJson) throw new Error('저장된 데이터가 비어 있습니다.');

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
      var rawValues = rawSheet.getDataRange().getValues();
      var rawJson = rawValues.map(function(r) { return r[0] || ''; }).join('');
      var data = JSON.parse(rawJson);
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        data: data
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 내보내기 (Export / SyncAll) 실행
    var data = payload.data;
    if (!data) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: '전송된 데이터가 없습니다.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 1. RawData 시트에 JSON 청크 분할 보관 (Google Sheets 50,000자 셀 제한 완벽 해결)
    var rawSheet = ss.getSheetByName('RawData');
    if (!rawSheet) {
      rawSheet = ss.insertSheet('RawData');
    }

    // 기존 데이터가 있다면 RawData_Backup 시트에 안전하게 이전 백업 보관
    var existingVals = rawSheet.getDataRange().getValues();
    if (existingVals.length > 0 && existingVals[0][0]) {
      var backupSheet = ss.getSheetByName('RawData_Backup');
      if (!backupSheet) {
        backupSheet = ss.insertSheet('RawData_Backup');
      }
      backupSheet.clear();
      backupSheet.getRange(1, 1, existingVals.length, 1).setValues(existingVals);
    }

    rawSheet.clear();
    var jsonString = JSON.stringify(data);
    var chunkSize = 40000; // 셀 제한 50,000자 미만 안전 마진
    var chunks = [];
    for (var i = 0; i < jsonString.length; i += chunkSize) {
      chunks.push([jsonString.substring(i, i + chunkSize)]);
    }
    if (chunks.length > 0) {
      rawSheet.getRange(1, 1, chunks.length, 1).setValues(chunks);
    }

    // 계좌/자산 ID -> 명칭 맵 생성
    var sourceNames = {};
    (data.accounts || []).forEach(function(a) { if (a.id) sourceNames[a.id] = a.name; });
    (data.assets || []).forEach(function(a) { if (a.id) sourceNames[a.id] = a.name; });

    // 2. 1.거래내역 시트 생성/갱신
    updateTransactionsSheet(ss, data.transactions || [], sourceNames);

    // 3. 2.계좌목록 시트 생성/갱신
    updateAccountsSheet(ss, data.accounts || []);

    // 4. 3.자산현황 시트 생성/갱신
    updateAssetsSheet(ss, data.assets || []);

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: '구글 시트에 모든 데이터(거래/계좌/자산)가 안전하게 동기화되었습니다!'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: '처리 중 오류 발생: ' + err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function updateTransactionsSheet(ss, txList, sourceNames) {
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
    var fromName = (tx.accountId && sourceNames[tx.accountId]) ? sourceNames[tx.accountId] : (tx.fromAccount || '');
    var toName = (tx.toAccountId && sourceNames[tx.toAccountId]) ? sourceNames[tx.toAccountId] : (tx.toAccount || '');

    rows.push([
      tx.datetime || tx.date || '',
      typeLabels[tx.type] || tx.type || '',
      tx.category || '',
      tx.memo || '',
      Number(tx.amount) || 0,
      fromName,
      toName
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#e8f3ff').setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }
}

function updateAccountsSheet(ss, accounts) {
  var sheet = ss.getSheetByName('2.계좌목록');
  if (!sheet) {
    sheet = ss.insertSheet('2.계좌목록', 1);
  }
  sheet.clear();
  var headers = ['계좌명', '금융기관', '분류', '현재잔액(원)', '등록/개설일'];
  var rows = [headers];

  accounts.forEach(function(acc) {
    rows.push([
      acc.name || '',
      acc.institution || '',
      acc.type || '입출금',
      Number(acc.balance) || 0,
      acc.date || ''
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
  var headers = ['자산명', '대분류', '평가금액(원)', '수량', '매수가', '현재가', '통화', '수익률/금리', '등록일'];
  var rows = [headers];

  var catLabels = {
    'STOCK': '투자(주식/ETF)',
    'INVEST': '투자(주식/ETF)',
    'SAVINGS': '예·적금',
    'DEPOSIT': '예·적금',
    'PENSION': '연금',
    'OTHER': '기타자산',
    'REAL_ESTATE': '부동산',
    'DEBT': '대출/부채'
  };

  assets.forEach(function(asset) {
    var rateStr = '';
    if (asset.category === 'STOCK') {
      rateStr = asset.dividendYield ? ('배당 ' + asset.dividendYield + '%') : '-';
    } else if (asset.interestRate !== undefined) {
      rateStr = '금리 ' + asset.interestRate + '%';
    }

    rows.push([
      asset.name || '',
      catLabels[asset.category] || asset.category || '',
      Number(asset.amount) || 0,
      asset.quantity || 1,
      asset.buyPrice || 0,
      asset.currentPrice || 0,
      asset.currency || 'KRW',
      rateStr,
      asset.date || ''
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#f3f0ff').setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }
}


/**
 * =========================================================================
 * 주식 현재가 / 환율 조회
 * 심볼 형식(앱에서 전송): KS:005930(코스피) / KQ:247540(코스닥) / KR:005930(국내 자동구분)
 *                         US:AAPL(미국) / FX:USDKRW(환율)
 * 1순위 야후 파이낸스 chart 엔드포인트(비공식), 2순위 GOOGLEFINANCE
 * 응답은 2분간 캐시하여 호출 횟수를 줄입니다.
 * =========================================================================
 */
var QUOTE_TOKEN_RE_ = /^(KS|KQ|KR|US|FX):[A-Za-z0-9.\-=^]{1,15}$/;

function getQuotes_(symbolsParam) {
  var tokens = String(symbolsParam || '').split(',')
    .map(function(s) { return s.trim(); })
    .filter(function(s) { return QUOTE_TOKEN_RE_.test(s); })
    .slice(0, 30);

  var quotes = {};
  var errors = {};
  var cache = CacheService.getScriptCache();
  var pending = [];

  tokens.forEach(function(tk) {
    var hit = cache.get('q3:' + tk);
    if (hit) {
      try { quotes[tk] = JSON.parse(hit); return; } catch (e) {}
    }
    pending.push(tk);
  });

  // 1) 한국 주식 (KR, KS, KQ): 네이버 모바일 증권 API 우선 조회 (초고속, 현재가+배당률+종목명)
  var krPending = pending.filter(function(tk) {
    return tk.indexOf('KR:') === 0 || tk.indexOf('KS:') === 0 || tk.indexOf('KQ:') === 0;
  });

  if (krPending.length > 0) {
    var krReqs = [];
    var krMeta = [];
    krPending.forEach(function(tk) {
      var code = tk.split(':')[1];
      krReqs.push({
        url: 'https://m.stock.naver.com/api/stock/' + encodeURIComponent(code) + '/basic',
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15' }
      });
      krMeta.push({ token: tk, code: code });
    });

    try {
      var krResps = UrlFetchApp.fetchAll(krReqs);
      krResps.forEach(function(r, i) {
        var metaItem = krMeta[i];
        if (!metaItem) return;
        var tk = metaItem.token;
        try {
          if (r.getResponseCode() !== 200) return;
          var json = JSON.parse(r.getContentText());
          var rawPrice = json.closePrice || (json.stockItemTotalInfos && json.stockItemTotalInfos.closePrice);
          var price = Number(String(rawPrice || '').replace(/,/g, ''));
          if (isFinite(price) && price > 0) {
            var rawDiv = json.dividendYield || (json.stockItemTotalInfos && json.stockItemTotalInfos.dividendYield);
            var divYield = rawDiv ? parseFloat(String(rawDiv).replace(/,/g, '').replace(/%/g, '')) : 0;
            if (isNaN(divYield) || divYield < 0) divYield = 0;
            quotes[tk] = {
              price: price,
              dividendYield: Math.round(divYield * 100) / 100,
              name: json.stockName || '',
              currency: 'KRW',
              source: 'naver',
              symbol: metaItem.code
            };
          }
        } catch (e) {}
      });
    } catch (e) {}
  }

  // 2) 미국 주식 (US) 및 네이버에서 누락된 한국 주식: 야후 파이낸스 Quote API 조회
  var yahooPending = pending.filter(function(tk) { return !quotes[tk]; });
  if (yahooPending.length > 0) {
    var yahooSymbols = [];
    var symToToken = {};
    yahooPending.forEach(function(tk) {
      toYahooSymbols_(tk).forEach(function(ys) {
        yahooSymbols.push(ys);
        symToToken[ys] = tk;
      });
    });

    if (yahooSymbols.length > 0) {
      try {
        var yUrl = 'https://query1.finance.yahoo.com/v7/finance/quote?symbols=' + encodeURIComponent(yahooSymbols.join(','));
        var yRes = UrlFetchApp.fetch(yUrl, {
          muteHttpExceptions: true,
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
        });
        if (yRes.getResponseCode() === 200) {
          var yData = JSON.parse(yRes.getContentText());
          var results = (yData && yData.quoteResponse && yData.quoteResponse.result) || [];
          results.forEach(function(item) {
            var tk = symToToken[item.symbol];
            if (!tk || quotes[tk]) return;
            var price = Number(item.regularMarketPrice);
            if (isFinite(price) && price > 0) {
              var dy = 0;
              if (item.trailingAnnualDividendYield !== undefined && item.trailingAnnualDividendYield !== null) {
                // 야후에서 trailingAnnualDividendYield는 종종 0.0244 형태로 옴
                var rawDy = Number(item.trailingAnnualDividendYield);
                dy = rawDy < 0.2 ? rawDy * 100 : rawDy;
              } else if (item.dividendYield !== undefined && item.dividendYield !== null) {
                dy = Number(item.dividendYield);
              }
              if (isNaN(dy) || dy < 0) dy = 0;

              quotes[tk] = {
                price: price,
                dividendYield: Math.round(dy * 100) / 100,
                name: item.shortName || item.longName || '',
                currency: item.currency || '',
                source: 'yahoo',
                symbol: item.symbol
              };
            }
          });
        }
      } catch (e) {}
    }
  }

  // 3) 야후 Quote로도 못 가져온 경우 야후 Chart API 보조
  var chartPending = pending.filter(function(tk) { return !quotes[tk]; });
  if (chartPending.length > 0) {
    var reqs = [];
    var meta = [];
    chartPending.forEach(function(tk) {
      toYahooSymbols_(tk).forEach(function(ys) {
        reqs.push({
          url: 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(ys) + '?interval=1d&range=1d',
          muteHttpExceptions: true,
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
        });
        meta.push({ token: tk, ys: ys });
      });
    });

    if (reqs.length > 0) {
      try {
        var resps = UrlFetchApp.fetchAll(reqs);
        resps.forEach(function(r, i) {
          var tk = meta[i].token;
          if (quotes[tk]) return;
          try {
            if (r.getResponseCode() !== 200) return;
            var m = JSON.parse(r.getContentText()).chart.result[0].meta;
            var price = Number(m.regularMarketPrice);
            if (isFinite(price) && price > 0) {
              quotes[tk] = {
                price: price,
                dividendYield: 0,
                name: m.shortName || '',
                currency: m.currency || '',
                source: 'yahoo-chart',
                symbol: meta[i].ys
              };
            }
          } catch (e) {}
        });
      } catch (e) {}
    }
  }

  // 4) 최종 누락 종목은 GOOGLEFINANCE 스프레드시트 수식으로 보조
  var missing = pending.filter(function(tk) { return !quotes[tk]; });
  if (missing.length > 0) {
    try { googleFinanceFallback_(missing, quotes); } catch (e) {}
  }

  pending.forEach(function(tk) {
    if (quotes[tk]) {
      try { cache.put('q3:' + tk, JSON.stringify(quotes[tk]), 180); } catch (e) {}
    } else {
      errors[tk] = '시세를 찾지 못했습니다';
    }
  });

  return { quotes: quotes, errors: errors };
}

function toYahooSymbols_(token) {
  var p = token.split(':');
  var kind = p[0];
  var code = p[1];
  if (kind === 'KS') return [code + '.KS'];
  if (kind === 'KQ') return [code + '.KQ'];
  if (kind === 'KR') return [code + '.KS', code + '.KQ'];
  if (kind === 'US') return [code.replace(/\./g, '-')];
  if (kind === 'FX') return [code + '=X'];
  return [];
}

function toGoogleSymbol_(token) {
  var p = token.split(':');
  var kind = p[0];
  var code = p[1];
  if (kind === 'KS' || kind === 'KQ' || kind === 'KR') return 'KRX:' + code;
  if (kind === 'US') return code;
  if (kind === 'FX') return 'CURRENCY:' + code;
  return '';
}

function googleFinanceFallback_(tokens, quotes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = '_quotes_tmp';
  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  var sh = ss.insertSheet(name);
  try {
    tokens.forEach(function(tk, i) {
      var sym = toGoogleSymbol_(tk);
      sh.getRange(i + 1, 1).setFormula('=GOOGLEFINANCE("' + sym + '","price")');
      // 배당률 yieldpct 수식 추가
      sh.getRange(i + 1, 2).setFormula('=GOOGLEFINANCE("' + sym + '","yieldpct")');
    });
    SpreadsheetApp.flush();
    Utilities.sleep(1500);
    var vals = sh.getRange(1, 1, tokens.length, 2).getValues();
    tokens.forEach(function(tk, i) {
      var price = Number(vals[i][0]);
      if (isFinite(price) && price > 0) {
        var dy = Number(vals[i][1]);
        if (isNaN(dy) || dy < 0) dy = 0;
        // GOOGLEFINANCE yieldpct는 백분율 비율(0.024) 또는 백분율(2.4)일 수 있음
        var yieldPct = dy < 0.2 && dy > 0 ? dy * 100 : dy;
        quotes[tk] = {
          price: price,
          dividendYield: Math.round(yieldPct * 100) / 100,
          name: '',
          currency: '',
          source: 'googlefinance',
          symbol: toGoogleSymbol_(tk)
        };
      }
    });
  } finally {
    ss.deleteSheet(sh);
  }
}
