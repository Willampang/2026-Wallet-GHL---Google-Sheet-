//birthdayHandle.gs
const TIANGAN = ['甲', '乙', '丙', '丁', '戊', '己', '庚', '辛', '壬', '癸'];
const DIZHI = ['子', '丑', '寅', '卯', '辰', '巳', '午', '未', '申', '酉', '戌', '亥'];

function doGet(e) {
  const p = e.parameter;
  const name = p.name || '';
  const qty = p.qty || '1';
  const row = p.row || '';
  const orderId = p.order || '';
  const token = p.token || '';
  const sheetName = p.sheet || 'Orders';  // Added sheet parameter
  
  if (!token || !row) {
    return HtmlService.createHtmlOutput(createErrorPage('无效的访问链接'));
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);  // Use dynamic sheet name
  
  if (!sh) {
    return HtmlService.createHtmlOutput(createErrorPage('系统错误'));
  }
  
  const rowId = parseInt(row);
  const storedLink = sh.getRange(rowId, 19).getValue();
  const goldenCardStatus = sh.getRange(rowId, 17).getValue();
  const goldenCardData = sh.getRange(rowId, 18).getValue();
  
  if (!storedLink) {
    return HtmlService.createHtmlOutput(createErrorPage('此链接已失效或无效'));
  }
  
  const urlMatch = storedLink.match(/token=([^&]+)/);
  const storedToken = urlMatch ? decodeURIComponent(urlMatch[1]) : null;
  
  if (!storedToken || storedToken !== token) {
    return HtmlService.createHtmlOutput(createErrorPage('此链接已失效或无效'));
  }
  
  if (goldenCardStatus === 'Complete') {
    return createResultsPage(name, goldenCardData, rowId, sh, sheetName);
  }
  
  return createBirthdayForm(name, qty, row, orderId, token, sheetName);
}

function processFormSubmission(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = data.sheetName || 'Orders';  // Get sheet name from data
    const sh = ss.getSheetByName(sheetName);
    const rowId = parseInt(data.rowId);
    const submittedToken = data.token || '';
    
    if (!sh) {
      return {
        success: false,
        error: '系统错误：找不到对应的订单表'
      };
    }
    
    const storedLink = sh.getRange(rowId, 19).getValue();
    const goldenCardStatus = sh.getRange(rowId, 17).getValue();
    
    if (!storedLink) {
      return {
        success: false,
        error: '访问令牌无效，请重新获取链接'
      };
    }
    
    const urlMatch = storedLink.match(/token=([^&]+)/);
    const storedToken = urlMatch ? decodeURIComponent(urlMatch[1]) : null;
    
    if (!storedToken || storedToken !== submittedToken) {
      return {
        success: false,
        error: '访问令牌无效，请重新获取链接'
      };
    }
    
    if (goldenCardStatus === 'Complete') {
      return {
        success: false,
        error: '您已经提交过生日资料了'
      };
    }
    
    const cards = [];
    const allCards = [];
    const detailedInfo = [];
    
    for (let i = 0; i < data.wallets.length; i++) {
      const wallet = data.wallets[i];
      const card = calculateCard(wallet.year, wallet.month, wallet.day, wallet.hour, wallet.hasTime);
      
      cards.push({
        walletNum: wallet.walletNum,
        recipient: wallet.recipient,
        goldenCard: card,
        hourName: wallet.hourName,
        birthday: wallet.birthday,
        birthtime: wallet.birthtime
      });
      
      allCards.push(card);
      
      detailedInfo.push({
        wallet: wallet.walletNum,
        recipient: wallet.recipient,
        birthday: wallet.birthday,
        birthtime: wallet.birthtime,
        hourName: wallet.hourName,
        card: card
      });
    }
    
    const formattedCards = formatCardsWithSeparator(allCards);
    
    const cache = CacheService.getScriptCache();
    const cacheKey = 'details_' + sheetName + '_' + rowId;  // Include sheet name in cache key
    const detailedInfoJson = JSON.stringify(detailedInfo);
    cache.put(cacheKey, detailedInfoJson, 86400);
    
    sh.getRange(rowId, 17).setValue('Complete');
    sh.getRange(rowId, 18).setValue(formattedCards);
    
    return {
      success: true,
      cards: cards
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================================
// 核心算法（星桥奇门 + GV修正版）
// ============================================================

function calculateCard(year, month, day, hourIndex, hasTime) {
  Logger.log('========================================');
  Logger.log('📅 输入: ' + year + '年' + month + '月' + day + '日');

  const dayPillar = getDayPillarFixed(year, month, day, hourIndex);
  const ganZhi = dayPillar.gan + dayPillar.zhi;
  Logger.log('日柱: ' + ganZhi);

  const isDayGanYin = isYinGan(dayPillar.gan);
  Logger.log('日干: ' + dayPillar.gan + ' (' + (isDayGanYin ? '阴干' : '阳干') + ')');

  const solarTerm = getSolarTerm(year, month, day);
  Logger.log('节气: ' + solarTerm.name);

  const juShu = getJuShuFromSolarTerm(solarTerm.name, solarTerm.isYangDun);
  Logger.log('局数: ' + juShu + '局 (' + (solarTerm.isYangDun ? '阳遁' : '阴遁') + ')');

  let palace = flyFromLiGong(juShu, solarTerm.isYangDun);
  Logger.log('基础飞宫: ' + palace);

  if (isDayGanYin) {
    palace = reverseFly(palace);
    Logger.log('阴干反飞: ' + palace);
  }

  const hourBranches = ['子','丑','寅','卯','辰','巳','午','未','申','酉','戌','亥'];
  const hourBranch = hourBranches[hourIndex % 12];

  if (solarTerm.isYangDun && ['午','未','申'].includes(hourBranch)) {
    palace = BAGONG[(BAGONG.indexOf(palace) + 1) % 8];
    Logger.log('阳遁午未申顺延 → ' + palace);
  } else if (!solarTerm.isYangDun && ['子','丑','寅'].includes(hourBranch)) {
    palace = BAGONG[(BAGONG.indexOf(palace) + 1) % 8];
    Logger.log('阴遁子丑寅顺延 → ' + palace);
  }

  palace = applyGVCorrection(palace);
  Logger.log('GV修正: ' + palace);

  Logger.log('✅ 最终命宫: ' + palace);
  Logger.log('========================================');
  return palace;
}

function getDayPillarFixed(year, month, day, hourIndex) {
  const baseJD = getJulianDay(2000, 1, 1);
  let targetJD = getJulianDay(year, month, day);

  if (hourIndex === 0) targetJD -= 1;

  const daysDiff = targetJD - baseJD;
  const ganIndex = ((daysDiff % 10) + 10) % 10;
  const zhiIndex = ((daysDiff % 12) + 4 + 12) % 12;
  return { gan: TIANGAN[ganIndex], zhi: DIZHI[zhiIndex] };
}

function getJulianDay(year, month, day) {
  if (month <= 2) {
    year = year - 1;
    month = month + 12;
  }
  
  const A = Math.floor(year / 100);
  const B = 2 - A + Math.floor(A / 4);
  
  const JD = Math.floor(365.25 * (year + 4716)) + 
             Math.floor(30.6001 * (month + 1)) + 
             day + B - 1524.5;
  
  return JD;
}

function isYinGan(gan) {
  return ['乙', '丁', '己', '辛', '癸'].indexOf(gan) !== -1;
}

function getSolarTerm(year, month, day) {
  const termDates = [
    {month: 1, day: 5, name: '小寒', isYangDun: true},
    {month: 1, day: 20, name: '大寒', isYangDun: true},
    {month: 2, day: 4, name: '立春', isYangDun: true},
    {month: 2, day: 19, name: '雨水', isYangDun: true},
    {month: 3, day: 5, name: '惊蛰', isYangDun: true},
    {month: 3, day: 20, name: '春分', isYangDun: true},
    {month: 4, day: 5, name: '清明', isYangDun: true},
    {month: 4, day: 20, name: '谷雨', isYangDun: true},
    {month: 5, day: 5, name: '立夏', isYangDun: true},
    {month: 5, day: 21, name: '小满', isYangDun: true},
    {month: 6, day: 6, name: '芒种', isYangDun: true},
    {month: 6, day: 21, name: '夏至', isYangDun: false},
    {month: 7, day: 7, name: '小暑', isYangDun: false},
    {month: 7, day: 23, name: '大暑', isYangDun: false},
    {month: 8, day: 8, name: '立秋', isYangDun: false},
    {month: 8, day: 23, name: '处暑', isYangDun: false},
    {month: 9, day: 8, name: '白露', isYangDun: false},
    {month: 9, day: 23, name: '秋分', isYangDun: false},
    {month: 10, day: 8, name: '寒露', isYangDun: false},
    {month: 10, day: 23, name: '霜降', isYangDun: false},
    {month: 11, day: 7, name: '立冬', isYangDun: false},
    {month: 11, day: 22, name: '小雪', isYangDun: false},
    {month: 12, day: 7, name: '大雪', isYangDun: false},
    {month: 12, day: 22, name: '冬至', isYangDun: true}
  ];
  
  let currentTerm = termDates[0];
  
  for (let i = 0; i < termDates.length; i++) {
    const term = termDates[i];
    if (month < term.month || (month === term.month && day < term.day)) {
      currentTerm = i > 0 ? termDates[i - 1] : termDates[termDates.length - 1];
      break;
    } else if (i === termDates.length - 1) {
      currentTerm = term;
    }
  }
  
  return {
    name: currentTerm.name,
    isYangDun: currentTerm.isYangDun
  };
}

function getJuShuFromSolarTerm(solarTermName, isYangDun) {
  const yangDunJuShu = {
    '冬至': 1, '小寒': 1, '大寒': 2,
    '立春': 2, '雨水': 3, '惊蛰': 3,
    '春分': 4, '清明': 4, '谷雨': 5,
    '立夏': 5, '小满': 6, '芒种': 6
  };
  
  const yinDunJuShu = {
    '夏至': 9, '小暑': 9, '大暑': 8,
    '立秋': 8, '处暑': 7, '白露': 7,
    '秋分': 6, '寒露': 6, '霜降': 5,
    '立冬': 5, '小雪': 4, '大雪': 4
  };
  
  if (isYangDun) {
    return yangDunJuShu[solarTermName] || 1;
  } else {
    return yinDunJuShu[solarTermName] || 9;
  }
}

const BAGONG = ['离', '坤', '兑', '乾', '坎', '艮', '震', '巽'];

function flyFromLiGong(juShu, isYangDun) {
  const steps = juShu - 1;
  
  if (isYangDun) {
    const index = steps % 8;
    return BAGONG[index];
  } else {
    const index = (8 - (steps % 8)) % 8;
    return BAGONG[index];
  }
}

function reverseFly(palace) {
  const reverseMap = {
    '离': '坎', '坎': '离',
    '震': '兑', '兑': '震',
    '巽': '乾', '乾': '巽',
    '艮': '坤', '坤': '艮'
  };
  return reverseMap[palace] || palace;
}

function applyGVCorrection(palace) {
  const currentIndex = BAGONG.indexOf(palace);
  const newIndex = (currentIndex + 1) % 8;
  return BAGONG[newIndex];
}

function formatCardsWithSeparator(cards) {
  const cardCount = {};
  
  for (let i = 0; i < cards.length; i++) {
    const card = cards[i];
    if (cardCount[card]) {
      cardCount[card]++;
    } else {
      cardCount[card] = 1;
    }
  }
  
  const formatted = [];
  for (const card in cardCount) {
    const count = cardCount[card];
    formatted.push(card + 'x' + count);
  }
  
  return formatted.join(' | ');
}

function formatDateFromString(dateStr) {
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    return parts[0] + '年' + parts[1] + '月' + parts[2] + '日';
  }
  return dateStr;
}

// ============================================================
// HTML生成函数
// ============================================================
function createResultsPage(name, goldenCardData, rowId, sheet, sheetName) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'details_' + sheetName + '_' + rowId;  // Include sheet name in cache key
    const cachedData = cache.get(cacheKey);
    
    let cardsInfo = [];
    
    if (cachedData) {
      try {
        cardsInfo = JSON.parse(cachedData);
      } catch (e) {
        Logger.log('Error parsing cached data: ' + e);
      }
    }
    
    let cardsHtml = '';
    
    if (cardsInfo && cardsInfo.length > 0) {
      for (let i = 0; i < cardsInfo.length; i++) {
        const info = cardsInfo[i];
        const birthdateFormatted = formatDateFromString(info.birthday);
        const birthtimeDisplay = info.birthtime !== '未提供' ? info.birthtime : '未提供';
        
        cardsHtml += '<div class="card-item">' +
          '<div class="card-header">' +
          '<span class="card-number">🎴 #【奇门遁甲 招财阵】' + info.wallet + '</span>' +
          '<span class="recipient-badge">' + info.recipient + '</span>' +
          '</div>' +
          '<div class="birthday-info">' +
          '<p>📅 ' + birthdateFormatted + '</p>' +
          '<p>🕐 ' + birthtimeDisplay + ' (' + info.hourName + ')</p>' +
          '</div>' +
          '<div class="golden-card">' +
          '<h2>' + info.card + '</h2>' +
          '</div>' +
          '</div>';
      }
    } else {
      const cards = goldenCardData.split(' | ');
      for (let i = 0; i < cards.length; i++) {
        cardsHtml += '<div class="card-item">' +
          '<div class="card-header">' +
          '<span class="card-number">🎴 #【奇门遁甲 招财阵】' + (i + 1) + '</span>' +
          '</div>' +
          '<div class="golden-card">' +
          '<h2>' + cards[i] + '</h2>' +
          '</div>' +
          '</div>';
      }
    }
    
    const html = '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">' +
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
      '<title>满金包 - 命宫结果</title>' +
      '<style>' +
      '*{margin:0;padding:0;box-sizing:border-box}' +
      'body{font-family:"Microsoft YaHei",Arial,sans-serif;background:#cca983;min-height:100vh;padding:20px}' +
      '.container{max-width:600px;margin:0 auto;background:white;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden}' +
      '.header{background:linear-gradient(135deg,#8a4f19 0%,#a0681f 100%);color:white;padding:40px 30px;text-align:center}' +
      '.header h1{font-size:48px;margin:0;font-weight:bold;letter-spacing:8px}' +
      '.header p{margin:12px 0 0 0;font-size:18px;letter-spacing:3px}' +
      '.results-content{padding:30px}' +
      '.card-item{background:white;border:2px solid #946c36;border-radius:12px;padding:20px;margin-bottom:20px}' +
      '.card-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:15px;border-bottom:2px solid #946c36;padding-bottom:10px}' +
      '.card-number{font-weight:bold;color:#333;font-size:16px}' +
      '.recipient-badge{background:#542e10;color:white;padding:8px 16px;border-radius:20px;font-weight:bold;font-size:14px}' +
      '.birthday-info{margin-bottom:15px;color:#333}' +
      '.birthday-info p{margin:8px 0;font-size:14px}' +
      '.golden-card{background:#c9a870;padding:25px;border-radius:8px;text-align:center}' +
      '.golden-card h2{color:white;font-size:36px;text-shadow:1px 1px 2px rgba(0,0,0,0.3);font-weight:bold;letter-spacing:4px}' +
      '.footer{background:#542e10;color:white;padding:20px;text-align:center;font-size:13px}' +
      '.footer p{margin:5px 0}' +
      '.footer-phones{display:flex;gap:15px;justify-content:center;margin-top:10px}' +
      '</style>' +
      '</head><body>' +
      '<div class="container">' +
      '<div class="header">' +
      '<h1>满金包</h1>' +
      '<p>奇门遁甲 · 命宫结果</p>' +
      '</div>' +
      '<div class="results-content">' +
      cardsHtml +
      '</div>' +
      '<div class="footer">' +
      '<p><strong>恭喜你！已获得专属【奇门遁甲 招财阵】！</strong></p>' +
      '<p><strong>这个赠品将会和钱包一起寄出。如果你有任何疑问，请联系我们的客服。</strong></p>' +
      '<div class="footer-phones">' +
      '<span>📞 +6013-928 4699</span>' +
      '<span>📞 +6013-530 8863</span>' +
      '</div>' +
      '</div>' +
      '</div>' +
      '</body></html>';
    
    return HtmlService.createHtmlOutput(html)
      .setTitle('满金包2026 - 命宫结果')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
  } catch (error) {
    Logger.log('Error in createResultsPage: ' + error);
    return HtmlService.createHtmlOutput(createErrorPage('加载结果时出错'));
  }
}

function createErrorPage(message) {
  const html = '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '<title>错误</title>' +
    '<style>' +
    'body{font-family:"Microsoft YaHei",Arial,sans-serif;background:#cca983;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}' +
    '.error-container{background:white;border-radius:20px;padding:40px;max-width:500px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,0.3)}' +
    'h2{color:#E63946;margin-bottom:20px;font-size:32px}' +
    'p{color:#333;font-size:18px;line-height:1.6}' +
    '</style>' +
    '</head><body>' +
    '<div class="error-container">' +
    '<h2>❌ 错误</h2>' +
    '<p>' + message + '</p>' +
    '</div>' +
    '</body></html>';
  
  return html;
}

function createBirthdayForm(name, qty, row, orderId, token, sheetName) {
  const qtyNum = parseInt(qty) || 1;
  let formGroups = '';
  
  for (let i = 1; i <= qtyNum; i++) {
    formGroups += '<div class="wallet-group">' +
      '<div class="wallet-header">' +
      '<h3>#【奇门遁甲 招财阵】' + i + '</h3>' +
      '</div>' +
      '<div class="form-group">' +
      '<label>👤 这个钱包是给谁使用的?</label>' +
      '<select id="recipient' + i + '" required>' +
      '<option value="">请选择...</option>' +
      '<option value="本人">本人 (Myself)</option>' +
      '<option value="爸爸">爸爸 (Father)</option>' +
      '<option value="妈妈">妈妈 (Mother)</option>' +
      '<option value="孩子">孩子 (Child)</option>' +
      '<option value="配偶">配偶 (Spouse)</option>' +
      '<option value="朋友">朋友 (Friend)</option>' +
      '<option value="其他">其他 (Other)</option>' +
      '</select>' +
      '</div>' +
      '<div class="form-group">' +
      '<label>📅 出生日期</label>' +
      '<input type="date" id="birthday' + i + '" placeholder="dd/mm/yyyy" required>' +
      '</div>' +
      '<div class="form-group">' +
      '<label>🕐 出生时间 (可选)</label>' +
      '<input type="time" id="birthtime' + i + '">' +
      '<small style="color:#666;display:block;margin-top:5px;">如果不知道准确时间，可以留空</small>' +
      '</div>' +
      '</div>';
  }
  
  const html = '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>满金包 - 生辰八字登记</title><style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:"Microsoft YaHei",Arial,sans-serif;background:#cca983;min-height:100vh;padding:20px}.container{max-width:600px;margin:0 auto;background:white;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden}.header{background:linear-gradient(135deg,#8a4f19 0%,#a0681f 100%);color:white;padding:40px 30px;text-align:center;border-radius:20px 20px 0 0}.header h1{font-size:48px;margin:0;font-weight:bold;letter-spacing:8px}.header p{margin:12px 0 0 0;font-size:18px;letter-spacing:3px}.customer-info{background:#b88f51;border-left:4px solid #946c36;padding:15px;margin:15px;border-radius:6px}.customer-info p{margin:8px 0;font-size:14px;color:white;font-weight:500}.security-warning{background:#fff9e6;border-left:4px solid #946c36;padding:12px 15px;margin:15px;border-radius:6px;font-size:13px}.security-warning p{margin:6px 0;color:#333}.form-section{padding:30px}.wallet-group{background:white;padding:25px;border-radius:10px;margin-bottom:20px;border:2px solid #b88f51}.wallet-header{border-bottom:3px solid #b88f51;padding-bottom:12px;margin-bottom:18px}.wallet-header h3{color:#542e10;font-size:16px;font-weight:bold}.form-group{margin-bottom:20px}label{display:block;font-weight:600;margin-bottom:8px;color:#542e10;font-size:15px}input,select{width:100%;padding:12px;border:2px solid #ddd;border-radius:8px;font-size:15px;background:white}input:focus,select:focus{outline:none;border-color:#b88f51;box-shadow:0 0 6px rgba(184,143,81,0.6)}.submit-btn{width:100%;padding:18px;background:#E63946;color:white;border:none;border-radius:10px;font-size:24px;font-weight:bold;cursor:pointer;margin-top:15px;transition:background 0.3s}.submit-btn:hover{background:#D62828;transform:translateY(-2px);box-shadow:0 6px 16px rgba(230,57,70,0.3)}.submit-btn:disabled{background:#ccc;cursor:not-allowed;transform:none}.loading-overlay{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.8);z-index:9999;justify-content:center;align-items:center}.loading-container{display:flex;flex-direction:column;align-items:center;justify-content:center}.spinner{width:60px;height:60px;border:4px solid rgba(255,255,255,0.3);border-top:4px solid white;border-radius:50%;animation:spin 1s linear infinite}.progress-bar{width:350px;height:10px;background:rgba(255,255,255,0.3);border-radius:10px;overflow:hidden;margin:25px auto}.progress-fill{height:100%;background:linear-gradient(90deg,#b88f51,#946c36,#542e10);border-radius:10px;animation:progress 1.5s ease-out forwards}.loading-text{color:white;font-size:18px;margin-top:25px;font-weight:bold}.footer-phones{display:flex;gap:15px;justify-content:center;margin-top:10px}.phone-item{display:flex;align-items:center;gap:8px;color:white}@keyframes spin{0%{transform:rotate(0deg)}100%{transform:rotate(360deg)}}@keyframes progress{0%{width:0%}100%{width:100%}}</style></head><body><div class="loading-overlay" id="loadingOverlay"><div class="loading-container"><div class="spinner"></div><div class="progress-bar"><div class="progress-fill"></div></div><div class="loading-text">✨ 正在计算您的命宫...</div></div></div><div class="container"><div class="header"><h1>满金包</h1><p>奇门遁甲 · 生辰八字登记</p></div><div class="customer-info"><p><strong>👤 姓名:</strong> ' + name + '</p><p><strong>🎁 数量:</strong> ' + qtyNum + ' 个钱包</p></div><div class="security-warning"><p><strong>隐私保护：</strong></p><p>• 你提供的资料（姓名、出生日期、出生时间、出生地点等）将被严格保密，不会对外公开或与第三方共享。</p><p>• 资料仅用于个人八字分析与能量评估，不作其他商业用途。</p><p>• 我们会安全保存资料，并于分析完成后加密或删除。</p><p>• 提交资料即表示你自愿提供并同意以上条款，分析结果仅供参考。</p></div><div class="form-section"><form id="birthdayForm">' + formGroups + '<button type="submit" class="submit-btn" id="submitBtn">马上提交计算命宫</button></form></div></div><script>const rowId="' + row + '";const qty=' + qtyNum + ';const token="' + token + '";const sheetName="' + sheetName + '";function timeToHour(t){if(!t)return 6;const h=parseInt(t.split(":")[0]);if(h>=23||h<1)return 0;if(h>=1&&h<3)return 1;if(h>=3&&h<5)return 2;if(h>=5&&h<7)return 3;if(h>=7&&h<9)return 4;if(h>=9&&h<11)return 5;if(h>=11&&h<13)return 6;if(h>=13&&h<15)return 7;if(h>=15&&h<17)return 8;if(h>=17&&h<19)return 9;if(h>=19&&h<21)return 10;if(h>=21&&h<23)return 11;return 6}const hourNames=["子时","丑时","寅时","卯时","辰时","巳时","午时","未时","申时","酉时","戌时","亥时"];function updateRecipientOptions(){const selectedValues=new Set();for(let i=1;i<=qty;i++){const select=document.getElementById("recipient"+i);if(select.value){selectedValues.add(select.value)}}for(let i=1;i<=qty;i++){const select=document.getElementById("recipient"+i);const options=select.querySelectorAll("option");options.forEach(option=>{if(option.value&&option.value!==""){if(selectedValues.has(option.value)&&option.value!==select.value){option.style.display="none"}else{option.style.display=""}}})}}for(let i=1;i<=qty;i++){document.getElementById("recipient"+i).addEventListener("change",updateRecipientOptions)}function formatDateFromString(dateStr){const parts=dateStr.split("-");if(parts.length===3){return parts[0]+"年"+parts[1]+"月"+parts[2]+"日"}return dateStr}function displayResults(cards){let cardsHtml="";for(let i=0;i<cards.length;i++){const card=cards[i];const birthdateFormatted=formatDateFromString(card.birthday);const birthtimeDisplay=card.birthtime!=="未提供"?card.birthtime:"未提供";cardsHtml+=\'<div class="card-item">\'+\'<div class="card-header">\'+\'<span class="card-number">🎴 #【奇门遁甲 招财阵】\'+card.walletNum+\'</span>\'+\'<span class="recipient-badge">\'+card.recipient+\'</span>\'+\'</div>\'+\'<div class="birthday-info">\'+\'<p>📅 \'+birthdateFormatted+\'</p>\'+\'<p>🕐 \'+birthtimeDisplay+\' (\'+card.hourName+\')</p>\'+\'</div>\'+\'<div class="golden-card">\'+\'<h2>\'+card.goldenCard+\'</h2>\'+\'</div>\'+\'</div>\'}const resultsHtml=\'<div class="results-content">\'+cardsHtml+\'</div>\'+\'<div class="footer">\'+\'<p><strong>恭喜你！已获得专属【奇门遁甲 招财阵】！</strong></p>\'+\'<p><strong>这个赠品将会和钱包一起寄出。如果你有任何疑问，请联系我们的客服。</strong></p>\'+\'<div class="footer-phones">\'+\'<span class="phone-item">📞 +6013-928 4699</span>\'+\'<span class="phone-item">📞 +6013-530 8863</span>\'+\'</div>\'+\'</div>\';const additionalStyles=\'<style>.results-content{padding:30px}.card-item{background:white;border:2px solid #946c36;border-radius:12px;padding:20px;margin-bottom:20px}.card-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:15px;border-bottom:2px solid #946c36;padding-bottom:10px}.card-number{font-weight:bold;color:#333;font-size:16px}.recipient-badge{background:#542e10;color:white;padding:8px 16px;border-radius:20px;font-weight:bold;font-size:14px}.birthday-info{margin-bottom:15px;color:#333}.birthday-info p{margin:8px 0;font-size:14px}.golden-card{background:#c9a870;padding:25px;border-radius:8px;text-align:center;max-width:100%}.golden-card h2{color:white;font-size:36px;text-shadow:1px 1px 2px rgba(0,0,0,0.3);font-weight:bold;letter-spacing:4px}.footer{background:#542e10;color:white;padding:20px;text-align:center;font-size:13px}.footer p{margin:5px 0}</style>\';document.head.insertAdjacentHTML("beforeend",additionalStyles);document.querySelector(".container").innerHTML=\'<div class="header">\'+\'<h1>满金包 2026</h1>\'+\'<p>奇门遁甲 · 命宫结果</p>\'+\'</div>\'+resultsHtml}document.getElementById("birthdayForm").addEventListener("submit",function(e){e.preventDefault();const submitBtn=document.getElementById("submitBtn");const loadingOverlay=document.getElementById("loadingOverlay");const wallets=[];for(let i=1;i<=qty;i++){const recipient=document.getElementById("recipient"+i).value;const birthday=document.getElementById("birthday"+i).value;const birthtime=document.getElementById("birthtime"+i).value;if(!recipient){alert("请选择钱包 #"+i+" 是给谁的");return}if(!birthday){alert("请填写钱包 #"+i+" 的出生日期");return}const dateObj=new Date(birthday+"T00:00:00");const year=dateObj.getFullYear();const month=dateObj.getMonth()+1;const day=dateObj.getDate();const hasTime=birthtime?true:false;const hourIndex=timeToHour(birthtime);wallets.push({walletNum:i,recipient:recipient,year:year,month:month,day:day,hour:hourIndex,hourName:hasTime?hourNames[hourIndex]:"未提供",birthday:birthday,birthtime:birthtime||"未提供",hasTime:hasTime})}submitBtn.disabled=true;loadingOverlay.style.display="flex";const data={wallets:wallets,rowId:rowId,qty:qty,token:token,sheetName:sheetName};google.script.run.withSuccessHandler(function(result){setTimeout(function(){if(result.success){loadingOverlay.style.display="none";displayResults(result.cards)}else{loadingOverlay.style.display="none";submitBtn.disabled=false;alert("提交失败："+result.error)}},1500)}).withFailureHandler(function(error){loadingOverlay.style.display="none";submitBtn.disabled=false;alert("提交失败："+error.message)}).processFormSubmission(data)});</script></body></html>';
  
  return HtmlService.createHtmlOutput(html)
    .setTitle('满金包2026 - 生辰八字登记')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
