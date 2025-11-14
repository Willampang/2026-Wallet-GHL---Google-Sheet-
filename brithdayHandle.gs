//birthdayHandle.gs - EXTERNAL BOS API VERSION
// This version calls Render server which then calls BOS API

// ============================================================
// EXTERNAL API CONFIGURATION
// ============================================================
const EXTERNAL_BOS_API_URL = 'https://bos-middleware.onrender.com/api/calculate_golden_card';

// ============================================================
// SHOPIFY CONFIGURATION
// ============================================================
const BDAY_SHOPIFY_CONFIG = {
  SHOP_URL: 'fsr2021.myshopify.com',
  ACCESS_TOKEN: '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',
  API_VERSION: '2024-01'
};

const GOLDEN_CARD_VARIANTS = {
  '震': '47294134386840',
  '巽': '47294134223000',
  '乾': '47294133665944',
  '离': '47294133600408',
  '坤': '47294132519064',
  '坎': '47294132224152',
  '艮': '47294132191384',
  '兑': '47294131011736'
};

// ============================================================
// WEB APP HANDLERS
// ============================================================
function doGet(e) {
  const p = e.parameter;
  const name = p.name || '';
  const row = p.row || '';
  const orderId = p.order || '';
  const token = p.token || '';
  const sheetName = p.sheet || 'Orders';
  
  if (!token || !row) {
    return HtmlService.createHtmlOutput(createErrorPage('无效的访问链接'));
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  
  if (!sh) {
    return HtmlService.createHtmlOutput(createErrorPage('系统错误'));
  }
  
  const rowId = parseInt(row);
  const storedLink = sh.getRange(rowId, 19).getValue();
  const goldenCardStatus = sh.getRange(rowId, 17).getValue();
  const goldenCardData = sh.getRange(rowId, 18).getValue();
  const orderSummary = sh.getRange(rowId, 12).getValue();
  
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
  
  const actualQty = smartSplitQty(orderSummary);
  Logger.log('📊 Actual wallet quantity: ' + actualQty);
  
  return createBirthdayForm(name, actualQty, row, orderId, token, sheetName);
}

function smartSplit(str) {
  const parts = [];
  let currentPart = '';
  let depth = 0;
  
  for (let i = 0; i < str.length; i++) {
    const char = str[i];
    
    if (char === '（' || char === '(') {
      depth++;
      currentPart += char;
    } else if (char === '）' || char === ')') {
      depth--;
      currentPart += char;
    } else if (char === '+' && depth === 0) {
      if (currentPart.trim()) {
        parts.push(currentPart.trim());
      }
      currentPart = '';
    } else {
      currentPart += char;
    }
  }
  
  if (currentPart.trim()) {
    parts.push(currentPart.trim());
  }
  
  return parts;
}

function smartSplitQty(orderSummary) {
  if (!orderSummary || orderSummary === '') return 0;
  
  let totalQty = 0;
  const parts = smartSplit(orderSummary);
  
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i].trim();
    let walletCount = 0;
    
    if (part.includes('F款') && part.includes('带财款') && part.includes('吸金款')) {
      walletCount = 2;
    } else if (part.includes('G款') && part.includes('带财款x2')) {
      walletCount = 2;
    } else if (part.includes('H款') && part.includes('吸金款x2')) {
      walletCount = 2;
    } else {
      const matches = part.match(/[xX×]\s*(\d+)\s*$/);
      if (matches && matches[1]) {
        walletCount = parseInt(matches[1]);
      }
    }
    
    totalQty += walletCount;
  }
  
  return totalQty;
}

// ============================================================
// EXTERNAL BOS API CALL
// ============================================================

function formatDateTimeForBOS(year, month, day, hourIndex, minute) {
  // 地支索引 → 该时辰的起始小时
  const hourStart = {
    0: 23,  // 子时 23:00-01:00
    1: 1,   // 丑时 01:00-03:00
    2: 3,   // 寅时 03:00-05:00
    3: 5,   // 卯时 05:00-07:00
    4: 7,   // 辰时 07:00-09:00
    5: 9,   // 巳时 09:00-11:00
    6: 11,  // 午时 11:00-13:00
    7: 13,  // 未时 13:00-15:00
    8: 15,  // 申时 15:00-17:00
    9: 17,  // 酉时 17:00-19:00
    10: 19, // 戌时 19:00-21:00
    11: 21  // 亥时 21:00-23:00
  };
  
  const hour24 = hourStart[hourIndex] || 12;
  
  let hour12 = hour24;
  let ampm = 'AM';
  
  if (hour24 >= 12) {
    ampm = 'PM';
    if (hour24 > 12) {
      hour12 = hour24 - 12;
    }
  }
  
  if (hour24 === 0 || hour24 === 23) {
    hour12 = 11;
    ampm = 'PM';
  }
  
  const dateStr = year + '-' + 
    String(month).padStart(2, '0') + '-' + 
    String(day).padStart(2, '0');
  
  const timeStr = String(hour12).padStart(2, '0') + ':' + 
    String(minute || 0).padStart(2, '0') + ampm;
  
  return dateStr + ' ' + timeStr;
}

function callExternalBOSAPI(walletData) {
  try {
    Logger.log('🌐 Calling external BOS API...');
    Logger.log('   URL: ' + EXTERNAL_BOS_API_URL);
    
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(walletData),
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(EXTERNAL_BOS_API_URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('📥 External API Response:');
    Logger.log('   Status: ' + responseCode);
    Logger.log('   Body: ' + responseText);
    
    if (responseCode !== 200) {
      return {
        success: false,
        error: 'External API returned status ' + responseCode
      };
    }
    
    const data = JSON.parse(responseText);
    return data;
    
  } catch (error) {
    Logger.log('❌ External API Error: ' + error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================================
// FORM SUBMISSION HANDLER
// ============================================================

function processFormSubmission(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = data.sheetName || 'Orders';
    const sh = ss.getSheetByName(sheetName);
    const rowId = parseInt(data.rowId);
    const submittedToken = data.token || '';
    
    if (!sh) {
      return { success: false, error: '系统错误：找不到对应的订单表' };
    }
    
    const storedLink = sh.getRange(rowId, 19).getValue();
    const goldenCardStatus = sh.getRange(rowId, 17).getValue();
    const shopifyOrderId = sh.getRange(rowId, 16).getValue();
    
    if (!storedLink) {
      return { success: false, error: '访问令牌无效' };
    }
    
    const urlMatch = storedLink.match(/token=([^&]+)/);
    const storedToken = urlMatch ? decodeURIComponent(urlMatch[1]) : null;
    
    if (!storedToken || storedToken !== submittedToken) {
      return { success: false, error: '访问令牌无效' };
    }
    
    if (goldenCardStatus === 'Complete') {
      return { success: false, error: '您已经提交过生日资料了' };
    }
    
    // Prepare data for external BOS API
    const walletsData = [];
    
    for (let i = 0; i < data.wallets.length; i++) {
      const wallet = data.wallets[i];
      
      Logger.log('\n🎴 Preparing wallet #' + wallet.walletNum);
      
      const hour = wallet.hour || 12;
      const minute = 0;
      const datetime = formatDateTimeForBOS(
        wallet.year, 
        wallet.month, 
        wallet.day, 
        hour, 
        minute
      );
      
      const gender = 'male'; // Default gender
      
      walletsData.push({
        walletNum: wallet.walletNum,
        recipient: wallet.recipient,
        name_cn: wallet.recipient,
        datetime: datetime,
        gender: gender,
        birthday: wallet.birthday,
        birthtime: wallet.birthtime,
        hourName: wallet.hourName
      });
    }
    
    // Call external API
    Logger.log('🚀 Sending to external BOS API...');
    const apiResponse = callExternalBOSAPI({
      wallets: walletsData,
      shopify_order_id: shopifyOrderId
    });
    
    if (!apiResponse.success) {
      Logger.log('❌ External API failed: ' + apiResponse.error);
      return { 
        success: false, 
        error: 'BOS API调用失败: ' + apiResponse.error 
      };
    }
    
    // Process results
    const cards = [];
    const allCards = [];
    const detailedInfo = [];
    
    const results = apiResponse.results || [];
    
    for (let i = 0; i < results.length; i++) {
      const result = results[i];
      const walletNum = result.walletNum || (i + 1);
      const goldenCard = result.goldenCard || '离';
      
      Logger.log('   Wallet #' + walletNum + ' → ' + goldenCard);
      
      const originalWallet = walletsData[i];
      
      cards.push({
        walletNum: walletNum,
        recipient: originalWallet.recipient,
        goldenCard: goldenCard,
        hourName: originalWallet.hourName,
        birthday: originalWallet.birthday,
        birthtime: originalWallet.birthtime
      });
      
      allCards.push(goldenCard);
      
      detailedInfo.push({
        wallet: walletNum,
        recipient: originalWallet.recipient,
        birthday: originalWallet.birthday,
        birthtime: originalWallet.birthtime,
        hourName: originalWallet.hourName,
        card: goldenCard
      });
    }
    
    const formattedCards = formatCardsWithSeparator(allCards);
    
    // Cache detailed info
    const cache = CacheService.getScriptCache();
    const cacheKey = 'details_' + sheetName + '_' + rowId;
    cache.put(cacheKey, JSON.stringify(detailedInfo), 86400);
    
    // Update Google Sheets
    sh.getRange(rowId, 17).setValue('Complete');
    sh.getRange(rowId, 18).setValue(formattedCards);
    
    // Add Golden Card to Shopify
    Logger.log('🛒 Adding Golden Card products to Shopify...');
    const addProductResult = addGoldenCardToShopifyOrder(shopifyOrderId, allCards);
    
    if (!addProductResult.success) {
      Logger.log('⚠️ Shopify update failed: ' + addProductResult.error);
      sh.getRange(rowId, 15).setValue('Golden Card calculated but Shopify failed');
    } else {
      Logger.log('✅ Successfully added to Shopify');
      sh.getRange(rowId, 15).setValue('✅ Golden Cards added to Shopify Order');
    }
    
    return {
      success: true,
      cards: cards,
      shopifyUpdate: addProductResult
    };
    
  } catch (error) {
    Logger.log('❌ Error: ' + error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// SHOPIFY INTEGRATION (GraphQL)
// ============================================================

function addGoldenCardToShopifyOrder(orderIdentifier, goldenCards) {
  try {
    const conversionResult = convertOrderNameToNumericId(orderIdentifier);
    if (!conversionResult.success) {
      return { success: false, error: conversionResult.error };
    }
    
    const graphqlOrderId = conversionResult.graphqlId;
    
    const cardQuantities = {};
    for (let i = 0; i < goldenCards.length; i++) {
      const card = goldenCards[i];
      cardQuantities[card] = (cardQuantities[card] || 0) + 1;
    }
    
    const lineItemsInput = [];
    for (const card in cardQuantities) {
      const variantId = GOLDEN_CARD_VARIANTS[card];
      if (variantId) {
        lineItemsInput.push({
          variantId: 'gid://shopify/ProductVariant/' + variantId,
          quantity: cardQuantities[card]
        });
      }
    }
    
    if (lineItemsInput.length === 0) {
      return { success: false, error: 'No valid products to add' };
    }
    
    const beginMutation = 'mutation orderEditBegin($id: ID!) { orderEditBegin(id: $id) { calculatedOrder { id } userErrors { field message } } }';
    const beginResult = executeGraphQLMutation(beginMutation, { id: graphqlOrderId });
    
    if (!beginResult.success) return beginResult;
    
    const calculatedOrderId = beginResult.data.orderEditBegin.calculatedOrder.id;
    
    const addMutation = 'mutation orderEditAddVariant($id: ID!, $variantId: ID!, $quantity: Int!) { orderEditAddVariant(id: $id, variantId: $variantId, quantity: $quantity) { calculatedLineItem { id } userErrors { field message } } }';
    
    for (let i = 0; i < lineItemsInput.length; i++) {
      executeGraphQLMutation(addMutation, {
        id: calculatedOrderId,
        variantId: lineItemsInput[i].variantId,
        quantity: lineItemsInput[i].quantity
      });
    }
    
    const commitMutation = 'mutation orderEditCommit($id: ID!) { orderEditCommit(id: $id, notifyCustomer: false, staffNote: "Added Golden Cards") { order { id } userErrors { field message } } }';
    const commitResult = executeGraphQLMutation(commitMutation, { id: calculatedOrderId });
    
    if (!commitResult.success) return commitResult;
    
    return { success: true, addedItems: lineItemsInput.length };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function convertOrderNameToNumericId(orderName) {
  try {
    const cleanOrderName = orderName.replace('#', '');
    const url = 'https://' + BDAY_SHOPIFY_CONFIG.SHOP_URL + '/admin/api/' + BDAY_SHOPIFY_CONFIG.API_VERSION + '/orders.json?name=' + encodeURIComponent(cleanOrderName) + '&status=any';
    
    const response = UrlFetchApp.fetch(url, {
      'method': 'get',
      'headers': {
        'X-Shopify-Access-Token': BDAY_SHOPIFY_CONFIG.ACCESS_TOKEN,
        'Content-Type': 'application/json'
      },
      'muteHttpExceptions': true
    });
    
    if (response.getResponseCode() !== 200) {
      return { success: false, error: 'Shopify API error' };
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data.orders || data.orders.length === 0) {
      return { success: false, error: 'Order not found' };
    }
    
    return {
      success: true,
      numericId: data.orders[0].id.toString(),
      graphqlId: 'gid://shopify/Order/' + data.orders[0].id
    };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function executeGraphQLMutation(mutation, variables) {
  try {
    const url = 'https://' + BDAY_SHOPIFY_CONFIG.SHOP_URL + '/admin/api/' + BDAY_SHOPIFY_CONFIG.API_VERSION + '/graphql.json';
    
    const response = UrlFetchApp.fetch(url, {
      'method': 'post',
      'headers': {
        'X-Shopify-Access-Token': BDAY_SHOPIFY_CONFIG.ACCESS_TOKEN,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify({ query: mutation, variables: variables }),
      'muteHttpExceptions': true
    });
    
    if (response.getResponseCode() !== 200) {
      return { success: false, error: 'HTTP ' + response.getResponseCode() };
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (data.errors) {
      return { success: false, error: JSON.stringify(data.errors) };
    }
    
    const mutationKey = Object.keys(data.data)[0];
    if (data.data[mutationKey].userErrors && data.data[mutationKey].userErrors.length > 0) {
      return { success: false, error: JSON.stringify(data.data[mutationKey].userErrors) };
    }
    
    return { success: true, data: data.data };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function formatCardsWithSeparator(cards) {
  const cardCount = {};
  for (let i = 0; i < cards.length; i++) {
    const card = cards[i];
    cardCount[card] = (cardCount[card] || 0) + 1;
  }
  const formatted = [];
  for (const card in cardCount) {
    formatted.push(card + 'x' + cardCount[card]);
  }
  return formatted.join(' | ');
}

function formatDateFromString(dateStr) {
  const parts = dateStr.split('-');
  return parts.length === 3 ? parts[0] + '年' + parts[1] + '月' + parts[2] + '日' : dateStr;
}

// ============================================================
// HTML GENERATION
// ============================================================

function createResultsPage(name, goldenCardData, rowId, sheet, sheetName) {
  try {
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get('details_' + sheetName + '_' + rowId);
    
    let cardsInfo = [];
    if (cachedData) {
      try { cardsInfo = JSON.parse(cachedData); } catch (e) {}
    }
    
    let cardsHtml = '';
    if (cardsInfo && cardsInfo.length > 0) {
      for (let i = 0; i < cardsInfo.length; i++) {
        const info = cardsInfo[i];
        cardsHtml += '<div class="card-item"><div class="card-header"><span class="card-number">🎴 #【奇门遁甲 招财阵】' + info.wallet + '</span><span class="recipient-badge">' + info.recipient + '</span></div><div class="birthday-info"><p>📅 ' + formatDateFromString(info.birthday) + '</p><p>🕐 ' + (info.birthtime !== '未提供' ? info.birthtime : '未提供') + ' (' + info.hourName + ')</p></div><div class="golden-card"><h2>' + info.card + '</h2></div></div>';
      }
    } else {
      const cards = goldenCardData.split(' | ');
      for (let i = 0; i < cards.length; i++) {
        cardsHtml += '<div class="card-item"><div class="card-header"><span class="card-number">🎴 #【奇门遁甲 招财阵】' + (i + 1) + '</span></div><div class="golden-card"><h2>' + cards[i] + '</h2></div></div>';
      }
    }
    
    return HtmlService.createHtmlOutput('<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>满金包 - 命宫结果</title><style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:"Microsoft YaHei",Arial,sans-serif;background:#cca983;min-height:100vh;padding:20px}.container{max-width:600px;margin:0 auto;background:white;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden}.header{background:linear-gradient(135deg,#8a4f19 0%,#a0681f 100%);color:white;padding:40px 30px;text-align:center}.header h1{font-size:48px;margin:0;font-weight:bold;letter-spacing:8px}.header p{margin:12px 0 0 0;font-size:18px;letter-spacing:3px}.results-content{padding:30px}.card-item{background:white;border:2px solid #946c36;border-radius:12px;padding:20px;margin-bottom:20px}.card-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:15px;border-bottom:2px solid #946c36;padding-bottom:10px}.card-number{font-weight:bold;color:#333;font-size:16px}.recipient-badge{background:#542e10;color:white;padding:8px 16px;border-radius:20px;font-weight:bold;font-size:14px}.birthday-info{margin-bottom:15px;color:#333}.birthday-info p{margin:8px 0;font-size:14px}.golden-card{background:#c9a870;padding:25px;border-radius:8px;text-align:center}.golden-card h2{color:white;font-size:36px;text-shadow:1px 1px 2px rgba(0,0,0,0.3);font-weight:bold;letter-spacing:4px}.footer{background:#542e10;color:white;padding:20px;text-align:center;font-size:13px}.footer p{margin:5px 0}.footer-phones{display:flex;gap:15px;justify-content:center;margin-top:10px}</style></head><body><div class="container"><div class="header"><h1>满金包</h1><p>奇门遁甲 · 命宫结果</p></div><div class="results-content">' + cardsHtml + '</div><div class="footer"><p><strong>恭喜你！已获得专属【奇门遁甲 招财阵】！</strong></p><p><strong>这个赠品将会和钱包一起寄出。如果你有任何疑问，请联系我们的客服。</strong></p><div class="footer-phones"><span>📞 +6013-928 4699</span><span>📞 +6013-530 8863</span></div></div></div></body></html>')
      .setTitle('满金包2026 - 命宫结果')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  } catch (error) {
    return HtmlService.createHtmlOutput(createErrorPage('加载结果时出错'));
  }
}

function createErrorPage(message) {
  return '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>错误</title><style>body{font-family:"Microsoft YaHei",Arial,sans-serif;background:#cca983;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}.error-container{background:white;border-radius:20px;padding:40px;max-width:500px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,0.3)}h2{color:#E63946;margin-bottom:20px;font-size:32px}p{color:#333;font-size:18px;line-height:1.6}</style></head><body><div class="error-container"><h2>❌ 错误</h2><p>' + message + '</p></div></body></html>';
}

function createBirthdayForm(name, qty, row, orderId, token, sheetName) {
  const qtyNum = parseInt(qty) || 1;
  let formGroups = '';
  
  for (let i = 1; i <= qtyNum; i++) {
    formGroups += '<div class="wallet-group"><div class="wallet-header"><h3>#【奇门遁甲 招财阵】' + i + '</h3></div><div class="form-group"><label>👤 这个钱包是给谁使用的?</label><select id="recipient' + i + '" required><option value="">请选择...</option><option value="本人">本人 (Myself)</option><option value="爸爸">爸爸 (Father)</option><option value="妈妈">妈妈 (Mother)</option><option value="孩子">孩子 (Child)</option><option value="配偶">配偶 (Spouse)</option><option value="朋友">朋友 (Friend)</option><option value="其他">其他 (Other)</option></select></div><div class="form-group"><label>📅 出生日期</label><input type="date" id="birthday' + i + '" required></div><div class="form-group"><label>🕐 出生时间 (可选)</label><input type="time" id="birthtime' + i + '"><small style="color:#666;display:block;margin-top:5px;">如果不知道准确时间，可以留空</small></div></div>';
  }
  
  return HtmlService.createHtmlOutput('<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>满金包 - 生辰八字登记</title><style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:"Microsoft YaHei",Arial,sans-serif;background:#cca983;min-height:100vh;padding:20px}.container{max-width:600px;margin:0 auto;background:white;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden}.header{background:linear-gradient(135deg,#8a4f19 0%,#a0681f 100%);color:white;padding:40px 30px;text-align:center}.header h1{font-size:48px;margin:0;font-weight:bold;letter-spacing:8px}.header p{margin:12px 0 0 0;font-size:18px;letter-spacing:3px}.customer-info{background:#b88f51;border-left:4px solid #946c36;padding:15px;margin:15px;border-radius:6px}.customer-info p{margin:8px 0;font-size:14px;color:white;font-weight:500}.security-warning{background:#fff9e6;border-left:4px solid #946c36;padding:12px 15px;margin:15px;border-radius:6px;font-size:13px}.security-warning p{margin:6px 0;color:#333}.form-section{padding:30px}.wallet-group{background:white;padding:25px;border-radius:10px;margin-bottom:20px;border:2px solid #b88f51}.wallet-header{border-bottom:3px solid #b88f51;padding-bottom:12px;margin-bottom:18px}.wallet-header h3{color:#542e10;font-size:16px;font-weight:bold}.form-group{margin-bottom:20px}label{display:block;font-weight:600;margin-bottom:8px;color:#542e10;font-size:15px}input,select{width:100%;padding:12px;border:2px solid #ddd;border-radius:8px;font-size:15px;background:white}input:focus,select:focus{outline:none;border-color:#b88f51;box-shadow:0 0 6px rgba(184,143,81,0.6)}.submit-btn{width:100%;padding:18px;background:#E63946;color:white;border:none;border-radius:10px;font-size:24px;font-weight:bold;cursor:pointer;margin-top:15px;transition:background 0.3s}.submit-btn:hover{background:#D62828;transform:translateY(-2px);box-shadow:0 6px 16px rgba(230,57,70,0.3)}.submit-btn:disabled{background:#ccc;cursor:not-allowed;transform:none}.loading-overlay{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.8);z-index:9999;justify-content:center;align-items:center}.loading-container{display:flex;flex-direction:column;align-items:center;justify-content:center}.spinner{width:60px;height:60px;border:4px solid rgba(255,255,255,0.3);border-top:4px solid white;border-radius:50%;animation:spin 1s linear infinite}.progress-bar{width:350px;height:10px;background:rgba(255,255,255,0.3);border-radius:10px;overflow:hidden;margin:25px auto}.progress-fill{height:100%;background:linear-gradient(90deg,#b88f51,#946c36,#542e10);border-radius:10px;animation:progress 1.5s ease-out forwards}.loading-text{color:white;font-size:18px;margin-top:25px;font-weight:bold}@keyframes spin{0%{transform:rotate(0deg)}100%{transform:rotate(360deg)}}@keyframes progress{0%{width:0%}100%{width:100%}}</style></head><body><div class="loading-overlay" id="loadingOverlay"><div class="loading-container"><div class="spinner"></div><div class="progress-bar"><div class="progress-fill"></div></div><div class="loading-text">✨ 正在计算您的命宫...</div></div></div><div class="container"><div class="header"><h1>满金包</h1><p>奇门遁甲 · 生辰八字登记</p></div><div class="customer-info"><p><strong>👤 姓名:</strong> ' + name + '</p><p><strong>🎁 数量:</strong> ' + qtyNum + ' 个钱包</p></div><div class="security-warning"><p><strong>隐私保护：</strong></p><p>• 你提供的资料（姓名、出生日期、出生时间、出生地点等）将被严格保密，不会对外公开或与第三方共享。</p><p>• 资料仅用于个人八字分析与能量评估，不作其他商业用途。</p><p>• 我们会安全保存资料，并于分析完成后加密或删除。</p><p>• 提交资料即表示你自愿提供并同意以上条款，分析结果仅供参考。</p></div><div class="form-section"><form id="birthdayForm">' + formGroups + '<button type="submit" class="submit-btn" id="submitBtn">马上提交计算命宫</button></form></div></div><script>const rowId="' + row + '";const qty=' + qtyNum + ';const token="' + token + '";const sheetName="' + sheetName + '";function timeToHour(t){if(!t)return 6;const h=parseInt(t.split(":")[0]);if(h>=23||h<1)return 0;if(h>=1&&h<3)return 1;if(h>=3&&h<5)return 2;if(h>=5&&h<7)return 3;if(h>=7&&h<9)return 4;if(h>=9&&h<11)return 5;if(h>=11&&h<13)return 6;if(h>=13&&h<15)return 7;if(h>=15&&h<17)return 8;if(h>=17&&h<19)return 9;if(h>=19&&h<21)return 10;if(h>=21&&h<23)return 11;return 6}const hourNames=["子时","丑时","寅时","卯时","辰时","巳时","午时","未时","申时","酉时","戌时","亥时"];function formatDateFromString(dateStr){const parts=dateStr.split("-");if(parts.length===3){return parts[0]+"年"+parts[1]+"月"+parts[2]+"日"}return dateStr}function displayResults(cards){let cardsHtml="";for(let i=0;i<cards.length;i++){const card=cards[i];const birthdateFormatted=formatDateFromString(card.birthday);const birthtimeDisplay=card.birthtime!=="未提供"?card.birthtime:"未提供";cardsHtml+=\'<div class="card-item">\'+\'<div class="card-header">\'+\'<span class="card-number">🎴 #【奇门遁甲 招财阵】\'+card.walletNum+\'</span>\'+\'<span class="recipient-badge">\'+card.recipient+\'</span>\'+\'</div>\'+\'<div class="birthday-info">\'+\'<p>📅 \'+birthdateFormatted+\'</p>\'+\'<p>🕐 \'+birthtimeDisplay+\' (\'+card.hourName+\')</p>\'+\'</div>\'+\'<div class="golden-card">\'+\'<h2>\'+card.goldenCard+\'</h2>\'+\'</div>\'+\'</div>\'}const resultsHtml=\'<div class="results-content">\'+cardsHtml+\'</div>\'+\'<div class="footer">\'+\'<p><strong>恭喜你！已获得专属【奇门遁甲 招财阵】！</strong></p>\'+\'<p><strong>这个赠品将会和钱包一起寄出。如果你有任何疑问，请联系我们的客服。</strong></p>\'+\'<div class="footer-phones">\'+\'<span class="phone-item">📞 +6013-928 4699</span>\'+\'<span class="phone-item">📞 +6013-530 8863</span>\'+\'</div>\'+\'</div>\';const additionalStyles=\'<style>.results-content{padding:30px}.card-item{background:white;border:2px solid #946c36;border-radius:12px;padding:20px;margin-bottom:20px}.card-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:15px;border-bottom:2px solid #946c36;padding-bottom:10px}.card-number{font-weight:bold;color:#333;font-size:16px}.recipient-badge{background:#542e10;color:white;padding:8px 16px;border-radius:20px;font-weight:bold;font-size:14px}.birthday-info{margin-bottom:15px;color:#333}.birthday-info p{margin:8px 0;font-size:14px}.golden-card{background:#c9a870;padding:25px;border-radius:8px;text-align:center;max-width:100%}.golden-card h2{color:white;font-size:36px;text-shadow:1px 1px 2px rgba(0,0,0,0.3);font-weight:bold;letter-spacing:4px}.footer{background:#542e10;color:white;padding:20px;text-align:center;font-size:13px}.footer p{margin:5px 0}</style>\';document.head.insertAdjacentHTML("beforeend",additionalStyles);document.querySelector(".container").innerHTML=\'<div class="header">\'+\'<h1>满金包 2026</h1>\'+\'<p>奇门遁甲 · 命宫结果</p>\'+\'</div>\'+resultsHtml}document.getElementById("birthdayForm").addEventListener("submit",function(e){e.preventDefault();const submitBtn=document.getElementById("submitBtn");const loadingOverlay=document.getElementById("loadingOverlay");const wallets=[];for(let i=1;i<=qty;i++){const recipient=document.getElementById("recipient"+i).value;const birthday=document.getElementById("birthday"+i).value;const birthtime=document.getElementById("birthtime"+i).value;if(!recipient){alert("请选择钱包 #"+i+" 是给谁的");return}if(!birthday){alert("请填写钱包 #"+i+" 的出生日期");return}const dateObj=new Date(birthday+"T00:00:00");const year=dateObj.getFullYear();const month=dateObj.getMonth()+1;const day=dateObj.getDate();const hasTime=birthtime?true:false;const hourIndex=timeToHour(birthtime);wallets.push({walletNum:i,recipient:recipient,year:year,month:month,day:day,hour:hourIndex,hourName:hasTime?hourNames[hourIndex]:"未提供",birthday:birthday,birthtime:birthtime||"未提供",hasTime:hasTime})}submitBtn.disabled=true;loadingOverlay.style.display="flex";const data={wallets:wallets,rowId:rowId,qty:qty,token:token,sheetName:sheetName};google.script.run.withSuccessHandler(function(result){setTimeout(function(){if(result.success){loadingOverlay.style.display="none";displayResults(result.cards)}else{loadingOverlay.style.display="none";submitBtn.disabled=false;alert("提交失败："+result.error)}},1500)}).withFailureHandler(function(error){loadingOverlay.style.display="none";submitBtn.disabled=false;alert("提交失败："+error.message)}).processFormSubmission(data)});</script></body></html>')
    .setTitle('满金包2026 - 生辰八字登记')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
