const BDAY_SETTINGS = {
  GMAIL_ADDRESS: 'customercare@mandarin.club',
  FORM_BASE_URL: 'https://script.google.com/macros/s/AKfycbzMxrnjDSq2u7bnf51WlrTHu-fH93LbIYFYEGx5Kw5LWagMMxpq8SLScNOP1E8dp3WAzg/exec',
  SHEET_NAME: 'Orders',
  BATCH_LIMIT: 5,
  SKIP_SINGLE_WALLET: true,
  COLUMNS: {
    TIME_STAMP: 1,
    ORDER_ID: 2,
    NAME: 3,
    EMAIL: 4,
    PHONE: 5,
    ADDRESS: 6,
    CITY: 7,
    STATE: 8,
    POSTCODE: 9,
    MAIN_PRODUCT: 10,
    QUANTITY: 11,
    ORDER_SUMMARY: 12,
    TOTAL_PRICE: 13,
    STATUS: 14,
    ERROR_MESSAGE: 15,
    SHOPIFY_ORDER_ID: 16,
    GOLDEN_CARD_STATUS: 17,
    GOLDEN_CARD: 18,
    FORM_ACCESS_TOKEN: 19
  }
};

function generateAccessToken(rowId, orderId) {
  const timestamp = new Date().getTime();
  const randomStr = Utilities.getUuid().substring(0, 8);
  const token = Utilities.base64Encode(rowId + '|' + orderId + '|' + timestamp + '|' + randomStr);
  return token;
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
  if (!orderSummary || orderSummary === '') {
    return 0;
  }
  
  let totalQty = 0;
  const parts = smartSplit(orderSummary);
  
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i].trim();
    const matches = part.match(/[xX×]\s*(\d+)\s*$/);
    
    if (matches && matches[1]) {
      const qty = parseInt(matches[1]);
      totalQty += qty;
      Logger.log('   📦 Found: ' + part + ' -> Qty: ' + qty);
    }
  }
  
  Logger.log('📊 Total wallets detected: ' + totalQty);
  return totalQty;
}

function processNextBdayOrder() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(BDAY_SETTINGS.SHEET_NAME);
    
    if (!sh) {
      Logger.log('❌ Sheet not found: ' + BDAY_SETTINGS.SHEET_NAME);
      return;
    }
    
    const lastRow = sh.getLastRow();
    Logger.log('🔍 Checking sheet: ' + BDAY_SETTINGS.SHEET_NAME + ', Last row: ' + lastRow);
    
    for (let i = 2; i <= lastRow; i++) {
      const goldenCardStatus = sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_STATUS).getValue();
      
      if (goldenCardStatus === '') {
        const name = sh.getRange(i, BDAY_SETTINGS.COLUMNS.NAME).getValue();
        const email = sh.getRange(i, BDAY_SETTINGS.COLUMNS.EMAIL).getValue();
        const orderSummary = sh.getRange(i, BDAY_SETTINGS.COLUMNS.ORDER_SUMMARY).getValue();
        const orderId = sh.getRange(i, BDAY_SETTINGS.COLUMNS.ORDER_ID).getValue();
        
        const qty = smartSplitQty(orderSummary);
        
        if (name && email && qty > 0) {
          Logger.log('📧 Processing next pending order - Row ' + i);
          Logger.log('   Order: ' + orderId + ' - ' + name + ' (Qty: ' + qty + ')');
          
          if (BDAY_SETTINGS.SKIP_SINGLE_WALLET && qty === 1) {
            sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_STATUS).setValue('Skipped');
            sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD).setValue('N/A - Single Wallet');
            Logger.log('⏭️ Skipped single wallet order (qty=1)');
            return;
          }
          
          const token = generateAccessToken(i, orderId);
          const formUrl = generateFormUrl(name, qty, i, orderId, token);
          
          if (sendBdayEmail(name, email, qty, i, orderId, formUrl)) {
            sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_STATUS).setValue('Pending');
            sh.getRange(i, BDAY_SETTINGS.COLUMNS.FORM_ACCESS_TOKEN).setValue(formUrl);
            Logger.log('✅ Email sent to ' + email);
          } else {
            sh.getRange(i, BDAY_SETTINGS.COLUMNS.ERROR_MESSAGE).setValue('Email Failed');
            Logger.log('❌ Failed to send email to: ' + email);
          }
          
          return;
        }
      }
    }
    
    Logger.log('✅ All Done - No pending orders found');
  } catch (error) {
    Logger.log('❌ Error in processNextBdayOrder: ' + error);
  }
}

function generateFormUrl(name, qty, rowId, orderId, token) {
  return BDAY_SETTINGS.FORM_BASE_URL + 
    '?name=' + encodeURIComponent(name) + 
    '&qty=' + qty + 
    '&row=' + rowId + 
    '&order=' + encodeURIComponent(orderId) +
    '&token=' + encodeURIComponent(token);
}

function processAllBdayOrders() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(BDAY_SETTINGS.SHEET_NAME);
    
    if (!sh) {
      Logger.log('❌ Sheet not found: ' + BDAY_SETTINGS.SHEET_NAME);
      SpreadsheetApp.getUi().alert('Error', 'Sheet not found: ' + BDAY_SETTINGS.SHEET_NAME, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const lastRow = sh.getLastRow();
    Logger.log('🔍 Checking sheet: ' + BDAY_SETTINGS.SHEET_NAME + ', Last row: ' + lastRow);
    
    if (lastRow < 2) {
      Logger.log('ℹ️ No orders to process');
      SpreadsheetApp.getUi().alert('Info', 'No orders to process', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const batchLimit = BDAY_SETTINGS.BATCH_LIMIT;
    Logger.log('📊 Batch limit set to: ' + batchLimit + ' emails');
    
    let processed = 0;
    let errors = 0;
    let skipped = 0;
    let singleWalletSkipped = 0;
    
    for (let i = 2; i <= lastRow; i++) {
      if (processed >= batchLimit) {
        Logger.log('⏸️ Batch limit reached (' + batchLimit + ' emails sent)');
        break;
      }
      
      const goldenCardStatus = sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_STATUS).getValue();
      
      if (goldenCardStatus === '') {
        const name = sh.getRange(i, BDAY_SETTINGS.COLUMNS.NAME).getValue();
        const email = sh.getRange(i, BDAY_SETTINGS.COLUMNS.EMAIL).getValue();
        const orderSummary = sh.getRange(i, BDAY_SETTINGS.COLUMNS.ORDER_SUMMARY).getValue();
        const orderId = sh.getRange(i, BDAY_SETTINGS.COLUMNS.ORDER_ID).getValue();
        
        const qty = smartSplitQty(orderSummary);
        
        if (name && email && qty > 0) {
          if (BDAY_SETTINGS.SKIP_SINGLE_WALLET && qty === 1) {
            Logger.log('\n⏭️ Skipping row ' + i + ' (single wallet order)');
            Logger.log('   Order: ' + orderId + ' - ' + name + ' (Qty: ' + qty + ')');
            sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_STATUS).setValue('Skipped');
            sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD).setValue('N/A - Single Wallet');
            singleWalletSkipped++;
            continue;
          }
          
          Logger.log('\n📧 Processing row ' + i + ' (' + (processed + 1) + ' of ' + batchLimit + ')');
          Logger.log('   Order: ' + orderId + ' - ' + name + ' (Qty: ' + qty + ')');
          
          try {
            const token = generateAccessToken(i, orderId);
            const formUrl = generateFormUrl(name, qty, i, orderId, token);
            
            if (sendBdayEmail(name, email, qty, i, orderId, formUrl)) {
              sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_STATUS).setValue('Pending');
              sh.getRange(i, BDAY_SETTINGS.COLUMNS.FORM_ACCESS_TOKEN).setValue(formUrl);
              Logger.log('✅ Email sent to ' + email);
              processed++;
              Utilities.sleep(1000);
            } else {
              sh.getRange(i, BDAY_SETTINGS.COLUMNS.ERROR_MESSAGE).setValue('Email Failed');
              Logger.log('❌ Failed to send email to: ' + email);
              errors++;
            }
          } catch (error) {
            errors++;
            Logger.log('❌ Error on row ' + i + ': ' + error);
            sh.getRange(i, BDAY_SETTINGS.COLUMNS.ERROR_MESSAGE).setValue('Error: ' + error.toString());
          }
        }
      } else if (goldenCardStatus === 'Pending' || goldenCardStatus === 'Complete' || goldenCardStatus === 'Skipped') {
        skipped++;
      }
    }
    
    let message = '🎉 Batch Processing Complete\n\n';
    message += '✅ Emails Sent: ' + processed + '\n';
    if (singleWalletSkipped > 0) {
      message += '⏭️ Single Wallets Auto-Skipped: ' + singleWalletSkipped + '\n';
    }
    message += '❌ Errors: ' + errors + '\n';
    if (skipped > 0) {
      message += '⏭️ Already Processed: ' + skipped + '\n';
    }
    message += '\n📊 Batch Limit: ' + batchLimit + ' emails per run';
    
    SpreadsheetApp.getUi().alert('✅ Process Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    Logger.log('❌ Error in processAllBdayOrders: ' + error);
    SpreadsheetApp.getUi().alert('Error', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function sendBdayEmail(name, email, qty, rowId, orderId, formUrl) {
  try {
    Logger.log('📧 Sending email with form URL');
    
    const subject = '🎁 完成您的' + qty + '个钱包订单 - 请填写生日资料';
    
    const htmlBody = '<!DOCTYPE html>' +
      '<html><head><meta charset="UTF-8"><style>' +
      'body{font-family:Arial,sans-serif;margin:0;padding:0;background:#f5f5f5}' +
      '.container{max-width:600px;margin:0 auto;background:white}' +
      '.header{background:#667eea;color:white;padding:30px;text-align:center}' +
      '.header h1{margin:0;font-size:28px}' +
      '.header p{margin:5px 0 0 0;font-size:14px}' +
      '.customer-info{background:#e3f2fd;border-left:4px solid #1976d2;padding:15px;margin:20px}' +
      '.customer-info p{margin:8px 0;font-size:14px;color:#333}' +
      '.content{padding:30px 20px}' +
      '.section{margin:20px 0}' +
      '.section-title{font-size:16px;font-weight:bold;color:#333;margin-bottom:10px}' +
      '.requirement{margin:8px 0;font-size:14px;color:#555}' +
      '.requirement-item{margin-left:20px}' +
      '.button-container{text-align:center;margin:30px 0}' +
      '.button{display:inline-block;padding:14px 40px;background:#667eea;color:white;text-decoration:none;border-radius:6px;font-weight:bold;font-size:16px}' +
      '.button:hover{background:#5568d3}' +
      '.warning{background:#fff3cd;border-left:4px solid #ffc107;padding:12px;margin:20px;font-size:13px;color:#856404}' +
      '.footer{background:#f8f8f8;padding:20px;text-align:center;border-top:1px solid #eee;font-size:12px;color:#666}' +
      '.footer p{margin:5px 0}' +
      '</style></head><body>' +
      '<div class="container">' +
      '<div class="header">' +
      '<h1>满金包 2026</h1>' +
      '<p>奇门遁甲 · 招财阵定制</p>' +
      '</div>' +
      '<div class="customer-info">' +
      '<p><strong>👤 客户姓名：</strong>' + name + '</p>' +
      '<p><strong>🎁 订购数量：</strong>' + qty + ' 个钱包</p>' +
      '</div>' +
      '<div class="content">' +
      '<div class="section">' +
      '<div class="section-title">你好' + name + '，</div>' +
      '<p style="font-size:14px;color:#555;line-height:1.6">感谢您的订购！为了为您计算专属的<strong>命宫</strong>和<strong>招财阵</strong>，我们需要您的生辰八字信息。</p>' +
      '</div>' +
      '<div class="section">' +
      '<div class="section-title">📋 请提供以下信息：</div>' +
      '<div class="requirement-item">' +
      '<p class="requirement">✓ 出生年月日</p>' +
      '<p class="requirement">✓ 出生时辰（可选）</p>' +
      '</div>' +
      '</div>' +
      '<div class="button-container">' +
      '<a href="' + formUrl + '" class="button">👉 马上填写</a>' +
      '</div>' +
      '<div class="warning">' +
      '<p><strong>⏰ 链接有效期：24小时</strong></p>' +
      '<p>此链接请在24小时内完成填写。完成后可随时通过此链接查看结果。</p>' +
      '</div>' +
      '<p style="font-size:13px;color:#666;margin-top:20px">系统将自动通过奇门遁甲算法计算您的命宫，并为您匹配最适合的招财阵。</p>' +
      '</div>' +
      '<div class="footer">' +
      '<p><strong>若有任何疑问，请联系我们：</strong></p>' +
      '<p>📞 +6013-928 4699</p>' +
      '<p>📞 +6013-530 8863</p>' +
      '<p style="margin-top:15px;color:#999">此邮件由满金包官方系统发送，请勿直接回复。</p>' +
      '</div>' +
      '</div></body></html>';
    
    const plainText = '满金包 2026 - 生辰八字信息填写\n\n' +
      '亲爱的 ' + name + '，\n\n' +
      '订购数量：' + qty + ' 个钱包\n\n' +
      '感谢您的订购！为了为您计算专属的命宫和招财阵，我们需要您的生辰八字信息。\n\n' +
      '请提供以下信息：\n' +
      '✓ 出生年月日\n' +
      '✓ 出生时辰（可选）\n\n' +
      '请点击以下链接填写（24小时内有效）：\n' + formUrl + '\n\n' +
      '⏰ 重要提示：\n' +
      '此链接请在24小时内完成填写。完成后可随时通过此链接查看结果。\n\n' +
      '若有任何疑问，请联系客服：\n' +
      '📞 +6013-928 4699\n' +
      '📞 +6013-530 8863\n\n' +
      '此邮件由满金包官方系统发送，请勿直接回复。';
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: plainText,
      htmlBody: htmlBody,
      name: 'Mandarin Club - 满金包',
      replyTo: BDAY_SETTINGS.GMAIL_ADDRESS,
      charset: 'UTF-8',
      noReply: false
    });
    
    return true;
  } catch (error) {
    Logger.log('❌ Error in sendBdayEmail: ' + error);
    return false;
  }
}

function authorizeBdayEmail() {
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: '✅ Email Authorization Successful - Mandarin Club',
    body: 'Your Google Apps Script now has permission to send emails!\n\nMandarin Club Birthday Form System',
    name: 'Mandarin Club - 满金包'
  });
  Logger.log('✅ Authorization complete!');
}
