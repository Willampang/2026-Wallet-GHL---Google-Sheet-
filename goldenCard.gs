//GoldenCard.gs - UPDATED VERSION
const BDAY_SETTINGS = {
  GMAIL_ADDRESS: 'customercare@mandarin.club',
  FORM_BASE_URL: 'https://script.google.com/macros/s/AKfycbyGABN4ZFSzpkqMfpvMN2h6sXxbOW6cGEX7R6av37FaTEk1LoJnE8w1cur14Bl6N-bRHg/exec',
  SHEET_NAMES: ['Orders', 'Orders_2', 'Orders_3'],
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
    GOLDEN_CARD_LINK: 19
  }
};

function generateAccessToken(rowId, orderId, sheetName) {
  const timestamp = new Date().getTime();
  const randomStr = Utilities.getUuid().substring(0, 8);
  const token = Utilities.base64Encode(sheetName + '|' + rowId + '|' + orderId + '|' + timestamp + '|' + randomStr);
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
    
    // Check for bundle products
    let walletCount = 0;
    
    // F款: 带财款x1+吸金款x1 = 2 wallets (D + E)
    if (part.includes('F款') && part.includes('带财款') && part.includes('吸金款')) {
      walletCount = 2;
      Logger.log('   🎁 Bundle F款 detected: 2 wallets (D + E)');
    }
    // G款: 带财款x2 = 2 wallets (D x2)
    else if (part.includes('G款') && part.includes('带财款x2')) {
      walletCount = 2;
      Logger.log('   🎁 Bundle G款 detected: 2 wallets (D x2)');
    }
    // H款: 吸金款x2 = 2 wallets (E x2)
    else if (part.includes('H款') && part.includes('吸金款x2')) {
      walletCount = 2;
      Logger.log('   🎁 Bundle H款 detected: 2 wallets (E x2)');
    }
    // Regular products: extract quantity from "x数字"
    else {
      const matches = part.match(/[xX×]\s*(\d+)\s*$/);
      if (matches && matches[1]) {
        walletCount = parseInt(matches[1]);
      }
    }
    
    totalQty += walletCount;
    Logger.log('   📦 Found: ' + part + ' -> Qty: ' + walletCount);
  }
  
  Logger.log('📊 Total wallets detected: ' + totalQty);
  return totalQty;
}

function processNextBdayOrder() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    for (let s = 0; s < BDAY_SETTINGS.SHEET_NAMES.length; s++) {
      const sheetName = BDAY_SETTINGS.SHEET_NAMES[s];
      const sh = ss.getSheetByName(sheetName);
      
      if (!sh) {
        Logger.log('⚠️ Sheet not found: ' + sheetName);
        continue;
      }
      
      const lastRow = sh.getLastRow();
      Logger.log('🔍 Checking sheet: ' + sheetName + ', Last row: ' + lastRow);
      
      for (let i = 2; i <= lastRow; i++) {
        const goldenCardStatus = sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_STATUS).getValue();
        
        if (goldenCardStatus === '') {
          const name = sh.getRange(i, BDAY_SETTINGS.COLUMNS.NAME).getValue();
          const email = sh.getRange(i, BDAY_SETTINGS.COLUMNS.EMAIL).getValue();
          const orderSummary = sh.getRange(i, BDAY_SETTINGS.COLUMNS.ORDER_SUMMARY).getValue();
          const shopifyOrderId = sh.getRange(i, BDAY_SETTINGS.COLUMNS.SHOPIFY_ORDER_ID).getValue();
          
          const qty = smartSplitQty(orderSummary);
          
          if (name && email && qty > 0 && shopifyOrderId) {
            Logger.log('📧 Processing next pending order - Sheet: ' + sheetName + ', Row ' + i);
            Logger.log('   Shopify Order: ' + shopifyOrderId + ' - ' + name + ' (Qty: ' + qty + ')');
            
            if (BDAY_SETTINGS.SKIP_SINGLE_WALLET && qty === 1) {
              sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_STATUS).setValue('Skipped');
              sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD).setValue('N/A - Single Wallet');
              Logger.log('⏭️ Skipped single wallet order (qty=1)');
              return;
            }
            
            const token = generateAccessToken(i, shopifyOrderId, sheetName);
            const formUrl = generateFormUrl(name, i, shopifyOrderId, token, sheetName);
            
            if (sendBdayEmail(name, email, qty, i, shopifyOrderId, formUrl)) {
              sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_STATUS).setValue('Pending');
              sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_LINK).setValue(formUrl);
              Logger.log('✅ Email sent to ' + email);
            } else {
              sh.getRange(i, BDAY_SETTINGS.COLUMNS.ERROR_MESSAGE).setValue('Email Failed');
              Logger.log('❌ Failed to send email to: ' + email);
            }
            
            return;
          } else {
            Logger.log('⚠️ Row ' + i + ' missing required data (name, email, qty>0, or Shopify Order ID)');
          }
        }
      }
    }
    
    Logger.log('✅ All Done - No pending orders found in any sheet');
  } catch (error) {
    Logger.log('❌ Error in processNextBdayOrder: ' + error);
  }
}

function generateFormUrl(name, rowId, orderId, token, sheetName) {
  return BDAY_SETTINGS.FORM_BASE_URL + 
    '?name=' + encodeURIComponent(name) + 
    '&row=' + rowId + 
    '&order=' + encodeURIComponent(orderId) +
    '&token=' + encodeURIComponent(token) +
    '&sheet=' + encodeURIComponent(sheetName);
}

function processAllBdayOrders() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const batchLimit = BDAY_SETTINGS.BATCH_LIMIT;
    
    let totalProcessed = 0;
    let totalErrors = 0;
    let totalSkipped = 0;
    let totalSingleWalletSkipped = 0;
    
    Logger.log('📊 Batch limit set to: ' + batchLimit + ' emails');
    
    for (let s = 0; s < BDAY_SETTINGS.SHEET_NAMES.length; s++) {
      const sheetName = BDAY_SETTINGS.SHEET_NAMES[s];
      const sh = ss.getSheetByName(sheetName);
      
      if (!sh) {
        Logger.log('⚠️ Sheet not found: ' + sheetName);
        continue;
      }
      
      const lastRow = sh.getLastRow();
      Logger.log('\n🔍 Processing sheet: ' + sheetName + ', Last row: ' + lastRow);
      
      if (lastRow < 2) {
        Logger.log('ℹ️ No orders to process in ' + sheetName);
        continue;
      }
      
      for (let i = 2; i <= lastRow; i++) {
        if (totalProcessed >= batchLimit) {
          Logger.log('⏸️ Batch limit reached (' + batchLimit + ' emails sent)');
          break;
        }
        
        const goldenCardStatus = sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_STATUS).getValue();
        
        if (goldenCardStatus === '') {
          const name = sh.getRange(i, BDAY_SETTINGS.COLUMNS.NAME).getValue();
          const email = sh.getRange(i, BDAY_SETTINGS.COLUMNS.EMAIL).getValue();
          const orderSummary = sh.getRange(i, BDAY_SETTINGS.COLUMNS.ORDER_SUMMARY).getValue();
          const shopifyOrderId = sh.getRange(i, BDAY_SETTINGS.COLUMNS.SHOPIFY_ORDER_ID).getValue();
          
          const qty = smartSplitQty(orderSummary);
          
          if (name && email && qty > 0 && shopifyOrderId) {
            if (BDAY_SETTINGS.SKIP_SINGLE_WALLET && qty === 1) {
              Logger.log('\n⏭️ Skipping ' + sheetName + ' row ' + i + ' (single wallet order)');
              Logger.log('   Shopify Order: ' + shopifyOrderId + ' - ' + name + ' (Qty: ' + qty + ')');
              sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_STATUS).setValue('Skipped');
              sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD).setValue('N/A - Single Wallet');
              totalSingleWalletSkipped++;
              continue;
            }
            
            Logger.log('\n📧 Processing ' + sheetName + ' row ' + i + ' (' + (totalProcessed + 1) + ' of ' + batchLimit + ')');
            Logger.log('   Shopify Order: ' + shopifyOrderId + ' - ' + name + ' (Qty: ' + qty + ')');
            
            try {
              const token = generateAccessToken(i, shopifyOrderId, sheetName);
              const formUrl = generateFormUrl(name, i, shopifyOrderId, token, sheetName);
              
              if (sendBdayEmail(name, email, qty, i, shopifyOrderId, formUrl)) {
                sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_STATUS).setValue('Pending');
                sh.getRange(i, BDAY_SETTINGS.COLUMNS.GOLDEN_CARD_LINK).setValue(formUrl);
                Logger.log('✅ Email sent to ' + email);
                totalProcessed++;
                Utilities.sleep(1000);
              } else {
                sh.getRange(i, BDAY_SETTINGS.COLUMNS.ERROR_MESSAGE).setValue('Email Failed');
                Logger.log('❌ Failed to send email to: ' + email);
                totalErrors++;
              }
            } catch (error) {
              totalErrors++;
              Logger.log('❌ Error on ' + sheetName + ' row ' + i + ': ' + error);
              sh.getRange(i, BDAY_SETTINGS.COLUMNS.ERROR_MESSAGE).setValue('Error: ' + error.toString());
            }
          } else {
            Logger.log('⚠️ ' + sheetName + ' row ' + i + ' missing required data');
          }
        } else if (goldenCardStatus === 'Pending' || goldenCardStatus === 'Complete' || goldenCardStatus === 'Skipped') {
          totalSkipped++;
        }
      }
      
      if (totalProcessed >= batchLimit) {
        break;
      }
    }
    
    let message = '🎉 Batch Processing Complete\n\n';
    message += '✅ Emails Sent: ' + totalProcessed + '\n';
    if (totalSingleWalletSkipped > 0) {
      message += '⏭️ Single Wallets Auto-Skipped: ' + totalSingleWalletSkipped + '\n';
    }
    message += '❌ Errors: ' + totalErrors + '\n';
    if (totalSkipped > 0) {
      message += '⏭️ Already Processed: ' + totalSkipped + '\n';
    }
    message += '\n📊 Batch Limit: ' + batchLimit + ' emails per run';
    message += '\n📋 Sheets Processed: ' + BDAY_SETTINGS.SHEET_NAMES.join(', ');
    
    SpreadsheetApp.getUi().alert('✅ Process Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    Logger.log('❌ Error in processAllBdayOrders: ' + error);
    SpreadsheetApp.getUi().alert('Error', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function sendBdayEmail(name, email, qty, rowId, orderId, formUrl) {
  try {
    Logger.log('📧 Sending email with form URL');
    
    const subject = '完成您的满金包订单 - 请填写生日资料 (订单 #' + orderId + ')';
    
    const htmlBody = '<!DOCTYPE html>' +
      '<html lang="zh-CN"><head><meta charset="UTF-8">' +
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
      '<title>满金包订单确认</title></head><body style="margin:0;padding:0;font-family:Arial,sans-serif;background-color:#f5f5f5">' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="background-color:#f5f5f5;padding:20px 0">' +
      '<tr><td align="center">' +
      '<table width="600" cellpadding="0" cellspacing="0" style="background-color:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1)">' +
      
      '<tr><td style="padding:0;text-align:center;border-radius:0">' +
      '<div style="background:#8a4f19;background:linear-gradient(135deg,#8a4f19 0%,#a0681f 100%);padding:30px;text-align:center;">' +
      '<h1 style="margin:0;color:#000000;font-size:28px;font-weight:bold;letter-spacing:0.5px;">满金包 2026</h1>' +
      '<p style="margin:8px 0 0 0;color:#000000;font-size:14px;font-weight:400;letter-spacing:1px;opacity:0.9;">奇门遁甲 · 招财阵定制</p>' +
      '</div>' +
      '</td></tr>' +
      
      '<tr><td style="padding:30px">' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f9f9f9;border-left:4px solid #b88f51;padding:15px;border-radius:4px">' +
      '<tr><td><p style="margin:8px 0;font-size:14px;color:#333"><strong>👤 尊敬的客户：</strong>' + name + '</p>' +
      '<p style="margin:8px 0;font-size:14px;color:#333"><strong>📦 订单编号：</strong>' + orderId + '</p>' +
      '<p style="margin:8px 0;font-size:14px;color:#333"><strong>🎁 订购数量：</strong>' + qty + ' 个钱包</p></td></tr>' +
      '</table>' +
      '</td></tr>' +
      
      '<tr><td style="padding:0 30px 20px 30px">' +
      '<h2 style="color:#333;font-size:18px;margin:0 0 15px 0">感谢您的订购！</h2>' +
      '<p style="font-size:14px;color:#555;line-height:1.6;margin:0 0 15px 0">为了为您计算专属的<strong>命宫</strong>和<strong>招财阵</strong>，我们需要您提供以下信息：</p>' +
      '<ul style="font-size:14px;color:#555;line-height:1.8;margin:0 0 20px 20px">' +
      '<li>出生年月日（必填）</li>' +
      '<li>出生时辰（选填，但可以提高准确度）</li>' +
      '</ul>' +
      '</td></tr>' +
      
      '<tr><td style="padding:0 30px 30px 30px;text-align:center">' +
      '<a href="' + formUrl + '" style="display:inline-block;padding:14px 40px;background:#E63946;color:#ffffff;text-decoration:none;border-radius:6px;font-weight:bold;font-size:16px">马上填写</a>' +
      '</td></tr>' +
      
      '<tr><td style="padding:0 30px 20px 30px">' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff9e6;border-left:4px solid #b88f51;padding:15px;border-radius:4px">' +
      '<tr><td><p style="margin:0 0 8px 0;font-size:14px;color:#8B4513"><strong>⏰ 重要提示：</strong></p>' +
      '<p style="margin:0;font-size:13px;color:#8B4513;line-height:1.6">此链接有效期为24小时，请尽快填写。完成后可随时通过此链接查看您的命宫结果。</p></td></tr>' +
      '</table>' +
      '</td></tr>' +
      
      '<tr><td style="background:#f9f9f9;padding:20px 30px;border-top:1px solid #e0e0e0;text-align:center">' +
      '<p style="margin:0 0 10px 0;font-size:13px;color:#666"><strong>客服联系方式</strong></p>' +
      '<p style="margin:5px 0;font-size:13px;color:#666">📞 +6013-928 4699 | +6013-530 8863</p>' +
      '<p style="margin:5px 0;font-size:13px;color:#666">📧 customercare@mandarin.club</p>' +
      '<p style="margin:15px 0 0 0;font-size:12px;color:#999">此邮件由 Mandarin Club 官方系统自动发送</p>' +
      '</td></tr>' +
      
      '</table>' +
      '</td></tr></table>' +
      '</body></html>';
    
    const plainText = '满金包 2026 - 订单确认\n\n' +
      '尊敬的 ' + name + ' 您好，\n\n' +
      '订单编号：' + orderId + '\n' +
      '订购数量：' + qty + ' 个钱包\n\n' +
      '感谢您的订购！为了为您计算专属的命宫和招财阵，我们需要您的生辰八字信息。\n\n' +
      '请提供以下信息：\n' +
      '✓ 出生年月日（必填）\n' +
      '✓ 出生时辰（选填）\n\n' +
      '请点击以下链接填写：\n' + formUrl + '\n\n' +
      '⏰ 重要提示：\n' +
      '此链接有效期为24小时，请尽快完成填写。\n\n' +
      '客服联系方式：\n' +
      '📞 +6013-928 4699 / +6013-530 8863\n' +
      '📧 customercare@mandarin.club\n\n' +
      '此邮件由 Mandarin Club 官方系统自动发送，请勿直接回复。\n' +
      '---\n' +
      'Mandarin Club\n' +
      'https://mandarin.club';
    
    const authorizedEmail = Session.getActiveUser().getEmail();
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: plainText,
      htmlBody: htmlBody,
      name: 'Mandarin Club',
      replyTo: 'customercare@mandarin.club',
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
    name: 'Mandarin Club'
  });
  Logger.log('✅ Authorization complete! Authorized email: ' + Session.getActiveUser().getEmail());
}
