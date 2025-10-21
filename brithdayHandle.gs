function doGet(e) {
  const p = e.parameter;
  const name = p.name || '';
  const qty = p.qty || '1';
  const row = p.row || '';
  const orderId = p.order || '';
  const token = p.token || '';
  
  if (!token || !row) {
    return HtmlService.createHtmlOutput(createErrorPage('无效的访问链接'));
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Orders');
  
  if (!sh) {
    return HtmlService.createHtmlOutput(createErrorPage('系统错误'));
  }
  
  const rowId = parseInt(row);
  const storedLink = sh.getRange(rowId, 19).getValue();
  const goldenCardStatus = sh.getRange(rowId, 17).getValue();
  const goldenCardData = sh.getRange(rowId, 18).getValue();
  
  if (!storedLink || !storedLink.includes(token)) {
    return HtmlService.createHtmlOutput(createErrorPage('此链接已失效或无效'));
  }
  
  if (goldenCardStatus === 'Complete') {
    return createResultsPage(name, goldenCardData, rowId, sh);
  }
  
  return createBirthdayForm(name, qty, row, orderId, token);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Orders');
    const rowId = parseInt(data.rowId);
    const submittedToken = data.token || '';
    
    const storedLink = sh.getRange(rowId, 19).getValue();
    const goldenCardStatus = sh.getRange(rowId, 17).getValue();
    
    if (!storedLink || !storedLink.includes(submittedToken)) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: '访问令牌无效，请重新获取链接'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (goldenCardStatus === 'Complete') {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: '您已经提交过生日资料了'
      })).setMimeType(ContentService.MimeType.JSON);
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
    const cacheKey = 'details_' + rowId;
    const detailedInfoJson = JSON.stringify(detailedInfo);
    cache.put(cacheKey, detailedInfoJson, 86400);
    
    sh.getRange(rowId, 17).setValue('Complete');
    sh.getRange(rowId, 18).setValue(formattedCards);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      cards: cards
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function createResultsPage(name, goldenCardData, rowId, sheet) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'details_' + rowId;
    const cachedData = cache.get(cacheKey);
    
    let detailedInfo = [];
    if (cachedData) {
      detailedInfo = JSON.parse(cachedData);
    }
    
    let cardsHtml = '';
    
    for (let i = 0; i < detailedInfo.length; i++) {
      const info = detailedInfo[i];
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
    
    const html = '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">' +
      '<meta name="viewport" content="width=device-width,initial-scale=1.0">' +
      '<title>满金包 - 您的命运分析结果</title>' +
      '<style>*{margin:0;padding:0;box-sizing:border-box}' +
      'body{font-family:"Microsoft YaHei",Arial,sans-serif;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);min-height:100vh;padding:20px}' +
      '.container{max-width:700px;margin:0 auto;background:white;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden}' +
      '.header{background:linear-gradient(135deg,#f6d365 0%,#fda085 100%);color:white;padding:40px;text-align:center}' +
      '.header h1{font-size:32px;margin-bottom:5px}' +
      '.header p{font-size:14px;opacity:0.9}' +
      '.content{padding:40px}' +
      '.result-summary{background:#e3f2fd;border-left:4px solid #1976d2;padding:15px;border-radius:8px;margin-bottom:30px}' +
      '.result-summary p{margin:5px 0;color:#1976d2}' +
      '.card-item{background:#f8f9fa;padding:25px;border-radius:10px;margin-bottom:20px;border-left:4px solid #667eea}' +
      '.card-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:15px}' +
      '.card-number{font-weight:bold;color:#667eea}' +
      '.recipient-badge{background:#e3f2fd;color:#1976d2;padding:4px 12px;border-radius:12px;font-size:12px}' +
      '.birthday-info{background:white;padding:12px;border-radius:6px;margin-bottom:15px;font-size:13px;color:#666}' +
      '.birthday-info p{margin:4px 0}' +
      '.golden-card{background:linear-gradient(135deg,#f6d365 0%,#fda085 100%);padding:25px;border-radius:10px;text-align:center}' +
      '.golden-card h2{font-size:32px;color:white;text-shadow:2px 2px 4px rgba(0,0,0,0.2)}' +
      '.footer{background:#f5f5f5;padding:30px;text-align:center;border-top:1px solid #eee}' +
      '.footer p{color:#666;font-size:14px;margin:10px 0}' +
      '</style></head><body>' +
      '<div class="container">' +
      '<div class="header">' +
      '<h1>🌟 满金包 2026</h1>' +
      '<p>奇门遁甲 · 命运分析结果</p>' +
      '</div>' +
      '<div class="content">' +
      '<div class="result-summary">' +
      '<p><strong>👤 客户：</strong>' + name + '</p>' +
      '<p><strong>📊 命宫结果：</strong>' + goldenCardData + '</p>' +
      '</div>' +
      cardsHtml +
      '<div class="footer">' +
      '<p><strong>感谢您的填写！</strong></p>' +
      '<p>系统已根据您的命宫为您计算专属的奇门遁甲招财阵。</p>' +
      '<p style="margin-top:15px;font-size:12px;color:#999">您可以随时通过此链接查看您的结果。</p>' +
      '</div>' +
      '</div></div></body></html>';
    
    return HtmlService.createHtmlOutput(html)
      .setTitle('满金包2026 - 命运分析结果')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  } catch (error) {
    return HtmlService.createHtmlOutput(createErrorPage('无法加载结果，请联系客服'));
  }
}

function formatDateFromString(dateStr) {
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    return parts[0] + '年' + parts[1] + '月' + parts[2] + '日';
  }
  return dateStr;
}

function createErrorPage(message) {
  const html = '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0">' +
    '<title>访问错误</title><style>' +
    'body{font-family:"Microsoft YaHei",Arial,sans-serif;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);min-height:100vh;padding:20px;display:flex;align-items:center;justify-content:center}' +
    '.container{max-width:500px;background:white;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,0.3);padding:40px;text-align:center}' +
    '.icon{font-size:64px;margin-bottom:20px}' +
    'h1{color:#e74c3c;font-size:24px;margin:20px 0}' +
    'p{color:#666;line-height:1.6;font-size:16px}' +
    '.contact{background:#f8f9fa;padding:15px;border-radius:8px;margin-top:20px;font-size:14px}' +
    '</style></head><body>' +
    '<div class="container">' +
    '<div class="icon">🔒</div>' +
    '<h1>' + message + '</h1>' +
    '<p>如需重新获取链接，请联系客服。</p>' +
    '<div class="contact">' +
    '<p><strong>客服联系方式</strong></p>' +
    '<p>📧 customercare@mandarin.club</p>' +
    '</div>' +
    '</div></body></html>';
  
  return html;
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

function calculateCard(year, month, day, hour, hasTime) {
  try {
    const normalizedHour = hasTime ? parseInt(hour) : 6;
    const yearGanZhi = getYearGanZhi(year);
    const monthGanZhi = getMonthGanZhi(month, yearGanZhi.ganIndex);
    const mingGongZhi = calculateMingGongZhi(monthGanZhi.zhiIndex, normalizedHour);
    const palace = zhiToPalace(mingGongZhi);
    return palace;
  } catch (error) {
    return '坎宫';
  }
}

function getYearGanZhi(year) {
  const ganIndex = (year - 4) % 10;
  const zhiIndex = (year - 4) % 12;
  return { ganIndex: ganIndex, zhiIndex: zhiIndex };
}

function getMonthGanZhi(month, yearGanIndex) {
  const monthZhiBase = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 0, 1];
  const zhiIndex = monthZhiBase[month - 1];
  const monthGanBase = [0, 2, 4, 6, 8];
  const ganOffset = monthGanBase[yearGanIndex % 5];
  const ganIndex = (ganOffset + zhiIndex) % 10;
  return { ganIndex: ganIndex, zhiIndex: zhiIndex };
}

function calculateMingGongZhi(monthZhi, hourZhi) {
  const mingGongZhi = (14 - monthZhi - hourZhi + 24) % 12;
  return mingGongZhi;
}

function zhiToPalace(zhiIndex) {
  const palaceMap = [
    '坎宫',
    '艮宫',
    '艮宫',
    '震宫',
    '巽宫',
    '巽宫',
    '离宫',
    '坤宫',
    '坤宫',
    '兑宫',
    '乾宫',
    '乾宫'
  ];
  return palaceMap[zhiIndex];
}

function getCardDescription(card) {
  return '';
}

function createBirthdayForm(name, qty, row, orderId, token) {
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
  
  const scriptUrl = ScriptApp.getService().getUrl();
  
  const html = '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>满金包 - 生辰八字登记</title><style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:"Microsoft YaHei",Arial,sans-serif;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);min-height:100vh;padding:20px}.container{max-width:600px;margin:0 auto;background:white;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden}.header{background:linear-gradient(135deg,#f6d365 0%,#fda085 100%);color:white;padding:30px;text-align:center}.header h1{font-size:28px;margin:0}header p{margin:5px 0 0 0;font-size:14px}.customer-info{background:#e3f2fd;border-left:4px solid #1976d2;padding:12px 15px;margin:15px;border-radius:6px}.customer-info p{margin:4px 0;font-size:13px;color:#1976d2}.security-warning{background:#fff3cd;border-left:4px solid #ffc107;padding:12px 15px;margin:15px;border-radius:6px;font-size:13px}.security-warning p{margin:4px 0;color:#856404}.form-section{padding:25px}.wallet-group{background:#f8f9fa;padding:20px;border-radius:8px;margin-bottom:20px}.wallet-header{border-bottom:2px solid #667eea;padding-bottom:10px;margin-bottom:15px}.wallet-header h3{color:#667eea;font-size:16px}.form-group{margin-bottom:18px}label{display:block;font-weight:600;margin-bottom:6px;color:#333;font-size:14px}input,select{width:100%;padding:10px;border:1px solid #ddd;border-radius:6px;font-size:14px;background:white}input:focus,select:focus{outline:none;border-color:#667eea;box-shadow:0 0 4px rgba(102,126,234,0.3)}.submit-btn{width:100%;padding:14px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;border:none;border-radius:8px;font-size:16px;font-weight:bold;cursor:pointer;margin-top:10px}.submit-btn:hover{transform:translateY(-2px);box-shadow:0 5px 15px rgba(102,126,234,0.4)}.submit-btn:disabled{background:#ccc;cursor:not-allowed}.loading-overlay{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.8);z-index:9999;justify-content:center;align-items:center}.loading-container{display:flex;flex-direction:column;align-items:center;justify-content:center}.spinner{width:50px;height:50px;border:4px solid rgba(255,255,255,0.3);border-top:4px solid white;border-radius:50%;animation:spin 1s linear infinite}.progress-bar{width:300px;height:8px;background:rgba(255,255,255,0.3);border-radius:10px;overflow:hidden;margin:20px auto}.progress-fill{height:100%;background:linear-gradient(90deg,#667eea,#764ba2);border-radius:10px;animation:progress 1.5s ease-out forwards}.loading-text{color:white;font-size:16px;margin-top:20px}@keyframes spin{0%{transform:rotate(0deg)}100%{transform:rotate(360deg)}}@keyframes progress{0%{width:0%}100%{width:100%}}</style></head><body><div class="loading-overlay" id="loadingOverlay"><div class="loading-container"><div class="spinner"></div><div class="progress-bar"><div class="progress-fill"></div></div><div class="loading-text">✨ 正在计算您的命宫...</div></div></div><div class="container"><div class="header"><h1>满金包</h1><p>奇门遁甲 · 生辰八字登记</p></div><div class="customer-info"><p><strong>👤 姓名:</strong> ' + name + '</p><p><strong>🎁 数量:</strong> ' + qtyNum + ' 个钱包</p></div><div class="security-warning"><p><strong>隐私保护：</strong></p><p>• 你提供的资料（姓名、出生日期、出生时间、出生地点等）将被严格保密，不会对外公开或与第三方共享。</p><p>• 资料仅用于个人八字分析与能量评估，不作其他商业用途。</p><p>• 我们会安全保存资料，并于分析完成后加密或删除。</p><p>• 提交资料即表示你自愿提供并同意以上条款，分析结果仅供参考。</p></div><div class="form-section"><form id="birthdayForm">' + formGroups + '<button type="submit" class="submit-btn" id="submitBtn">马上提交</button></form></div></div><script>const rowId="' + row + '";const qty=' + qtyNum + ';const scriptUrl="' + scriptUrl + '";const token="' + token + '";function timeToHour(t){if(!t)return 6;const h=parseInt(t.split(":")[0]);if(h>=23||h<1)return 0;if(h>=1&&h<3)return 1;if(h>=3&&h<5)return 2;if(h>=5&&h<7)return 3;if(h>=7&&h<9)return 4;if(h>=9&&h<11)return 5;if(h>=11&&h<13)return 6;if(h>=13&&h<15)return 7;if(h>=15&&h<17)return 8;if(h>=17&&h<19)return 9;if(h>=19&&h<21)return 10;if(h>=21&&h<23)return 11;return 6}const hourNames=["子时","丑时","寅时","卯时","辰时","巳时","午时","未时","申时","酉时","戌时","亥时"];document.getElementById("birthdayForm").addEventListener("submit",async function(e){e.preventDefault();const submitBtn=document.getElementById("submitBtn");const loadingOverlay=document.getElementById("loadingOverlay");const wallets=[];for(let i=1;i<=qty;i++){const recipient=document.getElementById("recipient"+i).value;const birthday=document.getElementById("birthday"+i).value;const birthtime=document.getElementById("birthtime"+i).value;if(!recipient){alert("请选择钱包 #"+i+" 是给谁的");return}if(!birthday){alert("请填写钱包 #"+i+" 的出生日期");return}const dateObj=new Date(birthday+"T00:00:00");const year=dateObj.getFullYear();const month=dateObj.getMonth()+1;const day=dateObj.getDate();const hasTime=birthtime?true:false;const hourIndex=timeToHour(birthtime);wallets.push({walletNum:i,recipient:recipient,year:year,month:month,day:day,hour:hourIndex,hourName:hasTime?hourNames[hourIndex]:"未提供",birthday:birthday,birthtime:birthtime||"未提供",hasTime:hasTime})}submitBtn.disabled=true;loadingOverlay.style.display="flex";const data={wallets:wallets,rowId:rowId,qty:qty,token:token};try{const response=await fetch(scriptUrl,{method:"POST",headers:{"Content-Type":"text/plain"},body:JSON.stringify(data)});await new Promise(resolve=>setTimeout(resolve,1500));const result=await response.json();if(result.success){loadingOverlay.style.display="none";window.location.reload()}else{loadingOverlay.style.display="none";submitBtn.disabled=false;alert("提交失败："+result.error)}}catch(error){loadingOverlay.style.display="none";submitBtn.disabled=false;alert("提交失败："+error.message)}});</script></body></html>';
  
  return HtmlService.createHtmlOutput(html)
    .setTitle('满金包2026 - 生辰八字登记')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
