function doGet(e) {
  checkAndCreateSheets(); 
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('阿婆居家照護紀錄')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=no');
}

function checkAndCreateSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Caregivers')) {
    const sh = ss.insertSheet('Caregivers'); sh.appendRow(['代號', '姓名', '狀態']); sh.appendRow(['W01', '阿蒂 (Ati)', 'Active']);
  }
  if (!ss.getSheetByName('DailyLogs')) {
    const sh = ss.insertSheet('DailyLogs'); sh.appendRow(['打卡時間', '照服員', '項目', '表定時間', '紀錄數值']);
  }
  
  const dlSheet = ss.getSheetByName('DailyLogs');
  const dlData = dlSheet.getDataRange().getValues();
  let hasInitialWeight = false;
  for (let i = 1; i < dlData.length; i++) { if (dlData[i][2] === '體重') { hasInitialWeight = true; break; } }
  if (!hasInitialWeight) {
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    dlSheet.appendRow([todayStr, '系統初始', '體重', "'初始", 63.2]);
  }

  if (!ss.getSheetByName('Settings')) {
    const sh = ss.insertSheet('Settings');
    sh.appendRow(['LongInsulinPlan', '07:00=12']);
    sh.appendRow(['ShortInsulinScale', '0-200=0\n201-250=4\n251-300=8\n301-350=8\n351-400=10\n401-=送醫院']);
    sh.appendRow(['SugarTimes', '04:00,07:00,09:30,12:00,15:00,18:00,21:00,23:40']); sh.appendRow(['BPTimes', '04:00,07:00,12:00,15:00,18:00,21:00,23:40']);
    sh.appendRow(['UrineColors', '淡黃(Kuning muda),深黃(Kuning tua),褐色(Coklat),茶色/深色(Gelap),血色(Berdarah),混濁(Keruh)']);
    sh.appendRow(['StoolColors', '褐色(Coklat),墨綠色(Hijau tua),黑色(Hitam),紅色(Merah),灰白色(Pucat),其他(Lainnya)']);
    sh.appendRow(['ExerciseTypes', '站立(Berdiri),走路(Berjalan)']); sh.appendRow(['PRNMedList', '安眠藥(Obat Tidur),止痛藥(Obat Nyeri),退燒藥(Obat Demam)']);
    sh.appendRow(['WeightDays', '0,1,3,5']); sh.appendRow(['WeightTime', "'08:00"]);
  } else {
    const sh = ss.getSheetByName('Settings'); const data = sh.getDataRange().getValues(); const keys = data.map(r => r[0]);
    if(!keys.includes('LongInsulinPlan')) sh.appendRow(['LongInsulinPlan', '07:00=12']);
    if(!keys.includes('ShortInsulinScale')) sh.appendRow(['ShortInsulinScale', '0-200=0\n201-250=4\n251-300=8\n301-350=8\n351-400=10\n401-=送醫院']);
    if(!keys.includes('SugarTimes')) sh.appendRow(['SugarTimes', '04:00,07:00,09:30,12:00,15:00,18:00,21:00,23:40']);
    if(!keys.includes('BPTimes')) sh.appendRow(['BPTimes', '04:00,07:00,12:00,15:00,18:00,21:00,23:40']);
    if(!keys.includes('UrineColors')) sh.appendRow(['UrineColors', '淡黃(Kuning muda),深黃(Kuning tua),褐色(Coklat),茶色/深色(Gelap),血色(Berdarah),混濁(Keruh)']);
    if(!keys.includes('StoolColors')) sh.appendRow(['StoolColors', '褐色(Coklat),墨綠色(Hijau tua),黑色(Hitam),紅色(Merah),灰白色(Pucat),其他(Lainnya)']);
    if(!keys.includes('ExerciseTypes')) sh.appendRow(['ExerciseTypes', '站立(Berdiri),走路(Berjalan)']);
    if(!keys.includes('PRNMedList')) sh.appendRow(['PRNMedList', '安眠藥(Obat Tidur),止痛藥(Obat Nyeri),退燒藥(Obat Demam)']);
    if(!keys.includes('WeightDays')) sh.appendRow(['WeightDays', '0,1,3,5']); if(!keys.includes('WeightTime')) sh.appendRow(['WeightTime', "'08:00"]);
  }
  if (!ss.getSheetByName('MedSettings')) {
    const sh = ss.insertSheet('MedSettings'); sh.appendRow(['Day', '07:30', '12:30', '18:30', '21:00']);
    const days = ['日', '一', '二', '三', '四', '五', '六']; for(let i=0; i<7; i++) sh.appendRow([i + ' (' + days[i] + ')', 0, 0, 0, 0]);
  }
  if (!ss.getSheetByName('ExtraMeds')) {
    const sh = ss.insertSheet('ExtraMeds'); sh.appendRow(['藥名', '星期', '時段', '狀態']);
  }
  if (!ss.getSheetByName('BreathingTraining')) {
    const sh = ss.insertSheet('BreathingTraining'); sh.appendRow(['日期', '次序', '開始時間', '結束時間', '訓練時數(分)', '血壓', '血氧', '心跳', '呼吸速率', '呼吸型態', '備註', '提早接回原因', '執行者']);
  }
  if (!ss.getSheetByName('Excretion')) {
    const sh = ss.insertSheet('Excretion'); sh.appendRow(['日期', '時間', '照服員', '類型(尿/便)', '第幾次', '顏色', '軟硬度(便)']);
  }
  if (!ss.getSheetByName('ExerciseTraining')) {
    const sh = ss.insertSheet('ExerciseTraining'); sh.appendRow(['日期', '次序', '照服員', '運動類型', '開始時間', '結束時間', '訓練時數(分)']);
  }
}

// 基礎設定與管理員
function verifyLogin(acc, pwd) { return (acc === 'admin' && pwd === '1234'); }
function getAllCaregiversAdmin() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Caregivers').getDataRange().getValues();
  let list = []; for (let i=1; i<data.length; i++) if (data[i][0]) list.push({ id: data[i][0], name: data[i][1], status: data[i][2] || 'Active' });
  return list;
}
function getActiveCaregivers() { return getAllCaregiversAdmin().filter(c => c.status !== 'Inactive'); }
function addCaregiver(id, name) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Caregivers').appendRow([id, name, 'Active']); return { success: true }; }
function toggleCaregiverStatus(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Caregivers'); const data = sheet.getDataRange().getValues();
  for (let i=1; i<data.length; i++) if (data[i][0] === id) { sheet.getRange(i+1, 3).setValue(data[i][2] === 'Active' ? 'Inactive' : 'Active'); return { success: true }; }
  return { success: false };
}
function getSystemSettings() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings').getDataRange().getDisplayValues(); let settings = {};
  for(let i=0; i<data.length; i++) settings[data[i][0]] = data[i][1]; return settings;
}
function saveSystemSettingsAdmin(settingsObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings'); sheet.clear();
  sheet.appendRow(['LongInsulinPlan', settingsObj.LongInsulinPlan]);
  sheet.appendRow(['ShortInsulinScale', settingsObj.ShortInsulinScale]);
  sheet.appendRow(['SugarTimes', "'" + settingsObj.SugarTimes]); sheet.appendRow(['BPTimes', "'" + settingsObj.BPTimes]);
  sheet.appendRow(['UrineColors', settingsObj.UrineColors]); sheet.appendRow(['StoolColors', settingsObj.StoolColors]);
  sheet.appendRow(['ExerciseTypes', settingsObj.ExerciseTypes]); sheet.appendRow(['PRNMedList', settingsObj.PRNMedList]); 
  sheet.appendRow(['WeightDays', "'" + settingsObj.WeightDays]); sheet.appendRow(['WeightTime', "'" + settingsObj.WeightTime]);
  return { success: true, message: '✅ 系統設定已全面更新！' };
}

// 藥物設定
function getMedSettingsAdmin() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MedSettings').getRange(2, 2, 7, 4).getValues(); }
function saveMedSettingsAdmin(dataMatrix) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MedSettings').getRange(2, 2, 7, 4).setValues(dataMatrix); return {success: true}; }
function getExtraMedsAdmin() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ExtraMeds').getDataRange().getDisplayValues();
  let list = []; for(let i=1; i<data.length; i++) if(data[i][3] === 'Active') list.push({rowIdx: i+1, name: data[i][0], day: data[i][1], time: data[i][2]}); return list;
}
function addExtraMedAdmin(name, day, time) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ExtraMeds').appendRow([name, day, "'" + time, 'Active']); return {success: true}; }
function removeExtraMedAdmin(rowIdx) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ExtraMeds').getRange(rowIdx, 4).setValue('Inactive'); return {success: true}; }

// 任務與打卡載入
// 任務與打卡載入
function getLastWeight() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DailyLogs').getDataRange().getValues();
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");

  for(let i = data.length - 1; i > 0; i--){ 
    if(data[i][2] === '體重') { 
      // 嘗試解析該筆紀錄的日期
      let rowDateStr = '';
      try {
        rowDateStr = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy/MM/dd");
      } catch(e) {
        rowDateStr = String(data[i][0]).split(' ')[0].replace(/-/g, '/');
      }

      //  關鍵修正：必須「不是今天」的體重，才夠資格被稱為「上次體重」
      if (rowDateStr !== todayStr) {
        return { 
          val: data[i][4], 
          date: Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "MM/dd") 
        };
      }
    } 
  }
  return { val: '63.2', date: '初始' }; // 如果都沒有更早的紀錄，給予初始值
}

function getDashboardData() {
  const nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "u"); let dayOfWeek = parseInt(nowStr); if (dayOfWeek === 7) dayOfWeek = 0; 
  const medSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MedSettings');
  const todayPillsRow = medSheet.getRange(dayOfWeek + 2, 2, 1, 4).getValues()[0]; 
  const times = ['07:30', '12:30', '18:30', '21:00'];
  let todayPills = {}; for(let i=0; i<4; i++) todayPills[times[i]] = todayPillsRow[i];

  const extraMedsList = getExtraMedsAdmin(); let todayExtra = {};
  extraMedsList.forEach(m => { if(m.day === 'All' || String(m.day) === String(dayOfWeek)) { if(!todayExtra[m.time]) todayExtra[m.time] = []; todayExtra[m.time].push(m.name); } });

  return { 
    logs: getTodayLogs(), settings: getSystemSettings(), todayPills: todayPills, todayExtra: todayExtra,
    excretion: getTodayExcretion(), breathing: getBreathingStatus(), exercise: getExerciseStatus(), lastWeight: getLastWeight() 
  };
}

function getTodayLogs() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DailyLogs').getDataRange().getDisplayValues();
  let logs = []; const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  for (let i = data.length - 1; i > 0; i--) {
    const ts = data[i][0] || '';
    if (ts.includes(todayStr) || ts.includes(todayStr.replace(/\//g, '-'))) logs.push({ time: ts.split(' ')[1] || ts, caregiver: data[i][1], category: data[i][2], scheduledTime: data[i][3], value: data[i][4] });
    else { if (ts.split(' ')[0].replace(/-/g, '/') < todayStr) break; }
  }
  return logs;
}
function addLog(caregiver, category, scheduledTime, value) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DailyLogs').appendRow([timestamp, caregiver, category, "'" + scheduledTime, value]); return { success: true };
}

// 排泄、呼吸與運動
function getTodayExcretion() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Excretion').getDataRange().getDisplayValues(); const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  let logs = []; let countUrine = 0; let countStool = 0;
  for(let i=1; i<data.length; i++) {
    if(data[i][0] === todayStr || data[i][0].replace(/\//g, '-') === todayStr.replace(/\//g, '-')) {
      logs.push({time: data[i][1], nurse: data[i][2], type: data[i][3], seq: data[i][4], color: data[i][5], texture: data[i][6]});
      if(data[i][3] === '尿尿 (BAK)') countUrine++; if(data[i][3] === '大便 (BAB)') countStool++;
    }
  }
  return { logs: logs.reverse(), nextUrineSeq: countUrine + 1, nextStoolSeq: countStool + 1 };
}
function addExcretionLog(nurse, type, seq, color, texture) {
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd"); const timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Excretion').appendRow([todayStr, "'" + timeStr, nurse, type, seq, color, texture]); return { success: true };
}

function getBreathingStatus() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BreathingTraining').getDataRange().getDisplayValues(); const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  let activeSession = null; let todayLogs = [];
  for(let i=1; i<data.length; i++) {
    // 🌟 強制日期比對 🌟
    let rDate = ''; try{ rDate = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy/MM/dd"); }catch(e){ rDate = data[i][0]; }
    if(rDate === todayStr || data[i][0].includes(todayStr) || data[i][0].replace(/\//g, '-').includes(todayStr.replace(/\//g, '-'))) {
      if(data[i][3] === '') activeSession = { rowIdx: i + 1, seq: data[i][1], startTime: data[i][2], caregiver: data[i][12] };
      else todayLogs.push({ seq: data[i][1], startTime: data[i][2], endTime: data[i][3], duration: data[i][4] });
    }
  }
  return { active: activeSession, countToday: todayLogs.length + (activeSession ? 1 : 0), logs: todayLogs };
}
function startBreathingTraining(caregiver) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BreathingTraining'); const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd"); const startTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm");
  const seq = getBreathingStatus().countToday + 1; sheet.appendRow([todayStr, seq, "'" + startTime, '', '', '', '', '', '', '', '', '', caregiver]); return { success: true };
}
function endBreathingTraining(rowIdx, startTime, bp, spo2, hr, rr, pattern, reason, note) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BreathingTraining'); const endTimeObj = new Date(); const endTime = Utilities.formatDate(endTimeObj, Session.getScriptTimeZone(), "HH:mm"); const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  let durationMins = Math.round((endTimeObj - new Date(todayStr + ' ' + startTime)) / 60000); if(durationMins < 0) durationMins = 0;
  sheet.getRange(rowIdx, 4, 1, 9).setValues([["'" + endTime, durationMins, bp, spo2, hr, rr, pattern, note, reason]]); return { success: true };
}

function getExerciseStatus() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ExerciseTraining').getDataRange().getDisplayValues(); const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  let activeSession = null; let todayLogs = [];
  for(let i=1; i<data.length; i++) {
    // 🌟 強制日期比對 🌟
    let rDate = ''; try{ rDate = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy/MM/dd"); }catch(e){ rDate = data[i][0]; }
    if(rDate === todayStr || data[i][0].includes(todayStr) || data[i][0].replace(/\//g, '-').includes(todayStr.replace(/\//g, '-'))) {
      if(data[i][5] === '') activeSession = { rowIdx: i + 1, seq: data[i][1], type: data[i][3], startTime: data[i][4], caregiver: data[i][2] };
      else todayLogs.push({ seq: data[i][1], type: data[i][3], startTime: data[i][4], endTime: data[i][5], duration: data[i][6] });
    }
  }
  return { active: activeSession, countToday: todayLogs.length + (activeSession ? 1 : 0), logs: todayLogs };
}
function startExercise(caregiver, type) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ExerciseTraining'); const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd"); const startTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm");
  const seq = getExerciseStatus().countToday + 1; sheet.appendRow([todayStr, seq, caregiver, type, "'" + startTime, '', '']); return { success: true };
}
function endExercise(rowIdx, startTime) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ExerciseTraining'); const endTimeObj = new Date(); const endTime = Utilities.formatDate(endTimeObj, Session.getScriptTimeZone(), "HH:mm"); const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  let durationMins = Math.round((endTimeObj - new Date(todayStr + ' ' + startTime)) / 60000); if(durationMins < 0) durationMins = 0;
  sheet.getRange(rowIdx, 6, 1, 2).setValues([["'" + endTime, durationMins]]); return { success: true };
}

// ==========================================
// 產出進階圖表報告資料 (修復日期比對)
// ==========================================
function getReportData(dateStr) {
  let targetDateStr = dateStr ? dateStr.replace(/-/g, '/') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 🌟 通用日期標準化工具，保證不管格式怎麼變都能對上 🌟
  function getStandardDate(rawDate) {
    if(!rawDate) return "";
    try {
      let dObj = new Date(rawDate);
      if(!isNaN(dObj)) return Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy/MM/dd");
    } catch(e) {}
    return String(rawDate).split(' ')[0].replace(/-/g, '/'); // 備用方案
  }
  
  const targetDateStandard = getStandardDate(targetDateStr);
  const targetDateObj = new Date(targetDateStandard);
  const sevenDaysAgoObj = new Date(targetDateObj.getTime() - 7 * 24 * 60 * 60 * 1000);
  const sevenDaysAgoStr = Utilities.formatDate(sevenDaysAgoObj, Session.getScriptTimeZone(), "yyyy/MM/dd");

  let weightTrend = []; let sugarTrend = []; let bpTrend = [];
  
  // 1. 讀取 DailyLogs
// 1. 讀取 DailyLogs
  const dlData = ss.getSheetByName('DailyLogs').getDataRange().getDisplayValues();
  for (let i = 1; i < dlData.length; i++) {
    const rowDateStr = getStandardDate(dlData[i][0]);
    const rowTimeStr = dlData[i][0].split(' ')[1] || '';
    const category = dlData[i][2]; 
    const val = dlData[i][4];
    
    // 🌟 統一篩選條件：嚴格限制在「目標日期」與「7天前」的區間內
    if (rowDateStr <= targetDateStandard && rowDateStr >= sevenDaysAgoStr) {
      const shortDate = rowDateStr.substring(5) + ' ' + rowTimeStr.substring(0,5);
      
      if (category === '體重') weightTrend.push({ x: shortDate, y: parseFloat(val) });
      if (category === '血糖') sugarTrend.push({ x: shortDate, y: parseFloat(val) });
      if (category === '血壓') {
        const parts = val.split('/');
        if (parts.length === 2) bpTrend.push({ x: shortDate, high: parseFloat(parts[0]), low: parseFloat(parts[1]) });
      }
    }
  }

  // 既然已經嚴格限制在 7 天內了，這行截取最後 10 或 7 筆的程式碼就可以直接刪掉，或是保留預防萬一
  // weightTrend = weightTrend.slice(-10); <-- 建議可以直接刪除這行
  
  // 🌟 修正 2：確保體重只取最後 7 筆資料
  weightTrend = weightTrend.slice(-7);
  // 2. 讀取 呼吸訓練
  let breathing = [];
  const brData = ss.getSheetByName('BreathingTraining').getDataRange().getDisplayValues();
  for(let i=1; i<brData.length; i++) {
    if(getStandardDate(brData[i][0]) === targetDateStandard && brData[i][3] !== '') {
      breathing.push({ seq: brData[i][1], start: brData[i][2], end: brData[i][3], mins: brData[i][4], bp: brData[i][5], spo2: brData[i][6], hr: brData[i][7], rr: brData[i][8], pattern: brData[i][9], note: brData[i][10], reason: brData[i][11], nurse: brData[i][12] });
    }
  }

  // 3. 讀取 運動訓練
  let exercise = [];
  const exData = ss.getSheetByName('ExerciseTraining').getDataRange().getDisplayValues();
  for(let i=1; i<exData.length; i++) {
    if(getStandardDate(exData[i][0]) === targetDateStandard && exData[i][5] !== '') {
      exercise.push({ seq: exData[i][1], nurse: exData[i][2], type: exData[i][3], start: exData[i][4], end: exData[i][5], mins: exData[i][6] });
    }
  }

  // 4. 讀取 排泄紀錄
  let excretion = [];
  const excData = ss.getSheetByName('Excretion').getDataRange().getDisplayValues();
  for(let i=1; i<excData.length; i++) {
    if(getStandardDate(excData[i][0]) === targetDateStandard) {
      excretion.push({ time: excData[i][1], nurse: excData[i][2], type: excData[i][3], seq: excData[i][4], color: excData[i][5], texture: excData[i][6] });
    }
  }

  return {
    targetDate: targetDateStandard, weightTrend: weightTrend, sugarTrend: sugarTrend, bpTrend: bpTrend,
    breathing: breathing, exercise: exercise, excretion: excretion
  };
}
