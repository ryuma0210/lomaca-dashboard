/**
 * ロマカレンタカー 予約管理 - Google Apps Script バックエンド
 * シート: ロマカレンタカー 予約管理マスタ - Gmail自動同期用
 * ID: 1StgPfSDtHv_vBanNrsxyhdWZThsu_j12PjSrXkTOW_4
 */

const SPREADSHEET_ID = '1StgPfSDtHv_vBanNrsxyhdWZThsu_j12PjSrXkTOW_4';
const SHEET_RESERVATIONS = '予約一覧';
const SHEET_VEHICLES = '車両マスタ';
const SHEET_LOCATIONS = '場所マスタ';

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const resSheet = ss.getSheetByName(SHEET_RESERVATIONS);
    const vehSheet = ss.getSheetByName(SHEET_VEHICLES);
    const locSheet = ss.getSheetByName(SHEET_LOCATIONS);

    const reservations = sheetToObjects(resSheet);
    const vehicles = vehSheet ? sheetToObjects(vehSheet) : [];
    const locations = locSheet ? sheetToObjects(locSheet) : [];

    // 最新1000件のみ、予約ID降順
    reservations.sort((a,b) => String(b['予約ID']||'').localeCompare(String(a['予約ID']||'')));
    
    return jsonResponse({
      reservations: reservations.slice(0, 1000),
      vehicles: vehicles,
      locations: locations,
      updatedAt: new Date().toISOString()
    });
  } catch (err) {
    return jsonResponse({ error: err.message, stack: err.stack }, 500);
  }
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_RESERVATIONS);
    if (!sheet) throw new Error('予約一覧シートが見つかりません');

    let payload = {};
    try {
      payload = JSON.parse(e.postData.contents);
    } catch {
      // フォーム送信の場合
      payload = e.parameter;
    }

    // 1. フィールド更新 (チェックボックスやメモなど)
    if (payload.id && payload.field) {
      return handleFieldUpdate(sheet, payload);
    }

    // 2. 新規予約追加
    if (payload.action === 'add' && payload.data) {
      return handleAddReservation(sheet, payload.data);
    }

    throw new Error('不明なリクエスト形式: ' + JSON.stringify(payload).slice(0,500));

  } catch (err) {
    return jsonResponse({ success: false, error: err.message }, 500);
  }
}

// フィールド1つを更新（壊れにくいように行検索を厳密に）
function handleFieldUpdate(sheet, {id, field, value}) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('予約ID');
  const fieldIdx = headers.indexOf(field);
  
  if (idIdx === -1) throw new Error('予約ID列が見つかりません');
  if (fieldIdx === -1) throw new Error('列 ' + field + ' が見つかりません');

  // IDで検索（文字列比較）
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idIdx]).trim() === String(id).trim()) {
      // TRUE/FALSE 正規化
      let v = value;
      if (['決済完了','LINE登録','確認済み'].includes(field)) {
        v = (value === true || value === 'TRUE' || value === 'true') ? 'TRUE' : 'FALSE';
      }
      sheet.getRange(r+1, fieldIdx+1).setValue(v);
      SpreadsheetApp.flush();
      return jsonResponse({ success: true, row: r+1, field, value: v });
    }
  }
  throw new Error('予約ID ' + id + ' が見つかりません');
}

// 新規予約追加
function handleAddReservation(sheet, data) {
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  
  // 予約ID自動採番 (yymmdd + 連番)
  if (!data['予約ID']) {
    data['予約ID'] = generateReservationId(sheet);
  }
  
  // 予約日
  if (!data['予約日']) {
    data['予約日'] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  }

  // 金額を数値化
  if (data['金額']) data['金額'] = Number(String(data['金額']).replace(/[^0-9]/g,'')) || 0;

  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);
  SpreadsheetApp.flush();

  return jsonResponse({ success: true, id: data['予約ID'], row: sheet.getLastRow() });
}

function generateReservationId(sheet) {
  const now = new Date();
  const prefix = Utilities.formatDate(now, 'Asia/Tokyo', 'yyMMdd');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return prefix + '0001';
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat().map(String);
  const todayIds = ids.filter(id => id.startsWith(prefix));
  const num = todayIds.length + 1;
  return prefix + String(num).padStart(4,'0');
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1).filter(row => row.some(c => c !== '')).map(row => {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
    return obj;
  });
}

function jsonResponse(obj, status) {
  const output = ContentService.createTextOutput(JSON.stringify(obj, null, 2));
  output.setMimeType(ContentService.MimeType.JSON);
  // CORS
  return output;
}

// Gmail自動取り込み用（既存のトリガーで動いている想定） - 日付正規化関数
function normalizeDateString(str) {
  if (!str) return '';
  // すでに Date型ならフォーマット
  if (str instanceof Date) {
    return Utilities.formatDate(str, 'Asia/Tokyo', "yyyy年MM月dd日（E）HH時mm分");
  }
  // 2026/9/27 2:00 -> 正規化
  const m = String(str).match(/(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})/);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2])-1, Number(m[3]), Number(m[4]), Number(m[5]));
    return Utilities.formatDate(d, 'Asia/Tokyo', "yyyy年MM月dd日（E）HH時mm分");
  }
  return str;
}
