const express = require('express');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 8080;
const BASE_URL = (process.env.PUBLIC_BASE_URL || 'https://newtaipeinoise.zeabur.app').replace(/\/$/, '');
const DASHBOARD_URL = process.env.DASHBOARD_URL || 'https://noise115.zeabur.app';
const FIELD_REPORT_URL = process.env.FIELD_REPORT_URL || 'https://out115.zeabur.app';
const HOTSPOT_URL = process.env.HOTSPOT_URL || 'https://ntpcnoisely.zeabur.app/login';
const GSHEET_XLSX_URL = process.env.GOOGLE_SHEET_XLSX_URL || process.env.GSHEET_XLSX_URL || '';
const TOP100_GOOGLE_SHEET_ID = process.env.TOP100_GOOGLE_SHEET_ID || process.env.HOTSPOT_GOOGLE_SHEET_ID || '';
const RESULTS_GOOGLE_SHEET_ID = process.env.RESULTS_GOOGLE_SHEET_ID || process.env.DASHBOARD_GOOGLE_SHEET_ID || '';
const FIELD_GOOGLE_SHEET_ID = process.env.FIELD_GOOGLE_SHEET_ID || process.env.FIELD_REPORT_GOOGLE_SHEET_ID || '';
function sheetExportUrlFromId(id){ return id ? `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx` : ''; }
const TOP100_GSHEET_XLSX_URL = process.env.TOP100_GSHEET_XLSX_URL || sheetExportUrlFromId(TOP100_GOOGLE_SHEET_ID);
const RESULTS_GSHEET_XLSX_URL = process.env.RESULTS_GSHEET_XLSX_URL || sheetExportUrlFromId(RESULTS_GOOGLE_SHEET_ID);
const FIELD_GSHEET_XLSX_URL = process.env.FIELD_GSHEET_XLSX_URL || sheetExportUrlFromId(FIELD_GOOGLE_SHEET_ID);
const GSHEET_SYNC_INTERVAL_MIN = Number(process.env.GSHEET_SYNC_INTERVAL_MIN || 0);
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || '';
const SESSION_SECRET = process.env.SESSION_SECRET || '';
const LINE_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN || '';
const LINE_SECRET = process.env.LINE_CHANNEL_SECRET || '';

const DATA_DIR = path.join(__dirname, 'data');
const STORE_PATH = path.join(DATA_DIR, 'store.json');
const upload = multer({ dest: path.join(__dirname, 'uploads') });
let latestDebug = { lastEvents: [], lastReply: null, lastError: null, startedAt: new Date().toISOString() };

function parseCookies(req){
  const out = {};
  const raw = req.headers.cookie || '';
  raw.split(';').forEach(part => {
    const idx = part.indexOf('=');
    if (idx > -1) out[part.slice(0,idx).trim()] = decodeURIComponent(part.slice(idx+1).trim());
  });
  return out;
}
function sign(value){
  const secret = SESSION_SECRET || 'runtime-only-session-secret';
  return crypto.createHmac('sha256', secret).update(value).digest('hex');
}
function makeSession(){
  const value = String(Date.now());
  return `${value}.${sign(value)}`;
}
function verifySession(cookieValue){
  if (!cookieValue) return false;
  const [value, sig] = String(cookieValue).split('.');
  if (!value || !sig) return false;
  const expected = sign(value);
  try {
    const ok = crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(expected));
    const ageMs = Date.now() - Number(value);
    return ok && Number.isFinite(ageMs) && ageMs >= 0 && ageMs < 12 * 60 * 60 * 1000;
  } catch { return false; }
}
function isAdmin(req){ return verifySession(parseCookies(req).ntpc_admin); }
function requireAdmin(req,res,next){
  if (!ADMIN_PASSWORD) return res.status(503).json({ok:false, error:'ADMIN_PASSWORD is not configured. Please set it in Zeabur environment variables.'});
  if (!isAdmin(req)) return res.status(401).json({ok:false, error:'Unauthorized'});
  next();
}
function authStatus(req){
  return { authenticated: isAdmin(req), adminConfigured: !!ADMIN_PASSWORD, sessionSecretConfigured: !!SESSION_SECRET };
}


app.use('/assets', express.static(path.join(__dirname, 'public/assets')));
app.use('/templates', express.static(path.join(__dirname, 'templates')));
app.use(express.urlencoded({ extended: true }));
app.use(express.json({ verify: (req, res, buf) => { req.rawBody = buf; } }));

function readStore(){
  try { return JSON.parse(fs.readFileSync(STORE_PATH, 'utf8')); }
  catch { return { annualGoal: 490, summary: { sessions:0, traffic:0, exceed:0, fines:0, notices:0 }, months:{}, districts:{}, plates:{}, equipment:[], hotspots:[], fieldReports:[], complaints:[] }; }
}
function writeStore(data){ fs.mkdirSync(DATA_DIR, {recursive:true}); fs.writeFileSync(STORE_PATH, JSON.stringify(data, null, 2)); }

function cell(row, keys){
  for(const k of keys){
    if(row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') return row[k];
  }
  return '';
}
function numValue(v){
  if(v === null || v === undefined || v === '') return 0;
  const n = Number(String(v).replace(/[,，%％]/g,'').trim());
  return Number.isFinite(n) ? n : 0;
}
function strValue(v){ return String(v ?? '').trim(); }
function monthKey(v){
  const raw = strValue(v);
  const m = raw.match(/(\d{1,2})/);
  return m ? String(Number(m[1])) : '';
}
function rowsOf(wb, names){
  for(const name of names){
    if(wb.Sheets[name]) return XLSX.utils.sheet_to_json(wb.Sheets[name], {defval:'', raw:false});
  }
  return [];
}
function rowsOfAll(wb){
  const out = [];
  for(const name of wb.SheetNames || []){
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], {defval:'', raw:false});
    rows.forEach(r => out.push({...r, _sheetName:name}));
  }
  return out;
}
function rowsOfNamedOrAll(wb, names){
  const rows = rowsOf(wb, names);
  return rows.length ? rows : rowsOfAll(wb);
}
function objAdd(target, key, row){
  if(!key) return;
  if(!target[key]) target[key] = {sessions:0, traffic:0, exceed:0, fines:0, notices:0};
  target[key].sessions += row.sessions;
  target[key].traffic += row.traffic;
  target[key].exceed += row.exceed;
  target[key].fines += row.fines;
  target[key].notices += row.notices;
}
function parseResultsRows(rows){
  const out = [];
  for(const r of rows){
    const district = strValue(cell(r, ['行政區','district','District']));
    const m = monthKey(cell(r, ['月份','month','Month','日期','執行日期']));
    const sessionsRaw = cell(r, ['執行場次','完成場次','場次數','sessions','完成場次數']);
    const sessions = sessionsRaw === '' ? 1 : numValue(sessionsRaw) || 1;
    const row = {
      month:m,
      district,
      sessions,
      traffic:numValue(cell(r, ['車流辨識','車流辨識件數','traffic','辨識車流'])),
      exceed:numValue(cell(r, ['超標件數','超標數','exceed','超標'])),
      fines:numValue(cell(r, ['告發件數','告發','fines','裁罰件數'])),
      notices:numValue(cell(r, ['通知到檢件數','通知到檢','notices','通檢件數']))
    };
    if(row.district || row.month || row.traffic || row.exceed || row.fines || row.notices) out.push(row);
  }
  return out;
}
function parseEquipmentRows(rows){
  return rows.map(r => ({
    id: strValue(cell(r, ['設備ID','設備編號','機台編號','id','deviceId'])),
    bitestDate: strValue(cell(r, ['中央比測日期','比測日期','bitestDate'])),
    soundMeterDate: strValue(cell(r, ['噪音計檢定日期','噪音計日期','soundMeterDate'])),
    windMeterDate: strValue(cell(r, ['風速計檢定日期','風速計日期','windMeterDate']))
  })).filter(x => x.id);
}
function parseHotspotRows(rows){
  const mapped = rows.map((r, idx) => {
    const route = strValue(cell(r, ['路段','點位名稱','建議點位','name','執勤地點']));
    const location = strValue(cell(r, ['代表位置','地址','location','設置地址','執勤地點']));
    const rank = numValue(cell(r, ['百大排名','排名','rank']));
    const district = strValue(cell(r, ['行政區','district']));
    const complaints = numValue(cell(r, ['檢舉量','中央陳情數','陳情數','complaints']));
    const localMailbox = numValue(cell(r, ['市政信箱月次數','市政信箱','地方陳情數','localMailbox']));
    const exceedVehicles = numValue(cell(r, ['超標車輛數','超標件數','超標數','exceedVehicles','超標車輛']));
    const executions = numValue(cell(r, ['歷年執行次數','執行次數','executions','執行場次']));
    const fines = numValue(cell(r, ['告發件數','告發','fines']));
    const notices = numValue(cell(r, ['通知到檢件數','通知到檢','notices']));
    const nightCases = numValue(cell(r, ['夜間案件','夜間件數']));
    const nightRatioRaw = cell(r, ['夜間比例','nightRatio']);
    const nightRatio = nightRatioRaw !== '' ? numValue(nightRatioRaw) : (exceedVehicles ? nightCases / exceedVehicles : 0);
    const year = strValue(cell(r, ['年份','年度','year']));
    const layer = strValue(cell(r, ['資料圖層','資料類型','點位類型']));
    const vendor = strValue(cell(r, ['廠商','vendor']));
    let score = numValue(cell(r, ['綜合分數','綜合得分','得分','score']));
    if(!score){
      // 無既有百大分數時，先依 Sheet 內可讀欄位建立可排序暫估分數；正式仍以局端百大表排名為準。
      score = complaints*0.4 + localMailbox*0.2 + exceedVehicles*0.25 + nightRatio*10 + Math.max(0, 10-executions)*0.5;
    }
    return {
      rank,
      grade: strValue(cell(r, ['分級','等級','grade'])),
      name: route || location,
      route,
      location,
      district,
      score: Math.round(score*100)/100,
      complaints,
      localMailbox,
      exceedVehicles,
      executions,
      fines,
      notices,
      nightCases,
      nightRatio,
      lat: numValue(cell(r, ['緯度','lat','latitude'])),
      lng: numValue(cell(r, ['經度','lng','longitude'])),
      scoringFocus: strValue(cell(r, ['評分重點','重點說明','scoringFocus'])) || [layer, year, vendor].filter(Boolean).join('／'),
      kpi: strValue(cell(r, ['KPI','點位KPI','KPI成效'])),
      precision: strValue(cell(r, ['執法精準率','精準率','precision'])),
      updatedAt: strValue(cell(r, ['最後更新時間','資料更新時間','updatedAt']))
    };
  }).filter(x => x.name || x.district);
  mapped.sort((a,b)=>{
    const ar = a.rank || 999999, br = b.rank || 999999;
    if(ar !== br) return ar - br;
    return (b.score||0) - (a.score||0);
  });
  mapped.forEach((x,i)=>{
    if(!x.rank) x.rank = i + 1;
    if(!x.grade) x.grade = x.rank <= 10 ? 'S' : x.rank <= 30 ? 'A' : x.rank <= 70 ? 'B' : 'C';
  });
  return mapped.slice(0,100);
}
function parseComplaintRows(rows){
  return rows.map(r => ({
    year: strValue(cell(r, ['年度','year'])),
    month: monthKey(cell(r, ['月份','month'])),
    district: strValue(cell(r, ['行政區','district'])),
    central: numValue(cell(r, ['中央陳情數','中央','centralComplaints'])),
    local: numValue(cell(r, ['地方陳情數','市政信箱','地方','localComplaints'])),
    total: numValue(cell(r, ['合計陳情數','陳情合計','totalComplaints'])),
    lastYear: numValue(cell(r, ['去年同期陳情數','去年同期','lastYearComplaints'])),
    updatedAt: strValue(cell(r, ['資料更新時間','updatedAt']))
  })).filter(x => x.district || x.month || x.total || x.central || x.local);
}
function parseFieldRows(rows){
  return rows.map(r => ({
    status: strValue(cell(r, ['狀態','回報狀態','status'])),
    date: strValue(cell(r, ['日期','回報日期','date'])),
    sessionId: strValue(cell(r, ['執行場次','場次ID','場次編號','sessionId'])),
    calibration: strValue(cell(r, ['校正值','現場校正值','calibration'])),
    device: strValue(cell(r, ['機台編號','設備機號','device'])),
    speedLimit: strValue(cell(r, ['路段限速','路段速限(km/h)','speedLimit'])),
    noiseStandard: strValue(cell(r, ['噪音標準','噪音標準dB(A)','noiseStandard'])),
    district: strValue(cell(r, ['行政區','district'])),
    location: strValue(cell(r, ['執勤地點','執行地點','location'])),
    lat: numValue(cell(r, ['執勤緯度','稽查點緯度','lat'])),
    lng: numValue(cell(r, ['執勤經度','稽查點經度','lng'])),
    signLocation: strValue(cell(r, ['告示牌位置','告示牌地址','signLocation'])),
    signLat: numValue(cell(r, ['告示牌緯度','signLat'])),
    signLng: numValue(cell(r, ['告示牌經度','signLng'])),
    distance: numValue(cell(r, ['距離公尺','告示牌距離(m)','distance'])),
    reporter: strValue(cell(r, ['回報人員','reporter'])),
    submittedAt: strValue(cell(r, ['送出時間','資料更新時間','submittedAt','updatedAt'])),
    reportId: strValue(cell(r, ['報表ID','回報ID','reportId'])),
    note: strValue(cell(r, ['備註','note']))
  })).filter(x => x.sessionId || x.device || x.district || x.location);
}
async function fetchWorkbookFromUrl(url){
  const r = await fetch(url, {headers:{'User-Agent':'NTPC Noise Management System'}});
  if(!r.ok) throw new Error(`Google Sheet fetch failed: ${r.status} ${url}`);
  const buf = Buffer.from(await r.arrayBuffer());
  return XLSX.read(buf, {type:'buffer'});
}
function updateResultsStore(store, resultRows){
  if(!resultRows.length) return false;
  const summary = {sessions:0, traffic:0, exceed:0, fines:0, notices:0};
  const months = {}, districts = {};
  for(const r of resultRows){
    summary.sessions += r.sessions; summary.traffic += r.traffic; summary.exceed += r.exceed; summary.fines += r.fines; summary.notices += r.notices;
    objAdd(months, r.month, r); objAdd(districts, r.district, r);
  }
  store.summary = {...(store.summary||{}), ...summary};
  store.months = months;
  store.districts = districts;
  return true;
}
async function syncFromLinkedSheets(){
  const store = readStore();
  const detail = { at: now(), source:'linked-google-sheets', resultsRows:0, hotspotRows:0, fieldRows:0, complaintRows:0, sheets:{} };
  if(RESULTS_GSHEET_XLSX_URL){
    const wb = await fetchWorkbookFromUrl(RESULTS_GSHEET_XLSX_URL);
    const rows = parseResultsRows(rowsOfAll(wb));
    updateResultsStore(store, rows);
    detail.resultsRows = rows.length;
    detail.sheets.results = wb.SheetNames;
  }
  if(TOP100_GSHEET_XLSX_URL){
    const wb = await fetchWorkbookFromUrl(TOP100_GSHEET_XLSX_URL);
    const rows = parseHotspotRows(rowsOfNamedOrAll(wb, ['前100大路段','百大點位_Output','點位資料_Input','百大點位','Hotspot_Output','TOP100']));
    if(rows.length) store.hotspots = rows;
    detail.hotspotRows = rows.length;
    detail.sheets.top100 = wb.SheetNames;
  }
  if(FIELD_GSHEET_XLSX_URL){
    const wb = await fetchWorkbookFromUrl(FIELD_GSHEET_XLSX_URL);
    const rows = parseFieldRows(rowsOfNamedOrAll(wb, ['外勤回報_Raw','外勤回報','回報彙整','field','工作表1']));
    if(rows.length) store.fieldReports = rows;
    detail.fieldRows = rows.length;
    detail.sheets.field = wb.SheetNames;
  }
  store.lastDataSync = detail;
  writeStore(store);
  return detail;
}
function importWorkbookToStore(wb, source='manual-upload'){
  const store = readStore();
  const resultRows = parseResultsRows(rowsOf(wb, ['成果彙整_Input','監測成果_Output','Session_Results_Input','成果彙整','成果']));
  if(resultRows.length){
    const summary = {sessions:0, traffic:0, exceed:0, fines:0, notices:0};
    const months = {}, districts = {};
    for(const r of resultRows){
      summary.sessions += r.sessions; summary.traffic += r.traffic; summary.exceed += r.exceed; summary.fines += r.fines; summary.notices += r.notices;
      objAdd(months, r.month, r); objAdd(districts, r.district, r);
    }
    store.summary = {...(store.summary||{}), ...summary};
    store.months = months;
    store.districts = districts;
  }
  const eqRows = parseEquipmentRows(rowsOf(wb, ['equipment_系統匯入','equipment','設備管理_Input','設備資料_中文輸入','設備管理']));
  if(eqRows.length) store.equipment = eqRows;
  const hotspotRows = parseHotspotRows(rowsOf(wb, ['百大點位_Output','點位資料_Input','百大點位','Hotspot_Output']));
  if(hotspotRows.length) store.hotspots = hotspotRows;
  const complaintRows = parseComplaintRows(rowsOf(wb, ['陳情資料_Input','陳情趨勢_Output','陳情資料','Complaint_Input']));
  if(complaintRows.length) store.complaints = complaintRows;
  store.lastDataSync = { at: now(), source, sheets: wb.SheetNames, resultRows: resultRows.length, equipmentRows: eqRows.length, hotspotRows: hotspotRows.length, complaintRows: complaintRows.length };
  writeStore(store);
  return store.lastDataSync;
}
async function syncFromGoogleSheet(){
  // V27: 若設定三套平台獨立 Google Sheet，優先同步三套來源；否則沿用單一 GOOGLE_SHEET_XLSX_URL。
  if(TOP100_GSHEET_XLSX_URL || RESULTS_GSHEET_XLSX_URL || FIELD_GSHEET_XLSX_URL){
    return await syncFromLinkedSheets();
  }
  if(!GSHEET_XLSX_URL) throw new Error('GOOGLE_SHEET_XLSX_URL / GSHEET_XLSX_URL 未設定。請使用 Google Sheet 匯出 xlsx URL。');
  const wb = await fetchWorkbookFromUrl(GSHEET_XLSX_URL);
  return importWorkbookToStore(wb, 'google-sheet');
}
function safeRate(a,b){ return b ? `${((Number(a||0)/Number(b||0))*100).toFixed(1)}%` : '0.0%'; }
function systemClassificationRows(){ return [
  {分類:'外勤回報', 系統連結:FIELD_REPORT_URL, 主要資料:'場次、機號、校正值、座標、照片、告示牌位置', 匯入表:'外勤回報_Raw', 輸出表:'外勤回報_Export', LINE查詢:'外勤回報／場次查詢', 更新頻率:'即時', 備註:'作為架設與現場佐證，不直接作為成果主表'},
  {分類:'成果查詢', 系統連結:DASHBOARD_URL, 主要資料:'車流、超標、告發、通知到檢、KPI、月份、行政區', 匯入表:'成果彙整_Input', 輸出表:'監測成果_Output', LINE查詢:'進度／月份／行政區／KPI', 更新頻率:'成果完成後立即更新', 備註:'LINE與Dashboard統一讀取成果主表'},
  {分類:'百大熱點', 系統連結:HOTSPOT_URL, 主要資料:'移動式點位、固定式點位、陳情數、執行成效、點位KPI、執法精準率', 匯入表:'點位資料_Input', 輸出表:'百大點位_Output', LINE查詢:'百大點位／行政區百大', 更新頻率:'每月或局端新資料提供後立即重算', 備註:'局端提供全市聲音照相數據後重算百大'},
  {分類:'設備管理', 系統連結:'後台匯入', 主要資料:'中央比測、噪音計檢定、風速計檢定、現場校正', 匯入表:'設備管理_Input', 輸出表:'設備提醒_Output', LINE查詢:'設備管理', 更新頻率:'證書或報告更新後立即匯入', 備註:'比測2年；噪音計1年；風速計1年'},
  {分類:'陳情趨勢', 系統連結:DASHBOARD_URL, 主要資料:'中央陳情、地方陳情、去年同期、同期比對', 匯入表:'陳情資料_Input', 輸出表:'陳情趨勢_Output', LINE查詢:'陳情趨勢／陳情同期比', 更新頻率:'局端提供或每月更新', 備註:'監測成果系統可同步呈現陳情與執法成效趨勢'},
  {分類:'法規新聞', 系統連結:'LINE 法規中心', 主要資料:'噪音管制法、聲音照相指引、NIEA P211.82B、新聞', 匯入表:'法規新聞_Input', 輸出表:'法規新聞_Output', LINE查詢:'法規中心／新聞／NIEA82B', 更新頻率:'每日08:00', 備註:'LINE回覆註明更新時間'}
]; }
function objRows(obj, labelName){ return Object.entries(obj||{}).map(([key,v])=>({[labelName]:key, 場次:v.sessions||0, 車流辨識:v.traffic||0, 超標件數:v.exceed||0, 告發件數:v.fines||0, 通知到檢件數:v.notices||0, 成案件數:(v.fines||0)+(v.notices||0), 告發率:safeRate(v.fines,v.exceed), 通檢率:safeRate(v.notices,v.exceed), KPI:v.sessions ? (((v.fines||0)+(v.notices||0))/(v.sessions||1)).toFixed(2) : '0.00'})); }
function buildExportSheets(){
  const s=readStore(); const a=s.summary||{};
  const updated = s.lastDataSync?.at || s.news?.updatedAt || new Date().toISOString();
  return {
    系統分類總覽: systemClassificationRows(),
    監測成果_Output: [{統計範圍:'全計畫累計', 年度目標:s.annualGoal||490, 完成場次:a.sessions||0, 車流辨識:a.traffic||0, 超標件數:a.exceed||0, 告發件數:a.fines||0, 通知到檢件數:a.notices||0, 成案件數:(a.fines||0)+(a.notices||0), 告發率:safeRate(a.fines,a.exceed), 通檢率:safeRate(a.notices,a.exceed), KPI:a.sessions ? (((a.fines||0)+(a.notices||0))/(a.sessions||1)).toFixed(2) : '0.00', 資料更新時間:updated}],
    月份統計_Output: objRows(s.months,'月份'),
    行政區統計_Output: objRows(s.districts,'行政區'),
    設備提醒_Output: (s.equipment||[]).map(e=>({設備ID:e.id||'', 中央比測日期:e.bitestDate||'', 噪音計檢定日期:e.soundMeterDate||'', 風速計檢定日期:e.windMeterDate||''})),
    百大點位_Output: (s.hotspots||[]).map(h=>({排名:h.rank||'', 行政區:h.district||'', 路段:h.route||h.name||'', 代表位置:h.location||'', 等級:h.grade||'', 綜合分數:h.score||'', 檢舉量:h.complaints||0, 市政信箱月次數:h.localMailbox||0, 超標車輛數:h.exceedVehicles||0, 歷年執行次數:h.executions||0, 夜間比例:h.nightRatio||0, 緯度:h.lat||'', 經度:h.lng||'', 評分重點:h.scoringFocus||'', 最後更新時間:h.updatedAt||updated})),
    陳情趨勢_Output: (s.complaints||[]).map(c=>({年度:c.year||'', 月份:c.month||'', 行政區:c.district||'', 中央陳情:c.central||0, 地方陳情:c.local||0, 合計陳情:(c.total||((c.central||0)+(c.local||0))), 去年同期:c.lastYear||0, 同期增減率:c.lastYear ? `${((((c.total||((c.central||0)+(c.local||0)))-c.lastYear)/c.lastYear)*100).toFixed(1)}%` : '未提供', 資料更新時間:c.updatedAt||updated})),
    法規新聞_Output: (s.news?.items||[]).map((n,i)=>({序號:i+1, 標題:n.title||'', 摘要:n.summary||'', 來源:n.source||'', 連結:n.url||'', 更新時間:s.news?.updatedAt||updated})),
    LINE查詢_Output: [
      {LINE指令:'進度', 資料來源:'監測成果_Output', 回覆:'全計畫場次、車流、超標、告發、通檢、KPI', 資料更新時間:updated},
      {LINE指令:'百大點位', 資料來源:'百大點位_Output', 回覆:'全市Top10與百大熱點平台連結', 資料更新時間:updated},
      {LINE指令:'陳情趨勢', 資料來源:'陳情趨勢_Output', 回覆:'中央/地方陳情與同期比', 資料更新時間:updated},
      {LINE指令:'設備管理', 資料來源:'設備提醒_Output', 回覆:'校正/檢定/比測到期提醒', 資料更新時間:updated},
      {LINE指令:'新聞', 資料來源:'法規新聞_Output', 回覆:'每日噪音車相關新聞摘要', 資料更新時間:s.news?.updatedAt||updated}
    ],
    資料更新紀錄: [{資料更新時間:updated, 來源:s.lastDataSync?.source||'store', 結果:JSON.stringify(s.lastDataSync||{}), 備註:'由V25匯出中心產生'}]
  };
}
function buildExportWorkbook(kind='all'){
  const sheets = buildExportSheets();
  const wb = XLSX.utils.book_new();
  const entries = kind==='all' ? Object.entries(sheets) : Object.entries(sheets).filter(([name])=>name.includes(kind)||name===kind);
  for(const [name,rows] of entries){ XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length?rows:[{提示:'目前無資料'}]), name.substring(0,31)); }
  return wb;
}
function sendWorkbook(res, wb, filename){
  const buf = XLSX.write(wb,{bookType:'xlsx',type:'buffer'});
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition',`attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
  res.end(buf);
}

function complaintTrendMsg(keyword=''){
  const s=readStore(); const rows=s.complaints||[];
  if(!rows.length) return textMessage('目前尚未匯入陳情資料。請先於後台匯入 Google Sheet 連動範本。', [{label:'開啟後台',uri:`${BASE_URL}/admin.html`}]);
  const district = (keyword.match(/[\u4e00-\u9fa5]{2,3}區/)||[])[0];
  const filtered = district ? rows.filter(r=>r.district===district) : rows;
  const latest = filtered.slice(-6);
  const lines = latest.map(r=>{
    const total = r.total || r.central + r.local;
    const change = r.lastYear ? `${(((total-r.lastYear)/r.lastYear)*100).toFixed(1)}%` : '未提供';
    return `${r.year||''}/${r.month||''} ${r.district||''}｜陳情${fmt(total)}件｜同期比：${change}`;
  }).join('\n');
  return textMessage(`【陳情趨勢】\n${district?district:'全市'}\n${lines}\n\n資料更新時間：${s.lastDataSync?.at || '未記錄'}`, [{label:'成果查詢',uri:DASHBOARD_URL}, {label:'百大點位',text:'百大點位'}]);
}
function hotspotRankDetailMsg(rank){
  const s=readStore(); const rows=s.hotspots||[];
  const r = rows.find(x => Number(x.rank) === Number(rank));
  if(!r) return textMessage(`目前查無百大第 ${rank} 名資料。`, [{label:'百大點位',text:'百大點位'}]);
  return textMessage(`【百大建議點位｜第${r.rank}名】
行政區：${r.district || '-'}
路段：${r.route || r.name || '-'}
代表位置：${r.location || '-'}
等級：${r.grade || '-'}
綜合分數：${r.score || '-'}

檢舉量：${fmt(r.complaints)}件
市政信箱月次數：${fmt(r.localMailbox)}次
超標車輛數：${fmt(r.exceedVehicles)}輛
歷年執行次數：${fmt(r.executions)}次
夜間比例：${r.nightRatio ? (Number(r.nightRatio)*100).toFixed(1)+'%' : '-'}
座標：${r.lat || '-'}, ${r.lng || '-'}

評分重點：${r.scoringFocus || '-'}

資料更新時間：${s.lastDataSync?.at || r.updatedAt || '未記錄'}`, [
    {label:'開啟百大熱點平台',uri:HOTSPOT_URL},
    {label:'百大Top10',text:'百大點位'},
    {label:`${r.district || ''}百大`,text:`${r.district || ''}百大`}
  ]);
}
function hotspotMethodMsg(){
  return textMessage(`【百大建議點位計算方式】
依本版「前100大加權排行－詳細版」呈現。

綜合分數 =
檢舉量 40%
＋市政信箱 20%
＋超標車輛 25%
＋夜間比例 10%
＋布點缺口 5%

正規化：
檢舉量、市政信箱月次數、超標車輛數、歷年執行次數先 log(1+x)，再以 min-max 轉為 0～1。

分級：
S：1～10名
A：11～30名
B：31～70名
C：71～100名`, [
    {label:'百大Top10',text:'百大點位'},
    {label:'開啟平台',uri:HOTSPOT_URL}
  ]);
}
function hotspotListMsg(keyword=''){
  const s=readStore(); const rows=s.hotspots||[];
  if(!rows.length) return hotspotMsg();
  const rankMatch = String(keyword||'').match(/(?:百大|第|排名)?\s*(\d{1,3})\s*(?:名|詳細|點位)?/);
  if(rankMatch && (keyword.includes('第') || keyword.includes('排名') || keyword.includes('詳細'))) return hotspotRankDetailMsg(Number(rankMatch[1]));
  const district = (keyword.match(/[\u4e00-\u9fa5]{2,3}區/)||[])[0];
  const filtered = district ? rows.filter(r=>r.district===district) : rows;
  const top = filtered.slice(0,10);
  const lines = top.map(r=>`#${r.rank||'-'}｜${r.district||''}｜${r.route||r.name||'未命名路段'}
${r.location ? '位置：'+r.location+'｜' : ''}${r.grade||'-'}級｜分數${r.score||'-'}｜檢舉${fmt(r.complaints)}｜超標${fmt(r.exceedVehicles)}
重點：${r.scoringFocus || '-'}`).join('\n\n');
  return textMessage(`【百大建議點位】
${district?district+' Top10':'全市 Top10'}
${lines || '目前沒有符合條件的點位'}

可輸入：
・百大計算方式
・第1名詳細
・淡水區百大

資料更新時間：${s.lastDataSync?.at || '未記錄'}`, [
    {label:'開啟百大熱點平台',uri:HOTSPOT_URL},
    {label:'百大計算方式',text:'百大計算方式'},
    {label:'陳情趨勢',text:'陳情趨勢'}
  ]);
}

function fmt(n){ return Number(n||0).toLocaleString('zh-TW'); }
function pct(a,b){ return b ? ((a/b)*100).toFixed(1)+'%' : '0.0%'; }
function kpi(obj){ return obj.sessions ? ((obj.fines + obj.notices) / obj.sessions).toFixed(2) : '0.00'; }
function rate(a,b){ return b ? ((a/b)*100).toFixed(1)+'%' : '0.0%'; }
function now(){ return new Date().toISOString(); }
function recordEvent(e){ latestDebug.lastEvents.unshift(e); latestDebug.lastEvents = latestDebug.lastEvents.slice(0,20); }

function verifySignature(req){
  if (!LINE_SECRET) return true;
  const signature = req.get('x-line-signature') || '';
  const hmac = crypto.createHmac('sha256', LINE_SECRET).update(req.rawBody || Buffer.from('')).digest('base64');
  try { return crypto.timingSafeEqual(Buffer.from(signature), Buffer.from(hmac)); } catch { return false; }
}
async function lineReply(replyToken, messages){
  if (!LINE_TOKEN || !replyToken) return { skipped:true };
  const payload = { replyToken, messages: Array.isArray(messages) ? messages : [messages] };
  const r = await fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${LINE_TOKEN}` },
    body: JSON.stringify(payload)
  });
  const text = await r.text();
  latestDebug.lastReply = { at: now(), status: r.status, body: text, sentMessages: payload.messages };
  if (!r.ok) throw new Error(`LINE Reply API ${r.status}: ${text}`);
  return { ok:true };
}
function textMessage(text, quickItems=[]){
  const m = { type:'text', text };
  if (quickItems.length) {
    m.quickReply = { items: quickItems.map(i => ({ type:'action', action: i.uri ? { type:'uri', label:i.label, uri:i.uri } : { type:'message', label:i.label, text:i.text || i.label } })) };
  }
  return m;
}
function menuQuick(){ return [
  {label:'成果查詢', uri:DASHBOARD_URL}, {label:'外勤回報', uri:FIELD_REPORT_URL}, {label:'月份統計', text:'月份選單'},
  {label:'行政區統計', text:'行政區選單'}, {label:'KPI報表', text:'KPI報表'}, {label:'法規中心', text:'法規中心'}, {label:'百大熱點', uri:HOTSPOT_URL}, {label:'設備管理', text:'設備管理'}
];}
function mainMenu(){
  return textMessage('新北市打擊噪音車管理系統\n\n請選擇要使用的功能。', menuQuick());
}
function progressMsg(){
  const s=readStore(); const a=s.summary||{}; const goal=s.annualGoal||490;
  return textMessage(`【全計畫執行成效】\n年度目標：${fmt(goal)}場\n已完成：${fmt(a.sessions)}場（${pct(a.sessions, goal)}）\n待執行：${fmt(Math.max(goal-a.sessions,0))}場\n\n車流辨識：${fmt(a.traffic)}件\n超標件數：${fmt(a.exceed)}件\n告發件數：${fmt(a.fines)}件\n通知到檢：${fmt(a.notices)}件\n告發率：${rate(a.fines,a.exceed)}\n通檢率：${rate(a.notices,a.exceed)}\nKPI成效：${kpi(a)}`, menuQuick());
}
function kpiMsg(){
  const s=readStore(); const a=s.summary||{};
  return textMessage(`【KPI報表】\n告發率：${rate(a.fines,a.exceed)}\n通檢率：${rate(a.notices,a.exceed)}\n達成率：${pct(a.sessions, s.annualGoal||490)}\nKPI成效：${kpi(a)}\n\n說明：KPI =（告發件數 + 通知到檢）/ 執行場次。`, [{label:'月份選單',text:'月份選單'}, {label:'行政區選單',text:'行政區選單'}, {label:'成果平台',uri:DASHBOARD_URL}]);
}
function monthMenu(){ return textMessage('請選擇月份查詢：', ['2','3','4','5','6'].map(m=>({label:`${m}月`, text:`${m}月份執行成效`}))); }
function districtMenu(){ return textMessage('請選擇行政區查詢：', ['土城區','淡水區','板橋區','新莊區','三重區','中和區'].map(d=>({label:d, text:`${d}執行成效`}))); }
function statMenu(){ return textMessage('統計查詢可選月份或行政區。', [{label:'月份查詢',text:'月份選單'}, {label:'行政區查詢',text:'行政區選單'}, {label:'成果平台', uri:DASHBOARD_URL}]); }
function monthStats(m){
  const s=readStore(); const d=s.months?.[m];
  if(!d) return textMessage(`目前沒有 ${m} 月資料。`, [{label:'月份選單',text:'月份選單'}]);
  return textMessage(`【${m}月份執行成效】\n執行場次：${fmt(d.sessions)}場\n車流辨識：${fmt(d.traffic)}件\n超標件數：${fmt(d.exceed)}件\n告發件數：${fmt(d.fines)}件\n通知到檢：${fmt(d.notices)}件\n告發率：${rate(d.fines,d.exceed)}\n通檢率：${rate(d.notices,d.exceed)}\nKPI成效：${kpi(d)}`, [{label:'月份選單',text:'月份選單'}, {label:'成果平台', uri:DASHBOARD_URL}]);
}
function districtStats(name){
  const s=readStore(); const d=s.districts?.[name];
  if(!d) return textMessage(`目前沒有「${name}」資料。`, [{label:'行政區選單',text:'行政區選單'}]);
  return textMessage(`【${name}執行成效】\n執行場次：${fmt(d.sessions)}場\n車流辨識：${fmt(d.traffic)}件\n超標件數：${fmt(d.exceed)}件\n告發件數：${fmt(d.fines)}件\n通知到檢：${fmt(d.notices)}件\n告發率：${rate(d.fines,d.exceed)}\n通檢率：${rate(d.notices,d.exceed)}\nKPI成效：${kpi(d)}`, [{label:'行政區選單',text:'行政區選單'}, {label:'成果平台', uri:DASHBOARD_URL}]);
}

function taipeiTimeText(dateInput){
  const d = dateInput ? new Date(dateInput) : new Date();
  try {
    return new Intl.DateTimeFormat('zh-TW', {
      timeZone:'Asia/Taipei', year:'numeric', month:'2-digit', day:'2-digit',
      hour:'2-digit', minute:'2-digit', hour12:false
    }).format(d).replace(/\//g,'/');
  } catch {
    return d.toISOString();
  }
}
function decodeXml(s){
  return String(s||'')
    .replace(/<!\[CDATA\[(.*?)\]\]>/gs,'$1')
    .replace(/&amp;/g,'&')
    .replace(/&lt;/g,'<')
    .replace(/&gt;/g,'>')
    .replace(/&quot;/g,'"')
    .replace(/&#39;/g,"'")
    .replace(/\s+/g,' ')
    .trim();
}
function stripHtml(s){ return decodeXml(String(s||'').replace(/<[^>]+>/g,' ')); }
function extractTag(item, tag){
  const m = String(item||'').match(new RegExp(`<${tag}[^>]*>([\\s\\S]*?)<\\/${tag}>`, 'i'));
  return m ? decodeXml(m[1]) : '';
}
function defaultNewsItems(){
  return [
    {title:'環境部噪音車管制與民眾檢舉專區', source:'環境部', url:'https://noisecar.moenv.gov.tw/', summary:'追蹤噪音車檢舉、通知到檢與管制政策。'},
    {title:'國家環境研究院與聲音照相檢測資訊', source:'國環院／環境部', url:'https://www.moenv.gov.tw/', summary:'追蹤檢測方法、設備比測、NIEA 方法與公告資訊。'},
    {title:'地方政府噪音車科技執法案例', source:'新聞彙整', url:'https://news.google.com/search?q=%E5%99%AA%E9%9F%B3%E8%BB%8A%20%E8%81%B2%E9%9F%B3%E7%85%A7%E7%9B%B8&hl=zh-TW&gl=TW&ceid=TW:zh-Hant', summary:'追蹤各縣市聲音照相、噪音車裁罰與修法相關新聞。'}
  ];
}
async function fetchGoogleNews(query, limit=4){
  const url = `https://news.google.com/rss/search?q=${encodeURIComponent(query)}&hl=zh-TW&gl=TW&ceid=TW:zh-Hant`;
  const r = await fetch(url, {headers:{'User-Agent':'Mozilla/5.0 NTPC Noise Bot'}});
  if(!r.ok) throw new Error(`news ${r.status}`);
  const xml = await r.text();
  const blocks = [...xml.matchAll(/<item>([\s\S]*?)<\/item>/gi)].map(m=>m[1]);
  return blocks.slice(0, limit).map(item => {
    const rawTitle = extractTag(item,'title');
    const link = extractTag(item,'link');
    const pubDate = extractTag(item,'pubDate');
    const desc = stripHtml(extractTag(item,'description'));
    let source = 'Google News';
    const sm = item.match(/<source[^>]*>([\s\S]*?)<\/source>/i);
    if(sm) source = decodeXml(sm[1]);
    return { title: rawTitle, source, url: link, publishedAt: pubDate ? new Date(pubDate).toISOString() : '', summary: desc.slice(0,120) };
  });
}
async function refreshNoiseNews(reason='scheduled'){
  const store = readStore();
  const queries = [
    '新北 噪音車 聲音照相',
    '環境部 噪音車 聲音照相',
    '國家環境研究院 NIEA P211.82B 噪音車',
    '噪音車 科技執法 修法'
  ];
  let items = [];
  try {
    for(const q of queries){
      const got = await fetchGoogleNews(q, 3);
      items.push(...got.map(x=>({...x, query:q})));
    }
    const seen = new Set();
    items = items.filter(x => {
      const key = `${x.title}|${x.source}`;
      if(!x.title || seen.has(key)) return false;
      seen.add(key); return true;
    }).slice(0, 8);
    if(!items.length) items = defaultNewsItems();
    store.news = { updatedAt: new Date().toISOString(), updatedAtTaipei: taipeiTimeText(), reason, items, error:null };
  } catch(e) {
    const old = store.news?.items?.length ? store.news.items : defaultNewsItems();
    store.news = { updatedAt: new Date().toISOString(), updatedAtTaipei: taipeiTimeText(), reason, items: old, error:String(e.message||e) };
  }
  writeStore(store);
  return store.news;
}
function scheduleDailyNewsRefresh(){
  const run = () => refreshNoiseNews('daily-08:00').catch(e => latestDebug.lastError = {at:now(), message:`news refresh failed: ${e.message||e}`});
  refreshNoiseNews('startup').catch(()=>{});
  function msUntilNextTaipei8(){
    const nowD = new Date();
    const twParts = new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Taipei',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false}).formatToParts(nowD).reduce((a,p)=>{a[p.type]=p.value;return a;},{});
    let targetUtc = Date.UTC(Number(twParts.year), Number(twParts.month)-1, Number(twParts.day), 0, 0, 0); // 08:00 Taipei = 00:00 UTC
    if(nowD.getTime() >= targetUtc) targetUtc += 24*60*60*1000;
    return Math.max(targetUtc - nowD.getTime(), 60*1000);
  }
  setTimeout(() => { run(); setInterval(run, 24*60*60*1000); }, msUntilNextTaipei8());
}
function getNewsCache(){
  const store = readStore();
  if(!store.news?.items?.length){
    store.news = { updatedAt:new Date().toISOString(), updatedAtTaipei:taipeiTimeText(), reason:'default', items:defaultNewsItems(), error:null };
    writeStore(store);
  }
  return store.news;
}
function formatNewsMessage(news){
  const items = (news.items||[]).slice(0,5);
  const lines = items.map((n,i)=>`■ ${i+1}. ${n.title}\n來源：${n.source || '新聞來源'}${n.publishedAt ? `｜發布：${taipeiTimeText(n.publishedAt)}` : ''}\n摘要：${n.summary || '請點選連結查看完整內容。'}`).join('\n\n');
  const err = news.error ? `\n\n⚠️ 今日自動擷取未完全成功，已保留最近一次新聞快取。` : '';
  return `【噪音車新聞】\n更新時間：${news.updatedAtTaipei || taipeiTimeText(news.updatedAt)}\n更新頻率：每日08:00自動更新；輸入「新聞」可即時查看。\n\n${lines}${err}`;
}
function lawCenter(){
  return textMessage('【法規中心】\n可查詢：噪音管制法、最新修法、聲音照相指引、NIEA P211.82B、噪音車新聞。\n\n常用指令：\n法條11、法條13、法條26、法條28、修法、聲音照相指引、NIEA82B、新聞', [
    {label:'最新修法',text:'修法'},
    {label:'法條26',text:'法條26'},
    {label:'聲音照相指引',text:'聲音照相指引'},
    {label:'NIEA 82B',text:'NIEA82B'},
    {label:'噪音車新聞',text:'新聞'},
    {label:'環境部專區',uri:'https://noisecar.moenv.gov.tw/'}
  ]);
}
function lawDetail(text){
  if(text.includes('11')) return textMessage('【噪音管制法 第11條】\n重點：車輛、機動車輛等噪音源應符合主管機關公告之噪音管制標準。\n\n實務：聲音照相執法會依車種、速限及噪音標準進行判定。');
  if(text.includes('13')) return textMessage('【噪音管制法 第13條】\n重點：主管機關得通知噪音源所有人或使用人到場檢驗、改善或提出說明。\n\n實務：對疑似噪音車可辦理通知到檢。');
  if(text.includes('26')) return textMessage('【噪音管制法 第26條】\n重點：違反噪音管制規定者，得依規定裁罰。\n\n噪音車修法重點：最高罰鍰提高至3萬6,000元；情節重大或一年內再犯，可吊扣牌照。');
  if(text.includes('28')) return textMessage('【噪音管制法 第28條】\n重點：未依通知檢驗、改善或提出資料者，可依規定處分。');
  return lawCenter();
}

function photoGuideMsg(){
  return textMessage('【聲音照相科技執法指引】\n一、用途：協助地方環保機關辦理機動車輛行駛噪音科技執法。\n二、核心流程：熱點分析 → 現勘評估 → 架設告示與設備 → 現場校正／對時 → 擷取影像與噪音資料 → 案件查證 → 告發或通知到檢。\n三、點位原則：以陳情熱點、歷史高噪音路段、可安全架設與可辨識車輛為優先。\n四、品質控管：確認噪音計、風速計、影像設備、時間同步、環境條件及車輛聲源判定。\n\n可再輸入：指引品管、指引點位、NIEA82B。', [
    {label:'NIEA 82B',text:'NIEA82B'},
    {label:'指引品管',text:'指引品管'},
    {label:'指引點位',text:'指引點位'},
    {label:'法規中心',text:'法規中心'}
  ]);
}
function guideQcMsg(){
  return textMessage('【聲音照相指引－品管重點】\n1. 執法前確認設備檢定、校正與比測狀態。\n2. 現場需確認風速、天候、背景噪音、車流與架設安全。\n3. 案件判定需結合噪音最大值、影像、車牌與聲源辨識。\n4. 若環境條件或量測系統確認不符要求，該時段資料不宜作為執法依據。\n5. 建議 LINE 設備管理同步追蹤：中央比測、噪音計檢定、風速計校正。', [
    {label:'設備管理',text:'設備管理'},
    {label:'NIEA 82B',text:'NIEA82B'},
    {label:'法規中心',text:'法規中心'}
  ]);
}
function guideSiteMsg(){
  return textMessage('【聲音照相指引－點位評估】\n建議優先檢核：\n1. 是否為噪音車陳情熱點或歷史高風險路段。\n2. 是否具備安全架設空間與穩定電力／通訊條件。\n3. 車牌、車道與聲源是否可清楚辨識。\n4. 是否避開高背景噪音、施工、強風或其他干擾源。\n5. 是否符合告示牌設置、現場安全與執法可及性。', [
    {label:'百大熱點',text:'百大點位'},
    {label:'成果查詢',uri:DASHBOARD_URL},
    {label:'指引品管',text:'指引品管'}
  ]);
}
function niea82bMsg(){
  return textMessage('【NIEA P211.82B】\n名稱：機動車輛行駛噪音量測方法－影像輔助法。\n狀態：環境部已公告訂定，自115/02/15生效。\n適用：機動車輛行駛於車道上產生噪音之量測。\n量測系統：可包含噪音計量測系統或陣列式聲音感應器等組合量測系統，並同步搭配影像輔助判定噪音源。\n資料重點：以行駛噪音最大值 Lmax 作為測試規範計算基礎。\n\n可再輸入：82B品管、82B比測、聲音照相指引。', [
    {label:'82B品管',text:'82B品管'},
    {label:'82B比測',text:'82B比測'},
    {label:'聲音照相指引',text:'聲音照相指引'},
    {label:'法規中心',text:'法規中心'}
  ]);
}
function niea82bQcMsg(){
  return textMessage('【NIEA P211.82B－品管與校正提醒】\n1. 風速計：每2年送至可追溯至國家量測標準的實驗室校正；器差不得超過 ±1.0 m/s，且至少一受校點需介於4～6 m/s。\n2. 噪音計量測系統及陣列式聲音感應器等組合量測系統：每2年送至符合 CNS 5799 試驗場地比測，結果須符合附表規定。\n3. 現場測試後需進行量測系統確認；若不符合品管要求，測試期間噪音數據無效。', [
    {label:'設備管理',text:'設備管理'},
    {label:'82B比測',text:'82B比測'},
    {label:'法規中心',text:'法規中心'}
  ]);
}
function niea82bCompareMsg(){
  return textMessage('【NIEA P211.82B－比測重點】\n比測目的：為執行噪音科技執法，確認候選方法與參考方法的一致性。\n測試情境：單一機車、單一汽車，以及多車或混合車流情境。\n速度條件：機車約40、50、60 km/h；汽車約50、60、70 km/h。\n多音源判定：陣列式聲音感應器等組合量測系統需能辨識多音源最大噪音位置。\n結果重點：測試規範計算皆以噪音最大值 Lmax 為基礎。', [
    {label:'82B品管',text:'82B品管'},
    {label:'聲音照相指引',text:'聲音照相指引'},
    {label:'法規中心',text:'法規中心'}
  ]);
}

function lawUpdate(){ return textMessage('【最新修法重點】\n1. 噪音車違規最高罰鍰提高至3萬6,000元。\n2. 情節重大者可吊扣牌照。\n3. 一年內再犯可加重處分。\n4. 後續執法需同步注意中央公告與地方裁罰基準。', [{label:'法條26',text:'法條26'}, {label:'噪音車新聞',text:'新聞'}]); }
function newsMsg(){
  const news = getNewsCache();
  // 若快取超過 23 小時，背景刷新；本次先回覆目前快取，避免 LINE webhook 逾時。
  const age = Date.now() - new Date(news.updatedAt || 0).getTime();
  if(!Number.isFinite(age) || age > 23*60*60*1000) refreshNoiseNews('on-demand').catch(()=>{});
  return textMessage(formatNewsMessage(news), [
    {label:'立即更新新聞',text:'更新新聞'},
    {label:'環境部專區',uri:'https://noisecar.moenv.gov.tw/'},
    {label:'法規中心',text:'法規中心'}
  ]);
}
function equipmentMsg(){
  const s=readStore(); const list=s.equipment||[]; const today=new Date();
  const rows = list.map(e=>{
    const dm = daysLeft(e.soundMeterDate, 365); const dw=daysLeft(e.windMeterDate,365); const db=daysLeft(e.bitestDate,730); const min=Math.min(dm,dw,db);
    const light = min < 0 ? '紅燈 已逾期' : min <= 30 ? '黃燈 30天內到期' : '綠燈 正常';
    return `${e.id}｜${light}\n比測剩餘：${db}天｜噪音計：${dm}天｜風速計：${dw}天`;
  }).join('\n\n');
  return textMessage(`【設備管理】\n比測週期：2年\n噪音計檢定：1年\n風速計檢定：1年\n\n${rows || '尚未匯入設備資料'}`, [{label:'匯入設備表',uri:`${BASE_URL}/admin.html`}, {label:'設備',text:'設備'}]);
}
function daysLeft(dateStr, cycleDays){
  const d = new Date(String(dateStr).replace(/\//g,'-')); if(isNaN(d)) return 9999;
  const due = new Date(d.getTime() + cycleDays*86400000);
  return Math.ceil((due - new Date())/86400000);
}
function fieldReportMsg(keyword=''){
  const s = readStore();
  const rows = s.fieldReports || [];
  if(!rows.length) return textMessage('目前尚未同步外勤回報架設點位資料。請確認 FIELD_GOOGLE_SHEET_ID 或 FIELD_GSHEET_XLSX_URL，並至後台按「從Google Sheet立即同步」。', [{label:'外勤平台', uri:FIELD_REPORT_URL}, {label:'後台同步', uri:`${BASE_URL}/admin.html`}]);
  const t = String(keyword||'').replace(/外勤|回報|架設點位|查詢|場次/g,'').trim();
  const district = (keyword.match(/[\u4e00-\u9fa5]{2,3}區/)||[])[0];
  const device = (keyword.match(/OE[_-]?ZB\d{3}/i)||[])[0];
  const session = (keyword.match(/S\d{1,4}/i)||[])[0];
  let filtered = rows;
  if(district) filtered = filtered.filter(r=>r.district===district);
  if(device) filtered = filtered.filter(r=>String(r.device).replace('-','_').toUpperCase().includes(device.replace('-','_').toUpperCase()));
  if(session) filtered = filtered.filter(r=>String(r.sessionId).toUpperCase()===session.toUpperCase());
  if(!district && !device && !session && t) filtered = filtered.filter(r => [r.sessionId,r.device,r.district,r.location,r.signLocation,r.reportId].join(' ').includes(t));
  const top = filtered.slice(0,8);
  const lines = top.map(r=>`${r.sessionId||'-'}｜${r.date||'-'}｜${r.district||'-'}｜${r.device||'-'}
地點：${r.location||'-'}
校正值：${r.calibration||'-'}｜限速${r.speedLimit||'-'}｜標準${r.noiseStandard||'-'}dB(A)
告示牌：${r.signLocation||'-'}｜距離${r.distance||'-'}m`).join('\n\n');
  return textMessage(`【外勤回報架設點位】
${lines || '查無符合條件的外勤回報資料。'}

可輸入：S01、OE_ZB001、淡水區外勤、架設點位
資料更新時間：${s.lastDataSync?.at || '未記錄'}`, [
    {label:'外勤回報平台', uri:FIELD_REPORT_URL},
    {label:'成果查詢', text:'進度'},
    {label:'百大點位', text:'百大點位'}
  ]);
}
function keywordSearchMsg(text){
  const t = String(text||'').replace(/搜尋|查詢/g,'').trim();
  if(!t) return textMessage('請輸入要搜尋的關鍵字，例如：搜尋 淡水、新莊區、OE_ZB001、S01、百大。', menuQuick());
  if(/S\d{1,4}|OE[_-]?ZB\d{3}|外勤|架設/.test(t)) return fieldReportMsg(t);
  if(t.includes('百大') || t.includes('熱點') || t.endsWith('區')) return hotspotListMsg(t.includes('百大') ? t : `${t}百大`);
  if(t.includes('成果') || t.includes('月份') || /\d+月/.test(t)) return routeText(t);
  return textMessage(`已收到關鍵字：「${t}」。\n可用：${t}百大、${t}執行成效、${t}外勤。`, [
    {label:`${t}百大`, text:`${t}百大`},
    {label:`${t}外勤`, text:`${t}外勤`},
    {label:'選單', text:'選單'}
  ]);
}
function platePrompt(){ return textMessage('請輸入車牌，例如：車牌 ABC-1234。', [{label:'範例 ABC-1234',text:'車牌 ABC-1234'}]); }
function plateQuery(text){
  const key = text.replace(/車牌|查詢|\s/g,'').toUpperCase(); const s=readStore(); const p=s.plates?.[key] || s.plates?.[key.replace('-','')];
  if(!p) return textMessage(`目前查無車牌 ${key} 的案件資料。`, [{label:'車號追蹤',text:'車號追蹤'}]);
  return textMessage(`【車號追蹤】\n車牌：${key}\n累犯次數：${p.repeat}次\n最高超標：${p.maxDb} dB\n最近日期：${p.lastDate}\n行政區：${p.district}\n告發件數：${p.fines}\n通知到檢：${p.notices}`, [{label:'成果平台',uri:DASHBOARD_URL}]);
}

function hotspotMsg(){
  return textMessage(`【新北市噪音車百大熱點分析平台】
可查看百大建議點位、熱區分布、優先布點與點位風險資訊。

請點下方按鈕開啟平台。`, [
    {label:'開啟百大熱點', uri:HOTSPOT_URL},
    {label:'成果查詢', uri:DASHBOARD_URL}
  ]);
}

function adminMsg(){ return textMessage(`【管理中心】\n後台網址：${BASE_URL}/admin.html\n\n後台密碼由 Zeabur 環境變數 ADMIN_PASSWORD 管理，不會在 LINE、GitHub 或前端頁面顯示。\n可操作：資料匯入、Rich Menu 更新、成果檢查、設備管理。`, [{label:'開啟後台',uri:`${BASE_URL}/admin.html`}]); }
function routeText(text){
  const t=String(text||'').trim();
  if(!t || t==='選單' || t==='menu') return mainMenu();
  if(t==='成果查詢') return textMessage('請開啟成果查詢平台，或使用月份／行政區快速查詢。', [{label:'成果平台',uri:DASHBOARD_URL}, {label:'月份選單',text:'月份選單'}, {label:'行政區選單',text:'行政區選單'}]);
  if(t==='外勤回報') return textMessage('請開啟外勤回報平台填寫場次、照片與座標；也可輸入「架設點位、S01、OE_ZB001、淡水區外勤」查詢已同步資料。', [{label:'外勤回報',uri:FIELD_REPORT_URL},{label:'架設點位',text:'架設點位'}]);
  if(['進度','KPI報表','KPI','kpi'].includes(t)) return t==='進度'?progressMsg():kpiMsg();
  if(t==='統計選單' || t==='統計查詢') return statMenu();
  if(t==='月份選單' || t==='月份') return monthMenu();
  if(t==='行政區選單' || t==='行政區') return districtMenu();
  if(t==='車號追蹤' || t==='車牌查詢') return platePrompt();
  if(t.includes('外勤') || t.includes('架設點位') || /^S\d{1,4}$/i.test(t) || /OE[_-]?ZB\d{3}/i.test(t)) return fieldReportMsg(t);
  if(/^\d+月份執行成效$/.test(t)) return monthStats(t.match(/^(\d+)/)[1]);
  if(/^\d+月$/.test(t)) return monthStats(t.match(/^(\d+)/)[1]);
  if(t.endsWith('區執行成效')) return districtStats(t.replace('執行成效',''));
  if(t.endsWith('區')) return districtStats(t);
  if(t==='法規中心') return lawCenter();
  if(t==='聲音照相指引' || t==='指引' || t==='聲音照相') return photoGuideMsg();
  if(t==='指引品管') return guideQcMsg();
  if(t==='指引點位' || t==='設置原則') return guideSiteMsg();
  if(/^NIEA\s*82B$/i.test(t) || /^NIEA\s*P?211\.82B$/i.test(t) || t==='NIEA82B' || t==='82B') return niea82bMsg();
  if(t==='82B品管' || t==='NIEA品管') return niea82bQcMsg();
  if(t==='82B比測' || t==='NIEA比測') return niea82bCompareMsg();
  if(t.startsWith('法條')) return lawDetail(t);
  if(t.includes('修法')) return lawUpdate();
  if(t==='更新新聞' || t==='重新整理新聞') { refreshNoiseNews('manual-line').catch(()=>{}); return textMessage('已收到更新新聞指令，系統會在背景重新整理。請約 5～10 秒後再輸入「新聞」查看最新快取。', [{label:'查看新聞',text:'新聞'}, {label:'法規中心',text:'法規中心'}]); }
  if(t.includes('新聞')) return newsMsg();
  if(t.includes('陳情趨勢') || t.includes('陳情同期比')) return complaintTrendMsg(t);
  if(t.includes('百大計算方式') || t.includes('百大公式') || t.includes('百大評分')) return hotspotMethodMsg();
  if(t.includes('百大點位') || t.includes('百大熱點') || t.includes('熱點分析') || t.endsWith('區百大') || /^第\d{1,3}名/.test(t) || /^排名\d{1,3}/.test(t)) return hotspotListMsg(t);
  if(t.startsWith('搜尋') || t.startsWith('查詢')) return keywordSearchMsg(t);
  if(t==='設備管理' || t==='設備') return equipmentMsg();
  if(t==='管理功能' || t==='管理中心') return equipmentMsg();
  if(/車牌|[A-Z]{2,4}-?\d{3,4}/i.test(t)) return plateQuery(t);
  return textMessage('感謝您的回覆🙂', menuQuick());
}

app.get('/healthz', (req,res)=>res.json({ok:true, service:'newtaipei-noise-control-system-v27-three-gsheet-line-search', hasAdminPassword:!!ADMIN_PASSWORD, hasSessionSecret:!!SESSION_SECRET}));
app.get('/api/line/test', (req,res)=>res.json({ok:true, service:'v27-three-gsheet-line-search', hasToken:!!LINE_TOKEN, hasSecret:!!LINE_SECRET, hasAdminPassword:!!ADMIN_PASSWORD, hasSessionSecret:!!SESSION_SECRET}));
app.get('/api/line/debug/latest', (req,res)=>res.json({ok:true, debug:latestDebug}));
app.get('/api/line/rich-menu-spec', (req,res)=>res.json({ok:true, image:`${BASE_URL}/assets/line-rich-menu.jpg`, spec: richMenuSpec()}));
app.get('/api/data/summary', (req,res)=>res.json({ok:true, data:readStore()}));
app.get('/api/legal/reference', (req,res)=>res.json({ok:true, data:{ lawCenter:['法條11','法條13','法條26','法條28','修法','新聞','更新新聞'], soundPhotoGuide:['聲音照相指引','指引品管','指引點位'], niea82b:['NIEA82B','82B品管','82B比測'] }}));
app.get('/api/legal/news', (req,res)=>res.json({ok:true, news:getNewsCache()}));
app.post('/api/admin/news/refresh', requireAdmin, async (req,res)=>{ try{ res.json({ok:true, news:await refreshNoiseNews('manual-admin')}); } catch(e){ res.status(500).json({ok:false, error:String(e.message||e)}); } });
app.get('/api/admin/status', (req,res)=>res.json({ok:true, ...authStatus(req)}));
app.post('/api/admin/login', express.urlencoded({ extended:true }), (req,res)=>{
  if(!ADMIN_PASSWORD) return res.status(503).json({ok:false, error:'後台密碼尚未在 Zeabur 環境變數 ADMIN_PASSWORD 設定'});
  const pwd = String(req.body?.password || '');
  const ok = pwd.length === ADMIN_PASSWORD.length && crypto.timingSafeEqual(Buffer.from(pwd), Buffer.from(ADMIN_PASSWORD));
  if(!ok) return res.status(401).json({ok:false, error:'密碼錯誤'});
  const secure = (BASE_URL.startsWith('https://')) ? '; Secure' : '';
  res.setHeader('Set-Cookie', `ntpc_admin=${encodeURIComponent(makeSession())}; HttpOnly; Path=/; Max-Age=43200; SameSite=Lax${secure}`);
  res.json({ok:true});
});
app.post('/api/admin/logout', (req,res)=>{ res.setHeader('Set-Cookie','ntpc_admin=; HttpOnly; Path=/; Max-Age=0; SameSite=Lax'); res.json({ok:true}); });
app.post('/api/line/webhook', async (req,res)=>{
  res.status(200).send('OK');
  const valid = verifySignature(req);
  latestDebug.lastSignature = { at:now(), valid, hasSignature:!!req.get('x-line-signature') };
  if(!valid) return;
  const events = req.body?.events || [];
  for (const ev of events){
    if(ev.type !== 'message' || ev.message?.type !== 'text') continue;
    const text = ev.message.text;
    recordEvent({ at:now(), type:'message', userId:ev.source?.userId, text });
    try { await lineReply(ev.replyToken, routeText(text)); }
    catch(err){ latestDebug.lastError = { at:now(), message:String(err.message||err), stack:String(err.stack||'') }; }
  }
});
function richMenuSpec(){
  return { size:{width:2500,height:1686}, selected:true, name:'新北噪音車V26百大建議點位版圖文選單', chatBarText:'管理選單', areas:[
    {bounds:{x:0,y:320,width:625,height:683}, action:{type:'uri', uri:DASHBOARD_URL}},
    {bounds:{x:625,y:320,width:625,height:683}, action:{type:'uri', uri:FIELD_REPORT_URL}},
    {bounds:{x:1250,y:320,width:625,height:683}, action:{type:'message', text:'車號追蹤'}},
    {bounds:{x:1875,y:320,width:625,height:683}, action:{type:'message', text:'KPI報表'}},
    {bounds:{x:0,y:1003,width:625,height:683}, action:{type:'message', text:'統計選單'}},
    {bounds:{x:625,y:1003,width:625,height:683}, action:{type:'message', text:'法規中心'}},
    {bounds:{x:1250,y:1003,width:625,height:683}, action:{type:'uri', uri:HOTSPOT_URL}},
    {bounds:{x:1875,y:1003,width:625,height:683}, action:{type:'message', text:'設備管理'}}
  ]};
}
async function lineApi(pathname, options={}){
  const r = await fetch(`https://api.line.me/v2/bot${pathname}`, { ...options, headers:{ ...(options.headers||{}), Authorization:`Bearer ${LINE_TOKEN}` } });
  const text = await r.text();
  if(!r.ok) throw new Error(`${pathname} ${r.status}: ${text}`);
  return text ? JSON.parse(text) : {};
}
app.post('/api/admin/rich-menu/setup', requireAdmin, async (req,res)=>{
  try{
    if(!LINE_TOKEN) throw new Error('LINE_CHANNEL_ACCESS_TOKEN 未設定');
    const list = await lineApi('/richmenu/list');
    for (const rm of (list.richmenus||[])) { if(String(rm.name||'').includes('新北噪音車')) await lineApi(`/richmenu/${rm.richMenuId}`, {method:'DELETE'}); }
    const created = await lineApi('/richmenu', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(richMenuSpec()) });
    const img = fs.readFileSync(path.join(__dirname,'public/assets/line-rich-menu.jpg'));
    const up = await fetch(`https://api-data.line.me/v2/bot/richmenu/${created.richMenuId}/content`, { method:'POST', headers:{ Authorization:`Bearer ${LINE_TOKEN}`, 'Content-Type':'image/jpeg'}, body:img });
    const upText = await up.text(); if(!up.ok) throw new Error(`upload ${up.status}: ${upText}`);
    await lineApi(`/user/all/richmenu/${created.richMenuId}`, { method:'POST' });
    res.json({ok:true, richMenuId:created.richMenuId});
  }catch(e){ res.status(500).json({ok:false, error:String(e.message||e)}); }
});
app.get('/api/admin/gsheet/status', requireAdmin, (req,res)=>res.json({ok:true, configured:!!GSHEET_XLSX_URL, lastDataSync:readStore().lastDataSync||null}));
app.post('/api/admin/gsheet/sync', requireAdmin, async (req,res)=>{ try{ res.json({ok:true, sync:await syncFromGoogleSheet()}); } catch(e){ res.status(500).json({ok:false, error:String(e.message||e)}); } });
app.post('/api/admin/upload-excel', requireAdmin, upload.single('file'), (req,res)=>{
  try{
    if(!req.file) throw new Error('未收到檔案');
    const wb = XLSX.readFile(req.file.path);
    const sync = importWorkbookToStore(wb, `excel-upload:${req.file.originalname}`);
    res.json({ok:true, sync});
  }catch(e){ res.status(500).json({ok:false, error:String(e.message||e)}); }
});

app.get('/api/admin/export/unified-xlsx', requireAdmin, (req,res)=>{ try{ sendWorkbook(res, buildExportWorkbook('all'), `新北市噪音車管理系統_全資料匯出_${new Date().toISOString().slice(0,10)}.xlsx`); }catch(e){ res.status(500).json({ok:false,error:String(e.message||e)}); } });
app.get('/api/admin/export/unified-json', requireAdmin, (req,res)=>{ res.json({ok:true, exportedAt:new Date().toISOString(), data:buildExportSheets()}); });
app.get('/api/admin/export/template-xlsx', requireAdmin, (req,res)=>res.download(path.join(__dirname,'templates','新北市噪音車管理系統_三平台匯入匯出範本_v25.xlsx')));
app.get('/api/admin/export/results-xlsx', requireAdmin, (req,res)=>sendWorkbook(res, buildExportWorkbook('監測成果'), '監測成果_匯出.xlsx'));
app.get('/api/admin/export/hotspots-xlsx', requireAdmin, (req,res)=>sendWorkbook(res, buildExportWorkbook('百大點位'), '百大點位_匯出.xlsx'));
app.get('/api/admin/export/equipment-xlsx', requireAdmin, (req,res)=>sendWorkbook(res, buildExportWorkbook('設備提醒'), '設備管理_匯出.xlsx'));
app.get('/api/admin/export/complaints-xlsx', requireAdmin, (req,res)=>sendWorkbook(res, buildExportWorkbook('陳情趨勢'), '陳情趨勢_匯出.xlsx'));

app.get('/', (req,res)=>res.redirect('/admin.html'));
app.get('/admin.html', (req,res)=>res.send(isAdmin(req) ? adminHtml() : loginHtml(req)));
app.get('/line-bot.html', (req,res)=>res.send(lineBotHtml()));
function htmlLayout(title, body){return `<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${title}</title><style>
body{margin:0;font-family:Arial,'Noto Sans TC','Microsoft JhengHei',sans-serif;background:radial-gradient(circle at 20% 0%,#dff3ff 0,#eef6ff 34%,#f8fbff 100%);color:#06295c}.hero{background:linear-gradient(120deg,#00265f,#005bc4 48%,#00a7ff);color:white;padding:34px 40px;border-radius:0 0 32px 32px;box-shadow:0 18px 42px #a9cde8;position:relative;overflow:hidden}.hero:after{content:'';position:absolute;right:-80px;bottom:-90px;width:420px;height:260px;background:radial-gradient(circle,#7ee8ff55,transparent 70%)}.wrap{max-width:1180px;margin:30px auto;padding:0 22px}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:20px}.card{background:linear-gradient(180deg,#ffffff,#f5fbff);border:1px solid #b8d9f5;border-radius:22px;padding:24px;box-shadow:0 14px 28px #bdd8ee}.btn{display:inline-block;border:0;border-radius:13px;background:linear-gradient(135deg,#005bd6,#00a7ff);color:white;padding:13px 18px;font-weight:800;cursor:pointer;text-decoration:none;box-shadow:0 8px 16px #b7d2ec}.btn2{background:linear-gradient(135deg,#009a77,#00d6b0)}.btn3{background:linear-gradient(135deg,#35475d,#718199)}.muted{color:#62748a}.code{background:#e9f2fb;border-radius:12px;padding:14px;overflow:auto}input[type=file]{padding:10px;background:#f6fbff;border:1px solid #bfd7ed;border-radius:10px}</style></head><body>${body}</body></html>`}
function loginHtml(req){ return htmlLayout('後台登入',`<div class="hero"><h1>新北市打擊噪音車管理系統</h1><p>後台登入｜密碼由 Zeabur 環境變數管理</p></div><div class="wrap"><div class="card"><h2>管理者登入</h2><p class="muted">系統不會在程式、GitHub、README 或前端頁面顯示管理密碼。</p><form id="login"><input type="password" name="password" placeholder="請輸入後台管理密碼" style="width:100%;box-sizing:border-box;padding:14px;border:1px solid #bfd7ed;border-radius:12px;font-size:16px"><br><br><button class="btn">登入</button></form><pre id="msg" class="code"></pre></div></div><script>document.getElementById('login').onsubmit=async e=>{e.preventDefault();const r=await fetch('/api/admin/login',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams(new FormData(e.target))});const j=await r.json();if(j.ok) location.href='/admin.html'; else document.getElementById('msg').textContent=j.error||'登入失敗';}</script>`)}
function adminHtml(){const s=readStore();return htmlLayout('後台管理',`<div class="hero"><h1>新北市打擊噪音車管理系統 V27</h1><p>後台管理｜LINE 圖文選單｜資料匯入｜系統檢查</p></div><div class="wrap"><div class="grid"><div class="card"><h2>系統狀態</h2><p>版本：V27 三平台 Google Sheet 分流同步 + LINE 關鍵字搜尋</p><p>Webhook：${BASE_URL}/api/line/webhook</p><p class="muted">管理密碼由 Zeabur ADMIN_PASSWORD 管理，頁面不顯示實際密碼。</p><button class="btn btn3" onclick="logout()">登出</button><a class="btn" href="/healthz" target="_blank">Health Check</a> <a class="btn btn3" href="/api/line/debug/latest" target="_blank">Debug</a></div><div class="card"><h2>Rich Menu</h2><p class="muted">重新產生科技風圖文選單並設為預設。</p><button class="btn" onclick="setupRichMenu()">一鍵建立／更新 LINE 圖文選單</button><pre id="rich" class="code"></pre></div><div class="card"><h2>成果摘要</h2><p>場次：${fmt(s.summary.sessions)} / ${fmt(s.annualGoal)}</p><p>車流：${fmt(s.summary.traffic)}</p><p>超標：${fmt(s.summary.exceed)}</p><p>告發：${fmt(s.summary.fines)}｜通檢：${fmt(s.summary.notices)}</p><p>KPI：${kpi(s.summary)}</p></div><div class="card"><h2>法規新聞</h2><p class="muted">每日08:00自動更新，並於 LINE 回覆內顯示更新時間。</p><button class="btn" onclick="refreshNews()">立即更新新聞</button><a class="btn btn3" href="/api/legal/news" target="_blank">查看新聞JSON</a><pre id="news" class="code"></pre></div><div class="card"><h2>資料匯入／Google Sheet連動</h2><p class="muted">匯入優化版範本後，系統會更新成果、月份、行政區、設備、百大點位與陳情趨勢。</p><p><a class="btn btn3" href="/templates/新北市噪音車管理系統_三平台匯入匯出範本_v25.xlsx">下載Google Sheet範本</a> <a class="btn btn3" href="/templates/前100大加權排行_詳細版_系統匯入範本_v26.xlsx">下載百大點位範本</a></p><form id="upload"><input type="file" name="file" accept=".xlsx,.xls"><button class="btn btn2">上傳並更新系統數據</button></form><hr><button class="btn" onclick="syncGSheet()">從Google Sheet立即同步</button><a class="btn btn3" href="/api/admin/gsheet/status" target="_blank">查看同步狀態</a><pre id="up" class="code"></pre></div></div><div class="card" style="margin-top:18px"><h2>最新匯出中心</h2><p class="muted">可匯出三平台分類後的統一資料、成果、百大點位、設備、陳情趨勢與範本。</p><a class="btn" href="/api/admin/export/unified-xlsx">匯出全資料Excel</a> <a class="btn btn2" href="/api/admin/export/template-xlsx">下載三平台匯入匯出範本</a> <a class="btn btn3" href="/api/admin/export/unified-json" target="_blank">匯出JSON</a><br><br><a class="btn btn3" href="/api/admin/export/results-xlsx">成果匯出</a> <a class="btn btn3" href="/api/admin/export/hotspots-xlsx">百大匯出</a> <a class="btn btn3" href="/api/admin/export/equipment-xlsx">設備匯出</a> <a class="btn btn3" href="/api/admin/export/complaints-xlsx">陳情匯出</a></div><div class="card" style="margin-top:18px"><h2>平台連結</h2><a class="btn" href="${DASHBOARD_URL}" target="_blank">成果查詢系統</a> <a class="btn btn2" href="${FIELD_REPORT_URL}" target="_blank">外勤回報平台</a> <a class="btn" href="${HOTSPOT_URL}" target="_blank">百大熱點平台</a> <a class="btn btn3" href="/line-bot.html">LINE設定頁</a></div></div><script>
async function setupRichMenu(){ const r=await fetch('/api/admin/rich-menu/setup',{method:'POST'}); document.getElementById('rich').textContent=JSON.stringify(await r.json(),null,2); }
document.getElementById('upload').onsubmit=async e=>{e.preventDefault(); const fd=new FormData(e.target); const r=await fetch('/api/admin/upload-excel',{method:'POST',body:fd}); document.getElementById('up').textContent=JSON.stringify(await r.json(),null,2)};
async function syncGSheet(){ const r=await fetch('/api/admin/gsheet/sync',{method:'POST'}); document.getElementById('up').textContent=JSON.stringify(await r.json(),null,2); }
async function refreshNews(){ const r=await fetch('/api/admin/news/refresh',{method:'POST'}); document.getElementById('news').textContent=JSON.stringify(await r.json(),null,2); }
async function logout(){ await fetch('/api/admin/logout',{method:'POST'}); location.href='/admin.html';}
</script>`)}
function lineBotHtml(){return htmlLayout('LINE BOT設定',`<div class="hero"><h1>LINE BOT 操作與設定</h1><p>Webhook、Rich Menu、常用指令</p></div><div class="wrap"><div class="grid"><div class="card"><h2>Webhook URL</h2><div class="code">${BASE_URL}/api/line/webhook</div><p>LINE Developers Verify 成功後，請確認 Use webhook 開啟。</p></div><div class="card"><h2>常用指令</h2><p>進度、KPI報表、統計選單、月份選單、行政區選單、陳情趨勢、百大點位、第1名詳細、淡水區百大、百大計算方式、架設點位、S01、OE_ZB001、淡水區外勤、法規中心、NIEA82B、設備管理、車牌 ABC-1234。</p></div><div class="card"><h2>圖文選單預覽</h2><img src="/assets/line-rich-menu.jpg" style="width:100%;border-radius:14px;border:1px solid #c9dff3"></div></div></div>`)}
scheduleDailyNewsRefresh();
if(GSHEET_SYNC_INTERVAL_MIN > 0 && (GSHEET_XLSX_URL || TOP100_GSHEET_XLSX_URL || RESULTS_GSHEET_XLSX_URL || FIELD_GSHEET_XLSX_URL)){
  syncFromGoogleSheet().catch(e=>latestDebug.lastError={at:now(), message:`initial gsheet sync failed: ${e.message||e}`});
  setInterval(()=>syncFromGoogleSheet().catch(e=>latestDebug.lastError={at:now(), message:`gsheet sync failed: ${e.message||e}`}), GSHEET_SYNC_INTERVAL_MIN*60*1000);
}
app.listen(PORT, ()=> console.log(`New Taipei V26 Hotspot Top100 Line running on :${PORT}`));
