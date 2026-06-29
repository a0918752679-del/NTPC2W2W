const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const vm = require('vm');
const XLSX = require('xlsx');
const crypto = require('crypto');
const { google } = require('googleapis');

const app = express();
app.use(cors());
app.use(express.json({ limit: '20mb' }));
app.use(express.static(path.join(__dirname, 'public')));
// templates downloads are protected below through /api/template and /templates/:name

function buildTemplateWorkbookBuffer() {
  const wb = XLSX.utils.book_new();
  const sheets = [
    ['監測點位匯入', HEADERS.monitor, [
      ['移動式','移動式聲音照相','115','聯韻','淡水區','新市一路一段','新北市淡水區新市一路一段','115.06.12','25.1865425','121.4332440','OE_ZB004','S199','3','0','1','5','2','範例資料']
    ]],
    ['陳情案件匯入', HEADERS.complaint, [
      ['中央噪音檢舉網','115','115.06.12','淡水區','新市一路一段','新北市淡水區新市一路一段','25.1865425','121.4332440','3','2','1','0','5','2','0.4','範例資料']
    ]],
    ['百大排行輸出', HEADERS.ranking, []],
    ['系統設定', HEADERS.settings, [
      ['WEIGHT_CENTRAL','30','中央噪音檢舉網權重'],
      ['WEIGHT_LOCAL','25','市政信箱權重'],
      ['WEIGHT_HOT','25','熱區密度權重'],
      ['WEIGHT_NIGHT','10','夜間比例權重'],
      ['WEIGHT_GAP','10','布點缺口權重'],
      ['COVERAGE_DISTANCE_M','500','涵蓋距離']
    ]]
  ];
  for (const [name, header, rows] of sheets) {
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

app.get('/api/template', requireAdmin, (req, res) => {
  const filename = encodeURIComponent('新北市聲音照相百大平台_固定匯入匯出範本清冊.xlsx');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
  res.send(buildTemplateWorkbookBuffer());
});

app.get('/templates/:name', requireAdmin, (req, res) => {
  const filePath = path.join(__dirname, 'templates', req.params.name);
  if (fs.existsSync(filePath)) return res.sendFile(filePath);
  const filename = encodeURIComponent('新北市聲音照相百大平台_固定匯入匯出範本清冊.xlsx');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
  res.send(buildTemplateWorkbookBuffer());
});

const VERSION = '1.0.6-admin-guard';
const PUBLIC_DIR = path.join(__dirname, 'public');
const DATA_JS = path.join(PUBLIC_DIR, 'data.js');


const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'Wayne0118';
const ADMIN_SESSION_SECRET = process.env.ADMIN_SESSION_SECRET || `${ADMIN_PASSWORD}:ntpc-noise-admin`;
const ADMIN_TOKEN_TTL_MS = Number(process.env.ADMIN_TOKEN_TTL_MS || 8 * 60 * 60 * 1000);

function b64url(input) {
  return Buffer.from(input).toString('base64url');
}
function safeJsonParse(text) {
  try { return JSON.parse(text); } catch (_) { return null; }
}
function signPayload(payload) {
  const body = b64url(JSON.stringify(payload));
  const sig = crypto.createHmac('sha256', ADMIN_SESSION_SECRET).update(body).digest('base64url');
  return `${body}.${sig}`;
}
function verifyAdminToken(token) {
  if (!token || typeof token !== 'string' || !token.includes('.')) return false;
  const [body, sig] = token.split('.', 2);
  const expected = crypto.createHmac('sha256', ADMIN_SESSION_SECRET).update(body).digest('base64url');
  const sigBuf = Buffer.from(sig || '');
  const expectedBuf = Buffer.from(expected);
  if (!sig || sigBuf.length !== expectedBuf.length || !crypto.timingSafeEqual(sigBuf, expectedBuf)) return false;
  const payload = safeJsonParse(Buffer.from(body, 'base64url').toString('utf8'));
  if (!payload || payload.scope !== 'admin') return false;
  return Number(payload.exp || 0) > Date.now();
}
function tokenFromRequest(req) {
  const auth = String(req.headers.authorization || '');
  if (auth.startsWith('Bearer ')) return auth.slice(7).trim();
  return String(req.headers['x-admin-token'] || req.query.admin_token || '').trim();
}
function hasAdminAccess(req) {
  return verifyAdminToken(tokenFromRequest(req));
}
function requireAdmin(req, res, next) {
  if (hasAdminAccess(req)) return next();
  return res.status(401).json({ error: '需要後端管理權限。請先輸入管理密碼登入。' });
}

const SHEETS = {
  monitor: process.env.SHEET_MONITOR || '監測點位匯入',
  complaint: process.env.SHEET_COMPLAINT || '陳情案件匯入',
  ranking: process.env.SHEET_RANKING || '百大排行輸出',
  settings: process.env.SHEET_SETTINGS || '系統設定'
};

const HEADERS = {
  monitor: ['資料圖層','資料類型','年份','廠商','行政區','路段','地址','日期','緯度','經度','設備機號','執行場次','監測天數','告發件數','通知到檢件數','超標車輛','夜間案件','備註'],
  complaint: ['資料來源','年份','日期','行政區','路段','地址','緯度','經度','案件量','中央檢舉量','市政信箱量','警察案件量','超標車輛','夜間案件','夜間比例','備註'],
  ranking: ['rank','grade','score','district','road','address','lat','lng','檢舉量','市政信箱量','超標車輛','夜間比例','歷年執行次數','最近架設點距離m','500m內固定點','500m內移動點','gap_level','reason','updated_at'],
  settings: ['設定項目','設定值','說明']
};

let cache = { data: loadEmbeddedData(), source: 'embedded', updatedAt: null };

function loadEmbeddedData() {
  try {
    const code = fs.readFileSync(DATA_JS, 'utf8');
    const sandbox = { window: {}, console };
    vm.createContext(sandbox);
    vm.runInContext(code, sandbox, { timeout: 3000 });
    return sandbox.window.NOISE_DATA || { summary: {}, points: [], candidates: [] };
  } catch (err) {
    console.error('loadEmbeddedData failed', err);
    return { summary: {}, points: [], candidates: [] };
  }
}

function getWeights() {
  return {
    central: Number(process.env.WEIGHT_CENTRAL || 30),
    local: Number(process.env.WEIGHT_LOCAL || 25),
    hot: Number(process.env.WEIGHT_HOT || 25),
    night: Number(process.env.WEIGHT_NIGHT || 10),
    gap: Number(process.env.WEIGHT_GAP || 10),
  };
}
function coverageDistance() { return Number(process.env.COVERAGE_DISTANCE_M || 500); }
function nowIso() { return new Date().toISOString(); }
function num(v, fallback = 0) { const n = Number(v); return Number.isFinite(n) ? n : fallback; }
function clean(v) { return v === undefined || v === null ? '' : String(v).trim(); }
function inferDistrict(address='') {
  const m = String(address).match(/([\u4e00-\u9fa5]{1,3}[區鄉鎮市])/);
  return m ? m[1] : '';
}
function get(row, names) {
  for (const n of names) if (row[n] !== undefined && row[n] !== '') return row[n];
  return '';
}
function distanceM(a, b) {
  const R = 6371000;
  const lat1 = num(a.lat) * Math.PI / 180, lat2 = num(b.lat) * Math.PI / 180;
  const dLat = (num(b.lat) - num(a.lat)) * Math.PI / 180;
  const dLng = (num(b.lng) - num(a.lng)) * Math.PI / 180;
  const s = Math.sin(dLat/2)**2 + Math.cos(lat1)*Math.cos(lat2)*Math.sin(dLng/2)**2;
  return 2 * R * Math.atan2(Math.sqrt(s), Math.sqrt(1-s));
}

function normalizeRows(values) {
  if (!values || values.length < 2) return [];
  const headers = values[0].map(clean);
  return values.slice(1).filter(r => r.some(x => clean(x))).map(r => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = r[i] ?? ''; });
    return obj;
  });
}

function monitorRowToPoint(r, i) {
  const lat = num(get(r, ['緯度','lat','latitude']), NaN);
  const lng = num(get(r, ['經度','lng','longitude']), NaN);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
  const layer = clean(get(r, ['資料圖層'])) || '聲音照相監測點位';
  const vendor = clean(get(r, ['廠商','vendor']));
  const address = clean(get(r, ['地址','地點']));
  return {
    id: `GS_MON_${i+1}`,
    layer: layer.includes('聲音照相') ? layer : `${layer}聲音照相監測點位`,
    subtype: clean(get(r, ['資料類型'])) || 'Google Sheet 匯入監測點',
    year: clean(get(r, ['年份','年度'])), vendor,
    district: clean(get(r, ['行政區'])) || inferDistrict(address),
    road: clean(get(r, ['路段'])), address,
    date: clean(get(r, ['日期'])), lat, lng,
    device_no: clean(get(r, ['設備機號'])), session_no: clean(get(r, ['執行場次'])),
    監測天數: num(get(r, ['監測天數']), 0),
    告發件數: num(get(r, ['告發件數']), 0),
    通知到檢件數: num(get(r, ['通知到檢件數']), 0),
    超標車輛: num(get(r, ['超標車輛']), 0),
    夜間案件: num(get(r, ['夜間案件']), 0),
    geo_source: 'Google Sheet'
  };
}

function complaintRowToPoint(r, i) {
  const lat = num(get(r, ['緯度','lat','latitude']), NaN);
  const lng = num(get(r, ['經度','lng','longitude']), NaN);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
  const src = clean(get(r, ['資料來源'])) || 'Google Sheet 陳情案件';
  const address = clean(get(r, ['地址','地點']));
  const central = num(get(r, ['中央檢舉量','檢舉量']), src.includes('中央') ? num(get(r, ['案件量']), 1) : 0);
  const local = num(get(r, ['市政信箱量']), src.includes('1999') || src.includes('市政') ? num(get(r, ['案件量']), 1) : 0);
  const police = num(get(r, ['警察案件量']), src.includes('警察') ? num(get(r, ['案件量']), 1) : 0);
  const hot = num(get(r, ['超標車輛']), police);
  const layer = src.includes('1999') || src.includes('市政') ? '市政信箱1999陳情點' : (src.includes('警察') ? '警察稽查案件' : '噪音檢舉網中央陳情點');
  return {
    id: `GS_CMP_${i+1}`,
    layer, subtype: src, source: 'Google Sheet',
    year: clean(get(r, ['年份','年度'])),
    district: clean(get(r, ['行政區'])) || inferDistrict(address),
    road: clean(get(r, ['路段'])), address,
    date: clean(get(r, ['日期'])), lat, lng,
    案件量: num(get(r, ['案件量']), Math.max(central + local + police, 1)),
    中央檢舉量: central, 市政信箱量: local, 警察案件量: police,
    超標車輛: hot, 夜間案件: num(get(r, ['夜間案件']), 0),
    夜間比例: num(get(r, ['夜間比例']), 0), geo_source: 'Google Sheet'
  };
}

function recalcTop100(points) {
  const monitors = points.filter(p => String(p.layer || '').includes('聲音照相監測點位') && Number.isFinite(num(p.lat, NaN)) && Number.isFinite(num(p.lng, NaN)));
  const complaints = points.filter(p => !String(p.layer || '').includes('聲音照相監測點位') && Number.isFinite(num(p.lat, NaN)) && Number.isFinite(num(p.lng, NaN)));
  const enriched = complaints.map(p => {
    let minD = Infinity, fixed = 0, mobile = 0;
    for (const m of monitors) {
      const d = distanceM(p, m);
      if (d < minD) minD = d;
      if (d <= coverageDistance()) {
        if (String(m.layer).includes('固定')) fixed++;
        if (String(m.layer).includes('移動')) mobile++;
      }
    }
    const gap = minD === Infinity || minD > coverageDistance();
    const central = num(p.中央檢舉量 ?? p.檢舉量, 0);
    const local = num(p.市政信箱量, 0);
    const hot = num(p.超標車輛 ?? p.警察案件量, 0);
    const nightRatio = num(p.夜間比例, 0) || (num(p.夜間案件,0) > 0 ? Math.min(1, num(p.夜間案件,0) / Math.max(num(p.案件量,1), 1)) : 0);
    return { ...p, 檢舉量: central, 市政信箱量: local, 超標車輛: hot, 夜間比例: nightRatio,
      歷年執行次數: fixed + mobile, 最近架設點距離m: minD === Infinity ? '' : Math.round(minD * 10) / 10,
      '500m內固定點': fixed, '500m內移動點': mobile, 布點缺口原始值: gap ? 1 : 0,
      gap_level: gap ? '布點缺口' : '已涵蓋' };
  });
  const max = key => Math.max(1, ...enriched.map(x => num(x[key], 0)));
  const maxCentral = max('檢舉量'), maxLocal = max('市政信箱量'), maxHot = max('超標車輛');
  const w = getWeights();
  return enriched.map(x => {
    const n1 = num(x.檢舉量,0) / maxCentral;
    const n2 = num(x.市政信箱量,0) / maxLocal;
    const n3 = num(x.超標車輛,0) / maxHot;
    const n4 = Math.max(0, Math.min(1, num(x.夜間比例,0)));
    const n5 = x.gap_level === '布點缺口' ? 1 : 0;
    const score = Math.round((n1*w.central + n2*w.local + n3*w.hot + n4*w.night + n5*w.gap) * 100) / 100;
    return { ...x, n1, n2, n3, n4, n5, score };
  }).sort((a,b) => b.score - a.score).slice(0, 100).map((x, i) => ({
    district: x.district || '', road: x.road || '', address: x.address || '', lat: x.lat, lng: x.lng,
    檢舉量: num(x.檢舉量,0), 市政信箱量: num(x.市政信箱量,0), 超標車輛: num(x.超標車輛,0), 夜間比例: num(x.夜間比例,0),
    歷年執行次數: num(x.歷年執行次數,0), 最近架設點距離m: x.最近架設點距離m,
    '500m內固定點': num(x['500m內固定點'],0), '500m內移動點': num(x['500m內移動點'],0),
    score: x.score, gap_level: x.gap_level,
    reason: `中央檢舉${num(x.檢舉量,0)}件、市政信箱1999 ${num(x.市政信箱量,0)}件、熱點/超標${num(x.超標車輛,0)}件；${coverageDistance()}公尺內固定${num(x['500m內固定點'],0)}點、移動${num(x['500m內移動點'],0)}點。`,
    rank: i + 1, grade: i < 10 ? 'S' : i < 30 ? 'A' : i < 60 ? 'B' : 'C'
  }));
}

function parseServiceAccountEnv(raw) {
  if (!raw) return null;
  let text = String(raw).trim();
  if ((text.startsWith('\"') && text.endsWith('\"')) || (text.startsWith("'") && text.endsWith("'"))) {
    text = text.slice(1, -1).trim();
  }

  const candidates = [text];

  // Zeabur 有時候使用者會貼到 base64 內容，例如 ewogICJ0eXBl...；此處自動解碼。
  const compact = text.replace(/\s+/g, '');
  if (/^[A-Za-z0-9+/=_-]+$/.test(compact) && !compact.startsWith('{')) {
    try {
      const normalized = compact.replace(/-/g, '+').replace(/_/g, '/');
      const decoded = Buffer.from(normalized, 'base64').toString('utf8').trim();
      if (decoded) candidates.push(decoded);
    } catch (_) {}
  }

  // 支援被雙重轉義的 JSON 字串。
  try {
    const unescaped = JSON.parse(text);
    if (typeof unescaped === 'string') candidates.push(unescaped.trim());
    if (unescaped && typeof unescaped === 'object') return normalizeCredentials(unescaped);
  } catch (_) {}

  for (const c of candidates) {
    try {
      return normalizeCredentials(JSON.parse(c));
    } catch (_) {}
  }

  const preview = text.slice(0, 18);
  throw new Error(`GOOGLE_SERVICE_ACCOUNT_JSON 格式錯誤。目前開頭為「${preview}...」。請貼服務帳戶 JSON 單行內容，或貼 base64 後的 JSON 內容。`);
}

function normalizeCredentials(credentials) {
  if (credentials && credentials.private_key) {
    credentials.private_key = String(credentials.private_key).replace(/\\n/g, '\n');
  }
  return credentials;
}

function authClient() {
  let credentials = null;
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    credentials = parseServiceAccountEnv(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  } else if (process.env.GOOGLE_CLIENT_EMAIL && process.env.GOOGLE_PRIVATE_KEY) {
    credentials = normalizeCredentials({
      type: 'service_account',
      project_id: process.env.GOOGLE_PROJECT_ID || undefined,
      client_email: process.env.GOOGLE_CLIENT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY
    });
  }
  if (!credentials) throw new Error('未設定 GOOGLE_SERVICE_ACCOUNT_JSON 或 GOOGLE_CLIENT_EMAIL / GOOGLE_PRIVATE_KEY。');
  if (!credentials.client_email || !credentials.private_key) throw new Error('服務帳戶 JSON 缺少 client_email 或 private_key。');
  return new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file'] });
}
async function sheetsApi() { const auth = await authClient().getClient(); return google.sheets({ version: 'v4', auth }); }
async function driveApi() { const auth = await authClient().getClient(); return google.drive({ version: 'v3', auth }); }
async function spreadsheetIdRequired() {
  const id = process.env.GOOGLE_SHEET_ID;
  if (!id) throw new Error('尚未設定 GOOGLE_SHEET_ID。請在 Zeabur 環境變數填入指定 Google Sheet ID。');
  return id;
}
async function ensureSheetTabs(spreadsheetId) {
  const sheets = await sheetsApi();
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const existing = new Set((meta.data.sheets || []).map(s => s.properties.title));
  const requests = Object.values(SHEETS).filter(t => !existing.has(t)).map(title => ({ addSheet: { properties: { title } } }));
  if (requests.length) await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
  await sheets.spreadsheets.values.update({ spreadsheetId, range: `${SHEETS.monitor}!A1:R1`, valueInputOption: 'RAW', requestBody: { values: [HEADERS.monitor] } });
  await sheets.spreadsheets.values.update({ spreadsheetId, range: `${SHEETS.complaint}!A1:P1`, valueInputOption: 'RAW', requestBody: { values: [HEADERS.complaint] } });
  await sheets.spreadsheets.values.update({ spreadsheetId, range: `${SHEETS.ranking}!A1:S1`, valueInputOption: 'RAW', requestBody: { values: [HEADERS.ranking] } });
  await sheets.spreadsheets.values.update({ spreadsheetId, range: `${SHEETS.settings}!A1:C6`, valueInputOption: 'RAW', requestBody: { values: [
    HEADERS.settings,
    ['COVERAGE_DISTANCE_M', String(coverageDistance()), '最近監測點距離大於此值視為布點缺口'],
    ['WEIGHT_CENTRAL', String(getWeights().central), '中央陳情權重'],
    ['WEIGHT_LOCAL', String(getWeights().local), '市政信箱權重'],
    ['WEIGHT_HOT', String(getWeights().hot), '超標/警察熱點權重'],
    ['WEIGHT_GAP', String(getWeights().gap), '布點缺口權重']
  ] } });
}

async function ensureSheetSize(spreadsheetId, sheetName, minRows, minCols) {
  const sheets = await sheetsApi();
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const target = (meta.data.sheets || []).find(s => s.properties && s.properties.title === sheetName);
  if (!target) return;
  const props = target.properties;
  const grid = props.gridProperties || {};
  const rowCount = Math.max(Number(grid.rowCount || 0), Number(minRows || 1));
  const columnCount = Math.max(Number(grid.columnCount || 0), Number(minCols || 1));
  if (rowCount !== grid.rowCount || columnCount !== grid.columnCount) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          updateSheetProperties: {
            properties: { sheetId: props.sheetId, gridProperties: { rowCount, columnCount } },
            fields: 'gridProperties(rowCount,columnCount)'
          }
        }]
      }
    });
  }
}

async function readSheetRows(spreadsheetId, sheetName) {
  const sheets = await sheetsApi();
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${sheetName}!A:Z` });
  return normalizeRows(res.data.values || []);
}
async function writeRanking(spreadsheetId, ranking) {
  const sheets = await sheetsApi();
  const rows = [HEADERS.ranking, ...ranking.map(r => HEADERS.ranking.map(h => h === 'updated_at' ? nowIso() : (r[h] ?? '')))];
  await ensureSheetSize(spreadsheetId, SHEETS.ranking, Math.max(rows.length + 20, 200), HEADERS.ranking.length);
  await sheets.spreadsheets.values.clear({ spreadsheetId, range: `${SHEETS.ranking}!A:S` });
  await sheets.spreadsheets.values.update({ spreadsheetId, range: `${SHEETS.ranking}!A1:S${rows.length}`, valueInputOption: 'RAW', requestBody: { values: rows } });
}

function systemPointToMonitorRow(p) {
  return HEADERS.monitor.map(h => {
    switch (h) {
      case '資料圖層': return p.layer || '';
      case '資料類型': return p.subtype || '';
      case '年份': return p.year || '';
      case '廠商': return p.vendor || '';
      case '行政區': return p.district || '';
      case '路段': return p.road || '';
      case '地址': return p.address || '';
      case '日期': return p.date || '';
      case '緯度': return p.lat ?? '';
      case '經度': return p.lng ?? '';
      case '設備機號': return p.device_no || p['設備機號'] || '';
      case '執行場次': return p.session_no || p['執行場次'] || '';
      case '監測天數': return p['監測天數'] ?? '';
      case '告發件數': return p['告發件數'] ?? '';
      case '通知到檢件數': return p['通知到檢件數'] ?? '';
      case '超標車輛': return p['超標車輛'] ?? '';
      case '夜間案件': return p['夜間案件'] ?? '';
      case '備註': return p.source || p.geo_source || '';
      default: return '';
    }
  });
}
function systemPointToComplaintRow(p) {
  return HEADERS.complaint.map(h => {
    switch (h) {
      case '資料來源': return p.layer || p.subtype || '';
      case '年份': return p.year || '';
      case '日期': return p.date || '';
      case '行政區': return p.district || '';
      case '路段': return p.road || '';
      case '地址': return p.address || '';
      case '緯度': return p.lat ?? '';
      case '經度': return p.lng ?? '';
      case '案件量': return p['案件量'] ?? '';
      case '中央檢舉量': return p['中央檢舉量'] ?? '';
      case '市政信箱量': return p['市政信箱量'] ?? '';
      case '警察案件量': return p['警察案件量'] ?? (String(p.layer||'').includes('警察') ? (p['案件量'] ?? p['超標車輛'] ?? '') : '');
      case '超標車輛': return p['超標車輛'] ?? '';
      case '夜間案件': return p['夜間案件'] ?? '';
      case '夜間比例': return p['夜間比例'] ?? '';
      case '備註': return [p.subtype, p.source, p.geo_source].filter(Boolean).join('｜');
      default: return '';
    }
  });
}
async function clearAndUpdate(spreadsheetId, sheetName, headers, rows) {
  const sheets = await sheetsApi();
  const colEnd = String.fromCharCode(64 + headers.length);
  const all = [headers, ...rows];
  await ensureSheetSize(spreadsheetId, sheetName, Math.max(all.length + 200, 2000), headers.length);
  await sheets.spreadsheets.values.clear({ spreadsheetId, range: `${sheetName}!A:${colEnd}` });
  const chunkSize = 500;
  for (let start = 0; start < all.length; start += chunkSize) {
    const chunk = all.slice(start, start + chunkSize);
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A${start + 1}:${colEnd}${start + chunk.length}`,
      valueInputOption: 'RAW',
      requestBody: { values: chunk }
    });
  }
}
async function writeSystemDataToSheet(spreadsheetId) {
  await ensureSheetTabs(spreadsheetId);
  const embedded = loadEmbeddedData();
  const points = embedded.points || [];
  const monitorPoints = points.filter(p => String(p.layer || '').includes('聲音照相監測點位'));
  const complaintPoints = points.filter(p => !String(p.layer || '').includes('聲音照相監測點位'));
  const ranking = recalcTop100(points);
  await clearAndUpdate(spreadsheetId, SHEETS.monitor, HEADERS.monitor, monitorPoints.map(systemPointToMonitorRow));
  await clearAndUpdate(spreadsheetId, SHEETS.complaint, HEADERS.complaint, complaintPoints.map(systemPointToComplaintRow));
  await writeRanking(spreadsheetId, ranking);
  const sheets = await sheetsApi();
  await ensureSheetSize(spreadsheetId, SHEETS.settings, 200, 8);
  await sheets.spreadsheets.values.update({ spreadsheetId, range: `${SHEETS.settings}!A1:C12`, valueInputOption: 'RAW', requestBody: { values: [
    HEADERS.settings,
    ['COVERAGE_DISTANCE_M', String(coverageDistance()), '最近監測點距離大於此值視為布點缺口'],
    ['WEIGHT_CENTRAL', String(getWeights().central), '中央陳情權重'],
    ['WEIGHT_LOCAL', String(getWeights().local), '市政信箱權重'],
    ['WEIGHT_HOT', String(getWeights().hot), '超標/警察熱點權重'],
    ['WEIGHT_NIGHT', String(getWeights().night), '夜間案件權重'],
    ['WEIGHT_GAP', String(getWeights().gap), '布點缺口權重'],
    ['SYSTEM_POINTS', String(points.length), '平台內建資料總筆數'],
    ['MONITOR_POINTS', String(monitorPoints.length), '監測點位筆數'],
    ['COMPLAINT_POINTS', String(complaintPoints.length), '陳情/警察熱點筆數'],
    ['TOP100_POINTS', String(ranking.length), '百大排行輸出筆數'],
    ['UPDATED_AT', nowIso(), '最近寫入時間']
  ] } });
  cache = { data: { summary: { ...(embedded.summary || {}), sheet_points: 0, exported_to_sheet_at: nowIso() }, points, candidates: ranking }, source: 'embedded_published_to_google_sheet', updatedAt: nowIso(), spreadsheetId };
  return { monitor: monitorPoints.length, complaint: complaintPoints.length, ranking: ranking.length, total: points.length };
}

function rankingWorkbookBuffer(ranking) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ranking), '百大排行輸出');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}


app.post('/api/admin/login', (req, res) => {
  const password = String(req.body?.password || '');
  const passBuf = Buffer.from(password);
  const expectedBuf = Buffer.from(ADMIN_PASSWORD);
  const ok = passBuf.length === expectedBuf.length && crypto.timingSafeEqual(passBuf, expectedBuf);
  if (!ok) return res.status(401).json({ error: '管理密碼錯誤。' });
  const now = Date.now();
  const token = signPayload({ scope: 'admin', iat: now, exp: now + ADMIN_TOKEN_TTL_MS });
  res.json({ ok: true, token, expiresAt: new Date(now + ADMIN_TOKEN_TTL_MS).toISOString() });
});
app.get('/api/admin/status', (req, res) => res.json({ authenticated: hasAdminAccess(req), version: VERSION }));

app.get('/api/health', (req, res) => res.json({ status: 'ok', version: VERSION, time: nowIso(), port: process.env.PORT || 8080 }));
app.get('/healthz', (req, res) => res.status(200).send('ok'));
app.get('/ready', (req, res) => res.status(200).json({ ready: true, version: VERSION }));
app.get('/api/data', (req, res) => res.json(cache));

app.get('/api/google/test', requireAdmin, async (req, res) => {
  try {
    const sheets = await sheetsApi();
    let spreadsheetId = process.env.GOOGLE_SHEET_ID || '';
    let canRead = false, canWrite = false;
    if (spreadsheetId) {
      await sheets.spreadsheets.get({ spreadsheetId }); canRead = true;
      await ensureSheetTabs(spreadsheetId);
      await ensureSheetSize(spreadsheetId, SHEETS.settings, 200, 8);
      await sheets.spreadsheets.values.update({ spreadsheetId, range: `${SHEETS.settings}!E1:E1`, valueInputOption: 'RAW', requestBody: { values: [[`test ${nowIso()}`]] } });
      canWrite = true;
    }
    res.json({ ok: true, spreadsheetId, canRead, canWrite });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/api/google/publish', requireAdmin, async (req, res) => {
  try {
    const spreadsheetId = await spreadsheetIdRequired();
    const counts = await writeSystemDataToSheet(spreadsheetId);
    res.json({ ok: true, spreadsheetId, counts, updatedAt: cache.updatedAt });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/api/sync', requireAdmin, async (req, res) => {
  try {
    const spreadsheetId = await spreadsheetIdRequired();
    await ensureSheetTabs(spreadsheetId);
    const [monitorRows, complaintRows] = await Promise.all([readSheetRows(spreadsheetId, SHEETS.monitor), readSheetRows(spreadsheetId, SHEETS.complaint)]);
    const base = loadEmbeddedData();
    const sheetPoints = [
      ...monitorRows.map(monitorRowToPoint).filter(Boolean),
      ...complaintRows.map(complaintRowToPoint).filter(Boolean)
    ];
    const points = [...(base.points || []), ...sheetPoints];
    const candidates = recalcTop100(points);
    await writeRanking(spreadsheetId, candidates);
    cache = { data: { summary: { ...(base.summary || {}), sheet_points: sheetPoints.length, updated_at: nowIso() }, points, candidates }, source: 'google_sheet', updatedAt: nowIso(), spreadsheetId };
    res.json(cache);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/export/ranking.xlsx', requireAdmin, (req, res) => {
  const buffer = rankingWorkbookBuffer(cache.data.candidates || []);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', "attachment; filename*=UTF-8''%E6%96%B0%E5%8C%97%E5%B8%82%E8%81%B2%E9%9F%B3%E7%85%A7%E7%9B%B8%E7%99%BE%E5%A4%A7%E6%8E%92%E8%A1%8C_%E7%B3%BB%E7%B5%B1%E5%8C%AF%E5%87%BA.xlsx");
  res.send(buffer);
});

app.get('*', (req, res) => res.sendFile(path.join(PUBLIC_DIR, 'index.html')));
const port = Number(process.env.PORT || 8080);
const host = process.env.HOST || '0.0.0.0';
app.listen(port, host, () => console.log(`NTPC noise platform running on ${host}:${port}`));
