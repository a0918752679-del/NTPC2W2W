# 新北市噪音車管理系統 V27

本版重點：LINE Bot 回覆改為分別讀取三套平台 Google Sheet 資料，並新增多組關鍵字搜尋。

## 三套 Google Sheet 來源

- 百大點位：`TOP100_GOOGLE_SHEET_ID`
- 監測成果：`RESULTS_GOOGLE_SHEET_ID`
- 外勤回報架設點位：`FIELD_GOOGLE_SHEET_ID`

## LINE Bot 可用指令

### 百大點位

- `百大點位`
- `百大熱點`
- `淡水區百大`
- `第1名詳細`
- `排名10`
- `百大計算方式`

### 監測成果

- `進度`
- `KPI報表`
- `2月份執行成效`
- `淡水區執行成效`
- `月份選單`
- `行政區選單`

### 外勤回報架設點位

- `架設點位`
- `外勤回報`
- `S01`
- `OE_ZB001`
- `淡水區外勤`

### 通用關鍵字

- `搜尋 淡水區`
- `查詢 OE_ZB001`
- `查詢 S01`

## Zeabur 環境變數

```env
ADMIN_PASSWORD=請設定新的後台密碼
SESSION_SECRET=請設定64字元以上隨機字串
LINE_CHANNEL_ACCESS_TOKEN=請填入LINE長期Access Token
LINE_CHANNEL_SECRET=請填入LINE Channel Secret
PUBLIC_BASE_URL=https://newtaipeinoise.zeabur.app
DASHBOARD_URL=https://noise115.zeabur.app
FIELD_REPORT_URL=https://out115.zeabur.app
HOTSPOT_URL=https://ntpcnoisely.zeabur.app/login
TOP100_GOOGLE_SHEET_ID=1WDKyCQJXIti67Wz4wH9MU8UB4Afht918tO25x5pY4jc
RESULTS_GOOGLE_SHEET_ID=1EfP7GoI87RRl1AUGegwPhqNvFN9xG-YXm9NoMSqm_O0
FIELD_GOOGLE_SHEET_ID=1BVZ4kEoKndO5OMAZmk8OLwplrzBL_Drt4xEmpzgejb8
GSHEET_SYNC_INTERVAL_MIN=60
```

## 部署後檢查

```text
https://newtaipeinoise.zeabur.app/healthz
```

應顯示：

```json
{"ok":true,"service":"newtaipei-noise-control-system-v27-three-gsheet-line-search"}
```

## 後台同步

進入：

```text
https://newtaipeinoise.zeabur.app/admin.html
```

登入後按：

```text
從 Google Sheet 立即同步
```

再至 LINE 測試：

```text
百大點位
進度
架設點位
淡水區百大
OE_ZB001
```
