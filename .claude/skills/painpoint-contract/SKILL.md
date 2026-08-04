---
name: painpoint-contract
description: >-
  產生「痛點科技」的服務合約書（繁體中文 .docx）。當使用者想要草擬、製作、產生痛點科技
  給客戶的合約、服務合約、委託合約，或要幫某個客戶專案（例如 Celine、隆志工業、宸大等）
  出一份可簽署的合約時，就使用這個 skill。即使使用者只說「幫我出一份合約」「這個客戶的合約」
  「做一份服務合約範本」而沒有明講痛點科技，只要情境是痛點科技對外的服務委託，也應觸發。
  可產生空白範本，或依客戶資料填好的正式版本。
---

# 痛點科技服務合約產生器

這個 skill 產出痛點科技對客戶的**服務合約書**(繁中 Word `.docx`),條文依痛點科技
的業務性質設計(OCR、流程自動化、知識庫、AI 客服等專案制服務),並內含個資保護條款。

輸出內容包含:主文 14 條 + 附件一(服務項目與報價明細)+ 附件二(付款排程)+ 甲乙雙方簽章欄。

## 使用方式

產生器是 `scripts/build_contract.js`。它相依 npm 的 `docx` 套件——如果 `require('docx')`
失敗,先在 skill 目錄執行 `npm install docx --no-save` 再跑。

**兩種模式:**

1. **空白範本**(不帶客戶資料)——所有要填的欄位以底線 `______` 呈現,供事後手填:
   ```bash
   node scripts/build_contract.js --out 痛點科技_服務合約書_範本.docx
   ```

2. **填好的客戶版本**——傳入一個 JSON 設定檔,腳本會把有值的欄位填入,沒給的欄位仍留底線:
   ```bash
   node scripts/build_contract.js --config client.json --out 痛點科技_Celine_服務合約書.docx
   ```

產生後把 `.docx` 交付給使用者(在 Cowork/remote 用 `SendUserFile`)。

## 設定檔格式

只需要填想帶入的欄位,其餘省略即可(省略的會自動留底線佔位)。範例:

```json
{
  "signDate": { "y": "115", "m": "8", "d": "4" },
  "client": {
    "name": "Celine 身心靈工作室",
    "taxId": "12345678",
    "rep": "王小明",
    "address": "台北市…"
  },
  "vendor": { "taxId": "87654321", "rep": "Akash", "address": "…" },
  "totalAmount": "300,000",
  "items": [
    { "name": "知識庫 / AI 客服建置", "qty": "1", "unit": "180,000", "subtotal": "180,000" },
    { "name": "教育訓練與導入輔導", "qty": "1", "unit": "20,000", "subtotal": "20,000" }
  ],
  "payments": [
    { "stage": "第一期（訂金）", "milestone": "合約簽訂後 7 日內", "ratio": "30%" },
    { "stage": "第二期（期中）", "milestone": "完成知識庫建置里程碑", "ratio": "40%" },
    { "stage": "第三期（尾款）", "milestone": "驗收合格後 7 日內", "ratio": "30%" }
  ],
  "terms": {
    "acceptanceDays": "7",
    "warrantyMonths": "3",
    "confidentialityYears": "3",
    "jurisdiction": "台灣台北地方法院"
  },
  "bank": { "name": "○○銀行 ○○分行", "account": "000-000-000000", "holder": "痛點科技" }
}
```

欄位對照:
- `client` / `vendor`:甲方(客戶)、乙方(痛點科技)基本資料。乙方公司名固定為「痛點科技」。
- `items`:附件一報價明細的每一列(合計/稅/總計由腳本自動不填,留底線讓人確認金額)。
- `payments`:附件二的付款期別。
- `terms`:散落在條文中的天數/月數/年數/管轄法院。
- `bank`:附件二的匯款資訊。

## 修改條文

若使用者要調整實際條文用詞(而非填空),直接編輯 `scripts/build_contract.js` 裡對應的
`clause(...)` / `item(...)` 段落。每一條都以 `clause(編號, 標題)` 起頭,底下用 `item(...)`
列出各款,結構清楚、容易對照修改。

## 驗證輸出

此環境若有 LibreOffice 可轉 PDF 預覽;若 `soffice` 因缺 Java 無法轉檔,可改以下法確認
文件完整(檢查 XML 良構與關鍵條文是否都在):

```bash
unzip -p out.docx word/document.xml > /tmp/d.xml
python3 -c "import xml.dom.minidom as m; m.parse('/tmp/d.xml'); print('ok')"
```
