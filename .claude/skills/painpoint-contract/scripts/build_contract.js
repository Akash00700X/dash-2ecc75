#!/usr/bin/env node
/*
 * 痛點科技 服務合約書產生器
 *
 *   node build_contract.js [--config client.json] [--out output.docx]
 *
 * 不帶 --config 時產生空白範本（欄位以底線佔位）。
 * 帶 --config 時，把設定檔中有值的欄位填入，未給的欄位仍留底線。
 */
const fs = require("fs");
const path = require("path");

let docx;
try {
  docx = require("docx");
} catch (e) {
  console.error(
    "找不到 docx 套件。請先在 skill 目錄執行：npm install docx --no-save"
  );
  process.exit(1);
}
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  PageBreak, Footer, PageNumber, LevelFormat,
} = docx;

// ---------- args ----------
function arg(flag, def) {
  const i = process.argv.indexOf(flag);
  return i !== -1 && process.argv[i + 1] ? process.argv[i + 1] : def;
}
const configPath = arg("--config", null);
const outPath = arg("--out", "痛點科技_服務合約書_範本.docx");
const cfg = configPath
  ? JSON.parse(fs.readFileSync(configPath, "utf-8"))
  : {};

// value-or-placeholder: 有值回傳值，否則回傳底線佔位
const BLANK = "______";
function v(val, blank = BLANK) {
  return val !== undefined && val !== null && String(val).trim() !== ""
    ? String(val)
    : blank;
}
const client = cfg.client || {};
const vendor = cfg.vendor || {};
const terms = cfg.terms || {};
const sign = cfg.signDate || {};
const bank = cfg.bank || {};

// ---------- style constants ----------
const NAVY = "1F3864";
const GREY = "595959";
const LIGHT = "EDF0F5";

// ---------- text helpers ----------
function h1(text) {
  return new Paragraph({
    spacing: { before: 260, after: 140 },
    children: [new TextRun({ text, bold: true, size: 26, color: NAVY })],
  });
}
function clause(num, title) {
  return new Paragraph({
    spacing: { before: 220, after: 90 },
    children: [new TextRun({ text: `第 ${num} 條　${title}`, bold: true, size: 24, color: NAVY })],
  });
}
function body(text, indent = false) {
  return new Paragraph({
    spacing: { after: 90, line: 300 },
    indent: indent ? { left: 360 } : undefined,
    children: [new TextRun({ text, size: 22, color: "222222" })],
  });
}
function item(text) {
  return new Paragraph({
    numbering: { reference: "clause-list", level: 0 },
    spacing: { after: 70, line: 300 },
    children: [new TextRun({ text, size: 22, color: "222222" })],
  });
}

// ---------- table helpers ----------
function cell(text, { w, bold = false, shade = false, color = "222222", align } = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: shade ? { type: ShadingType.CLEAR, fill: LIGHT, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ alignment: align, children: [new TextRun({ text, bold, size: 20, color })] })],
  });
}

// 報價明細
const quoteCols = [3400, 1200, 1600, 1600];
function quoteTable() {
  const rows = [new TableRow({
    tableHeader: true,
    children: [
      cell("服務項目", { w: quoteCols[0], bold: true, shade: true, color: NAVY }),
      cell("數量", { w: quoteCols[1], bold: true, shade: true, color: NAVY, align: AlignmentType.CENTER }),
      cell("單價（NT$）", { w: quoteCols[2], bold: true, shade: true, color: NAVY, align: AlignmentType.RIGHT }),
      cell("小計（NT$）", { w: quoteCols[3], bold: true, shade: true, color: NAVY, align: AlignmentType.RIGHT }),
    ],
  })];
  const items = (cfg.items && cfg.items.length)
    ? cfg.items
    : [
        { name: "【例】OCR 發票辨識模組建置", qty: "1" },
        { name: "【例】知識庫 / AI 客服建置", qty: "1" },
        { name: "【例】流程自動化（RPA）開發", qty: "1" },
        { name: "【例】教育訓練與導入輔導", qty: "" },
        { name: "", qty: "" },
      ];
  for (const it of items) {
    rows.push(new TableRow({ children: [
      cell(v(it.name, ""), { w: quoteCols[0] }),
      cell(v(it.qty, ""), { w: quoteCols[1], align: AlignmentType.CENTER }),
      cell(v(it.unit, ""), { w: quoteCols[2], align: AlignmentType.RIGHT }),
      cell(v(it.subtotal, ""), { w: quoteCols[3], align: AlignmentType.RIGHT }),
    ]}));
  }
  // 合計 / 稅 / 總計：金額留底線讓人再確認
  for (const label of ["合計（未稅）", "營業稅 5%", "總計（含稅）"]) {
    rows.push(new TableRow({ children: [
      cell(label, { w: quoteCols[0] + quoteCols[1] + quoteCols[2], bold: true, shade: true, align: AlignmentType.RIGHT }),
      cell("", { w: quoteCols[3], align: AlignmentType.RIGHT }),
    ]}));
  }
  return new Table({ columnWidths: quoteCols, width: { size: quoteCols.reduce((a, b) => a + b, 0), type: WidthType.DXA }, rows });
}

// 付款排程
const payCols = [2000, 3000, 2800];
function paymentTable() {
  const rows = [new TableRow({ tableHeader: true, children: [
    cell("期別", { w: payCols[0], bold: true, shade: true, color: NAVY }),
    cell("付款條件 / 里程碑", { w: payCols[1], bold: true, shade: true, color: NAVY }),
    cell("比例 / 金額", { w: payCols[2], bold: true, shade: true, color: NAVY, align: AlignmentType.RIGHT }),
  ]})];
  const pays = (cfg.payments && cfg.payments.length)
    ? cfg.payments
    : [
        { stage: "第一期（訂金）", milestone: "合約簽訂後 ___ 日內", ratio: "___ %" },
        { stage: "第二期（期中）", milestone: "完成 ______________ 里程碑", ratio: "___ %" },
        { stage: "第三期（尾款）", milestone: "驗收合格後 ___ 日內", ratio: "___ %" },
      ];
  for (const p of pays) {
    rows.push(new TableRow({ children: [
      cell(v(p.stage, ""), { w: payCols[0] }),
      cell(v(p.milestone, ""), { w: payCols[1] }),
      cell(v(p.ratio, ""), { w: payCols[2], align: AlignmentType.RIGHT }),
    ]}));
  }
  return new Table({ columnWidths: payCols, width: { size: payCols.reduce((a, b) => a + b, 0), type: WidthType.DXA }, rows });
}

// 簽章區
const sigCols = [4400, 4400];
function signatureTable() {
  const line = (t) => new Paragraph({ spacing: { before: 200, after: 40 }, children: [new TextRun({ text: t, size: 22, color: "222222" })] });
  const mk = (title, rows) => new TableCell({
    width: { size: sigCols[0], type: WidthType.DXA },
    margins: { top: 120, bottom: 120, left: 160, right: 160 },
    children: [new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: title, bold: true, size: 24, color: NAVY })] }), ...rows.map(line)],
  });
  return new Table({
    columnWidths: sigCols,
    width: { size: sigCols[0] + sigCols[1], type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: "BBBBBB" },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: "BBBBBB" },
      left: { style: BorderStyle.SINGLE, size: 6, color: "BBBBBB" },
      right: { style: BorderStyle.SINGLE, size: 6, color: "BBBBBB" },
      insideVertical: { style: BorderStyle.SINGLE, size: 6, color: "BBBBBB" },
    },
    rows: [new TableRow({ children: [
      mk("甲方（委託方）", [
        `公司名稱：${v(client.name, "________________________")}`,
        `統一編號：${v(client.taxId, "________________________")}`,
        `代表人：${v(client.rep, "__________________________")}`,
        `地　址：${v(client.address, "__________________________")}`,
        "簽章：", "", "日期：______年______月______日",
      ]),
      mk("乙方（受託方）", [
        "公司名稱：痛點科技",
        `統一編號：${v(vendor.taxId, "________________________")}`,
        `代表人：${v(vendor.rep, "__________________________")}`,
        `地　址：${v(vendor.address, "__________________________")}`,
        "簽章：", "", "日期：______年______月______日",
      ]),
    ]})],
  });
}

// ---------- assemble document ----------
const doc = new Document({
  numbering: { config: [{
    reference: "clause-list",
    levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.START, style: { paragraph: { indent: { left: 620, hanging: 260 } } } }],
  }]},
  styles: { default: { document: { run: { font: "Microsoft JhengHei", size: 22 } } } },
  sections: [{
    properties: { page: { margin: { top: 1200, bottom: 1200, left: 1200, right: 1200 } } },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: "痛點科技　服務合約書　　第 ", size: 16, color: GREY }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GREY }),
        new TextRun({ text: " 頁 / 共 ", size: 16, color: GREY }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: GREY }),
        new TextRun({ text: " 頁", size: 16, color: GREY }),
      ],
    })] }) },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "痛點科技", bold: true, size: 30, color: NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "服 務 合 約 書", bold: true, size: 40, color: NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 220 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 6 } }, children: [new TextRun({ text: "SERVICE AGREEMENT", size: 18, color: GREY })] }),

      body(`本服務合約書（以下簡稱「本合約」）由下列雙方於中華民國 ${v(sign.y)} 年 ${v(sign.m)} 月 ${v(sign.d)} 日簽訂，雙方同意遵守下列各項約定：`),

      new Paragraph({ spacing: { before: 120, after: 90 }, children: [
        new TextRun({ text: "甲方（委託方）：", bold: true, size: 22, color: "222222" }),
        new TextRun({ text: `${v(client.name, "________________________")}（以下簡稱「甲方」）`, size: 22, color: "222222" }),
      ]}),
      new Paragraph({ spacing: { after: 90 }, children: [
        new TextRun({ text: "乙方（受託方）：", bold: true, size: 22, color: "222222" }),
        new TextRun({ text: "痛點科技（以下簡稱「乙方」）", size: 22, color: "222222" }),
      ]}),

      clause(1, "契約標的與服務範圍"),
      item("乙方同意依本合約及附件一「服務項目與報價明細」之約定，為甲方提供人工智慧、流程自動化、光學字元辨識（OCR）、知識庫建置、AI 客服及相關資訊系統建置與導入之服務（以下簡稱「本服務」）。"),
      item("本服務之具體項目、規格、交付標的與數量，以附件一及雙方另行確認之需求規格書為準。"),
      item("凡未載明於附件一或需求規格書之工作項目，均不屬本合約服務範圍，如需增加應依第 10 條辦理。"),

      clause(2, "合約期間與服務時程"),
      item("本合約有效期間自簽約日起至全部服務交付驗收合格之日止；如屬持續性維護服務者，另依第 9 條約定辦理。"),
      item("本服務預計自______年______月______日起至______年______月______日止完成，分階段之里程碑與時程以附件一為準。"),
      item("因甲方未及時提供資料、決策或配合事項致時程延誤者，交付期限順延，乙方不負遲延責任。"),

      clause(3, "合約金額與付款方式"),
      item(`本合約總金額為新臺幣（NT$）${v(cfg.totalAmount, "__________________")} 元整（含 5% 營業稅），詳如附件一。`),
      item("付款方式與期別如附件二「付款排程」所示。"),
      item("甲方應於乙方開立發票並提出請款後 ______ 日內，以匯款方式支付至乙方指定帳戶。逾期未付者，每逾一日按應付金額萬分之______計付遲延利息。"),

      clause(4, "交付與驗收"),
      item(`乙方應於各階段完成後以書面（含電子郵件）通知甲方交付，甲方應於收到通知後 ${v(terms.acceptanceDays)} 個工作日內完成驗收。`),
      item("甲方於前述期限內未提出書面異議者，視為驗收合格。"),
      item("驗收如有不符約定之處，甲方應具體列明缺失，乙方應於合理期間內修正後再行交付驗收。"),

      clause(5, "甲方之義務與協力"),
      item("甲方應指派專責窗口，並及時提供本服務所需之資料、樣本、系統權限、測試環境及必要之決策。"),
      item("甲方對其所提供資料之合法性、正確性及來源正當性負責，並確保未侵害第三人之權利。"),

      clause(6, "智慧財產權"),
      item("乙方為履行本服務所自行開發之通用工具、框架、模型、程式模組及既有技術（Know-how），其智慧財產權仍歸乙方所有。"),
      item("就乙方專為甲方客製化開發並交付之成果，於甲方付清全部款項後，其使用權歸屬甲方；除另有書面約定外，乙方保留前款既有元件之權利。"),
      item("任一方不得移除或變更他方於交付物上之著作權標示或商標。"),

      clause(7, "保密義務"),
      item("雙方對於因履行本合約而知悉他方之營業秘密、技術資料、客戶資訊及其他經標示為機密之資訊（以下簡稱「保密資訊」），應盡善良管理人之注意義務予以保密，非經他方書面同意不得洩漏或供作本合約以外之用途。"),
      item(`本條義務於本合約終止或屆滿後 ${v(terms.confidentialityYears)} 年內仍繼續有效。`),

      clause(8, "個人資料保護"),
      item("雙方應遵守《個人資料保護法》及相關法令。乙方於本服務範圍內所處理之個人資料，僅得於甲方指示及本服務目的範圍內蒐集、處理及利用。"),
      item("本服務結束後，乙方應依甲方指示刪除、銷毀或返還所持有之個人資料，但法令另有規定者不在此限。"),

      clause(9, "保固與維護"),
      item(`乙方就交付之客製化成果提供自驗收合格日起 ${v(terms.warrantyMonths)} 個月之免費保固，保固範圍以修正非因甲方操作不當或第三方系統變更所致之瑕疵為限。`),
      item("保固期滿後之維護、更新及技術支援，雙方得另訂維護合約辦理。"),

      clause(10, "變更與追加"),
      item("任一方就服務範圍、規格或時程提出變更者，應以書面提出，經雙方確認變更內容、費用及時程調整並簽署變更單後，始生效力。"),

      clause(11, "契約終止與違約"),
      item("任一方違反本合約且經他方以書面通知後 ______ 日內仍未改善者，他方得終止本合約。"),
      item("本合約終止時，甲方應就乙方已完成部分之服務按比例給付報酬；乙方應將已完成之成果交付甲方。"),
      item("因可歸責於一方之事由致他方受損害者，該方應負損害賠償責任。"),

      clause(12, "責任限制"),
      item("除因故意或重大過失外，乙方就本合約所生之賠償責任總額，以甲方依本合約已實際支付予乙方之金額為上限。"),
      item("任一方均不對他方之間接、附隨、衍生性損害（包括營業損失、利潤損失）負責。"),

      clause(13, "不可抗力"),
      item("因天災、戰爭、疫情、政府命令、網路或雲端服務中斷等不可抗力事由致無法履約者，受影響之一方不負遲延或不履行之責任，但應即時通知他方並採取合理減損措施。"),

      clause(14, "其他約定"),
      item("本合約之通知均應以書面或電子郵件為之，並以雙方於本合約所載或另行書面指定之聯絡方式為送達。"),
      item("本合約之修改或補充，非經雙方書面同意不生效力。"),
      item("本合約如有部分條款無效，不影響其他條款之效力。"),
      item(`本合約以中華民國法律為準據法；因本合約所生之爭議，雙方同意以 ${v(terms.jurisdiction, "__________")} 為第一審管轄法院。`),
      item("本合約一式二份，甲乙雙方各執一份為憑。附件一、附件二為本合約之一部分，與本文具同等效力。"),

      new Paragraph({ spacing: { before: 320, after: 200 }, children: [new TextRun({ text: "—— 立合約書人 ——", bold: true, size: 22, color: NAVY })] }),
      signatureTable(),

      new Paragraph({ children: [new PageBreak()] }),
      h1("附件一　服務項目與報價明細"),
      body("（本表列示本合約之服務項目、數量與金額；標示【例】者為範例，實際簽約時請刪除或替換。）"),
      quoteTable(),
      new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "備註：", bold: true, size: 20, color: NAVY })] }),
      body("1. 報價有效期限至______年______月______日止。", true),
      body("2. 上列金額如未特別載明，均為未稅價，另加 5% 營業稅。", true),
      body("3. 交付方式、環境需求與時程里程碑詳如需求規格書。", true),

      new Paragraph({ children: [new PageBreak()] }),
      h1("附件二　付款排程"),
      paymentTable(),
      new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "匯款資訊：", bold: true, size: 20, color: NAVY })] }),
      body(`戶名：${v(bank.holder, "______________________")}　銀行：${v(bank.name, "______________________")}`, true),
      body(`帳號：${v(bank.account, "______________________")}`, true),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log(`已產生：${path.resolve(outPath)}（${buf.length} bytes）`);
});
