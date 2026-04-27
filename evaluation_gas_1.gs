// ============================================================
// 評価スプレッドシート自動生成スクリプト
// 使い方: このコードをApps Scriptに貼り付けて「createEvaluationSheet」を実行
// ============================================================

function createEvaluationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存シートを全削除（確認なし）
  const existingSheets = ss.getSheets();
  existingSheets.forEach(s => {
    if (s.getName() !== "_temp_") ss.deleteSheet(s);
  });

  // ページ名一覧
  const pageNames = ["ページ1","ページ2","ページ3","ページ4","ページ5","ページ6","ページ7","ページ8"];
  const PAGE_COLORS = ["#3A7CA5","#2E86AB","#1D6FA4","#2563A8","#1B6CA8","#2076B4","#165C8A","#0F4C75"];

  const HEADER_BG   = "#2E4057";
  const HEADER_FG   = "#FFFFFF";
  const ROW_ODD     = "#F7F9FC";
  const ROW_EVEN    = "#FFFFFF";
  const RATING_BG   = "#EAF4FB";
  const SUMMARY_BG  = "#1A3A5C";

  const pageSheets = [];

  // ── ページ1〜8 作成 ──────────────────────────────────────
  pageNames.forEach((pageName, pi) => {
    const ws = ss.insertSheet(pageName);
    pageSheets.push(ws);

    // ページタイトル行
    ws.getRange("A1:E1").merge()
      .setValue(pageName)
      .setBackground(PAGE_COLORS[pi])
      .setFontColor(HEADER_FG)
      .setFontSize(14)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    ws.setRowHeight(1, 36);

    // 列ヘッダー行
    const headers = ["項目", "評価①", "評価②", "評価③", "コメント"];
    headers.forEach((h, ci) => {
      ws.getRange(2, ci + 1)
        .setValue(h)
        .setBackground(HEADER_BG)
        .setFontColor(HEADER_FG)
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setBorder(true,true,true,true,false,false,"#888888", SpreadsheetApp.BorderStyle.SOLID);
    });
    ws.setRowHeight(2, 24);

    // データ行（9行）
    for (let ri = 0; ri < 9; ri++) {
      const row = ri + 3;
      const fill = ri % 2 === 0 ? ROW_ODD : ROW_EVEN;

      // 項目セル
      ws.getRange(row, 1)
        .setValue(`項目${ri + 1}`)
        .setBackground(fill)
        .setFontWeight("bold")
        .setHorizontalAlignment("left")
        .setVerticalAlignment("middle")
        .setBorder(true,true,true,true,false,false,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

      // 評価セル①②③
      for (let ci = 1; ci <= 3; ci++) {
        const cell = ws.getRange(row, ci + 1);
        cell.setBackground(RATING_BG)
            .setFontSize(11)
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setBorder(true,true,true,true,false,false,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

        // ドロップダウン
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(["S","A","B","C","D"], true)
          .setAllowInvalid(false)
          .setHelpText("S, A, B, C, D から選択してください")
          .build();
        cell.setDataValidation(rule);
      }

      // コメントセル
      ws.getRange(row, 5)
        .setBackground(fill)
        .setHorizontalAlignment("left")
        .setVerticalAlignment("middle")
        .setWrap(true)
        .setBorder(true,true,true,true,false,false,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

      ws.setRowHeight(row, 26);
    }

    // 列幅
    ws.setColumnWidth(1, 140);
    ws.setColumnWidth(2, 80);
    ws.setColumnWidth(3, 80);
    ws.setColumnWidth(4, 80);
    ws.setColumnWidth(5, 260);

    // 枠線・グリッドを非表示
    // (任意) ws.setHiddenGridlines(true);
  });

  // ── まとめシート 作成 ────────────────────────────────────
  const ws = ss.insertSheet("まとめ");

  // タイトル (A1 〜 最終列)
  // 列数: 1(項目) + 8ページ×3列 = 25列
  const TOTAL_COLS = 1 + pageNames.length * 3;
  ws.getRange(1, 1, 1, TOTAL_COLS).merge()
    .setValue("評価まとめ")
    .setBackground(SUMMARY_BG)
    .setFontColor(HEADER_FG)
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  ws.setRowHeight(1, 40);

  // 項目ヘッダー (行2)
  ws.getRange(2, 1)
    .setValue("項目")
    .setBackground(HEADER_BG)
    .setFontColor(HEADER_FG)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // ページ名ヘッダー (行2) & 評価①②③サブヘッダー (行3)
  pageNames.forEach((pageName, pi) => {
    const startCol = 2 + pi * 3;
    ws.getRange(2, startCol, 1, 3).merge()
      .setValue(pageName)
      .setBackground(PAGE_COLORS[pi])
      .setFontColor(HEADER_FG)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true,true,true,true,false,false,"#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

    ["評価①","評価②","評価③"].forEach((h, si) => {
      ws.getRange(3, startCol + si)
        .setValue(h)
        .setBackground(HEADER_BG)
        .setFontColor(HEADER_FG)
        .setFontSize(9)
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setBorder(true,true,true,true,false,false,"#888888", SpreadsheetApp.BorderStyle.SOLID);
    });
  });

  ws.getRange(3, 1)
    .setBackground(HEADER_BG)
    .setBorder(true,true,true,true,false,false,"#888888", SpreadsheetApp.BorderStyle.SOLID);

  ws.setRowHeight(2, 26);
  ws.setRowHeight(3, 22);

  // データ行（9行）
  for (let ri = 0; ri < 9; ri++) {
    const row = ri + 4;
    const fill = ri % 2 === 0 ? ROW_ODD : ROW_EVEN;

    ws.getRange(row, 1)
      .setValue(`項目${ri + 1}`)
      .setBackground(fill)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true,true,true,true,false,false,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

    pageNames.forEach((pageName, pi) => {
      const startCol = 2 + pi * 3;
      const srcRow = ri + 3; // ページシートのデータ開始行

      ["B","C","D"].forEach((col, si) => {
        const formula = `='${pageName}'!${col}${srcRow}`;
        ws.getRange(row, startCol + si)
          .setFormula(formula)
          .setBackground(RATING_BG)
          .setFontSize(11)
          .setFontWeight("bold")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle")
          .setBorder(true,true,true,true,false,false,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
      });
    });

    ws.setRowHeight(row, 26);
  }

  // まとめシートの列幅
  ws.setColumnWidth(1, 80);
  for (let pi = 0; pi < pageNames.length; pi++) {
    for (let si = 0; si < 3; si++) {
      ws.setColumnWidth(2 + pi * 3 + si, 70);
    }
  }

  // まとめシートを先頭へ移動
  ss.moveActiveSheet(0);
  ss.setActiveSheet(ss.getSheetByName("まとめ"));
  ss.moveActiveSheet(ss.getNumSheets()); // 末尾 → 最後

  SpreadsheetApp.getUi().alert("✅ 作成完了！\nページ1〜8 と まとめシート が生成されました。");
}
