// ============================================================
// 評価スプレッドシート自動生成スクリプト（高速版）
// 使い方: このコードをApps Scriptに貼り付けて「createEvaluationSheet」を実行
// ============================================================

N=5

function createEvaluationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 一時シートを用意（既存なら再利用、なければ新規作成）
  let tempSheet = ss.getSheetByName("_temp_");
  if (!tempSheet) tempSheet = ss.insertSheet("_temp_");
  const sheetsToDelete = ss.getSheets().filter(s => s.getName() !== "_temp_");
  sheetsToDelete.forEach(s => ss.deleteSheet(s));
  SpreadsheetApp.flush();

  // ★ ここを変えるだけで列数・ページ数・項目数を調整できます ★
  const RATING_COUNT = N;   // 評価列の数（①②③…好きな数に変更可）
  const ITEMS        = 9;   // 各ページの項目数
  const pageNames    = ["ページ1","ページ2","ページ3","ページ4","ページ5","ページ6","ページ7","ページ8"];

  const PAGE_COLORS  = ["#3A7CA5","#2E86AB","#1D6FA4","#2563A8","#1B6CA8","#2076B4","#165C8A","#0F4C75"];
  const HEADER_BG    = "#2E4057";
  const HEADER_FG    = "#FFFFFF";
  const ROW_ODD      = "#F7F9FC";
  const ROW_EVEN     = "#FFFFFF";
  const RATING_BG    = "#EAF4FB";
  const SUMMARY_BG   = "#1A3A5C";

  // 評価列ヘッダー（①②③④…）を動的生成
  const circledNums  = ["①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩"];
  const ratingHeaders = Array.from({length: RATING_COUNT}, (_, i) => `評価${circledNums[i]}`);

  // スプレッドシートの列文字（B, C, D, E, ...）を動的生成
  const colLetters = Array.from({length: RATING_COUNT}, (_, i) =>
    String.fromCharCode(66 + i) // B=66, C=67, ...
  );

  // 各ページの総列数（項目1列 + 評価N列 + コメント1列）
  const PAGE_TOTAL_COLS = 1 + RATING_COUNT + 1;
  const commentCol = 2 + RATING_COUNT; // コメント列番号

  const dvRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["S","A","B","C","D"], true)
    .setAllowInvalid(false)
    .setHelpText("S, A, B, C, D から選択してください")
    .build();

  // ── ページ1〜8 作成 ──────────────────────────────────────
  pageNames.forEach((pageName, pi) => {
    const ws = ss.insertSheet(pageName);

    // タイトル行
    ws.getRange(1, 1, 1, PAGE_TOTAL_COLS).merge()
      .setValue(pageName)
      .setBackground(PAGE_COLORS[pi])
      .setFontColor(HEADER_FG).setFontSize(14).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.setRowHeight(1, 36);

    // ヘッダー行（一括）
    const headerValues = [["項目", ...ratingHeaders, "コメント"]];
    ws.getRange(2, 1, 1, PAGE_TOTAL_COLS).setValues(headerValues)
      .setBackground(HEADER_BG).setFontColor(HEADER_FG)
      .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setBorder(true,true,true,true,true,true,"#888888", SpreadsheetApp.BorderStyle.SOLID);
    ws.setRowHeight(2, 24);

    // データ行を一括セット
    const emptyRow = Array(PAGE_TOTAL_COLS).fill("");
    const itemLabels = Array.from({length: ITEMS}, (_, i) => [`項目${i+1}`, ...emptyRow.slice(1)]);
    ws.getRange(3, 1, ITEMS, PAGE_TOTAL_COLS).setValues(itemLabels);

    // 項目列書式
    ws.getRange(3, 1, ITEMS, 1)
      .setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setBorder(true,true,true,true,false,false,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

    // 評価列書式＋ドロップダウン（一括）
    ws.getRange(3, 2, ITEMS, RATING_COUNT)
      .setBackground(RATING_BG).setFontSize(11).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setBorder(true,true,true,true,true,true,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID)
      .setDataValidation(dvRule);

    // コメント列書式
    ws.getRange(3, commentCol, ITEMS, 1)
      .setHorizontalAlignment("left").setVerticalAlignment("middle").setWrap(true)
      .setBorder(true,true,true,true,false,false,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

    // 奇数・偶数行の背景色
    for (let ri = 0; ri < ITEMS; ri++) {
      const fill = ri % 2 === 0 ? ROW_ODD : ROW_EVEN;
      ws.getRange(ri + 3, 1).setBackground(fill);
      ws.getRange(ri + 3, commentCol).setBackground(fill);
      ws.setRowHeight(ri + 3, 26);
    }

    // 列幅
    ws.setColumnWidths(1, 1, 140);
    ws.setColumnWidths(2, RATING_COUNT, 80);
    ws.setColumnWidths(commentCol, 1, 260);
  });

  // ── まとめシート 作成 ────────────────────────────────────
  const ws = ss.insertSheet("まとめ");
  const TOTAL_COLS = 1 + pageNames.length * RATING_COUNT;

  // タイトル行
  ws.getRange(1, 1, 1, TOTAL_COLS).merge()
    .setValue("評価まとめ")
    .setBackground(SUMMARY_BG).setFontColor(HEADER_FG)
    .setFontSize(16).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(1, 40);

  // 行2: 項目ヘッダー
  ws.getRange(2, 1)
    .setValue("項目").setBackground(HEADER_BG).setFontColor(HEADER_FG)
    .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,false,false,"#888888", SpreadsheetApp.BorderStyle.SOLID);

  // 行2: ページ名（結合）
  pageNames.forEach((pageName, pi) => {
    const sc = 2 + pi * RATING_COUNT;
    ws.getRange(2, sc, 1, RATING_COUNT).merge()
      .setValue(pageName).setBackground(PAGE_COLORS[pi]).setFontColor(HEADER_FG)
      .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setBorder(true,true,true,true,false,false,"#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);
  });
  ws.setRowHeight(2, 26);

  // 行3: サブヘッダー（一括）
  ws.getRange(3, 1).setBackground(HEADER_BG)
    .setBorder(true,true,true,true,false,false,"#888888", SpreadsheetApp.BorderStyle.SOLID);
  const subHeaders = [];
  pageNames.forEach(() => ratingHeaders.forEach(h => subHeaders.push(h)));
  ws.getRange(3, 2, 1, pageNames.length * RATING_COUNT)
    .setValues([subHeaders])
    .setBackground(HEADER_BG).setFontColor(HEADER_FG).setFontSize(9).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#888888", SpreadsheetApp.BorderStyle.SOLID);
  ws.setRowHeight(3, 22);

  // データ行（数式を一括セット）
  const dataRows = [];
  for (let ri = 0; ri < ITEMS; ri++) {
    const row = [`項目${ri + 1}`];
    const srcRow = ri + 3;
    pageNames.forEach(pageName => {
      colLetters.forEach(col => row.push(`='${pageName}'!${col}${srcRow}`));
    });
    dataRows.push(row);
  }
  ws.getRange(4, 1, ITEMS, TOTAL_COLS).setValues(dataRows);

  // まとめデータ書式（一括）
  ws.getRange(4, 1, ITEMS, TOTAL_COLS)
    .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  ws.getRange(4, 1, ITEMS, 1).setBackground(HEADER_BG).setFontColor(HEADER_FG);
  ws.getRange(4, 2, ITEMS, pageNames.length * RATING_COUNT).setBackground(RATING_BG).setFontSize(11);

  for (let ri = 0; ri < ITEMS; ri++) {
    ws.setRowHeight(ri + 4, 26);
  }

  // 列幅
  ws.setColumnWidths(1, 1, 80);
  ws.setColumnWidths(2, pageNames.length * RATING_COUNT, 70);

  // 一時シート削除・まとめを末尾へ
  ss.deleteSheet(tempSheet);
  ss.setActiveSheet(ws);
  ss.moveActiveSheet(ss.getNumSheets());

  SpreadsheetApp.getUi().alert("✅ 作成完了！\nページ1〜8 と まとめシート が生成されました。");
}
