/**
 * 🤖 万能フォーム集計システム v10.45 "Timeline Architect X24"
 * Based on v10.44
 * * * 【v10.45 修正内容】
 * - SOS Highlight Integration: 全校集計の記述回答まとめシートにおいて、
 * 設定パネル(B31, B32)のSOS設定（列・ワード）と連動し、
 * 該当するSOSワードを含む回答セルを自動的に「赤字・太字」でハイライトする機能を追加。
 */

const CONFIG_SHEET_NAME = "集計設定パネル";
const RESULT_SHEET_NAME = "集計結果";
const TEXT_SHEET_NAME = "📝記述回答まとめ";
const PERSONAL_SHEET_NAME = "🖨️個人カルテ";
const ALL_SCHOOL_SHEET_NAME = "🏫全校集計レポート";
const MASTER_SHEET_NAME = "名簿マスタ";
const APP_TITLE = "📊 フォーム集計システム v15";

// ★設定: 履歴参照リミット
const MAX_RECORDS = 50000; 
// ★設定: 印刷時の1名あたりの行数 (40行固定)
const PAGE_BREAK_ROWS = 40;
// ★設定: 1行あたりの高さ(ピクセル) ※ここで行の高さを調整
const ROW_HEIGHT_PX = 23; 

// 設定行の定義 (全体集計用)
const FILTER_ROW_A = 7;
const FILTER_ROW_B = 10;
const FILTER_ROW_C = 13;
const CROSS_AXIS_LABEL_ROW = 17;
const CROSS_AXIS_VAL_ROW = 18;

// 学校用設定エリアの開始行
const SCHOOL_CONFIG_START_ROW = 25;
// 25: header
// 26: class (対象クラス)
// 27: key col (ID/Email)
// 28: date col (日付 or 回)
// 29: (empty)
// 30: SOS Header
// 31: SOS Col
// 32: SOS Word
// 33: (empty)
// 34: Chart Header
// 35-42: Radar 1-8
// 43: Unit Selector (抽出単位)
// 44-55: Compare Points 1-12

const SCHOOL_DATE_COMPARE_START_ROW = 44; 

// ==================================================
// 🚪 1. トリガー & メニュー制御
// ==================================================

function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  let menu = ui.createMenu(APP_TITLE)
    .addItem('1. ⚙️ 初期設定', 'initConfiguration')
    .addSeparator()
    .addItem('2. 📊 全体集計実行', 'runUniversalAnalysis')
    .addSeparator();

  if (masterSheet) {
    menu.addItem('3. 🏫 名簿マスタ管理 (更新)', 'enableSchoolMode')
        .addItem('4. 🖨️ 個人カルテ・SOS作成', 'runPersonalAnalysis')
        // ★ここに新機能を挿入
        .addItem('5. 🏫 クラス集計 (時系列・抽出)', 'runClassMatrixAnalysis') 
        // ★既存機能を繰り下げ
        .addItem('6. 🏫 全校集計 (時系列マトリクス)', 'runAllSchoolAnalysis'); 
  } else {
    menu.addItem('3. 🏫 学校機能・名簿管理モードON', 'enableSchoolMode');
  }

  menu.addToUi();
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    // 名簿マスタの変更監視 -> 設定パネルへ警告
    if (sheetName === MASTER_SHEET_NAME) {
      const configSheet = e.source.getSheetByName(CONFIG_SHEET_NAME);
      if (configSheet) {
        const classSelectCell = configSheet.getRange(SCHOOL_CONFIG_START_ROW + 1, 2);
        classSelectCell.setValue("⚠️名簿変更検知: メニュー[3]で更新してください")
                       .setFontColor("red")
                       .setFontWeight("bold")
                       .clearDataValidations();
      }
      return;
    }

    // 設定パネルの操作監視
    if (sheetName === CONFIG_SHEET_NAME) {
      if (col === 2) {
        // B3: 対象シート変更 -> 全リセット＆更新
        if (row === 3) {
          detectAnswerSheetColumns_(sheet, SCHOOL_CONFIG_START_ROW);
          updateClassDropdown_(sheet);
          updateQuestionDropdowns_(sheet); 
          updateDateDropdown_(sheet); 
        }
        
        // 条件設定列
        if ([FILTER_ROW_A, FILTER_ROW_B, FILTER_ROW_C, CROSS_AXIS_LABEL_ROW].includes(row)) {
          updateQuestionDropdowns_(sheet); 
          if (row !== CROSS_AXIS_LABEL_ROW) {
            updateValueDropdown_(sheet, row);
          }
        }

        // 学校SOS設定 (行31)
        const schoolSosRow = SCHOOL_CONFIG_START_ROW + 6;
        if (row === schoolSosRow) {
          updateValueDropdown_(sheet, row);
        }

        // レーダー項目の変更監視 (行35-42)
        const radarStart = SCHOOL_CONFIG_START_ROW + 10;
        const radarEnd = radarStart + 8;
        if (row >= radarStart && row < radarEnd) {
          updateQuestionDropdowns_(sheet);
        }

        // ★日付(回)列 or 単位セレクタ変更 -> プルダウン更新
        const dateColRow = SCHOOL_CONFIG_START_ROW + 3; // 行28
        const unitSelectorRow = SCHOOL_DATE_COMPARE_START_ROW - 1; // 行43
        
        // B44～B55の変更も監視して、重複除外をリアルタイム反映
        const isComparePointRow = (row >= SCHOOL_DATE_COMPARE_START_ROW && row < SCHOOL_DATE_COMPARE_START_ROW + 12);

        if (row === dateColRow || row === unitSelectorRow || isComparePointRow) {
           updateDateDropdown_(sheet);
        }
      }
    }
  } catch (err) {
    console.error("onEdit Error: " + err.message);
  }
}

// ==================================================
// ⚙️ 2. 初期設定 (Hybrid UI) - Revised v10.46
// ==================================================

function initConfiguration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    
    if (!configSheet) {
      configSheet = ss.insertSheet(CONFIG_SHEET_NAME, 0);
    }
    
    if (configSheet.getLastRow() > 5) {
       const res = ui.alert('確認', '設定パネルを初期化しますか？\n（入力済みの値はクリアされます）', ui.ButtonSet.YES_NO);
       if (res == ui.Button.NO) return;
    }
    configSheet.clear();
    
    // --- レイアウト定義 ---
    const layout = [
      ["📊 フォーム集計システム 設定パネル", ""], // 1
      ["【基本設定】", ""], // 2
      ["① 対象シート名(回答)", "▼シートを選択"], // 3
      ["集計対象の列(質問)", "自動取得"], // 4
      ["", ""], // 5
      ["【全体集計: フィルタリング】", "※任意で絞り込みできます（空欄OK）"], // 6 ★案内文追加
      ["条件A (列名)", "▼質問を選択"], // 7
      ["　値 (一致)", "-"], // 8
      ["", ""], // 9
      ["条件B (列名)", "▼質問を選択"], // 10
      ["　値 (一致)", "-"], // 11
      ["", ""], // 12
      ["条件C (列名)", "▼質問を選択"], // 13
      ["　値 (一致)", "-"], // 14
      ["", ""], // 15
      ["【全体集計: 詳細設定】", ""], // 16
      ["比較分析する列 (横軸)", "▼質問を選択"], // 17
      ["※選択すると右側に詳細表を作成", ""], // 18
      ["全体集計:タイムスタンプ単位", "▼自動(しない)"], // 19 ★New: 日付集計設定
      ["", ""]  // 20
    ];
    
    configSheet.getRange(1, 1, layout.length, 2).setValues(layout);
    
    // スタイル適用
    configSheet.getRange("A1:B1").merge().setFontSize(14).setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
    
    const mainConfigRange = configSheet.getRange("A3:B4");
    mainConfigRange.setBorder(true, true, true, true, true, true, "red", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    configSheet.getRange("A3:A4").setFontWeight("bold").setBackground("#FFEBEE"); 
    configSheet.getRange("B3:B4").setFontWeight("bold").setBackground("#FFFFFF");
    
    configSheet.getRange("A2").setFontWeight("bold").setBackground("#EFEFEF");
    configSheet.getRange("A6").setFontWeight("bold").setBackground("#EFEFEF");
    configSheet.getRange(16, 1).setFontWeight("bold").setBackground("#D9EAD3"); 
// A19:B19を黒い太枠で囲む
    configSheet.getRange("A19:B19")
      .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);


    // ★UI: B6の案内文を目立たない色に
    configSheet.getRange("B6").setFontColor("gray").setFontSize(8);
    
    [7, 10, 13, 17].forEach(r => {
      configSheet.getRange(r, 2).setBackground("#FFF2CC");
    }); 
    
    [8, 11, 14].forEach(r => {
      configSheet.getRange(r, 2).setBackground("#FFFFFF").setBorder(null, null, true, null, null, null);
    }); 

    configSheet.getRange(18, 1).setFontSize(8).setFontColor("gray");

    // ★行19: 日付単位プルダウン (New)
    const dateUnitCell = configSheet.getRange(19, 2);
    const dateRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["▼自動(しない)", "【年別】", "【月別】", "【日付別】"])
        .build();
    dateUnitCell.setDataValidation(dateRule).setBackground("#FFF2CC");

    configSheet.setColumnWidth(1, 200);
    configSheet.setColumnWidth(2, 400);

    // シート一覧プルダウン
    const sheets = ss.getSheets().filter(s => ![CONFIG_SHEET_NAME, RESULT_SHEET_NAME, TEXT_SHEET_NAME, MASTER_SHEET_NAME, PERSONAL_SHEET_NAME, ALL_SCHOOL_SHEET_NAME].includes(s.getName()));
    const sheetNames = sheets.map(s => s.getName());
    
    if (sheetNames.length > 0) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(sheetNames).build();
      const targetCell = configSheet.getRange("B3");
      targetCell.setDataValidation(rule).setValue(sheetNames[0]);
      
      SpreadsheetApp.flush(); 
      updateQuestionDropdowns_(configSheet); 
    } else {
      configSheet.getRange("B3").setValue("フォームの回答シートがありません");
    }
    
    // 学校用エリア
    initSchoolConfigArea_(configSheet);
    
    if (!ss.getSheetByName(MASTER_SHEET_NAME)) {
      const maxRows = configSheet.getMaxRows();
      if (maxRows >= SCHOOL_CONFIG_START_ROW) {
        configSheet.hideRows(SCHOOL_CONFIG_START_ROW, maxRows - SCHOOL_CONFIG_START_ROW + 1);
      }
    } else {
      updateClassDropdown_(configSheet);
    }
    
    ui.alert("初期設定が完了しました。\n赤枠の「対象シート」を選択してください。");

  } catch (e) {
    Browser.msgBox("⚠️ 初期設定中にエラーが発生しました:\n" + e.message);
    console.error(e.stack);
  }
}

function initSchoolConfigArea_(sheet) {
  const startRow = SCHOOL_CONFIG_START_ROW;
  sheet.getRange(startRow, 1, 60, 2).clear(); 

  const schoolLayout = [
    ["🏫 学校・クラス・個人カルテ設定", ""], // 25
    ["対象クラス", "▼名簿から自動生成"], // 26
    ["回答シートの「Key(ID/Email)」列", ""], // 27
    ["回答シートの「日付・回」列", "▼自動判定"], // 28 (Updated)
    ["", ""], // 29
    ["【SOS検知設定】", ""], // 30
    ["🚨 SOS判定する質問(列)", "▼ここから質問を選択"], // 31
    ["🚨 反応する言葉(部分一致)", "（例）つらい、苦しい、休みたい"], // 32
    ["", ""], // 33
    ["【カルテ出力設定】", ""], // 34
    ["レーダー項目1", ""], // 35
    ["レーダー項目2", ""], 
    ["レーダー項目3", ""], 
    ["レーダー項目4", ""], 
    ["レーダー項目5", ""], 
    ["レーダー項目6", ""], 
    ["レーダー項目7", ""], 
    ["レーダー項目8", ""], // 42
    ["【比較データの抽出単位】", "【日付別】"], // 43
    ["比較対象ポイント 1", ""], // 44
    ["比較対象ポイント 2", ""],
    ["比較対象ポイント 3", ""],
    ["比較対象ポイント 4", ""],
    ["比較対象ポイント 5", ""],
    ["比較対象ポイント 6", ""],
    ["比較対象ポイント 7", ""],
    ["比較対象ポイント 8", ""],
    ["比較対象ポイント 9", ""],
    ["比較対象ポイント 10", ""],
    ["比較対象ポイント 11", ""],
    ["比較対象ポイント 12", ""]
  ];
  
  sheet.getRange(startRow, 1, schoolLayout.length, 2).setValues(schoolLayout);
  sheet.getRange(startRow, 1, 1, 2).merge().setFontSize(12).setFontWeight("bold").setBackground("#34A853").setFontColor("white");
  
  sheet.getRange(startRow + 5, 1).setFontWeight("bold").setBackground("#E6F4EA"); // SOS Header (30)
  sheet.getRange(startRow + 9, 1).setFontWeight("bold").setBackground("#E6F4EA"); // Chart Header (34)
  sheet.getRange(startRow + 18, 1).setFontWeight("bold").setBackground("#E6F4EA"); // Compare Header (43)

  sheet.getRange(startRow + 6, 2).setBackground("#FFF2CC"); // SOS Col (31)
  sheet.getRange(startRow + 7, 2).setBackground("#FFFFFF").setBorder(null, null, true, null, null, null); 
  
  // 日付列設定 (行28)
  sheet.getRange(startRow + 3, 2).setBackground("#FFF2CC");

  // レーダー項目エリア (行35-42)
  sheet.getRange(startRow + 10, 2, 8, 1).setBackground("#F3F3F3");
  
  // 単位セレクタ (行43)
  const unitCell = sheet.getRange(SCHOOL_DATE_COMPARE_START_ROW - 1, 2);
  const unitRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["【日付別】", "【月別】", "【年別】"])
    .build();
  unitCell.setDataValidation(unitRule)
          .setValue("【日付別】")
          .setBackground("#FFF2CC")
          .setFontWeight("bold");

  // 日付比較エリア (行44～55)
  sheet.getRange(SCHOOL_DATE_COMPARE_START_ROW, 2, 12, 1).setBackground("#FFFFFF");

  SpreadsheetApp.flush();

  detectAnswerSheetColumns_(sheet, startRow);
  updateQuestionDropdowns_(sheet); 
  updateDateDropdown_(sheet);
}

function enableSchoolMode() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    let masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    
    if (!masterSheet) {
      masterSheet = ss.insertSheet(MASTER_SHEET_NAME);
      const headers = [["Account(Email/ID)", "学年", "組", "番号", "氏名", "ふりがな（任意）", "性別（任意）"]];
      masterSheet.getRange("A1:G1").setValues(headers).setFontWeight("bold").setBackground("#FFF2CC");
      masterSheet.setFrozenRows(1);
      
      const sample = [
        ["st01@ex.com", 1, 1, 1, "相川 翔", "あいかわ しょう", "男"],
        ["st02@ex.com", 1, 1, 2, "井上 真", "いのうえ まこと", "女"],
        ["st03@ex.com", 1, 2, 1, "上野 樹里", "うえの じゅり", "女"],
        ["st04@ex.com", 1, 10, 1, "遠藤 憲一", "えんどう けんいち", "男"],
        ["st05@ex.com", 1, "ひまわり", 1, "大谷 翔平", "おおたに しょうへい", "男"],
        ["st06@ex.com", 2, "A", 1, "加藤 茶", "かとう ちゃ", "男"],
        ["st07@ex.com", 2, "B", 1, "北川 景子", "きたがわ けいこ", "女"],
        ["st08@ex.com", 2, "特2", 1, "久保田 利伸", "くぼた としのぶ", "男"],
        ["st09@ex.com", 2, "コスモス", 1, "小池 栄子", "こいけ えいこ", "女"],
        ["st10@ex.com", 3, "I", 1, "佐藤 健", "さとう たける", "男"],
        ["st11@ex.com", 3, "II", 1, "鈴木 亮平", "すずき りょうへい", "男"],
        ["st12@ex.com", 3, "い", 1, "高橋 一生", "たかはし いっせい", "男"],
        ["st13@ex.com", 3, "ろ", 1, "千鳥 ノブ", "ちどり のぶ", "男"],
        ["st14@ex.com", "全", "ひまわり", 2, "妻夫木 聡", "つまぶき さとし", "男"],
        ["st15@ex.com", "全", "特2", 2, "寺田 心", "てらだ こころ", "男"]
      ];
      masterSheet.getRange(2, 1, sample.length, sample[0].length).setValues(sample);
      
      SpreadsheetApp.flush();
      Browser.msgBox("「名簿マスタ」シートを作成しました。");
    }
    
    if (configSheet) {
      initSchoolConfigArea_(configSheet);
      const maxRows = configSheet.getMaxRows();
      configSheet.showRows(SCHOOL_CONFIG_START_ROW, maxRows - SCHOOL_CONFIG_START_ROW + 1);
      updateClassDropdown_(configSheet);
    }
    
    onOpen(); 
    Browser.msgBox("学校機能モードを有効化しました。");
    
  } catch (e) {
    Browser.msgBox("⚠️ 学校モード有効化中にエラーが発生しました:\n" + e.message);
  }
}


// ==================================================
// 📊 4. 全体集計実行 (Universal Analysis) - Revised v10.46
// ==================================================

function runUniversalAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    if (!configSheet) { Browser.msgBox("先に「1. 初期設定」を実行してください。"); return; }

    const targetSheetName = configSheet.getRange("B3").getValue();
    const dataSheet = ss.getSheetByName(targetSheetName);
    if (!dataSheet) { Browser.msgBox(`エラー: 対象シート「${targetSheetName}」が見つかりません。`); return; }

    // ★New: 日付チャート単位の取得 (B19)
    const dateUnitVal = configSheet.getRange(19, 2).getValue();
    const isDateChartEnabled = dateUnitVal && dateUnitVal !== "▼自動(しない)" && !String(dateUnitVal).startsWith("▼");
    let dateFormat = "yyyy/MM/dd";
    if (dateUnitVal === "【年別】") dateFormat = "yyyy";
    if (dateUnitVal === "【月別】") dateFormat = "yyyy/MM";

    const totalLastRow = dataSheet.getLastRow();
    const lastCol = dataSheet.getLastColumn();
    if (totalLastRow < 2) { Browser.msgBox("データがありません。"); return; }

    const headers = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    let startRow = 2;
    let numRows = totalLastRow - 1;
    if (numRows > MAX_RECORDS) {
      startRow = totalLastRow - MAX_RECORDS + 1;
      numRows = MAX_RECORDS;
    }
    const body = dataSheet.getRange(startRow, 1, numRows, lastCol).getValues();

    let filters = [];
    [FILTER_ROW_A, FILTER_ROW_B, FILTER_ROW_C].forEach(r => {
      let cName = configSheet.getRange(r, 2).getValue();
      let cVal = configSheet.getRange(r+1, 2).getValue();
      if (cName && cVal !== "" && !String(cName).startsWith("▼")) {
        filters.push({ name: cName, value: String(cVal) });
      }
    });

    const uniqueFilterNames = new Set(filters.map(f => f.name));
    if (uniqueFilterNames.size !== filters.length) {
      Browser.msgBox("⚠️ エラー: 同じ列で複数のフィルタリング条件を指定することはできません。");
      return;
    }

    let targetRows = body;
    let filterLogArr = [];

    if (filters.length > 0) {
      const validFilters = filters.map(f => {
        const idx = headers.indexOf(f.name);
        return { index: idx, value: f.value, name: f.name };
      }).filter(f => f.index !== -1); 

      if (validFilters.length > 0) {
          targetRows = body.filter(row => {
            return validFilters.every(f => String(row[f.index]) === f.value);
          });
          filterLogArr = validFilters.map(f => `${f.name}=${f.value}`);
      }
    }

    if (targetRows.length === 0) {
      Browser.msgBox(`条件に一致するデータはありませんでした。`);
      return;
    }

    ss.toast("集計を開始します...", "処理中", 10);

    let resultSheet = ss.getSheetByName(RESULT_SHEET_NAME);
    if (resultSheet) {
      const existingCharts = resultSheet.getCharts();
      existingCharts.forEach(c => resultSheet.removeChart(c));
      resultSheet.clear();
    } else {
      resultSheet = ss.insertSheet(RESULT_SHEET_NAME);
    }

    let textSheet = ss.getSheetByName(TEXT_SHEET_NAME);
    if (textSheet) {
      textSheet.clear();
    } else {
      textSheet = ss.insertSheet(TEXT_SHEET_NAME);
      textSheet.setTabColor("yellow");
    }
    textSheet.getRange(1, 1).setValue("📝 自由記述回答まとめ (最新順)").setFontSize(14).setFontWeight("bold");
    let textSheetCurrentCol = 1; 

    let currentRow = 1;

    resultSheet.getRange(currentRow, 1).setValue(`集計レポート: ${targetSheetName}`).setFontWeight("bold");
    currentRow++;
    resultSheet.getRange(currentRow, 1).setValue(`絞り込み: ${filterLogArr.join(" AND ") || "（全件）"}`);
    currentRow++;
    resultSheet.getRange(currentRow, 1).setValue(`対象件数: ${targetRows.length}件`);
    currentRow += 2;

    let chartConfigs = [];

    for (let col = 1; col < headers.length; col++) {
      const question = headers[col];
      if (!question) continue;

      const colValues = targetRows.map(r => r[col]).filter(v => v !== "" && v != null);
      if (colValues.length === 0) continue;

      const colType = analyzeColumnType_(colValues, question);

      // ★Modified: 日付(TIMESTAMP)の扱い変更
      // 設定がOFFならスキップ、ONなら通過させる
      if (colType === 'SKIP') continue;
      if (colType === 'TIMESTAMP' && !isDateChartEnabled) continue;

      if (colType === 'FREE_TEXT') {
        textSheet.getRange(3, textSheetCurrentCol).setValue(question)
          .setFontWeight("bold").setBackground("#f3f3f3").setBorder(true, true, true, true, null, null);
        
        const responses = colValues.reverse(); 
        if (responses.length > 0) {
          // ★記述回答での日付フォーマット統一
          const formattedRes = responses.map(v => {
             if (v instanceof Date) return [Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy/MM/dd")];
             return [v];
          });
          textSheet.getRange(4, textSheetCurrentCol, formattedRes.length, 1).setValues(formattedRes);
        }
        
        textSheet.setColumnWidth(textSheetCurrentCol, 300); 
        textSheetCurrentCol += 1; 
        continue;
      }

      let counts = {};
      let totalScore = 0;
      let numericCount = 0;

      colValues.forEach(val => {
        let strVal = String(val);

        // ★Modified: 日付フォーマットの適用
        if (val instanceof Date) {
            strVal = Utilities.formatDate(val, Session.getScriptTimeZone(), dateFormat);
        } else if (colType === 'TIMESTAMP') {
            const d = new Date(val);
            if(!isNaN(d)) strVal = Utilities.formatDate(d, Session.getScriptTimeZone(), dateFormat);
        }

        // ★Strict Fix: 数値判定の厳格化 (parseFloat -> Number)
        const num = Number(strVal);
        if (!isNaN(num) && strVal.trim() !== "") { 
          totalScore += num; 
          numericCount++; 
        }

        if (strVal.includes(',') && strVal.length > 2) {
          strVal.split(',').map(s => s.trim()).forEach(p => { 
            if(p) counts[p] = (counts[p] || 0) + 1; 
          });
        } else {
          counts[strVal] = (counts[strVal] || 0) + 1;
        }
      });

      if (numericCount > 0 && numericCount > (targetRows.length * 0.5)) {
        resultSheet.getRange(currentRow, 1).setNote(`平均: ${(totalScore / numericCount).toFixed(2)}`);
      }

      const uniqueKeys = Object.keys(counts);

      // ★Safety: 項目が多すぎる場合の集約処理 (Top 20 + Others)
      let finalKeys = [];
      let finalCounts = {};
      
      if (uniqueKeys.length > 20) {
        // カウント順にソート
        const sortedAll = uniqueKeys.sort((a, b) => counts[b] - counts[a]);
        const top19 = sortedAll.slice(0, 19);
        const others = sortedAll.slice(19);
        
        top19.forEach(k => {
           finalKeys.push(k);
           finalCounts[k] = counts[k];
        });
        
        let otherSum = 0;
        others.forEach(k => otherSum += counts[k]);
        if (otherSum > 0) {
          finalKeys.push("その他");
          finalCounts["その他"] = otherSum;
        }
      } else {
        // 通常ソート
        finalKeys = uniqueKeys.sort((a, b) => counts[b] - counts[a]);
        finalCounts = counts;
      }

      resultSheet.getRange(currentRow, 1).setValue(`Q${col}. ${question}`).setFontWeight("bold");
      currentRow++;

      resultSheet.getRange(currentRow, 1, 1, 3).setValues([["回答", "件数", "割合"]])
        .setBackground("#e0e0e0").setFontWeight("bold");
      currentRow++;

      const startDataRow = currentRow;
      finalKeys.forEach(key => {
          const cnt = finalCounts[key];
          let pct = targetRows.length > 0 ? Math.round((cnt / targetRows.length) * 100) + "%" : "0%";
          resultSheet.getRange(currentRow, 1, 1, 3).setValues([[key, cnt, pct]]);
          currentRow++;
      });

      chartConfigs.push({
          title: `Q${col}. ${question}`,
          startRow: startDataRow, 
          rowCount: finalKeys.length,
          type: finalKeys.length <= 6 ? "PIE" : "BAR",
          anchorRow: startDataRow - 2
      });

      currentRow += 2;
    }

    resultSheet.setColumnWidth(1, 300);
    resultSheet.setColumnWidth(4, 400); 

   // ... (runUniversalAnalysisの前半部分はそのまま) ...

    try { 
      generateUniversalCharts_(resultSheet, chartConfigs);
    } catch (e) { 
      console.error(e);
    }

    // ==========================================
    // ▼▼▼ ここからロジック修正 (Fix for Issue ① & ②) ▼▼▼
    // ==========================================
    
    // 1. 次の開始行を現在の最終行から安全に取得
    let nextStartRow = resultSheet.getLastRow() + 3;

    // 2. 詳細クロス集計 (B17設定ありの場合)
    const crossAxisColName = configSheet.getRange(CROSS_AXIS_LABEL_ROW, 2).getValue();
    
    if (crossAxisColName && !String(crossAxisColName).startsWith("▼")) {
      const crossIdx = headers.indexOf(crossAxisColName);
      if (crossIdx !== -1) {
        const isTimestamp = /タイムスタンプ|Timestamp|日時|Date/i.test(crossAxisColName);
        let modeMsg = "";
        if (isTimestamp) {
             if(dateFormat === "yyyy") modeMsg = "【年別推移モード】";
             else if(dateFormat === "yyyy/MM") modeMsg = "【月別推移モード】";
             else modeMsg = "【日別推移モード】";
        }
        
        ss.toast(`詳細クロス集計を作成中... ${modeMsg}`, "分析中", 20);
        Utilities.sleep(100);

        // ★修正: 戻り値を確実に受け取り、かつエラー時も停止させない
        try {
          const crossResultRow = renderCrossTabulation_(resultSheet, headers, targetRows, crossIdx, crossAxisColName, 8, isTimestamp, dateFormat);
          // もし有効な行数が返ってきたら更新、そうでなければ元のまま
          if (crossResultRow && crossResultRow > nextStartRow) {
            nextStartRow = crossResultRow;
          }
        } catch (e) {
          console.warn("CrossTab Error: " + e.message);
          // エラーが出ても次の処理に進むため、行だけ少し空ける
          nextStartRow = resultSheet.getLastRow() + 5;
        }
      }
    }

    // 安全マージン（グラフ重複防止のため念のため空ける）
    nextStartRow += 2;

    // 3. 相関分析マトリクス実行
    try {
      // 念のため再度最終行チェック（グラフ等の浮動要素対策）
      const checkRow = resultSheet.getLastRow() + 3;
      if (checkRow > nextStartRow) nextStartRow = checkRow;

      const corrResultRow = generateCorrelationMatrix_(resultSheet, headers, targetRows, nextStartRow);
      if (corrResultRow) nextStartRow = corrResultRow;
    } catch (e) { 
      console.warn("Correlation Error", e); 
      // エラー表示をシートに出す（デバッグ用）
      resultSheet.getRange(nextStartRow, 1).setValue("⚠️ 相関分析エラー: データ不足または形式不一致");
      nextStartRow += 2;
    }

    // 4. 抽出生データテーブル出力
    try {
      renderRawDataTable_(resultSheet, headers, targetRows, nextStartRow);
    } catch (e) { 
      console.warn("RawData Error", e); 
    }

    // ▲▲▲ ロジック修正ここまで ▲▲▲
    // ==========================================

    resultSheet.activate();
    ss.toast("集計完了！記述回答は別シートにまとめました。", "完了", 5);
    Browser.msgBox(`全体集計完了！\n記述回答は「${TEXT_SHEET_NAME}」を確認してください。`);

  } catch (e) {

    Browser.msgBox("⚠️ 全体集計中にエラーが発生しました:\n" + e.message);
  }
}



// ==================================================
// 🖨️ 5. 個人カルテ・SOS作成 (v10.46 High-Speed Batch Edition)
// ==================================================

function runPersonalAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(CONFIG_SHEET_NAME);
  const master = ss.getSheetByName(MASTER_SHEET_NAME);

  try {
    if (!config) throw new Error("設定パネルが見つかりません。初期設定を実行してください。");

    // --- 1. 設定情報の取得 ---
    const targetSheetName = config.getRange("B3").getValue();
    const targetClass = config.getRange(SCHOOL_CONFIG_START_ROW + 1, 2).getValue();
    const ansKeyCol = config.getRange(SCHOOL_CONFIG_START_ROW + 2, 2).getValue(); // Row 27
    const dateColStr = config.getRange(SCHOOL_CONFIG_START_ROW + 3, 2).getValue(); // Row 28

    const sosColName = config.getRange(SCHOOL_CONFIG_START_ROW + 6, 2).getValue(); // Row 31
    const sosValue = config.getRange(SCHOOL_CONFIG_START_ROW + 7, 2).getValue(); // Row 32

    const timeUnit = config.getRange(SCHOOL_DATE_COMPARE_START_ROW - 1, 2).getValue(); // Row 43

    // レーダー項目取得
    const radarCols = [];
    for (let i = 0; i < 8; i++) {
      const val = config.getRange(SCHOOL_CONFIG_START_ROW + 10 + i, 2).getValue(); // Row 35-42
      if (val) radarCols.push(val);
    }

    if (radarCols.length === 0) {
      Browser.msgBox("⚠️ 設定エラー: レーダーチャートの項目が1つも選択されていません。");
      return;
    }

    // 比較対象ポイントリストの取得 (B44～B55)
    // ★Fix: .getDisplayValues() を使用して「見た目の文字」をそのまま取得する
    // これにより「10月」が勝手に日付型変換されて不一致になる問題を回避
    const comparePointsRaw = config.getRange(SCHOOL_DATE_COMPARE_START_ROW, 2, 12, 1).getDisplayValues().flat();
    const comparePoints = comparePointsRaw.filter(s => s !== "");

    const isDateMode = ["【日付別】", "【月別】", "【年別】"].includes(timeUnit);

    if (!master || !targetClass || !ansKeyCol || String(ansKeyCol).startsWith("▼")) {
      Browser.msgBox("⚠️ 設定エラー:\n学校・カルテ設定の必須項目（対象クラス、Key列など）が正しく選択されていません。");
      return;
    }

    // --- 2. 生徒データの抽出 ---
    const masterData = master.getDataRange().getValues();
    const mGradeIdx = 1, mClassIdx = 2, mNumIdx = 3, mNameIdx = 4, mKeyIdx = 0, mGenderIdx = 6;

    let targetStudents = [];
    if (targetClass.startsWith("(全学年)")) {
      const tClass = targetClass.replace("(全学年)", "");
      targetStudents = masterData.slice(1).filter(row => String(row[mClassIdx]) === tClass);
    } else {
      const match = targetClass.match(/^(.+)年(.+)組$/);
      if (match) {
        targetStudents = masterData.slice(1).filter(row => String(row[mGradeIdx]) === match[1] && String(row[mClassIdx]) === match[2]);
      }
    }

    if (targetStudents.length === 0) {
      Browser.msgBox(`クラス「${targetClass}」の生徒が見つかりません。`);
      return;
    }

    // --- 3. 回答データのマッピング ---
    const dataSheet = ss.getSheetByName(targetSheetName);
    if (!dataSheet) throw new Error(`回答シート「${targetSheetName}」が見つかりません。`);

    const dHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    const allResponses = dataSheet.getDataRange().getValues().slice(1);

    let ansKeyColIdx = -1;
    const kIdx = dHeaders.indexOf(ansKeyCol);
    if (kIdx > -1) ansKeyColIdx = kIdx;
    else ansKeyColIdx = letterToColumn_(ansKeyCol) - 1;

    if (ansKeyColIdx < 0) throw new Error("Key列の指定が不正です。設定パネルを確認してください。");

    // 日付(時系列)列の特定
    let dateColIdx = 0; // デフォルトはA列
    if (dateColStr && !String(dateColStr).startsWith("▼")) {
      const idx = dHeaders.indexOf(dateColStr);
      if (idx > -1) dateColIdx = idx;
      else dateColIdx = letterToColumn_(dateColStr) - 1;
    }
    if (dateColIdx < 0) dateColIdx = 0;

    let sosIdx = sosColName ? dHeaders.indexOf(sosColName) : -1;
    const radarIndices = radarCols.map(name => dHeaders.indexOf(name));

    if (radarIndices.some(idx => idx === -1)) {
      throw new Error("選択されたレーダー項目の一部が、回答シート内に見つかりません。");
    }

    let responseMap = {};
    allResponses.forEach(row => {
      const val = row[ansKeyColIdx];
      const key = val != null ? String(val).trim() : "";
      if (key === "") return;

      if (!responseMap[key]) {
        responseMap[key] = [];
      }
      responseMap[key].push(row);
    });

    // --- 4. シート初期化 (高速化のため一度削除して作り直す) ---
    let pSheet = ss.getSheetByName(PERSONAL_SHEET_NAME);
    if (pSheet) ss.deleteSheet(pSheet);
    pSheet = ss.insertSheet(PERSONAL_SHEET_NAME);

    // --- 5. バッチ処理用メモリ確保 ---
    // 生徒数 × 1人あたりの行数 = 全体の行数
    const totalRows = targetStudents.length * PAGE_BREAK_ROWS;
    const maxCols = 30; // 安全のため多めに確保

    // 全セルの値を格納する巨大な配列
    const allValues = new Array(totalRows).fill(null).map(() => new Array(maxCols).fill(""));
    // 書式情報の配列
    const allBackgrounds = new Array(totalRows).fill(null).map(() => new Array(maxCols).fill(null));
    const allFontWeights = new Array(totalRows).fill(null).map(() => new Array(maxCols).fill("normal"));
    const allFontColors = new Array(totalRows).fill(null).map(() => new Array(maxCols).fill("black"));
    const allBorders = []; // 枠線適用箇所リスト {r, c, h, w, color}
    const allMerges = [];  // セル結合リスト range string

    const chartQueue = [];
    let printedCount = 0;

    ss.toast(`${targetStudents.length}名分のデータを処理中...`, "高速生成モード", 60);

    // --- 6. 生徒ループ (メモリ内処理) ---
    targetStudents.forEach((student, sIndex) => {
      const startRowIdx = sIndex * PAGE_BREAK_ROWS; // 0始まりの配列インデックス
      const currentRowNum = startRowIdx + 1; // 1始まりのシート行番号

      const acct = student[mKeyIdx], name = student[mNameIdx];
      const grade = student[mGradeIdx], cls = student[mClassIdx], num = student[mNumIdx], gender = student[mGenderIdx];

      let myResponses = responseMap[String(acct).trim()] || [];

      // ソートロジック
      if (isDateMode) {
        myResponses.sort((a, b) => new Date(a[dateColIdx]) - new Date(b[dateColIdx]));
      } else {
        myResponses.sort((a, b) => String(a[dateColIdx]).localeCompare(String(b[dateColIdx]), undefined, { numeric: true }));
      }

      // SOSチェック
      let isSos = false;
      if (sosIdx !== -1 && sosValue && myResponses.length > 0) {
        if (String(myResponses[myResponses.length - 1][sosIdx]) === String(sosValue)) {
          isSos = true;
        }
      }

      printedCount++;

      // --- A. ヘッダー情報 (配列へ書き込み) ---
      const genderText = gender ? `(${gender})` : "";
      const titleText = `【カルテ】${grade}年${cls}組${num}番 氏名: ${name} ${genderText}` + (isSos ? " ⚠️SOS" : "");
      
      allValues[startRowIdx][0] = titleText;
      allFontWeights[startRowIdx][0] = "bold";
      // ※フォントサイズ変更は後で一括で行うか、標準のままにする（高速化のため標準推奨だが、最後に範囲指定で変更可）
      
      // 背景色設定 (SOSなら赤)
      const headerBg = isSos ? "#FCE8E6" : "#E8F0FE";
      for(let c=0; c<14; c++) allBackgrounds[startRowIdx][c] = headerBg;

      // 結合予約
      allMerges.push(pSheet.getRange(currentRowNum, 1, 1, 14)); // A~N

      if (isSos) {
        // 枠線予約
        allBorders.push({ r: currentRowNum, c: 1, h: 2, w: 8, color: "red" });
      }

      const countText = myResponses.length > 0 ? `${myResponses.length}回` : "なし";
      let lastDateStr = "-";
      if (myResponses.length > 0) {
        const rawD = myResponses[myResponses.length - 1][dateColIdx];
        if (rawD instanceof Date) {
          lastDateStr = Utilities.formatDate(rawD, Session.getScriptTimeZone(), "yyyy/MM/dd");
        } else {
          lastDateStr = String(rawD);
        }
      }
      allValues[startRowIdx + 1][0] = `最終更新: ${lastDateStr} / ${countText}`;
      allMerges.push(pSheet.getRange(currentRowNum + 1, 1, 1, 14));

      // --- B. レーダーチャート用データ ---
      const chartBaseRelRow = 3; // 相対行 3 (currentRowNum + 3)
      
      if (myResponses.length > 0 && radarCols.length > 0) {
        const generations = myResponses.slice(-3).reverse();
        const shortRadarCols = radarCols.map(c => c.length > 9 ? c.substring(0, 9) : c);
        
        // ヘッダー
        allValues[startRowIdx + chartBaseRelRow][0] = "";
        shortRadarCols.forEach((colName, idx) => {
           allValues[startRowIdx + chartBaseRelRow][idx + 1] = colName;
           allBackgrounds[startRowIdx + chartBaseRelRow][idx + 1] = "#eee";
        });

        // データ行
        generations.forEach((gen, gIdx) => {
           const rowPos = startRowIdx + chartBaseRelRow + 1 + gIdx;
           
           const rawDate = gen[dateColIdx];
           let dateLabel = "回不明";
           if (rawDate instanceof Date) {
             dateLabel = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MM/dd");
           } else if (rawDate) {
             dateLabel = String(rawDate);
           }
           const genLabel = gIdx === 0 ? `最新(${dateLabel})` : (gIdx === 1 ? `前回(${dateLabel})` : `前々回(${dateLabel})`);
           
           allValues[rowPos][0] = genLabel;

           radarIndices.forEach((rIdx, rColPos) => {
             const vRaw = gen[rIdx];
             const vNum = Number(vRaw);
             const vClean = (typeof vRaw === 'string') ? vRaw.replace(/[\r\n]+/g, ' ') : vRaw;
             
             // 値セット
             const finalVal = isNaN(vNum) || String(vRaw).trim() === "" ? vClean : vNum;
             allValues[rowPos][rColPos + 1] = finalVal;

             // SOSハイライト (メモリ上)
             if (sosIdx !== -1 && sosValue && rIdx === sosIdx) {
               if (String(vRaw).includes(String(sosValue))) {
                 allFontColors[rowPos][rColPos + 1] = "red";
                 allFontWeights[rowPos][rColPos + 1] = "bold";
               }
             }
           });
        });

        // チャート予約
        const rRange = pSheet.getRange(currentRowNum + chartBaseRelRow, 1, 1 + generations.length, shortRadarCols.length + 1);
        chartQueue.push({
          type: "RADAR",
          range: rRange,
          posRow: currentRowNum + chartBaseRelRow,
          posCol: shortRadarCols.length + 2,
          title: `直近バランス推移`
        });
      }

      // --- C. 推移表 & 推移グラフ用データ ---
      if (comparePoints.length > 0 && myResponses.length > 0) {
        const trendBaseRelRow = chartBaseRelRow + 18; // 相対行 21
        const shortRadarColsForTrend = radarCols.map(c => c.length > 9 ? c.substring(0, 9) : c);
        
        // ヘッダー行
        allValues[startRowIdx + trendBaseRelRow][0] = timeUnit;
        shortRadarColsForTrend.forEach((c, idx) => {
          allValues[startRowIdx + trendBaseRelRow][idx + 1] = c;
          allBackgrounds[startRowIdx + trendBaseRelRow][idx + 1] = "#fafafa";
        });

        const dateFormat = timeUnit === "【月別】" ? "yyyy/MM" : (timeUnit === "【年別】" ? "yyyy" : "yyyy/MM/dd");
        let colSums = new Array(radarCols.length).fill(0);
        let colCounts = new Array(radarCols.length).fill(0);
        let validRowsCount = 0;

        comparePoints.forEach((pt, ptIdx) => {
          const matched = myResponses.filter(r => {
            const val = r[dateColIdx];
            // ★Fix: 文字列同士の比較を優先（設定パネルの「10月」とデータ側の「10月」を一致させる）
            const strVal = String(val).trim();
            const strPt = String(pt).trim();
            if (strVal === strPt) return true;

            // 日付比較フォールバック
            if (isDateMode) {
              const rd = new Date(val);
              if (!isNaN(rd)) {
                return Utilities.formatDate(rd, Session.getScriptTimeZone(), dateFormat) === pt;
              }
            }
            return false;
          });

          if (matched.length > 0) {
            const targetRow = matched[matched.length - 1];
            const rowPos = startRowIdx + trendBaseRelRow + 1 + validRowsCount;
            validRowsCount++;

            // ラベル
            let label = pt;
             if (timeUnit === "【日付別】" && targetRow[dateColIdx] instanceof Date) {
               const dObj = new Date(targetRow[dateColIdx]);
               label = `${dObj.getMonth() + 1}/${dObj.getDate()}`;
             }
            allValues[rowPos][0] = label;
            allBackgrounds[rowPos][0] = "#fafafa";

            radarIndices.forEach((rIdx, i) => {
              const vRaw = targetRow[rIdx];
              const vNum = Number(vRaw);
              
              if (!isNaN(vNum) && String(vRaw).trim() !== "") {
                colSums[i] += vNum;
                colCounts[i]++;
                allValues[rowPos][i + 1] = vNum;
              } else {
                const vClean = (typeof vRaw === 'string') ? vRaw.replace(/[\r\n]+/g, ' ') : vRaw;
                allValues[rowPos][i + 1] = vClean;
              }
              allBackgrounds[rowPos][i + 1] = "#fafafa";

              // SOS Check
              if (sosIdx !== -1 && sosValue && rIdx === sosIdx) {
                if (String(vRaw).includes(String(sosValue))) {
                   allFontColors[rowPos][i + 1] = "red";
                   allFontWeights[rowPos][i + 1] = "bold";
                }
              }
            });
          }
        });

        if (validRowsCount > 0) {
           // 平均行
           const avgRowPos = startRowIdx + trendBaseRelRow + 1 + validRowsCount;
           allValues[avgRowPos][0] = "平均";
           allBackgrounds[avgRowPos][0] = "#e6e6e6";
           allFontWeights[avgRowPos][0] = "bold";

           for(let i=0; i<radarCols.length; i++) {
             const val = colCounts[i] > 0 ? parseFloat((colSums[i] / colCounts[i]).toFixed(1)) : "-";
             allValues[avgRowPos][i + 1] = val;
             allBackgrounds[avgRowPos][i + 1] = "#e6e6e6";
             allFontWeights[avgRowPos][i + 1] = "bold";
           }

           // チャート予約
           const chartRange = pSheet.getRange(currentRowNum + trendBaseRelRow, 1, validRowsCount + 1, shortRadarColsForTrend.length + 1);
           chartQueue.push({
             type: "MULTI_LINE",
             range: chartRange,
             posRow: currentRowNum + trendBaseRelRow,
             posCol: shortRadarColsForTrend.length + 2,
             title: "パラメータ比較推移"
           });
        }
      }
      
      // 空白アンカー (ページ区切り用)
      const anchorRowIdx = startRowIdx + PAGE_BREAK_ROWS - 1;
      allValues[anchorRowIdx][0] = " ";
      
    }); // End Student Loop

    // --- 7. 一括書き込み (The Batch Write) ---
    if (printedCount > 0) {
      ss.toast("シートへの書き込みを開始します...", "出力中");
      
      const fullRange = pSheet.getRange(1, 1, totalRows, maxCols);
      
      // 値、背景、文字色、太字を一気に適用
      fullRange.setValues(allValues);
      fullRange.setBackgrounds(allBackgrounds);
      fullRange.setFontColors(allFontColors);
      fullRange.setFontWeights(allFontWeights);
      
      // 折り返し設定と行の高さ設定
      fullRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      pSheet.setRowHeights(1, totalRows, ROW_HEIGHT_PX);

      // 結合処理 (ここはループが必要だがAPIコールは軽い)
      allMerges.forEach(rng => rng.merge());

      // 枠線処理
      allBorders.forEach(b => {
        pSheet.getRange(b.r, b.c, b.h, b.w).setBorder(true, true, true, true, null, null, b.color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      });

      // 最後に一度だけFlush
      SpreadsheetApp.flush();

      // --- 8. チャート生成 (一括) ---
      ss.toast("グラフを生成しています...", "仕上げ");
      generatePersonalCharts_(pSheet, chartQueue);

      pSheet.activate();
      Browser.msgBox(`${printedCount}名分のカルテを高速作成しました。`);
    } else {
      Browser.msgBox("対象者が0名でした。");
    }

  } catch (e) {
    Browser.msgBox("⚠️ エラーが発生しました:\n" + e.message + "\n\n(設定を確認するか、管理者に問い合わせてください)");
    console.error(e.stack);
  }
}



// ==================================================
// 🏫 5. クラス集計 (Class Matrix & Chrono-Graph) ★Fix
// ==================================================

function runClassMatrixAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(CONFIG_SHEET_NAME);
  const master = ss.getSheetByName(MASTER_SHEET_NAME);

  try {
    if (!config) throw new Error("設定パネルが見つかりません。");

    // 1. 設定読み込み
    const targetSheetName = config.getRange("B3").getValue();
    const targetClass = config.getRange(SCHOOL_CONFIG_START_ROW + 1, 2).getValue();
    const ansKeyCol = config.getRange(SCHOOL_CONFIG_START_ROW + 2, 2).getValue();
    const dateColStr = config.getRange(SCHOOL_CONFIG_START_ROW + 3, 2).getValue();
    const sosColName = config.getRange(SCHOOL_CONFIG_START_ROW + 6, 2).getValue();
    const sosValue = config.getRange(SCHOOL_CONFIG_START_ROW + 7, 2).getValue();
    const timeUnit = config.getRange(SCHOOL_DATE_COMPARE_START_ROW - 1, 2).getValue();

    if (!targetClass || targetClass === "") { Browser.msgBox("対象クラスが選択されていません。"); return; }
    if (!ansKeyCol || String(ansKeyCol).startsWith("▼")) { Browser.msgBox("Key列が正しく設定されていません。"); return; }

    // 2. 比較ポイント取得 (B44:B55)
    // ★Fix: getDisplayValues() で見た目の文字列（"10月"など）をそのまま取得
    const comparePointsRaw = config.getRange(SCHOOL_DATE_COMPARE_START_ROW, 2, 12, 1).getDisplayValues().flat();
    const comparePoints = comparePointsRaw.filter(s => s !== "");

    const isDateMode = ["【日付別】", "【月別】", "【年別】"].includes(timeUnit);
    const fmt = timeUnit === "【月別】" ? "yyyy/MM" : (timeUnit === "【年別】" ? "yyyy" : "yyyy/MM/dd");

    const isTimelineMode = comparePoints.length > 0;
    const modeName = isTimelineMode ? "時系列比較モード" : "最新スナップショットモード";
    ss.toast(`${modeName}で集計中...`, "処理開始", 10);

    // 3. マスタ & 回答データ取得
    const masterData = master.getDataRange().getValues();
    const mGradeIdx = 1, mClassIdx = 2, mNumIdx = 3, mNameIdx = 4, mKeyIdx = 0;

    // 対象生徒抽出
    let targetStudents = [];
    if (targetClass.startsWith("(全学年)")) {
      const tClass = targetClass.replace("(全学年)", "");
      targetStudents = masterData.slice(1).filter(row => String(row[mClassIdx]) === tClass);
    } else {
      const match = targetClass.match(/^(.+)年(.+)組$/);
      if (match) {
        targetStudents = masterData.slice(1).filter(row => String(row[mGradeIdx]) === match[1] && String(row[mClassIdx]) === match[2]);
      }
    }
    targetStudents.sort((a, b) => Number(a[mNumIdx]) - Number(b[mNumIdx]));

    if (targetStudents.length === 0) { Browser.msgBox(`クラス「${targetClass}」の生徒が見つかりません。`); return; }

    const dataSheet = ss.getSheetByName(targetSheetName);
    const dHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    const allResponses = dataSheet.getDataRange().getValues().slice(1);

    // インデックス特定
    let ansKeyColIdx = -1;
    const kIdx = dHeaders.indexOf(ansKeyCol);
    if (kIdx > -1) ansKeyColIdx = kIdx; else ansKeyColIdx = letterToColumn_(ansKeyCol) - 1;

    let dateColIdx = 0;
    if (dateColStr && !String(dateColStr).startsWith("▼")) {
      const idx = dHeaders.indexOf(dateColStr);
      if (idx > -1) dateColIdx = idx; else dateColIdx = letterToColumn_(dateColStr) - 1;
    }
    let sosIdx = sosColName ? dHeaders.indexOf(sosColName) : -1;

    // 生徒ごとの回答マップ生成 (studentId -> [responses])
    let responseMap = {};
    allResponses.forEach(row => {
      const keyRaw = row[ansKeyColIdx];
      const key = keyRaw != null ? String(keyRaw).trim() : "";
      if (!key) return;
      if (!responseMap[key]) responseMap[key] = [];
      responseMap[key].push(row);
    });

    // 4. シート作成
    const resultSheetName = `🏫クラス集計_${targetClass}`;
    let cSheet = ss.getSheetByName(resultSheetName);
    if (cSheet) ss.deleteSheet(cSheet);
    cSheet = ss.insertSheet(resultSheetName);

    // 5. データ集計 & 描画準備
    const questionIndices = [];
    dHeaders.forEach((h, i) => {
      if (i === ansKeyColIdx || i === dateColIdx) return;
      if (/氏名|名前|出席番号|番号|クラス|学年|組|性別|Timestamp|タイムスタンプ/.test(h)) return;
      questionIndices.push({ index: i, title: h });
    });

    // Zone A (Dashboard): A-I (Width=9) -> B:Pie, D:Line, F-H:Table
    // Zone B (Divider): J (Width=1) -> Gray
    // Zone C (Matrix): K〜 (MatrixStart=11)
    const MATRIX_START_COL = 11; // K列
    const matrixEndCol = MATRIX_START_COL + targetStudents.length + 1;
    const graphDataStartCol = matrixEndCol + 2;

    // 必要な列数を計算して拡張
    const requiredCols = graphDataStartCol + (questionIndices.length * 3) + 10;
    if (cSheet.getMaxColumns() < requiredCols) {
      cSheet.insertColumnsAfter(cSheet.getMaxColumns(), requiredCols - cSheet.getMaxColumns());
    }

    // レイアウト設定 (Spacerとメイン列)
    cSheet.setColumnWidth(1, 20); // A (spacer)
    cSheet.setColumnWidth(2, 375); // B (Graphs - Wide)
    cSheet.setColumnWidth(3, 20); // C (spacer)
    cSheet.setColumnWidth(4, 20); // D (spacer)
    cSheet.setColumnWidth(5, 20); // E (spacer)
    cSheet.setColumnWidth(6, 80); // F (Table Date)
    cSheet.setColumnWidth(7, 50); // G (Table Avg)
    cSheet.setColumnWidth(8, 50); // H (Table Count)
    cSheet.setColumnWidth(9, 20); // I (Spacer)
    cSheet.setColumnWidth(10, 20); // J (Divider)
    cSheet.getRange("J:J").setBackground("#EFEFEF"); // Divider Color

    cSheet.getRange(1, 1).setValue(`🏫 クラス集計レポート: ${targetClass} (${modeName})`).setFontSize(14).setFontWeight("bold");

    // === 左側レーン: ダッシュボード (推移 & 合計 & 表) ===
    let graphCurrentRow = 3;
    const chartQueue = [];

    questionIndices.forEach((q, qIndex) => {
      const qTitle = q.title;
      const qIdx = q.index;

      const allValues = [];
      const trendData = []; // [{ label: "4/1", avg: 3.5, count: 30 }]

      const pointsToAnalyze = isTimelineMode ? comparePoints : ["最新"];

      pointsToAnalyze.forEach(pt => {
        let ptValues = [];
        targetStudents.forEach(stu => {
          const key = String(stu[mKeyIdx]).trim();
          const history = responseMap[key] || [];
          if (history.length === 0) return;

          let targetRow = null;
          if (isTimelineMode) {
            // ★Fix: ハイブリッド比較ロジック（文字列一致優先）
            targetRow = history.find(r => {
              const val = r[dateColIdx];
              const strVal = String(val).trim();
              const strPt = String(pt).trim();
              if (strVal === strPt) return true; // 文字列として一致

              if (isDateMode && val instanceof Date) {
                 return Utilities.formatDate(val, Session.getScriptTimeZone(), fmt) === pt;
              }
              return false;
            });
          } else {
            targetRow = history[history.length - 1]; // Latest
          }

          if (targetRow) {
            const v = targetRow[qIdx];
            if (v !== "" && v != null) {
              ptValues.push(v);
              allValues.push(v);
            }
          }
        });

        let numSum = 0, numCnt = 0;
        ptValues.forEach(v => {
          const n = parseFloat(v);
          if (!isNaN(n)) { numSum += n; numCnt++; }
        });
        const avg = numCnt > 0 ? (numSum / numCnt) : 0;
        trendData.push({ label: pt, avg: avg, count: ptValues.length });
      });

      // 質問タイトル
      const qBlockRow = graphCurrentRow;
      cSheet.getRange(qBlockRow, 2).setValue(`Q. ${qTitle}`).setFontWeight("bold");

      // Dynamic Column for Graph Data (Zone D)
      const hiddenColBase = graphDataStartCol + (qIndex * 3);

      // 1. 合計円グラフ (B列 上)
      const dist = {};
      allValues.forEach(v => { dist[v] = (dist[v]||0)+1; });
      const sortedDist = Object.keys(dist).sort((a,b)=>dist[b]-dist[a]);

      // グラフの高さを1.5倍(300px)として行数を計算 (23px/行 -> 約13行)
      const chartRows = 14;

      if (sortedDist.length > 0) {
        const hiddenRow = 1;
        const distData = [["Label", "Count"], ...sortedDist.map(k => [k, dist[k]])];
        // 計算用データエリアに枠線を付けて見やすく
        const dataRange = cSheet.getRange(hiddenRow, hiddenColBase, distData.length, 2);
        dataRange.setValues(distData).setBorder(true, true, true, true, true, true).setBackground("#FDFDFD");

        const pieRange = cSheet.getRange(hiddenRow, hiddenColBase, distData.length, 2);
        chartQueue.push({
          type: "PIE", range: pieRange,
          posRow: qBlockRow + 1, posCol: 2, // B列
          title: "期間合計構成比", width: 375, height: 300 // 1.5x Size
        });
      }

      // 2. 推移グラフ (B列 下) & 推移表 (F-H列)
      if (isTimelineMode) {
        const hiddenRow = 20;
        const trendRows = [["Point", "Average"], ...trendData.map(d => [d.label, d.avg])];
        cSheet.getRange(hiddenRow, hiddenColBase, trendRows.length, 2).setValues(trendRows);
        const lineRange = cSheet.getRange(hiddenRow, hiddenColBase, trendRows.length, 2);

        chartQueue.push({
          type: "LINE", range: lineRange,
          posRow: qBlockRow + 1 + chartRows, posCol: 2, // 円グラフの下に配置
          title: "平均値推移", width: 375, height: 300 // 1.5x Size
        });

        // 推移表 (F列)
        const tableHeader = [["集計日", "平均", "数"]];
        const tableBody = trendData.map(d => [d.label, d.avg.toFixed(2), d.count]);

        const tableRange = cSheet.getRange(qBlockRow + 1, 6, 1 + tableBody.length, 3); // F列(6)
        tableRange.setValues([...tableHeader, ...tableBody]);
        tableRange.setBorder(true, true, true, true, true, true).setFontSize(8).setHorizontalAlignment("center");
        cSheet.getRange(qBlockRow + 1, 6, 1, 3).setBackground("#E0E0E0").setFontWeight("bold");
      }

      // 次のブロックまでの間隔 (タイトル1 + 円グラフ14 + 折れ線14 + 余白2)
      graphCurrentRow += 1 + chartRows + chartRows + 2;
    });


    // === 右側レーン: 生徒別マトリクス (ブロック積み上げ) ===
    let matrixCurrentRow = 3;
    const pointsToRender = isTimelineMode ? comparePoints : ["【最新の回答状況】"];

    pointsToRender.forEach(pt => {
      const headerLabel = isTimelineMode ? `📅 ${pt} の記録` : pt;
      cSheet.getRange(matrixCurrentRow, MATRIX_START_COL).setValue(headerLabel)
        .setFontSize(11).setFontWeight("bold").setBackground("#34A853").setFontColor("white");
      cSheet.getRange(matrixCurrentRow, MATRIX_START_COL, 1, targetStudents.length + 1).merge();
      matrixCurrentRow++;

      // 生徒名の標準表示
      const stuNames = targetStudents.map(s => `${s[mNumIdx]}.${s[mNameIdx]}`);

      const matrixHeader = ["質問項目", ...stuNames];
      const headerRange = cSheet.getRange(matrixCurrentRow, MATRIX_START_COL, 1, matrixHeader.length);
      headerRange.setValues([matrixHeader])
        .setBackground("#E6F4EA").setFontWeight("bold").setBorder(true, true, true, true, true, true)
        .setVerticalAlignment("top")
        .setFontSize(9);

      // 回転なし(0°)を明示
      cSheet.getRange(matrixCurrentRow, MATRIX_START_COL + 1, 1, targetStudents.length).setTextRotation(0);
      // 列幅を 50px に拡張
      cSheet.setColumnWidths(MATRIX_START_COL + 1, targetStudents.length, 50);

      matrixCurrentRow++;

      const matrixRows = [];
      const sosCoords = [];

      questionIndices.forEach((q, qRowIdx) => {
        const rowData = [q.title];

        targetStudents.forEach((stu, stuIdx) => {
          const key = String(stu[mKeyIdx]).trim();
          const history = responseMap[key] || [];
          let val = "-";

          if (history.length > 0) {
            let targetRow = null;
            if (isTimelineMode) {
              // ★Fix: ハイブリッド比較ロジック（マトリクス用）
              targetRow = history.find(r => {
                 const v = r[dateColIdx];
                 const strVal = String(v).trim();
                 const strPt = String(pt).trim();
                 if (strVal === strPt) return true;

                 if (isDateMode && v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), fmt) === pt;
                 return false;
              });
            } else {
              targetRow = history[history.length - 1];
            }

            if (targetRow) {
              val = targetRow[q.index];
              if (sosIdx !== -1 && sosValue && q.index === sosIdx) {
                if (String(val).includes(String(sosValue))) {
                  sosCoords.push({ r: qRowIdx, c: stuIdx + 1 });
                }
              }
            }
          }
          rowData.push(val);
        });
        matrixRows.push(rowData);
      });

      if (matrixRows.length > 0) {
        const r = cSheet.getRange(matrixCurrentRow, MATRIX_START_COL, matrixRows.length, matrixHeader.length);
        r.setValues(matrixRows).setBorder(true, true, true, true, true, true);

        // 回答エリアの書式設定 (Clip, Left, Middle)
        r.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
         .setHorizontalAlignment("left")
         .setVerticalAlignment("middle");

        if (matrixRows.length > 1) r.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
        // 1. デフォルトのスタイル定義（黒文字・背景なし）
        const numRows = matrixRows.length;
        const numCols = matrixHeader.length;
        const fontColors = Array(numRows).fill(null).map(() => Array(numCols).fill("black"));
        const fontWeights = Array(numRows).fill(null).map(() => Array(numCols).fill("normal"));
        const backgrounds = Array(numRows).fill(null).map(() => Array(numCols).fill(null));

        // 2. SOS座標の箇所だけスタイルを上書き
        sosCoords.forEach(coord => {
          if (coord.r < numRows && coord.c < numCols) {
            fontColors[coord.r][coord.c] = "red";
            fontWeights[coord.r][coord.c] = "bold";
            backgrounds[coord.r][coord.c] = "#FFCCCC";
          }
        });

        // 3. 対象範囲を取得してAPIを叩いて一括適用
        const targetRange = cSheet.getRange(matrixCurrentRow, MATRIX_START_COL, numRows, numCols);
        targetRange.setFontColors(fontColors);
        targetRange.setFontWeights(fontWeights);
        targetRange.setBackgrounds(backgrounds);

        matrixCurrentRow += matrixRows.length;
      }
      matrixCurrentRow += 2;
    });

// =================================================================
    // ★機能追加: 項目別・時系列変化マトリクス (Item-Centric Evolution)
    // 概要: 比較ポイントがある場合のみ、項目を主軸にした時系列表を追加出力
    // =================================================================
    if (isTimelineMode && comparePoints.length > 0) {
      // 1. 比較対象リスト構築 (最新 + 過去)
      const chronoPoints = [
        { label: "今回 (最新)", val: "LATEST" }, // マーカー
        ...comparePoints.map(p => ({ label: p, val: p }))
      ];

      // 2. セクション区切り
      matrixCurrentRow += 3;
      cSheet.getRange(matrixCurrentRow, MATRIX_START_COL).setValue("▼ 項目別 時系列変化 (Item Evolution Mode)");
      cSheet.getRange(matrixCurrentRow, MATRIX_START_COL, 1, targetStudents.length + 1)
            .setBackground("#4285F4") // Google Blue
            .setFontColor("white")
            .setFontWeight("bold");
      matrixCurrentRow += 2;

      // 3. 質問項目ごとにループ
      questionIndices.forEach(q => {
        // 見出し (Q. 質問文)
        cSheet.getRange(matrixCurrentRow, MATRIX_START_COL).setValue(`Q. ${q.title}`);
        cSheet.getRange(matrixCurrentRow, MATRIX_START_COL)
              .setFontWeight("bold")
              .setFontColor("#1a73e8")
              .setFontSize(10);
        matrixCurrentRow++;

        // データ準備
        const tableData = [];
        const sosHighlightCoords = []; // SOSハイライト用座標

        // [A] ヘッダー行: [ "時期", 生徒名... ]
        const tableHeader = ["時期"];
        targetStudents.forEach(s => tableHeader.push(`${s[mNumIdx]}.${s[mNameIdx]}`));
        tableData.push(tableHeader);

        // [B] データ行: 時期ごとにループ
        chronoPoints.forEach((point, pIdx) => {
          const row = [point.label];

          targetStudents.forEach((stu, sIdx) => {
            const key = String(stu[mKeyIdx]).trim();
            const history = responseMap[key] || [];
            
            let val = "-";
            let targetRow = null;

            if (history.length > 0) {
              if (point.val === "LATEST") {
                targetRow = history[history.length - 1];
              } else {
                // ハイブリッド比較ロジック (既存処理を流用)
                targetRow = history.find(r => {
                  const v = r[dateColIdx];
                  const strVal = String(v).trim();
                  const strPt = String(point.val).trim();
                  if (strVal === strPt) return true;
                  if (isDateMode && v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), fmt) === point.val;
                  return false;
                });
              }

              if (targetRow) {
                val = targetRow[q.index];
                if (val === "" || val == null) val = " - ";

                // SOS判定 (該当する場合、座標を記憶)
                if (sosIdx !== -1 && sosValue && q.index === sosIdx) {
                  if (String(val).includes(String(sosValue))) {
                    // ヘッダーが1行あるので +1
                    sosHighlightCoords.push({ r: pIdx + 1, c: sIdx + 1 });
                  }
                }
              }
            }
            row.push(val);
          });
          tableData.push(row);
        });

        // [C] 書き込み
        if (tableData.length > 0) {
          const numRows = tableData.length;
          const numCols = tableData[0].length;
          const range = cSheet.getRange(matrixCurrentRow, MATRIX_START_COL, numRows, numCols);
          
          range.setValues(tableData)
               .setBorder(true, true, true, true, true, true)
               .setVerticalAlignment("middle")
               .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // はみ出し防止

          // スタイル: ヘッダー
          cSheet.getRange(matrixCurrentRow, MATRIX_START_COL, 1, numCols)
                .setBackground("#E8F0FE").setFontWeight("bold").setHorizontalAlignment("center");
          // スタイル: 左端(時期)
          cSheet.getRange(matrixCurrentRow + 1, MATRIX_START_COL, numRows - 1, 1)
                .setBackground("#F1F3F4").setFontWeight("bold");

          // [D] SOSハイライト適用
          if (sosHighlightCoords.length > 0) {
            const fontColors = range.getFontColors();
            const fontWeights = range.getFontWeights();
            const bgColors = range.getBackgrounds();
            
            sosHighlightCoords.forEach(coord => {
              if(coord.r < numRows && coord.c < numCols) {
                fontColors[coord.r][coord.c] = "red";
                fontWeights[coord.r][coord.c] = "bold";
                bgColors[coord.r][coord.c] = "#FFCCCC";
              }
            });
            range.setFontColors(fontColors).setFontWeights(fontWeights).setBackgrounds(bgColors);
          }

          matrixCurrentRow += numRows + 1; // 間隔
        }
      });
      
      // 最後に余白
      matrixCurrentRow += 1;
    }

    chartQueue.forEach(cq => {
      let builder = cSheet.newChart()
        .addRange(cq.range)
        .setOption('title', cq.title)
        .setPosition(cq.posRow, cq.posCol, 0, 0)
        .setOption('width', cq.width)
        .setOption('height', cq.height);

      if (cq.type === "PIE") builder = builder.setChartType(Charts.ChartType.PIE);
      if (cq.type === "LINE") builder = builder.setChartType(Charts.ChartType.LINE).setOption('legend', {position: 'bottom'});
      cSheet.insertChart(builder.build());
    });

    // 仕上げ: K列(質問文)は折り返し
    cSheet.getRange("K:K").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    cSheet.setColumnWidth(MATRIX_START_COL, 200);

    Browser.msgBox(`クラス集計が完了しました。\nシート: ${resultSheetName}\n※生徒名の縦書き設定は、必要に応じて手動で行ってください。`);

  } catch (e) {
    Browser.msgBox("⚠️ クラス集計エラー: " + e.message);
    console.error(e.stack);
  }
}

// ==================================================
// 🏫 6. 全校集計実行 (All School Analysis) ★Fix
// ==================================================

function runAllSchoolAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(CONFIG_SHEET_NAME);
  const master = ss.getSheetByName(MASTER_SHEET_NAME);

  try {
    // --- 1. 設定読み込み ---
    if (!config) throw new Error("設定パネルが見つかりません。");
    if (!master) throw new Error("名簿マスタが見つかりません。");

    const targetSheetName = config.getRange("B3").getValue();
    const ansKeyCol = config.getRange(SCHOOL_CONFIG_START_ROW + 2, 2).getValue();
    const dateColStr = config.getRange(SCHOOL_CONFIG_START_ROW + 3, 2).getValue();
    const timeUnit = config.getRange(SCHOOL_DATE_COMPARE_START_ROW - 1, 2).getValue();

    // ★v10.45: SOS設定読み込み
    const sosColName = config.getRange(SCHOOL_CONFIG_START_ROW + 6, 2).getValue(); // Row 31
    const sosWord = config.getRange(SCHOOL_CONFIG_START_ROW + 7, 2).getValue(); // Row 32

    if (!ansKeyCol || String(ansKeyCol).startsWith("▼")) {
      Browser.msgBox("⚠️ 設定エラー: 「Key列」が正しく設定されていません。");
      return;
    }

    const isDateMode = ["【日付別】", "【月別】", "【年別】"].includes(timeUnit);
    let timePoints = [];
    
    // 比較ポイント取得 (B44～B55)
    for (let i = 0; i < 12; i++) {
      const d = config.getRange(SCHOOL_DATE_COMPARE_START_ROW + i, 2).getValue();
      if (d) {
         if (d instanceof Date && isDateMode) {
             const fmt = timeUnit === "【月別】" ? "yyyy/MM" : (timeUnit === "【年別】" ? "yyyy" : "yyyy/MM/dd");
             timePoints.push(Utilities.formatDate(d, Session.getScriptTimeZone(), fmt));
         } else {
             timePoints.push(String(d).trim());
         }
      }
    }

    // 期間設定がない場合、「全期間のみ」モードとして動作
    let isAllTimeMode = false;
    if (timePoints.length === 0) {
       isAllTimeMode = true;
    }

    // --- 2. データ準備 (クラスリスト作成: v10.38 Special Class Logic) ---
    const masterData = master.getDataRange().getValues(); 
    // ★v10.44: 氏名インデックス(mNameIdx=4)を追加
    const mKeyIdx = 0, mGradeIdx = 1, mClassIdx = 2, mNameIdx = 4;

    const studentClassMap = new Map();
    // ★v10.44: 氏名マップを追加
    const studentNameMap = new Map();
    const classSet = new Set();

    masterData.slice(1).forEach(row => {
      const sKey = String(row[mKeyIdx]).trim();
      const sGrade = row[mGradeIdx];
      const sClass = String(row[mClassIdx]).trim();
      // ★v10.44: 氏名取得
      const sName = String(row[mNameIdx] || "").trim();

      if (!sKey || !sGrade || !sClass) return;

      let classLabel = "";
      const isStandardClass = !isNaN(sClass) || /^[A-Z]$|^[IVX]+$/i.test(sClass);
      
      if (isStandardClass) {
          classLabel = `${sGrade}年${sClass}組`;
      } else {
          if (sClass.startsWith("(全学年)")) {
              classLabel = sClass;
          } else {
              classLabel = `(全学年)${sClass}`;
          }
      }
      
      studentClassMap.set(sKey, classLabel);
      // ★v10.44: マップ登録
      studentNameMap.set(sKey, sName);
      classSet.add(classLabel);
    });

    const sortedClasses = Array.from(classSet).sort(); 

    // --- 3. 回答データ取得 ---
    const dataSheet = ss.getSheetByName(targetSheetName);
    if (!dataSheet) throw new Error("回答シートが見つかりません。");
    const dHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    const allResponses = dataSheet.getDataRange().getValues().slice(1);

    let ansKeyColIdx = -1;
    const kIdx = dHeaders.indexOf(ansKeyCol);
    if (kIdx > -1) ansKeyColIdx = kIdx;
    else ansKeyColIdx = letterToColumn_(ansKeyCol) - 1;

    let dateColIdx = 0;
    if (dateColStr && !String(dateColStr).startsWith("▼")) {
        const idx = dHeaders.indexOf(dateColStr);
        if (idx > -1) dateColIdx = idx;
        else dateColIdx = letterToColumn_(dateColStr) - 1; 
    }

    const dateFormat = timeUnit === "【月別】" ? "yyyy/MM" : (timeUnit === "【年別】" ? "yyyy" : "yyyy/MM/dd");

    // --- 4. 出力シート準備 ---
    let reportSheet = ss.getSheetByName(ALL_SCHOOL_SHEET_NAME);
    if (reportSheet) ss.deleteSheet(reportSheet);
    reportSheet = ss.insertSheet(ALL_SCHOOL_SHEET_NAME);

    // 記述回答シート
    let textSheet = ss.getSheetByName(TEXT_SHEET_NAME);
    if (textSheet) textSheet.clear();
    else textSheet = ss.insertSheet(TEXT_SHEET_NAME);
    textSheet.setTabColor("orange");
    textSheet.getRange(1, 1).setValue("📝 全校・記述回答まとめ").setFontSize(14).setFontWeight("bold");
    let textSheetCol = 1;

    // ★v10.39: 総合グラフ用データ保持配列
    // [{qTitle: "Q1...", averages: { "4月": 4.5, "5月": 4.2... }}]
    const globalTrendData = [];

    // ★v10.44: 実行結果カウンタ
    let countNumericTable = 0; // 数値表が作成された数
    let countTextOnly = 0;     // 記述回答まとめになった数

    // --- 5. 集計 & 出力ループ ---
    let currentOutputRow = 1;
    
    // タイトル
    const modeTitle = isAllTimeMode ? "全期間平均のみ" : "時系列マトリクス";
    reportSheet.getRange(currentOutputRow, 1).setValue(`🏫 全校集計レポート (${modeTitle})`)
      .setFontSize(14).setFontWeight("bold").setFontColor("#34A853");
    currentOutputRow += 2;

    // 質問ごとにループ
    for (let col = 1; col < dHeaders.length; col++) {
      const qTitle = dHeaders[col];
      if (!qTitle) continue;
      
      // Smart Column Filter
      if (/学年|組|クラス|番号|出席番号|氏名|名前|Name|ID|Email|メール/i.test(qTitle)) {
          continue;
      }

      // 数値判定
      let numericCount = 0;
      let totalCount = 0;
      const validResponses = []; // {class, val, time, key}

      allResponses.forEach(row => {
         const sKey = String(row[ansKeyColIdx]).trim();
         const sClass = studentClassMap.get(sKey);
         if (!sClass) return; 

         const v = row[col];
         if (v !== "" && v != null) {
            totalCount++;
            
            // ★Fix: 厳格な数値判定 (単位付き数値を弾く)
            // 以前: if (!isNaN(parseFloat(v))) numericCount++;
            // 変更: Number()を使用して "4回" などをNaNとして扱う
            const vStr = String(v).trim();
            if (!isNaN(Number(vStr)) && vStr !== "") {
                numericCount++;
            }
            
            // 時期判定
            let timeLabel = "ALL"; // Default
            if (!isAllTimeMode) {
                const rDateVal = row[dateColIdx];
                if (isDateMode) {
                    const rd = new Date(rDateVal);
                    if (!isNaN(rd)) {
                        timeLabel = Utilities.formatDate(rd, Session.getScriptTimeZone(), dateFormat);
                    }
                } else {
                    timeLabel = String(rDateVal).trim();
                }
            }
            
            validResponses.push({ cls: sClass, val: v, time: timeLabel, key: sKey });
         }
      });
      
      if (totalCount === 0) continue;

      // 記述式分岐 (数値回答率8割未満)
      if ((numericCount / totalCount) < 0.8) {
         // ★v10.44: カウントアップ
         countTextOnly++;

         // テキストまとめ出力
         textSheet.getRange(3, textSheetCol).setValue(`Q. ${qTitle}`)
           .setFontWeight("bold").setBackground("#f3f3f3").setBorder(true, true, true, true, null, null);
         
         // ★v10.44: 氏名付きフォーマットに変更 [1年1組 相川 翔] ...
         const textRows = validResponses.map(r => {
             const sName = studentNameMap.get(r.key) || "";
             return [`[${r.cls} ${sName}] ${r.val}`];
         });

         if (textRows.length > 0) {
             const targetRange = textSheet.getRange(4, textSheetCol, textRows.length, 1);
             targetRange.setValues(textRows);

             // --- ★v10.45: SOSハイライト処理 ---
             if (sosColName && sosWord && qTitle === sosColName) {
                 const sosKeyword = String(sosWord).trim();
                 if (sosKeyword) {
                     // ハイライト用配列作成
                     const fontColors = [];
                     const fontWeights = [];
                     
                     textRows.forEach(row => {
                         const cellText = String(row[0]);
                         if (cellText.includes(sosKeyword)) {
                             fontColors.push(["red"]);
                             fontWeights.push(["bold"]);
                         } else {
                             fontColors.push(["black"]);
                             fontWeights.push(["normal"]);
                         }
                     });
                     
                     // 一括適用
                     targetRange.setFontColors(fontColors).setFontWeights(fontWeights);
                 }
             }
             // ------------------------------------
         }
         textSheet.setColumnWidth(textSheetCol, 300);
         textSheetCol++;
         
         // ★Fix: 全校集計レポート側にも「移動案内」看板を設置
         reportSheet.getRange(currentOutputRow, 1).setValue(`Q. ${qTitle}`)
           .setFontWeight("bold").setFontColor("#333333");
         currentOutputRow++;
         
         reportSheet.getRange(currentOutputRow, 1).setValue("➡ この項目の回答（回数・年月・自由記述等）は「📝記述回答まとめ」シートに集約しました。")
           .setFontSize(10).setFontColor("gray").setFontStyle("italic");
         
         currentOutputRow += 2; // 行間を空ける

         continue; 
      }

      // --- マトリクスデータ生成 (数値のみ) ---
      // ★v10.44: カウントアップ
      countNumericTable++;

      const matrix = {};
      const allStats = {}; 
      
      sortedClasses.forEach(cls => {
         matrix[cls] = {};
         if (!isAllTimeMode) {
             timePoints.forEach(tp => matrix[cls][tp] = {sum: 0, count: 0});
         }
         matrix[cls]['ALL_TOTAL'] = {sum: 0, count: 0};
      });

      if (!isAllTimeMode) {
          timePoints.forEach(tp => allStats[tp] = {sum: 0, count: 0});
      }
      allStats['ALL_TOTAL'] = {sum: 0, count: 0};

      validResponses.forEach(d => {
         const valNum = parseFloat(d.val);
         if (isNaN(valNum)) return;

         if (matrix[d.cls]) {
             if (!isAllTimeMode && matrix[d.cls][d.time]) {
                 matrix[d.cls][d.time].sum += valNum;
                 matrix[d.cls][d.time].count++;
             }
             matrix[d.cls]['ALL_TOTAL'].sum += valNum;
             matrix[d.cls]['ALL_TOTAL'].count++;
         }

         if (!isAllTimeMode && allStats[d.time]) {
             allStats[d.time].sum += valNum;
             allStats[d.time].count++;
         }
         allStats['ALL_TOTAL'].sum += valNum;
         allStats['ALL_TOTAL'].count++;
      });

      // --- ★v10.39: 総合グラフ用データ収集 ---
      if (!isAllTimeMode) {
          const averages = {};
          timePoints.forEach(tp => {
              const s = allStats[tp];
              averages[tp] = s.count > 0 ? parseFloat((s.sum / s.count).toFixed(2)) : null;
          });
          globalTrendData.push({ title: qTitle, averages: averages });
      }

      // --- 表書き出し ---
      reportSheet.getRange(currentOutputRow, 1).setValue(`Q.${qTitle}`).setFontWeight("bold");
      currentOutputRow++;
      
      let tableHeader = ["クラス名"];
      if (!isAllTimeMode) {
          tableHeader = [...tableHeader, ...timePoints, "全期間平均"];
      } else {
          tableHeader.push("全期間平均");
      }

      reportSheet.getRange(currentOutputRow, 1, 1, tableHeader.length)
                 .setValues([tableHeader])
                 .setBackground("#E8F0FE").setFontWeight("bold").setBorder(true, true, true, true, true, true);
      
      currentOutputRow++;
      const startTableBodyRow = currentOutputRow;

      // 全校平均行
      const allRowVals = ["🏫 全校平均"];
      if (!isAllTimeMode) {
          timePoints.forEach(tp => {
              const s = allStats[tp];
              allRowVals.push(s.count > 0 ? (s.sum / s.count).toFixed(2) : "-");
          });
      }
      const allS = allStats['ALL_TOTAL'];
      allRowVals.push(allS.count > 0 ? (allS.sum / allS.count).toFixed(2) : "-");
      
      reportSheet.getRange(currentOutputRow, 1, 1, allRowVals.length).setValues([allRowVals])
                 .setFontWeight("bold").setBackground("#FFF2CC");
      currentOutputRow++;

      // クラス行
      const classRows = [];
      sortedClasses.forEach(cls => {
          const rowVals = [cls];
          if (!isAllTimeMode) {
              timePoints.forEach(tp => {
                  const d = matrix[cls][tp];
                  rowVals.push(d.count > 0 ? (d.sum / d.count).toFixed(2) : "-");
              });
          }
          const totalD = matrix[cls]['ALL_TOTAL'];
          rowVals.push(totalD.count > 0 ? (totalD.sum / totalD.count).toFixed(2) : "-");
          classRows.push(rowVals);
      });

      if (classRows.length > 0) {
          reportSheet.getRange(currentOutputRow, 1, classRows.length, classRows[0].length)
                     .setValues(classRows);
          
          reportSheet.getRange(startTableBodyRow - 1, 1, classRows.length + 2, tableHeader.length)
                     .setBorder(true, true, true, true, true, true);
          
          currentOutputRow += classRows.length;
      }

      currentOutputRow += 2; 
    }

    reportSheet.autoResizeColumns(1, 15);
    reportSheet.setColumnWidth(1, 150); 

    // --- ★v10.39: 総合サマリ表 & 大型グラフ生成 ---
    if (!isAllTimeMode && globalTrendData.length > 0) {
        currentOutputRow += 2;
        // 区切り線
        reportSheet.getRange(currentOutputRow, 1, 1, 10).setBorder(true, null, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        currentOutputRow++;

        reportSheet.getRange(currentOutputRow, 1).setValue("📈 総合推移サマリ (全項目平均)")
            .setFontSize(14).setFontWeight("bold").setFontColor("#E91E63");
        currentOutputRow += 2;

        // グラフ出力位置確保 (グラフ用に20行ほど空ける)
        const chartPositionRow = currentOutputRow;
        currentOutputRow += 25; 

        // サマリ表ヘッダー
        const summaryHeader = ["項目名(質問)", ...timePoints];
        reportSheet.getRange(currentOutputRow, 1, 1, summaryHeader.length)
            .setValues([summaryHeader])
            .setBackground("#FCE8E6").setFontWeight("bold").setBorder(true, true, true, true, true, true);
        currentOutputRow++;

        const startSummaryRow = currentOutputRow;
        const summaryRows = [];

        globalTrendData.forEach(item => {
            const rowVals = [item.title];
            timePoints.forEach(tp => {
                rowVals.push(item.averages[tp] !== null ? item.averages[tp] : "");
            });
            summaryRows.push(rowVals);
        });

        if (summaryRows.length > 0) {
            const range = reportSheet.getRange(currentOutputRow, 1, summaryRows.length, summaryRows[0].length);
            range.setValues(summaryRows).setBorder(true, true, true, true, true, true);
            
            // グラフ生成 (横長)
            // データ範囲: ヘッダー行 + データ行
            const chartDataRange = reportSheet.getRange(startSummaryRow - 1, 1, summaryRows.length + 1, summaryHeader.length);
            
            const bigChart = reportSheet.newChart()
                .setChartType(Charts.ChartType.LINE)
                .addRange(chartDataRange)
                .setPosition(chartPositionRow, 1, 0, 0)
                .setOption('title', '全校・総合コンディション推移 (全項目平均)')
                .setOption('width', 1000) 
                .setOption('height', 450)
                // ★修正: 行列入れ替えとテキストラベル強制
                .setTransposeRowsAndColumns(true) 
                .setOption('treatLabelsAsText', true) 
                .setOption('useFirstColumnAsDomain', true)
                // ----------------------------------
                .setNumHeaders(1)
                .setOption('legend', {position: 'right'})
                .setOption('hAxis', {title: '時期'})
                .setOption('vAxis', {title: '平均スコア'})
                .build();
            
            reportSheet.insertChart(bigChart);
        }
    }
    
    // ★v10.44: 結果案内分岐
    if (countNumericTable === 0 && countTextOnly > 0) {
        Browser.msgBox("⚠️ 推移レポートは作成されませんでした\n\n" +
            "抽出されたデータがすべて文字列（選択式など）だったため、平均値の推移表は作成されませんでした。\n" +
            "回答内容は『📝記述回答まとめ』シートに全て反映されていますので、そちらをご確認ください。");
    } else {
        ss.toast("全校集計レポートを作成しました。", "完了", 5);
        reportSheet.activate();
    }

  } catch (e) {
    Browser.msgBox("⚠️ 全校集計中にエラーが発生しました:\n" + e.message);
    console.error(e.stack);
  }
}



// ==================================================
// 🧩 Helper: All Functions (Reorganized)
// ==================================================

function letterToColumn_(letter) {
  if (!letter || letter === "") return -1;
  let column = 0, length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function columnToLetter_(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// 
function updateQuestionDropdowns_(configSheet) {
  try {
    const targetSheetName = configSheet.getRange("B3").getValue();
    if (!targetSheetName) return;
    
    const ss = configSheet.getParent();
    const targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) return;
    
    const headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    const lastRow = targetSheet.getLastRow();
    let columnTypes = [];
    if (lastRow > 1) {
      const sampleData = targetSheet.getRange(2, 1, Math.min(lastRow - 1, 50), headers.length).getValues();
      for (let c = 0; c < headers.length; c++) {
        const colVals = sampleData.map(r => r[c]);
        columnTypes[c] = analyzeColumnType_(colVals, headers[c]);
      }
    } else {
       columnTypes = new Array(headers.length).fill('CATEGORY');
    }

    // 現在選択されている値を全て取得（排他制御用）
    const currentSelections = {};
    [FILTER_ROW_A, FILTER_ROW_B, FILTER_ROW_C, CROSS_AXIS_LABEL_ROW].forEach(r => {
      const val = configSheet.getRange(r, 2).getValue();
      if (val && !String(val).startsWith("▼")) {
        currentSelections[r] = val;
      }
    });

    // レーダー項目の現在の選択
    const radarStart = SCHOOL_CONFIG_START_ROW + 10;
    const currentRadarSelections = {};
    for (let i = 0; i < 8; i++) {
      const row = radarStart + i;
      const val = configSheet.getRange(row, 2).getValue();
      if (val) currentRadarSelections[row] = val;
    }

    // ★修正: allowTimestamp引数を削除し、常にTIMESTAMPを含めるように変更
    const setupFilterDropdown = (targetRow, allowNumberSkip) => {
      let candidates = [];
      headers.forEach((h, i) => {
        // SKIPタイプ（IDや個人名など）の処理
        if (columnTypes[i] === 'SKIP') {
            if (allowNumberSkip && /番号|出席番号|No\.|ナンバー|number|ID/i.test(h)) {
                candidates.push(h);
            }
        } 
        // 記述回答(FREE_TEXT)以外は追加（TIMESTAMPもCATEGORYもここに含まれる）
        else if (columnTypes[i] !== 'FREE_TEXT') {
            candidates.push(h);
        }
      });

      // ★修正: B17(横軸)での「記述回答」除外のみ残し、タイムスタンプ除外ロジックを削除
      if (targetRow === CROSS_AXIS_LABEL_ROW) {
         headers.forEach((h, i) => {
             if (columnTypes[i] === 'FREE_TEXT') {
                 const idx = candidates.indexOf(h);
                 if (idx > -1) candidates.splice(idx, 1);
             }
         });
         // ※ここで以前あった「タイムスタンプ除外」コードを削除しました
      }

      // 排他制御: 他のフィルタや軸で選ばれている項目を除外
      const others = Object.keys(currentSelections)
        .filter(r => Number(r) !== targetRow)
        .map(r => currentSelections[r]);
      candidates = candidates.filter(h => !others.includes(h));

      if (candidates.length > 0) {
        const rule = SpreadsheetApp.newDataValidation()
             .requireValueInList(candidates)
             .setAllowInvalid(true)
             .setHelpText("リストから選択するか、空白のままにしてください。")
             .build();
        configSheet.getRange(targetRow, 2).setDataValidation(rule);
      } else {
        configSheet.getRange(targetRow, 2).clearDataValidations().setNote("⚠️ 選択可能な項目がありません");
      }
    };

    const setupSchoolDropdown = (targetRows, isRadar = false, isDate = false) => {
      let candidates = [];
      headers.forEach((h, i) => {
          candidates.push(h); 
      });
      if (isDate) {
        const datePattern = /日付|日時|Date|Time|Timestamp|タイムスタンプ|年月|年|月|回/i;
        candidates = candidates.filter(h => datePattern.test(h));
      } else if (isRadar) {
         const excludePattern = /氏名|名前|出席番号|番号|ID|Key|Email|メール|mail|address|Timestamp|タイムスタンプ|日付|Date|Time|学年|組|クラス|性別|Gender|作成|感想|自由|記述|コメント/i;
         candidates = candidates.filter(h => !excludePattern.test(h));
      }

      const ruleBuilder = SpreadsheetApp.newDataValidation();
      targetRows.forEach(r => {
        let myCandidates = [...candidates];
        if (isRadar) {
           const others = Object.keys(currentRadarSelections)
             .filter(rowKey => Number(rowKey) !== r)
             .map(rowKey => currentRadarSelections[rowKey]);
           myCandidates = myCandidates.filter(h => !others.includes(h));
        }
        
        if (myCandidates.length > 0) {
          const rule = ruleBuilder.requireValueInList(myCandidates)
             .setAllowInvalid(true)
             .setHelpText("リストから選択するか、直接入力してください。")
             .build();
          configSheet.getRange(r, 2).setDataValidation(rule);
        } else {
          configSheet.getRange(r, 2).clearDataValidations().setNote("⚠️ 候補なし");
        }
      });
    };

    // ★修正: 引数を減らして呼び出し（タイムスタンプは常に許可されるため）
    // フィルタ設定 (B7, B10, B13)
    setupFilterDropdown(FILTER_ROW_A, false); 
    setupFilterDropdown(FILTER_ROW_B, false); 
    setupFilterDropdown(FILTER_ROW_C, false);
    // クロス集計軸 (B17)
    setupFilterDropdown(CROSS_AXIS_LABEL_ROW, true); 

    const schoolTargetRows = [];
    schoolTargetRows.push(SCHOOL_CONFIG_START_ROW + 6);
    // SOS (31)
    setupSchoolDropdown(schoolTargetRows, false); 
    
    // Radar 1-8 (35-42)
    const radarTargetRows = [];
    for(let k=0; k<8; k++) radarTargetRows.push(SCHOOL_CONFIG_START_ROW + 10 + k); 
    setupSchoolDropdown(radarTargetRows, true, false);
    // Date Col (28)
    const dateRow = SCHOOL_CONFIG_START_ROW + 3;
    setupSchoolDropdown([dateRow], false, true);

  } catch (err) {
    console.error("updateQuestionDropdowns_ Error: " + err.message);
  }
}


// ★日付選択プルダウン更新 (v10.24: 型一致ロジック追加)
function updateDateDropdown_(configSheet) {
  try {
    const ss = configSheet.getParent();
    const targetSheetName = configSheet.getRange("B3").getValue();
    if (!targetSheetName) return;

    const targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) return;

    const lastRow = targetSheet.getLastRow();
    if (lastRow < 2) return;

    // 設定された日付列を取得 (行28)
    const dateColName = configSheet.getRange(SCHOOL_CONFIG_START_ROW + 3, 2).getValue(); 
    // 単位を取得 (行43)
    const timeUnitCell = configSheet.getRange(SCHOOL_DATE_COMPARE_START_ROW - 1, 2);
    let timeUnit = timeUnitCell.getValue();

    let dateColIdx = 0; // default A
    const headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    if (dateColName && !String(dateColName).startsWith("▼")) {
        const idx = headers.indexOf(dateColName);
        if (idx > -1) dateColIdx = idx;
    }

    const rawDates = targetSheet.getRange(2, dateColIdx + 1, lastRow - 1, 1).getValues().flat();
    const valSet = new Set();
    let isDateSeries = false;

    // データ走査
    rawDates.forEach(d => {
       if (d instanceof Date) {
          isDateSeries = true;
       }
    });

    // ★UI Automation: B43(単位)の自動切り替え
    if (isDateSeries) {
        // 日付モードならプルダウンセット
        if (!["【日付別】", "【月別】", "【年別】"].includes(timeUnit)) {
             timeUnit = "【日付別】";
             const rule = SpreadsheetApp.newDataValidation()
                .requireValueInList(["【日付別】", "【月別】", "【年別】"]).build();
             timeUnitCell.setDataValidation(rule).setValue(timeUnit);
        }
    } else {
        // 回数モードなら項目名固定
        const fixedUnit = dateColName || "回数";
        if (timeUnit !== fixedUnit) {
            timeUnit = fixedUnit;
            // 入力規則解除して値セット
            timeUnitCell.clearDataValidations().setValue(fixedUnit);
        }
    }

    // リスト作成
    let fmt = "yyyy/MM/dd";
    if (timeUnit === "【月別】") fmt = "yyyy/MM";
    if (timeUnit === "【年別】") fmt = "yyyy";

    rawDates.forEach(d => {
       if (d instanceof Date && isDateSeries) {
          valSet.add(Utilities.formatDate(d, Session.getScriptTimeZone(), fmt));
       } else if (d) {
          // 文字列/回数
          valSet.add(String(d).trim());
       }
    });

    // 降順ソート
    const masterList = Array.from(valSet).sort((a, b) => {
        const da = new Date(a);
        const db = new Date(b);
        if (!isNaN(da) && !isNaN(db)) return db - da;
        return String(b).localeCompare(String(a), undefined, {numeric: true});
    });

    // ★重複防止ロジック: 現在選択されている値を収集
    const currentSelections = {};
    for (let i = 0; i < 12; i++) {
        const row = SCHOOL_DATE_COMPARE_START_ROW + i;
        const val = configSheet.getRange(row, 2).getValue();
        if (val) {
             // ★v10.24 Fix: Date型なら文字列化して格納
             if (val instanceof Date) {
                  currentSelections[row] = Utilities.formatDate(val, Session.getScriptTimeZone(), fmt);
             } else {
                  currentSelections[row] = String(val).trim();
             }
        }
    }

    // ★各セルごとに候補リストを生成してセット
    const baseRule = SpreadsheetApp.newDataValidation();
    
    for (let i = 0; i < 12; i++) {
        const targetRow = SCHOOL_DATE_COMPARE_START_ROW + i;
        // 自分以外のセルで選ばれている値を除外
        const otherSelectedValues = Object.keys(currentSelections)
            .filter(r => Number(r) !== targetRow)
            .map(r => currentSelections[r]);
        
        const availableOptions = masterList.filter(item => !otherSelectedValues.includes(item));

        if (availableOptions.length > 0) {
            const rule = baseRule.requireValueInList(availableOptions).build();
            configSheet.getRange(targetRow, 2).setDataValidation(rule);
        } else {
             // 選択肢がない場合
             configSheet.getRange(targetRow, 2).clearDataValidations();
        }
    }

  } catch (e) {
    console.warn("Date Dropdown Error", e);
  }
}

function updateClassDropdown_(configSheet) {
  const ss = configSheet.getParent();
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) return;
  
  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return;
  
  const values = masterSheet.getRange(2, 2, lastRow - 1, 2).getValues();
  const classSet = new Set();
  
  values.forEach(row => {
    const grade = row[0];
    const shClass = String(row[1]); 
    if (grade === "" || shClass === "") return;
    
    const isStandard = !isNaN(shClass) || shClass.length === 1 || /^[IVXivx]+$/.test(shClass);
    if (isStandard) { 
      classSet.add(`${grade}年${shClass}組`); 
    } else { 
      classSet.add(`(全学年)${shClass}`); 
    }
  });
  
  const classList = Array.from(classSet).sort();
  if (classList.length > 0) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(classList).build();
    const cell = configSheet.getRange(SCHOOL_CONFIG_START_ROW + 1, 2);
    cell.setDataValidation(rule).setFontColor("black").setFontWeight("normal");
  }
}

function updateValueDropdown_(configSheet, activeRow) {
  const ss = configSheet.getParent();
  const targetSheetName = configSheet.getRange("B3").getValue();
  const targetColName = configSheet.getRange(activeRow, 2).getValue();
  const valueCell = configSheet.getRange(activeRow + 1, 2);

  valueCell.clearContent().clearDataValidations();
  if (!targetSheetName || !targetColName) return;
  if (String(targetColName).startsWith("▼")) return; 

  const dataSheet = ss.getSheetByName(targetSheetName);
  if (!dataSheet) return;
  const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(targetColName);
  if (colIndex === -1) return;

  // ★追加: B19(時系列単位)の設定を取得してフォーマットを決定
  const dateUnitVal = configSheet.getRange(19, 2).getValue();
  let dateFormat = "yyyy/MM/dd";
  if (dateUnitVal === "【年別】") dateFormat = "yyyy";
  if (dateUnitVal === "【月別】") dateFormat = "yyyy/MM";

  const lastRow = dataSheet.getLastRow();
  let startRow = 2;
  let numRows = lastRow - 1;
  if (numRows > MAX_RECORDS) { 
    startRow = lastRow - MAX_RECORDS + 1; 
    numRows = MAX_RECORDS;
  }

  const colValues = dataSheet.getRange(startRow, colIndex+1, numRows, 1).getValues().flat();
  
  // ★修正: Date型の場合、設定したフォーマットで文字列化してからリストにする
  const uniqueValues = [...new Set(colValues)]
    .filter(v => v !== "" && v != null)
    .map(v => {
        if (v instanceof Date) {
            return Utilities.formatDate(v, Session.getScriptTimeZone(), dateFormat);
        }
        return String(v);
    })
    .sort()
    .slice(0, 500);

  if (uniqueValues.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(uniqueValues)
        .setAllowInvalid(true)
        .setHelpText("リストから選択するか、直接入力してください。")
        .build();
    valueCell.setDataValidation(rule);
  } else {
    valueCell.setNote("⚠️ 候補なし");
  }
}


function generateUniversalCharts_(sheet, chartConfigs) {
  if (!chartConfigs || chartConfigs.length === 0) return;
  
  chartConfigs.forEach(cfg => {
    const range = sheet.getRange(cfg.startRow, 1, cfg.rowCount, 3);
    let chartBuilder = sheet.newChart()
      .addRange(range)
      .setOption('title', cfg.title)
      .setPosition(cfg.anchorRow, 4, 0, 0)
      .setOption('width', 400)
      .setOption('height', 250);
      
    if (cfg.type === "PIE") { 
      chartBuilder = chartBuilder.setChartType(Charts.ChartType.PIE); 
    } else { 
      chartBuilder = chartBuilder.setChartType(Charts.ChartType.BAR); 
    }
    
    sheet.insertChart(chartBuilder.build());
  });
}

function generatePersonalCharts_(sheet, queue) {
  if (!queue || queue.length === 0) return;
  
  queue.forEach(q => {
    let builder = sheet.newChart()
      .addRange(q.range)
      .setOption('title', q.title)
      .setPosition(q.posRow, q.posCol, 0, 0); 
      
    if (q.type === "RADAR") {
        builder = builder.setChartType(Charts.ChartType.RADAR)
          .setTransposeRowsAndColumns(true)
          .setNumHeaders(1) // ★追加: これで「列Aを見出し」として認識させます
          .setOption('useFirstColumnAsDomain', true) // ★念のため: これも合わせ技で入れると完璧です
          .setOption('width', 400)
          .setOption('height', 350);
      }
else if (q.type === "MULTI_LINE") {
      // ★折れ線グラフの設定強化 (行ヘッダーの強制認識)
      builder = builder.setChartType(Charts.ChartType.LINE)
         .setTransposeRowsAndColumns(false) // 行と列を入れ替えない（通常）
         .setNumHeaders(1) // ★先頭行をヘッダーとして明示
         .setOption('useFirstColumnAsDomain', true) // ★1列目をX軸ラベルとして使用
         .setOption('legend', {position: 'right'})
         .setOption('width', 500)
         .setOption('height', 300);
    } else { 
      builder = builder.setChartType(Charts.ChartType.LINE)
         .setOption('legend', {position: 'bottom'}); 
    }
    
    try { 
      sheet.insertChart(builder.build()); 
    } catch(e) { 
      console.warn("Chart Error", e); 
    }
  });
}

function detectAnswerSheetColumns_(configSheet, startRow) {
  const ss = configSheet.getParent();
  const targetSheetName = configSheet.getRange("B3").getValue();
  
  let keyCol = "", dateCol = "";
  let keyMsg = "▼列文字(A,B..)を入力";
  let dateMsg = "▼自動判定";

  if (targetSheetName) {
    const targetSheet = ss.getSheetByName(targetSheetName);
    if (targetSheet) {
      const lastCol = targetSheet.getLastColumn();
      if (lastCol > 0) {
        const headers = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0];
        
        // Key列 (Row 27)
        const keyIndex = headers.findIndex(h => /ID|Email|Account|アカウント|No|Key|コード|番号|メール/i.test(String(h)));
        if (keyIndex > -1) keyCol = headers[keyIndex]; 
        else keyMsg = "⚠️見当たりません"; 

        // 日付(回)列 (Row 28)
      // ★v10.46: 「月」「年」も自動判定対象に追加
      const dateIndex = headers.findIndex(h => /日付|日時|Date|Time|Timestamp|タイムスタンプ|年月|年|月|回/i.test(String(h)));
      if (dateIndex > -1) dateCol = headers[dateIndex];
      else dateCol = headers[0];

      }
    }
  }
  
  configSheet.getRange(startRow + 2, 2).setValue(keyCol || keyMsg).setFontColor(keyCol ? "black" : "red");
  configSheet.getRange(startRow + 3, 2).setValue(dateCol || dateMsg).setFontColor(dateCol ? "black" : "blue");
}

function analyzeColumnType_(values, headerName) {
  // ★日付判定ロジック強化 (Logic Hardening)
  // ヘッダー名に「日付」「Date」「タイムスタンプ」が含まれていたら即TIMESTAMP認定
  if (headerName && /日付|日時|Date|Time|Timestamp|タイムスタンプ/i.test(headerName)) {
      return 'TIMESTAMP';
  }

  if (headerName && /氏名|名前|なまえ|Name|name|フルネーム|番号|出席番号|No\.|ナンバー|number|ID/i.test(headerName)) {
    return 'SKIP';
  }

  if (!values || values.length === 0) return 'CATEGORY';
  
  const sampleSize = Math.min(values.length, 100);
  const sample = values.slice(0, sampleSize).map(String);
  
  let emailCount = 0;
  let totalLen = 0;
  const uniqueSet = new Set();
  let commaCount = 0;
  let dateCount = 0;

  sample.forEach(str => {
    if (str.includes('@')) emailCount++;
    if (str.includes(',') || str.includes('、')) commaCount++; 
    
    // 中身による日付判定
    if (!isNaN(Date.parse(str)) && (str.includes('/') || str.includes('-'))) {
        dateCount++;
    }

    totalLen += str.length;
    uniqueSet.add(str);
  });

  if (dateCount / sample.length > 0.8) return 'TIMESTAMP';
  if (emailCount / sample.length > 0.3) return 'SKIP';
  if (commaCount / sample.length > 0.3) return 'CATEGORY'; 

  const uniqueRatio = uniqueSet.size / sample.length;
  if (uniqueRatio > 0.8) return 'FREE_TEXT'; 
  
  return 'CATEGORY';
}

// ★引数 dateFormat を末尾に追加
function renderCrossTabulation_(sheet, headers, data, crossIdx, crossName, startCol, isTimestamp, dateFormat) {
  // フォーマットのデフォルト値設定
  const fmt = dateFormat || "yyyy/MM/dd";

  // --- A. 横軸（グループ）のキー生成 ---
  const getGroupKey = (row) => {
    const val = row[crossIdx];
    if (!val) return null;
    
    // ★修正: 固定の"yyyy/MM"ではなく、受け取ったfmtを使用する
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), fmt);
    }
    return String(val);
  };
  
  // ソートロジック
  const groups = [...new Set(data.map(row => getGroupKey(row)).filter(v => v))].sort((a, b) => {
    return String(a).localeCompare(String(b), undefined, {numeric: true, sensitivity: 'base'});
  });

  if (groups.length === 0) return;

  const output = [];
  // ヘッダー生成 (フォーマットに合わせてラベルを変える)
  let modeLabel = "";
  if (isTimestamp) {
     if (fmt === "yyyy") modeLabel = ":年別";
     else if (fmt === "yyyy/MM") modeLabel = ":月別";
     else modeLabel = ":日別";
  }
  
  const headerRow = [`【詳細比較${modeLabel}】質問項目`, "選択肢", ...groups];
  output.push(headerRow);

  const isStrictNumber = (val) => {
      if (val === "" || val === null) return false;
      if (val instanceof Date) return false; 
      const s = String(val).trim();
      if (s === "") return false;
      if (s.includes('/') || s.includes(':') || s.includes('-')) return false;
      const n = Number(s);
      return !isNaN(n);
  };

  const averageTrendData = [];

  // --- B. 各質問についてループ ---
  for (let i = 1; i < headers.length; i++) {
    if (i === crossIdx) continue;
    const qTitle = headers[i];
    if (!qTitle) continue;
    
    const colValues = data.map(r => r[i]).filter(v => v !== "" && v != null);
    if (colValues.length === 0) continue;
    
    const colType = analyzeColumnType_(colValues, qTitle);
    if (colType === 'SKIP' || colType === 'FREE_TEXT' || colType === 'TIMESTAMP') continue;

    const isAttributeCol = /学年|組|クラス|番号|出席番号|No\.|ID|コード|性別|Gender|氏名|名前|Name/i.test(qTitle);

    const pairs = data.map(row => ({
      val: row[i], 
      ans: String(row[i]), 
      group: getGroupKey(row)
    })).filter(p => p.ans && p.group && p.ans !== "");

    if (pairs.length === 0) continue;

    let numericCount = 0;
    pairs.forEach(p => { if (isStrictNumber(p.val)) numericCount++; });
    const isNumericQuestion = !isAttributeCol && (numericCount / pairs.length) > 0.8;

    if (isNumericQuestion) {
        const stats = {};
        groups.forEach(g => stats[g] = {sum: 0, count: 0});
        pairs.forEach(p => {
            if (isStrictNumber(p.val) && stats[p.group]) {
                const vNum = Number(String(p.val).trim()); 
                stats[p.group].sum += vNum;
                stats[p.group].count++;
            }
        });
        const averages = {};
        groups.forEach(g => {
            const s = stats[g];
            averages[g] = s.count > 0 ? parseFloat((s.sum / s.count).toFixed(2)) : null;
        });
        averageTrendData.push({ title: qTitle, averages: averages });
    }

    const uniqueAnswers = [...new Set(pairs.map(p => p.ans))].sort();
    if (uniqueAnswers.length > 50) continue; 

    const counts = {};
    uniqueAnswers.forEach(ans => {
      counts[ans] = {};
      groups.forEach(g => counts[ans][g] = 0);
    });
    pairs.forEach(p => {
      if (counts[p.ans] && counts[p.ans][p.group] !== undefined) {
        counts[p.ans][p.group]++;
      }
    });

    let isFirst = true;
    uniqueAnswers.forEach(ans => {
      const rowData = [isFirst ? qTitle : "", ans];
      groups.forEach(g => {
        rowData.push(counts[ans][g] || 0);
      });
      output.push(rowData);
      isFirst = false;
    });
    output.push(new Array(headerRow.length).fill(""));
  }

  // --- F. 出力処理 ---
  if (output.length > 0) {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    // 既存データのクリア (ヘッダーより下、開始列より右をクリア)
    if (maxCols >= startCol) {
      // 安全策: 行数が少ない場合はクリア範囲を調整
      const clearRows = maxRows > 1 ? maxRows - 1 : 1;
      try {
        sheet.getRange(1, startCol, clearRows, maxCols - startCol + 1).clearContent().clearFormat();
      } catch (e) { /* 範囲外エラー抑制 */ }
    }

    // クロス集計表の出力
    sheet.getRange(1, startCol).setValue(`🔍 詳細クロス集計（軸: ${crossName}）`)
         .setFontSize(12).setFontWeight("bold").setFontColor("#0b5394");
    
    const range = sheet.getRange(4, startCol, output.length, output[0].length);
    range.setValues(output);
    range.setBorder(true, true, true, true, true, true);
    
    // スタイル適用
    sheet.getRange(4, startCol, 1, output[0].length).setBackground("#c9daf8").setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(4, startCol, output.length, 1).setBackground("#f3f3f3").setFontWeight("bold");
    
    // 列幅調整
    sheet.setColumnWidth(startCol, 200); 
    sheet.setColumnWidth(startCol + 1, 150);
    for (let k = 0; k < groups.length; k++) {
      sheet.setColumnWidth(startCol + 2 + k, 70);
    }
    
    let currentOutputRow = 4 + output.length + 2;

    // --- G. 平均値推移表 & グラフ (Trend & GAP Analysis) ---
    // ここでグラフ描画と、次の開始行の計算を行う
    if (averageTrendData.length > 0) {
        sheet.getRange(currentOutputRow, startCol).setValue(`📈 平均値比較推移（軸: ${crossName}）`)
             .setFontSize(12).setFontWeight("bold").setFontColor("#E91E63");
        currentOutputRow += 2;

        const summaryHeader = ["項目名(質問)", ...groups];
        sheet.getRange(currentOutputRow, startCol, 1, summaryHeader.length)
             .setValues([summaryHeader])
             .setBackground("#FCE8E6").setFontWeight("bold").setBorder(true, true, true, true, true, true);
        currentOutputRow++;
        
        const startAvgRow = currentOutputRow;
        const avgRows = [];

        averageTrendData.forEach(item => {
            const rowVals = [item.title];
            groups.forEach(g => {
                rowVals.push(item.averages[g] !== null ? item.averages[g] : "");
            });
            avgRows.push(rowVals);
        });

        if (avgRows.length > 0) {
            const avgRange = sheet.getRange(currentOutputRow, startCol, avgRows.length, avgRows[0].length);
            avgRange.setValues(avgRows).setBorder(true, true, true, true, true, true);
            
            // 1. 折れ線グラフ (Trend Chart)
            const chartRow = currentOutputRow + avgRows.length + 2;
            const chartDataRange = sheet.getRange(startAvgRow - 1, startCol, avgRows.length + 1, summaryHeader.length);
            
            const trendChart = sheet.newChart()
                .setChartType(Charts.ChartType.LINE)
                .addRange(chartDataRange)
                .setPosition(chartRow, startCol, 0, 0)
                .setOption('title', `詳細クロス集計: 平均値推移 (${crossName})`)
                .setOption('width', 1000) 
                .setOption('height', 400)
                .setTransposeRowsAndColumns(true) 
                .setOption('treatLabelsAsText', true) 
                .setOption('useFirstColumnAsDomain', true)
                .setNumHeaders(1)
                .setOption('legend', {position: 'right'})
                .setOption('vAxis', {title: '平均スコア'})
                .build();
            sheet.insertChart(trendChart);

            // 2. GAP分析グラフ (Gap Chart)
            // 全項目の平均値に対するGAPを可視化
            const validAvgs = Object.values(averageTrendData[0].averages).filter(v => v !== null); // 簡易的に最初の項目のデータ構造を利用
            if (validAvgs.length > 0) {
               // 全体平均算出 (単純平均)
               let globalSum = 0, globalCnt = 0;
               averageTrendData.forEach(d => {
                   Object.values(d.averages).forEach(v => { if(v!==null){ globalSum+=v; globalCnt++; }});
               });
               const globalAvg = globalCnt > 0 ? globalSum / globalCnt : 0;

               // GAPデータ作成
               const gapData = [["Group", "GAP (vs Total Avg)"]];
               groups.forEach(g => {
                   let gSum = 0, gCnt = 0;
                   averageTrendData.forEach(d => {
                       if(d.averages[g] !== null) { gSum += d.averages[g]; gCnt++; }
                   });
                   const gAvg = gCnt > 0 ? gSum / gCnt : 0;
                   gapData.push([g, parseFloat((gAvg - globalAvg).toFixed(2))]);
               });

               // データ書き出し (グラフの裏側エリアを使用)
               const gapDataRow = chartRow;
               const gapDataCol = startCol + summaryHeader.length + 2; 
               const gapRange = sheet.getRange(gapDataRow, gapDataCol, gapData.length, 2);
               gapRange.setValues(gapData);

               const gapChartRow = chartRow + 21; // 折れ線グラフの下
               const gapChart = sheet.newChart()
                  .setChartType(Charts.ChartType.COLUMN)
                  .addRange(gapRange)
                  .setPosition(gapChartRow, startCol, 0, 0)
                  .setOption('title', `GAP分析: 全体平均(${globalAvg.toFixed(2)})との乖離`)
                  .setOption('width', 1000)
                  .setOption('height', 300)
                  .setOption('legend', {position: 'none'})
                  .setOption('colors', ['#FF5722'])
                  .build();
               sheet.insertChart(gapChart);
            }
        }
    }

    // ★修正ポイント: グラフを描画した場合、その高さを考慮して次の開始位置を決定する
    // これにより、後続の「相関分析」などがグラフと重なるのを防ぐ
    let nextStartRow = currentOutputRow;
    if (averageTrendData.length > 0) {
        // ヒートマップ + 折れ線(20行) + GAP(15行) + 余白
        nextStartRow = currentOutputRow + 45; 
    }
    
    return nextStartRow;

  } else {
    // ★修正ポイント: データがなく出力しなかった場合でも、有効な行番号を返す
    // これを返さないと呼び出し元で undefined になりエラー停止する
    return Math.max(4, sheet.getLastRow() + 2);
  }
} // End function


// ==================================================
// 🆕 拡張機能: 相関分析 & 生データ出力 & GAP計算
// ==================================================

/**
 * 拡張機能: 相関分析マトリクス生成 (v10.46 Modified)
 * - ヘッダー: 0度/折り返し
 * - 除外: メールアドレス等を強化
 * - UI: ガイドパネル追加
 */
function generateCorrelationMatrix_(sheet, headers, body, startRow) {
  // 1. 数値列の特定とデータ抽出
  const numericData = []; // [{title: "Q1...", values: [1, 5, 3...]}]
  const numRows = body.length;
  if (numRows < 2) return startRow;

  headers.forEach((h, colIdx) => {
    // 【修正】除外キーワードを強化（学年、組、HR、Class等を追加）
    if (/学年|組|クラス|HR|Grade|Class|氏名|名前|出席番号|番号|No\.|ID|コード|Timestamp|タイムスタンプ|メール|Email|address|account/i.test(h)) return;
    
    const rawVals = body.map(r => r[colIdx]);
    
    // 【修正】数値判定の厳格化 (parseFloatをやめ、Numberを使用)
    let nCnt = 0;
    const nVals = [];
    rawVals.forEach(v => {
      const s = String(v).trim();
      // 日付スラッシュや時刻コロンが含まれる場合は数値扱いしない
      if (s === "" || s.includes('/') || s.includes(':')) {
          nVals.push(null);
      } else {
          const n = Number(s);
          if (!isNaN(n)) { nCnt++; nVals.push(n); } else { nVals.push(null); }
      }
    });

    // 8割以上が数値の場合のみ採用
    if (nCnt / numRows > 0.8) {
      numericData.push({ title: h, values: nVals });
    }
  });


  // 比較対象が2つ未満なら作成しない
  if (numericData.length < 2) return startRow;

  // 2. マトリクス計算 (Pearson)
  const size = numericData.length;
  const matrix = Array(size).fill(null).map(() => Array(size).fill(""));

  for (let i = 0; i < size; i++) {
    for (let j = 0; j < size; j++) {
      if (i === j) {
        matrix[i][j] = "-";
      } else {
        const r = calculateCorrelation_(numericData[i].values, numericData[j].values);
        matrix[i][j] = r !== null ? parseFloat(r.toFixed(2)) : "";
      }
    }
  }

  // 3. 出力処理
  let currentRow = startRow;
  sheet.getRange(currentRow, 1).setValue("📈 相関分析マトリクス (相関係数)")
       .setFontSize(12).setFontWeight("bold").setFontColor("#673AB7");
  currentRow += 2;

  // ヘッダー (横)
  const titles = numericData.map(d => d.title);
  
  // ★Fix: 0度回転 & 折り返し設定 & 列幅固定
  const headerRange = sheet.getRange(currentRow, 2, 1, size);
  headerRange.setValues([titles])
       .setBackground("#EDE7F6")
       .setFontWeight("bold")
       .setTextRotation(0) // 0度に戻す
       .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // 折り返し
       .setVerticalAlignment("middle")
       .setHorizontalAlignment("center")
       .setBorder(true, true, true, true, true, true);
  
  // 列幅を適度なサイズ(100px)に固定して見やすくする
  sheet.setColumnWidths(2, size, 100);

  // データ本体出力
  const outRows = [];
  for(let i=0; i<size; i++){
    outRows.push([titles[i], ...matrix[i]]);
  }
  
  sheet.getRange(currentRow + 1, 1, size, size + 1)
       .setValues(outRows)
       .setBorder(true, true, true, true, true, true)
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle");
       
  // 左端列(項目名)も折り返し設定
  sheet.getRange(currentRow + 1, 1, size, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // 4. 条件付き書式 (ヒートマップ)
  const dataRange = sheet.getRange(currentRow + 1, 2, size, size);
  const rules = sheet.getConditionalFormatRules();

  // 正の相関 (赤)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0.4)
    .setBackground("#FFCDD2") // 薄い赤
    .setFontColor("#B71C1C")
    .setRanges([dataRange])
    .build());

  // 負の相関 (青)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(-0.4)
    .setBackground("#BBDEFB") // 薄い青
    .setFontColor("#0D47A1")
    .setRanges([dataRange])
    .build());

  sheet.setConditionalFormatRules(rules);

 // 5. ★New: 「見方」ガイドパネルの作成 (表の右側に配置)
  const guideStartCol = 2 + size + 1; // 表の右隣+1列空ける
  const guideRange = sheet.getRange(currentRow, guideStartCol, 7, 3);
  
  // ガイド用データ
  const guideData = [
    ["💡 相関係数の見方", "", ""],
    ["数値", "意味", "色"],
    ["0.7 ～ 1.0", "強い正の相関 (比例)", "赤"],
    ["0.4 ～ 0.7", "正の相関あり", "薄赤"],
    ["-0.4 ～ 0.4", "相関なし", "白"],
    ["-0.7 ～ -0.4", "負の相関あり (反比例)", "薄青"],
    ["-1.0 ～ -0.7", "強い負の相関", "青"]
  ];
  
  // ガイド書き込み & 書式
  guideRange.setValues(guideData);
  sheet.getRange(currentRow, guideStartCol, 1, 3).merge().setFontWeight("bold").setBackground("#f3f3f3");
  sheet.getRange(currentRow + 1, guideStartCol, 1, 3).setFontWeight("bold").setBackground("#e0e0e0");
  
  // 枠線
  guideRange.setBorder(true, true, true, true, true, true);

  // ★修正: 幅を自動調整ではなく、指定サイズ（広め）に固定
  // 数値: 150px, 意味: 150px, 色: 150px
  sheet.setColumnWidth(guideStartCol, 150);     // 数値列
  sheet.setColumnWidth(guideStartCol + 1, 300); // 意味列（ここを大きく）
  sheet.setColumnWidth(guideStartCol + 2, 100); // 色列

  return currentRow + size + 4;
}



/**
 * 拡張B: 抽出生データテーブル出力
 */
function renderRawDataTable_(sheet, headers, body, startRow) {
  if (!body || body.length === 0) return startRow;

  let currentRow = startRow;
  sheet.getRange(currentRow, 1).setValue("🔍 抽出データ・ローデータ一覧 (フィルタ適用済)")
       .setFontSize(12).setFontWeight("bold").setFontColor("#333333");
  currentRow += 1;

  // ヘッダー出力
  sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers])
       .setBackground("#666666").setFontColor("white").setFontWeight("bold");
  
  // データ出力 (最大10000行まで安全策)
  const safeRows = body.length > 10000 ? 10000 : body.length;
  if (safeRows > 0) {
    sheet.getRange(currentRow + 1, 1, safeRows, headers.length).setValues(body.slice(0, safeRows))
         .setBorder(true, true, true, true, true, true);
  }

  if (body.length > 10000) {
    sheet.getRange(currentRow + 1 + safeRows, 1).setValue("※表示制限: 10,000件までを表示しています");
  }

  return currentRow + safeRows + 3;
}

/**
 * Helper: ピアソンの積率相関係数算出
 */
function calculateCorrelation_(x, y) {
  let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0, sumY2 = 0;
  let n = 0;
  for (let i = 0; i < x.length; i++) {
    if (x[i] !== null && y[i] !== null) {
      sumX += x[i];
      sumY += y[i];
      sumXY += x[i] * y[i];
      sumX2 += x[i] * x[i];
      sumY2 += y[i] * y[i];
      n++;
    }
  }
  if (n === 0) return null;
  const numerator = (n * sumXY) - (sumX * sumY);
  const denominator = Math.sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
  if (denominator === 0) return 0;
  return numerator / denominator;
}

