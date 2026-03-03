/**
 * 介護記録作成アシスタント - Google Apps Script版
 * 
 * 【主な機能】
 * - Gemini API統合でAI記録文生成
 * - Google Sheetsへの自動保存（履歴管理）
 * - 介護用語辞書の管理
 * - RESTful API提供（フロントエンド用）
 */

// =============================================================================
// 設定・初期化
// =============================================================================

/**
 * Webアプリとして表示するHTMLを返す
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('介護記録作成アシスタント')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Gemini APIキーを取得
 * @return {string} APIキー
 */
function getGeminiApiKey() {
  const properties = PropertiesService.getUserProperties();
  return properties.getProperty('GEMINI_API_KEY') || '';
}

/**
 * Gemini APIキーを保存
 * @param {string} apiKey - APIキー
 */
function setGeminiApiKey(apiKey) {
  const properties = PropertiesService.getUserProperties();
  properties.setProperty('GEMINI_API_KEY', apiKey);
  return { success: true, message: 'APIキーを保存しました' };
}

/**
 * 介護記録履歴用スプレッドシートを取得または作成
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getRecordsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('介護記録履歴');
  
  if (!sheet) {
    sheet = ss.insertSheet('介護記録履歴');
    // ヘッダー行を作成
    sheet.appendRow([
      'ID', 
      'タイトル', 
      '作成日時', 
      '更新日時', 
      '記録タイプ', 
      '入力テキスト', 
      '整理済テキスト', 
      '最終記録文',
      '職員名',
      '利用者名'
    ]);
    
    // ヘッダー行をフリーズ
    sheet.setFrozenRows(1);
    
    // ヘッダー行を太字にする
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
  }
  
  return sheet;
}

/**
 * 用語辞書用スプレッドシートを取得または作成
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getDictionarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('用語辞書');
  
  if (!sheet) {
    sheet = ss.insertSheet('用語辞書');
    // ヘッダー行を作成
    sheet.appendRow(['誤認識語', '正しい表記', 'カテゴリ', '登録日時']);
    
    // デフォルト辞書を追加
    const defaultDict = getDefaultDictionary();
    defaultDict.forEach(entry => {
      sheet.appendRow([
        entry.wrong,
        entry.correct,
        entry.category || 'デフォルト',
        new Date().toISOString()
      ]);
    });
    
    // ヘッダー行をフリーズ
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  
  return sheet;
}

/**
 * デフォルト介護用語辞書
 * @return {Array<Object>}
 */
function getDefaultDictionary() {
  return [
    // ADL・IADL
    { wrong: 'えーでぃーえる', correct: 'ADL', category: 'ADL・IADL' },
    { wrong: 'あいえーでぃーえる', correct: 'IADL', category: 'ADL・IADL' },
    
    // バイタル
    { wrong: 'ばいたる', correct: 'バイタル', category: 'バイタル' },
    { wrong: 'けつあつ', correct: '血圧', category: 'バイタル' },
    { wrong: 'たいおん', correct: '体温', category: 'バイタル' },
    { wrong: 'みゃくはく', correct: '脈拍', category: 'バイタル' },
    { wrong: 'さっそ', correct: 'SpO2', category: 'バイタル' },
    { wrong: 'えすぴーおーつー', correct: 'SpO2', category: 'バイタル' },
    
    // 移動・移乗
    { wrong: 'いじょう', correct: '移乗', category: '移動・移乗' },
    { wrong: 'いどう', correct: '移動', category: '移動・移乗' },
    { wrong: 'ほこう', correct: '歩行', category: '移動・移乗' },
    { wrong: 'くるまいす', correct: '車椅子', category: '移動・移乗' },
    { wrong: 'しるばーかー', correct: 'シルバーカー', category: '移動・移乗' },
    
    // 食事
    { wrong: 'しょくじかいじょ', correct: '食事介助', category: '食事' },
    { wrong: 'けんげ', correct: '嚥下', category: '食事' },
    { wrong: 'えんげ', correct: '嚥下', category: '食事' },
    { wrong: 'すいぶんほきゅう', correct: '水分補給', category: '食事' },
    
    // 排泄
    { wrong: 'はいせつかいじょ', correct: '排泄介助', category: '排泄' },
    { wrong: 'おむつ', correct: 'おむつ', category: '排泄' },
    { wrong: 'ぽーたぶるといれ', correct: 'ポータブルトイレ', category: '排泄' },
    
    // 入浴・清潔
    { wrong: 'にゅうよく', correct: '入浴', category: '入浴・清潔' },
    { wrong: 'こうくうけあ', correct: '口腔ケア', category: '入浴・清潔' },
    { wrong: 'せいしき', correct: '清拭', category: '入浴・清潔' },
    
    // 医療
    { wrong: 'ないふくやく', correct: '内服薬', category: '医療' },
    { wrong: 'ちゅうしゃ', correct: '注射', category: '医療' },
    
    // 認知症
    { wrong: 'にんちしょう', correct: '認知症', category: '認知症' },
    { wrong: 'びーぴーえすでぃー', correct: 'BPSD', category: '認知症' },
    
    // リハビリ
    { wrong: 'りはびり', correct: 'リハビリ', category: 'リハビリ' },
    { wrong: 'きのうくんれん', correct: '機能訓練', category: 'リハビリ' }
  ];
}

// =============================================================================
// Gemini API連携
// =============================================================================

/**
 * Gemini APIでテキストを生成
 * @param {string} prompt - プロンプト
 * @return {string} 生成されたテキスト
 */
function callGeminiApi(prompt) {
  const apiKey = getGeminiApiKey();
  
  if (!apiKey) {
    throw new Error('Gemini APIキーが設定されていません。設定画面で設定してください。');
  }
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  
  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 2048
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.error) {
      throw new Error(`Gemini API Error: ${json.error.message}`);
    }
    
    return json.candidates[0].content.parts[0].text;
  } catch (error) {
    throw new Error(`API呼び出しエラー: ${error.message}`);
  }
}

/**
 * Gemini API接続テスト
 * @param {string} apiKey - テストするAPIキー
 * @return {Object} テスト結果
 */
function testGeminiConnection(apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  
  const payload = {
    contents: [{
      parts: [{
        text: 'こんにちは'
      }]
    }]
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.error) {
      return { 
        success: false, 
        message: `接続失敗: ${json.error.message}` 
      };
    }
    
    return { 
      success: true, 
      message: '✅ 接続成功！APIは正常に動作しています' 
    };
  } catch (error) {
    return { 
      success: false, 
      message: `❌ 接続失敗: ${error.message}` 
    };
  }
}

// =============================================================================
// 記録文生成
// =============================================================================

/**
 * 音声入力テキストを一次整理
 * @param {string} inputText - 入力テキスト
 * @return {string} 整理済みテキスト
 */
function organizeText(inputText) {
  const prompt = `
あなたは介護記録作成を支援する専門アシスタントです。

以下は介護現場で職員が音声入力した文字起こしデータです。
内容を解釈しすぎず、話された事実を正確に整理してください。

【整理ルール】
・推測・感情・評価は書かない
・「〜と思われる」「〜のようだ」は使わない
・話された内容のみを整理する
・同じ内容は統合する
・時系列が分かる場合は維持する

【出力形式】
■ 利用者の状態・変化
■ 実施した支援内容
■ 観察された事実
■ 職員の対応
■ 特記事項（事故・ヒヤリハット・注意点があれば）

【文字起こし】
${inputText}
`;

  return callGeminiApi(prompt);
}

/**
 * 整理済みテキストから記録文を生成
 * @param {string} organizedText - 整理済みテキスト
 * @param {string} recordType - 記録タイプ（standard/monitoring/plan）
 * @return {string} 記録文
 */
function generateRecord(organizedText, recordType) {
  let prompt = '';
  
  if (recordType === 'monitoring') {
    prompt = `
以下の情報をもとに、モニタリング記録として適切な文章を作成してください。

【条件】
・計画に対する実施状況が分かること
・「できた／できなかった」を明確に
・評価は事実ベースで簡潔に

【出力形式】
・目標に対する状況
・支援の実施内容
・利用者の反応
・今後の対応（必要があれば）

【入力】
${organizedText}
`;
  } else if (recordType === 'plan') {
    prompt = `
以下の内容をもとに、介護支援計画書に使用できる文章を作成してください。

【重要】
・具体的かつ現実的
・抽象表現は禁止
・誰が・いつ・何をするかが分かる

【出力項目】
■ 短期目標
■ 支援内容
■ 留意事項

【入力】
${organizedText}
`;
  } else {
    prompt = `
あなたは介護記録作成の専門家です。
以下の整理済みメモをもとに、介護記録として適切な文章を作成してください。

【前提】
・この文章は業務記録として保存されます
・第三者（監査・家族）が読む可能性があります

【記述ルール】
・簡潔で客観的
・主語を明確にする
・曖昧な表現は避ける
・事実のみを記載する
・1文は長くしすぎない

【出力形式】
① 経過観察記録（記録用文章）
② 必要があれば注意事項（1行）

【入力】
${organizedText}
`;
  }
  
  return callGeminiApi(prompt);
}

/**
 * ワンステップで記録文を生成
 * @param {string} inputText - 入力テキスト
 * @param {string} recordType - 記録タイプ
 * @return {Object} 結果
 */
function generateRecordOneStep(inputText, recordType) {
  try {
    // 一次整理
    const organizedText = organizeText(inputText);
    
    // 記録文生成
    const finalText = generateRecord(organizedText, recordType || 'standard');
    
    return {
      success: true,
      organizedText: organizedText,
      finalText: finalText
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

// =============================================================================
// 履歴管理
// =============================================================================

/**
 * 記録を保存
 * @param {Object} record - 記録データ
 * @return {Object} 保存結果
 */
function saveRecord(record) {
  try {
    const sheet = getRecordsSheet();
    const id = record.id || Utilities.getUuid();
    const now = new Date().toISOString();
    
    // 既存レコードを検索
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        rowIndex = i + 1;
        break;
      }
    }
    
    const row = [
      id,
      record.title || '無題の記録',
      record.createdAt || now,
      now, // 更新日時
      record.recordType || 'standard',
      record.inputText || '',
      record.organizedText || '',
      record.finalText || '',
      record.staffName || '',
      record.userName || ''
    ];
    
    if (rowIndex > 0) {
      // 更新
      sheet.getRange(rowIndex, 1, 1, 10).setValues([row]);
    } else {
      // 新規作成
      sheet.appendRow(row);
    }
    
    return {
      success: true,
      id: id,
      message: '記録を保存しました'
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 全記録を取得
 * @return {Array<Object>} 記録のリスト
 */
function getAllRecords() {
  try {
    const sheet = getRecordsSheet();
    const data = sheet.getDataRange().getValues();
    
    // ヘッダー行をスキップ
    const records = [];
    for (let i = 1; i < data.length; i++) {
      records.push({
        id: data[i][0],
        title: data[i][1],
        createdAt: data[i][2],
        updatedAt: data[i][3],
        recordType: data[i][4],
        inputText: data[i][5],
        organizedText: data[i][6],
        finalText: data[i][7],
        staffName: data[i][8],
        userName: data[i][9]
      });
    }
    
    // 更新日時の降順でソート
    records.sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
    
    return {
      success: true,
      records: records
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 記録を削除
 * @param {string} id - 記録ID
 * @return {Object} 削除結果
 */
function deleteRecord(id) {
  try {
    const sheet = getRecordsSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return {
          success: true,
          message: '記録を削除しました'
        };
      }
    }
    
    return {
      success: false,
      error: '記録が見つかりませんでした'
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

// =============================================================================
// 用語辞書管理
// =============================================================================

/**
 * 全辞書データを取得
 * @return {Array<Object>} 辞書エントリのリスト
 */
function getAllDictionaryEntries() {
  try {
    const sheet = getDictionarySheet();
    const data = sheet.getDataRange().getValues();
    
    const entries = [];
    for (let i = 1; i < data.length; i++) {
      entries.push({
        wrong: data[i][0],
        correct: data[i][1],
        category: data[i][2],
        createdAt: data[i][3]
      });
    }
    
    return {
      success: true,
      entries: entries
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 辞書エントリを追加
 * @param {string} wrong - 誤認識語
 * @param {string} correct - 正しい表記
 * @param {string} category - カテゴリ
 * @return {Object} 追加結果
 */
function addDictionaryEntry(wrong, correct, category) {
  try {
    const sheet = getDictionarySheet();
    sheet.appendRow([
      wrong,
      correct,
      category || 'カスタム',
      new Date().toISOString()
    ]);
    
    return {
      success: true,
      message: '用語を追加しました'
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 辞書エントリを削除
 * @param {string} wrong - 誤認識語
 * @return {Object} 削除結果
 */
function deleteDictionaryEntry(wrong) {
  try {
    const sheet = getDictionarySheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === wrong) {
        sheet.deleteRow(i + 1);
        return {
          success: true,
          message: '用語を削除しました'
        };
      }
    }
    
    return {
      success: false,
      error: '用語が見つかりませんでした'
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 辞書を初期状態にリセット
 * @return {Object} リセット結果
 */
function resetDictionary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('用語辞書');
    
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    
    // 再作成
    getDictionarySheet();
    
    return {
      success: true,
      message: '辞書を初期状態にリセットしました'
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}
