// Firestore の初期化
function initializeFirestore() {
    return FirestoreApp.getFirestore(
      FIREBASE_CONFIG.client_email,
      FIREBASE_CONFIG.private_key,
      FIREBASE_CONFIG.project_id
    );
  }
  
  // ログを「ログ記録」シートに書き込む関数
  function logToSheet(logSheet, functionName, logMessage) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    logSheet = ss.getSheetByName("ログ記録") || ss.insertSheet("ログ記録");
  
    const timestamp = new Date();
    logSheet.appendRow([timestamp, functionName, logMessage]);
  }
  
  // キャッシング機能を使って店舗データを取得する関数
  function getStoreDataWithCache(storeId) {
    const cache = CacheService.getScriptCache(); // スクリプト全体で使えるキャッシュ
    const cachedData = cache.get(storeId);
  
    if (cachedData) {
      Logger.log(`キャッシュから店舗ID: ${storeId} のデータを取得`);
      return Promise.resolve(JSON.parse(cachedData)); // キャッシュデータは文字列で保存されるため、JSON.parseでオブジェクトに変換
    } else {
      const firestore = initializeFirestore();
      const documentPath = `stores/${storeId}`;  // パスを変数に保存して確認
      Logger.log(`Firestoreからデータを取得中: ${documentPath}`);  // パスをログに表示
  
      // 正しいパスを指定してFirestoreからデータを取得
      return firestore.getDocument(documentPath).then((doc) => {
        if (doc && doc.fields) {
          const storeData = {
            name: doc.fields.name.stringValue,
            address: doc.fields.address.stringValue
          };
  
          // キャッシュに保存
          cache.put(storeId, JSON.stringify(storeData), 3600); // キャッシュは1時間保存（3600秒）
          Logger.log(`Firestoreから店舗ID: ${storeId} のデータを取得し、キャッシュに保存`);
          return storeData;
        } else {
          Logger.log(`店舗ID: ${storeId} のデータが見つかりません`);
          return null;
        }
      }).catch((error) => {
        Logger.log('Firestoreの読み取り中にエラーが発生しました: ' + error.message);
        throw new Error('Firestore読み取りエラー: ' + error.message);
      });
    }
  }
  
  // タイムゾーン変換関数
  function convertToJST(date) {
    const jstOffset = 9 * 60; // 日本標準時のオフセットを分に変換
    const jstDate = new Date(date.getTime() + (jstOffset * 60000));
    return jstDate;
  }
  
  /// チェックされた行のみFirestoreに同期する関数
  function syncCheckedStoreInfoToFirestore() {
    const firestore = initializeFirestore();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName("ログ記録"); // ログシート取得
    const storeSheet = ss.getSheetByName("店舗情報");
  
    const lastRow = storeSheet.getLastRow();  // 最終行を取得
    const dataRange = storeSheet.getRange(2, 1, lastRow - 1, 5); // データ範囲（A列〜E列）を取得
    const data = dataRange.getValues();  // 全データを取得
    const checkboxes = storeSheet.getRange(2, 4, lastRow - 1, 1).getValues();  // D列のチェックボックスの状態を取得
  
    Logger.log(`データ: ${JSON.stringify(data)}`);
    Logger.log(`チェックボックスの状態: ${JSON.stringify(checkboxes)}`);
  
    const now = new Date(); // 現在時刻
    let isAnyChecked = false; // チェックボックスが1つでも選択されているかどうかのフラグ
    const storeDataList = []; // 空のリストを初期化
  
    // データを確認しながらstoreDataListに追加
    data.forEach((row, index) => {
      const storeId = row[0];  // 店舗ID
      const storeName = row[1];  // 店舗名
      const storeAddress = row[2];  // 店舗住所
      const isChecked = checkboxes[index][0];  // D列のチェックボックスの状態
  
      Logger.log(`storeId: ${storeId}, storeName: ${storeName}, storeAddress: ${storeAddress}, isChecked: ${isChecked}`);
  
      // チェックされている行のみ同期
      if (isChecked) {
        isAnyChecked = true;  // 少なくとも1つのチェックが選択されていることを示す
        const storeData = {
          storeId: storeId,
          name: storeName,
          address: storeAddress
        };
  
        storeDataList.push(storeData); // 同期対象のデータリストに追加
  
        // チェックボックスをクリア（チェックを外す）
        storeSheet.getRange(index + 2, 4).setValue(false);
  
        // 最終同期時間をE列に記録
        storeSheet.getRange(index + 2, 5).setValue(now);
      }
    });
  
    Logger.log(`storeDataList: ${JSON.stringify(storeDataList)}`);
  
    if (!isAnyChecked) {
      SpreadsheetApp.getUi().alert('同期する店舗が選択されていません。チェックボックスを選択してください。');
      logToSheet(logSheet, "syncCheckedStoreInfoToFirestore", '同期する店舗が選択されていません。');
      return;
    }
  
    // storeDataListが空でないか確認する
    if (storeDataList.length > 0) {
      // 個別書き込みで Firestore にデータを同期
      updateFirestoreIndividually(storeDataList);
    } else {
      Logger.log("同期対象のデータがありません。");
      logToSheet(logSheet, "syncCheckedStoreInfoToFirestore", "同期対象のデータがありません。");
    }
  
    Logger.log('Firestore への更新が完了しました。');
    logToSheet(logSheet, "syncCheckedStoreInfoToFirestore", 'Firestore への更新が完了しました。');
  }
  
  // Firestoreの補充情報をスプレッドシートの「報告情報」に同期する関数（キャッシング適用済）
  async function updateReportInfoSheet() {
    const firestore = initializeFirestore();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName("ログ記録"); // ログシート取得
    const reportSheet = ss.getSheetByName("報告情報");
  
    // シートを初期化
    reportSheet.clear();
    reportSheet.getRange(1, 1, 1, 4).setValues([["店舗ID", "店舗名", "所在地", "報告日時"]]);
    reportSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    reportSheet.setFrozenRows(1);
  
    try {
      // Firestoreのrestock-reportsコレクションからデータを取得
      const documents = firestore.getDocuments("restock-reports"); // collectionの代わりにgetDocumentsを使用
  
      if (!documents || documents.length === 0) {  
        Logger.log('Firestoreからのドキュメントが取得できませんでした。');
        logToSheet(logSheet, "updateReportInfoSheet", 'Firestoreからのドキュメントが取得できませんでした。');
        return;
      }
  
      let newData = [];
  
      documents.forEach(doc => {
        const storeId = doc.name.split('/').pop();  // ドキュメント名はstoreIdのみ
        const reportedAt = new Date(doc.fields.reportedAt.timestampValue);  // reportedAtフィールドを使用してタイムスタンプを取得
  
        Logger.log(`Firestoreからデータを取得中: stores/${storeId}`);  // パスをログに表示
  
        // キャッシュを利用して店舗情報を取得
        const storeData = getStoreDataWithCache(storeId);
        if (storeData) {
          newData.push([storeId, storeData.name, storeData.address, reportedAt]);
  
          // 新しいデータをスプレッドシートに追加
          if (newData.length > 0) {
            reportSheet.getRange(2, 1, newData.length, 4).setValues(newData);
          }
  
          reportSheet.autoResizeColumns(1, 4);
          Logger.log(`${newData.length}件の補充情報をスプレッドシートに記録しました。`);
          logToSheet(logSheet, "updateReportInfoSheet", `${newData.length}件の補充情報をスプレッドシートに記録しました。`);
        }
      });
    } catch (error) {
      Logger.log('Firestoreからの補充情報取得中にエラーが発生しました: ' + error.message);
      logToSheet(logSheet, "updateReportInfoSheet", 'Firestoreからの補充情報取得中にエラーが発生しました: ' + error.message);
    }
  }
  
  function updateFirestoreIndividually(storeDataList) {
    if (!storeDataList || storeDataList.length === 0) {
      Logger.log('storeDataList が未定義または空です。');
      return; // データがない場合は処理を終了
    }
  
    const firestore = initializeFirestore();
    const cache = CacheService.getScriptCache();  // キャッシュも更新するために使用
  
    storeDataList.forEach((storeData) => {
      if (storeData && storeData.storeId) {
        // Firestoreへの個別更新
        firestore.updateDocument(`stores/${storeData.storeId}`, {
          name: storeData.name,
          address: storeData.address
        });
  
        // キャッシュも更新
        cache.put(storeData.storeId, JSON.stringify(storeData), 3600); // キャッシュは1時間保存
      } else {
        Logger.log('storeData または storeId が無効です。');
      }
    });
  
    Logger.log('Firestoreの更新が成功しました');
  }
  
  
  
  // 日次レポート送信関数
  async function sendDailyInventoryReport() {
    try {
      const firestore = initializeFirestore();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName("ログ記録"); // ログシート取得
      const reportSheet = ss.getSheetByName("報告情報");
  
      const today = new Date();
      today.setHours(0, 0, 0, 0);
  
      let lastRow = reportSheet.getLastRow();
      Logger.log(`現在の最終行: ${lastRow}`);
      logToSheet(logSheet, "sendDailyInventoryReport", `現在の最終行: ${lastRow}`);
  
      // Firestoreからドキュメントを取得するための変更
      const documents = firestore.getDocuments("restock-reports"); // collectionではなくgetDocumentsを使用
  
      if (!documents || documents.length === 0) {  
        Logger.log('Firestoreからのドキュメントが取得できませんでした。');
        logToSheet(logSheet, "sendDailyInventoryReport", 'Firestoreからのドキュメントが取得できませんでした。');
        return;
      }
  
      Logger.log(`取得したドキュメント数: ${documents.length}`);
      logToSheet(logSheet, "sendDailyInventoryReport", `取得したドキュメント数: ${documents.length}`);
      let reportContent = '';
      let newData = [];
  
      documents.forEach(async (doc) => {
        const storeId = doc.name; // ドキュメントIDを取得
        const reportedAtField = doc.fields.reportedAt; // reportedAtフィールドを取得
        const reportedAtJST = convertToJST(new Date(reportedAtField.timestampValue)); // タイムスタンプを日本標準時に変換
  
        Logger.log(`処理中の店舗ID: ${storeId}, reportedAtJST: ${reportedAtJST}`);
        logToSheet(logSheet, "sendDailyInventoryReport", `処理中の店舗ID: ${storeId}, reportedAtJST: ${reportedAtJST}`);
  
        if (reportedAtJST >= today) {
          const storeDoc = firestore.getDocument(`stores/${storeId}`);
          if (storeDoc && storeDoc.fields) {
            const storeName = storeDoc.fields.name.stringValue || '不明';
            const storeAddress = storeDoc.fields.address.stringValue || '不明';
  
            reportContent += `
  店舗ID: ${storeId}
  店舗名: ${storeName}
  所在地: ${storeAddress}
  報告日時: ${reportedAtJST.toLocaleString('ja-JP')}
  ---------------------`;
  
            newData.push([storeId, storeName, storeAddress, reportedAtJST.toLocaleString('ja-JP')]);
          }
        }
      });
  
      if (reportContent) {
        MailApp.sendEmail({
          to: APP_CONFIG.reportRecipientEmail,
          subject: `ゼロ在庫店舗日次報告 ${today.toLocaleDateString('ja-JP')}`,
          body: reportContent
        });
        logToSheet(logSheet, "sendDailyInventoryReport", `日次レポートが送信されました: ${today.toLocaleDateString('ja-JP')}`);
      } else {
        Logger.log("本日の報告はありません。");
        logToSheet(logSheet, "sendDailyInventoryReport", "本日の報告はありません。");
      }
    } catch (error) {
      MailApp.sendEmail({
        to: APP_CONFIG.reportRecipientEmail,
        subject: "日次在庫報告エラー",
        body: `レポート生成中にエラーが発生しました: ${error.message}`
      });
      Logger.log('エラーが発生しました: ' + error.message);
      const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ログ記録");
      logToSheet(logSheet, "sendDailyInventoryReport", 'エラーが発生しました: ' + error.message);
    }
  }
  
  
  