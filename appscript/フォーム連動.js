  // Firestoreの初期化（認証情報を必要に応じて記載）
  function initializeFirestore() {
    return FirestoreApp.getFirestore(
      FIREBASE_CONFIG.client_email,
      FIREBASE_CONFIG.private_key,
      FIREBASE_CONFIG.project_id
    );
  }

// スプレッドシートの「フォームの回答」シートからデータを取得して処理する関数
function onFormSubmit() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const formSheet = ss.getSheetByName("フォームの回答");
    const storeSheet = ss.getSheetByName("店舗情報");
  
    // ログシートを取得
    const logSheet = ss.getSheetByName("ログ記録");
  
    // フォームの回答シートのデータを取得
    const lastFormRow = formSheet.getLastRow();
    if (lastFormRow < 2) {
      logToSheet(logSheet, "onFormSubmit", "フォームの回答がありません");
      return;
    }
  
    const formData = formSheet.getRange(2, 1, lastFormRow - 1, 5).getValues(); // ヘッダーを除くデータ
    const firestore = initializeFirestore(); // Firestoreの初期化
    let storeCounter = 1; // 店舗IDの連番を開始
  
    // 店舗情報シートの最終行を取得
    const lastStoreRow = storeSheet.getLastRow();
    const storeIds = storeSheet.getRange(2, 1, lastStoreRow - 1, 1).getValues().flat(); // ヘッダーを除く店舗ID列
  
    formData.forEach(row => {
      const timestamp = row[0];  // A列: タイムスタンプ
      const storeName = row[1];  // B列: 店舗名
      const location = row[2];   // C列: 所在地
      const installationFrequency = row[3];  // D列: 何回目の設置ですか？
      const continuityFeasibility = row[4];  // E列: 継続設置は可能か
  
      // 店舗IDを生成
      const installationCode = installationFrequency === '初めて' ? '1' : installationFrequency === '複数回' ? '2' : '3';
      const continuityCode = continuityFeasibility === '自分が設置にいけば' ? '1' : continuityFeasibility === '別のメンバーが行っても可能' ? '2' : '3';
      const storeId = `${installationCode}${continuityCode}${String(storeCounter).padStart(3, '0')}`;
  
      // Firestore内に同じIDのドキュメントが存在するかを確認
      if (checkIfDocumentExists(firestore, `stores/${storeId}`)) {
        logToSheet(logSheet, "onFormSubmit", `店舗ID: ${storeId} は既にFirestoreに存在しています。処理をスキップします。`);
        storeCounter++;  // スキップした場合もカウンタは進める
        return;
      }
  
      // 空白の行を探してデータを追加
      let foundEmptyRow = false;
      for (let i = 0; i < storeIds.length; i++) {
        if (!storeIds[i]) { // 店舗IDが空白の場合
          storeSheet.getRange(i + 2, 1).setValue(storeId);  // 店舗IDをA列に追加
          storeSheet.getRange(i + 2, 2).setValue(storeName);  // 店舗名をB列に追加
          storeSheet.getRange(i + 2, 3).setValue(location);   // 所在地をC列に追加
          foundEmptyRow = true;
          break;
        }
      }
  
      // 空の行が見つからなかった場合は、最後の行に追加
      if (!foundEmptyRow) {
        const newRow = lastStoreRow + 1;
        storeSheet.getRange(newRow, 1).setValue(storeId);  // 店舗IDをA列に追加
        storeSheet.getRange(newRow, 2).setValue(storeName);  // 店舗名をB列に追加
        storeSheet.getRange(newRow, 3).setValue(location);   // 所在地をC列に追加
      }
  
      // Firestoreにデータを追加
      try {
        firestore.createDocument(`stores/${storeId}`, {
          name: storeName,
          address: location,
          installationFrequency: installationFrequency,
          continuityFeasibility: continuityFeasibility,
          timestamp: timestamp
        });
        logToSheet(logSheet, "onFormSubmit", `店舗ID: ${storeId} がFirestoreに追加されました。`);
      } catch (error) {
        logToSheet(logSheet, "onFormSubmit", `Firestoreへのデータ追加中にエラーが発生しました: ${error.message}`);
      }
  
      storeCounter++; // 店舗IDの連番をインクリメント
    });
  }
  
  function logToSheet(logSheet, functionName, logMessage) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    logSheet = ss.getSheetByName("ログ記録") || ss.insertSheet("ログ記録");
  
    const timestamp = new Date();
    logSheet.appendRow([timestamp, functionName, logMessage]);
  }
  // Firestore内にドキュメントが存在するかどうかを確認する関数
  function checkIfDocumentExists(firestore, docPath) {
    try {
      const doc = firestore.getDocument(docPath);
      return !!doc;  // ドキュメントが存在すればtrue、存在しなければfalse
    } catch (error) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName("ログ記録");
      logToSheet(logSheet, "checkIfDocumentExists", `ドキュメント確認中にエラーが発生しました: ${error.message}`);
      return false;  // エラーが発生した場合もfalseを返す
    }
  }
  
  // シリアルナンバーを生成する関数
  function generateSerialNumber() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('店舗情報');
    const lastRow = sheet.getLastRow();
    const lastStoreId = sheet.getRange(lastRow, 1).getValue();  // A列に店舗IDがあると仮定
    const lastSerial = parseInt(lastStoreId.slice(-3)) || 0;  // 末尾3桁を連番と見なす
    const newSerial = (lastSerial + 1).toString().padStart(3, '0');  // 連番を生成し3桁にフォーマット
    return newSerial;
  }
  
