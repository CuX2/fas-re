  // Firestore の初期化
  function initializeFirestore() {
    return FirestoreApp.getFirestore(
      FIREBASE_CONFIG.client_email,
      FIREBASE_CONFIG.private_key,
      FIREBASE_CONFIG.project_id
    );
  }

// カスタムHTMLダイアログを作成
function showCustomAlert() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('AlertDialog')
      .setWidth(300)
      .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '削除確認');
  }
  
  // チェックされた店舗情報を削除する関数
  function deleteCheckedStoreInfo() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const storeSheet = ss.getSheetByName("店舗情報");
  
    const lastRow = storeSheet.getLastRow(); // 最終行を取得
    const checkboxes = storeSheet.getRange(2, 4, lastRow - 1, 1).getValues(); // D列のチェックボックスの状態を取得（D列が4番目）
  
    let selectedRows = [];
    
    // チェックボックスが選択されている行を取得
    for (let i = 0; i < checkboxes.length; i++) {
      if (checkboxes[i][0]) { // チェックが入っている場合
        selectedRows.push(i + 2); // 行番号を保存（ヘッダー行を考慮して +2）
      }
    }
  
    // チェックされている行がなければ警告ダイアログを表示
    if (selectedRows.length === 0) {
      SpreadsheetApp.getUi().alert('削除する店舗が選択されていません。');
      return;
    }
  
    // カスタムHTMLダイアログを表示して削除確認
    showCustomAlert();
  }
  
  // HTML ダイアログから呼ばれる関数
  function confirmDelete() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const storeSheet = ss.getSheetByName("店舗情報");
  
    const lastRow = storeSheet.getLastRow(); // 最終行を取得
    const checkboxes = storeSheet.getRange(2, 4, lastRow - 1, 1).getValues(); // D列のチェックボックスの状態を取得
    const firestore = initializeFirestore(); // Firestoreを初期化
    
    let deleteCount = 0; // 削除された行数をカウントする
    let rowsToDelete = []; // 削除予定の行を保持
  
    // チェックされている行の店舗情報を削除
    for (let i = 0; i < checkboxes.length; i++) {
      if (checkboxes[i][0]) { // チェックが入っている場合
        const storeId = storeSheet.getRange(i + 2, 1).getValue(); // A列からstoreIdを取得
  
        if (storeId) {
          // Firestoreのstoresコレクションから該当する店舗情報を削除
          try {
            firestore.deleteDocument(`stores/${storeId}`);
            Logger.log(`ドキュメント削除: ${storeId}`);
            deleteCount++; // 削除カウントを増やす
            rowsToDelete.push(i + 2); // 削除する行番号を記録
          } catch (error) {
            Logger.log(`Firestore削除エラー: ${error.message}`);
          }
        } else {
          Logger.log(`storeId が見つかりません: 行 ${i + 2}`);
        }
      }
    }
  
    // 行削除処理を実行
    if (deleteCount > 0) {
      rowsToDelete.reverse().forEach(row => {
        storeSheet.deleteRow(row);
      });
  
      const ui = SpreadsheetApp.getUi();
      ui.alert(`${deleteCount} 件のデータが削除されました。`);
    } else {
      Logger.log('削除するデータがありませんでした。');
      SpreadsheetApp.getUi().alert('削除するデータがありませんでした。');
    }
  }
  
