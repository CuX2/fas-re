// Firestoreに同期する処理
function syncCheckedStoreInfoToFirestore() {
  const firestore = initializeFirestore();  // Firestoreを初期化
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("ログ記録"); // ログシート取得
  const storeSheet = ss.getSheetByName("店舗情報");

  const lastRow = storeSheet.getLastRow();  // 最終行を取得
  const dataRange = storeSheet.getRange(2, 1, lastRow - 1, 5); // データ範囲を取得
  const data = dataRange.getValues();  // データを取得
  const checkboxes = storeSheet.getRange(2, 4, lastRow - 1, 1).getValues();  // チェックボックスの状態を取得

  Logger.log(`データ: ${JSON.stringify(data)}`);
  Logger.log(`チェックボックスの状態: ${JSON.stringify(checkboxes)}`);

  const now = new Date(); // 現在時刻
  let isAnyChecked = false;
  const storeDataList = [];

  // チェックされた行を確認し、同期リストに追加
  data.forEach((row, index) => {
    const storeId = row[0];
    const storeName = row[1];
    const storeAddress = row[2];
    const isChecked = checkboxes[index][0];

    Logger.log(`storeId: ${storeId}, storeName: ${storeName}, storeAddress: ${storeAddress}, isChecked: ${isChecked}`);

    if (isChecked) {
      isAnyChecked = true;
      storeDataList.push({ storeId, name: storeName, address: storeAddress });

      storeSheet.getRange(index + 2, 4).setValue(false);
      storeSheet.getRange(index + 2, 5).setValue(now);
    }
  });

  if (!isAnyChecked) {
    SpreadsheetApp.getUi().alert('同期する店舗が選択されていません。チェックボックスを選択してください。');
    logToSheet(logSheet, "syncCheckedStoreInfoToFirestore", '同期する店舗が選択されていません。');
    return;
  }

  if (storeDataList.length > 0) {
    updateFirestoreIndividually(storeDataList);
  } else {
    Logger.log("同期対象のデータがありません。");
    logToSheet(logSheet, "syncCheckedStoreInfoToFirestore", "同期対象のデータがありません。");
  }

  Logger.log('Firestore への更新が完了しました。');
  logToSheet(logSheet, "syncCheckedStoreInfoToFirestore", 'Firestore への更新が完了しました。');
}

// 個別にFirestoreを更新
function updateFirestoreIndividually(storeDataList) {
  const firestore = initializeFirestore();
  const cache = CacheService.getScriptCache();

  storeDataList.forEach((storeData) => {
    firestore.updateDocument(`stores/${storeData.storeId}`, {
      name: storeData.name,
      address: storeData.address
    });

    cache.put(storeData.storeId, JSON.stringify(storeData), 3600);
  });

  Logger.log('Firestoreの更新が成功しました');
}
