async function sendDailyInventoryReport() {
  try {
    const firestore = initializeFirestore();  // Firestoreを初期化
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName("ログ記録"); // ログシート取得
    const reportSheet = ss.getSheetByName("報告情報");

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let lastRow = reportSheet.getLastRow();
    Logger.log(`現在の最終行: ${lastRow}`);
    logToSheet(logSheet, "sendDailyInventoryReport", `現在の最終行: ${lastRow}`);

    // Firestoreからドキュメントを取得
    const documents = firestore.getDocuments("restock-reports");

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
      const storeId = doc.name;
      const reportedAtField = doc.fields.reportedAt;
      const reportedAtJST = convertToJST(new Date(reportedAtField.timestampValue));

      Logger.log(`処理中の店舗ID: ${storeId}, reportedAtJST: ${reportedAtJST}`);
      logToSheet(logSheet, "sendDailyInventoryReport", `処理中の店舗ID: ${storeId}, reportedAtJST: ${reportedAtJST}`);

      // 日付だけで比較するために時間を無視
      const reportedAtDate = new Date(reportedAtJST);
      reportedAtDate.setHours(0, 0, 0, 0);

      Logger.log(`今日の日付: ${today}`);
      Logger.log(`報告された日付: ${reportedAtDate}`);
      Logger.log(`比較結果: ${reportedAtDate.getTime() === today.getTime()}`);

      if (reportedAtDate.getTime() === today.getTime()) {
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
