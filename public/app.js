// DOM elements
const storeInfoElement = document.getElementById('store-info');
const errorMessageElement = document.getElementById('error-message');
const reportButton = document.getElementById('report-button');
const thankYouElement = document.getElementById('thank-you');
const closeTabButton = document.getElementById('close-tab-button');
const manualInputElement = document.getElementById('manual-input');
const manualStoreIdInput = document.getElementById('manual-store-id');
const manualSubmitButton = document.getElementById('manual-submit-button');

// Initial setup for hiding elements
errorMessageElement.style.display = 'none';
reportButton.style.display = 'none';
thankYouElement.style.display = 'none';
manualInputElement.style.display = 'none';

// Get store ID from URL parameters
let urlParams = new URLSearchParams(window.location.search);
let storeId = urlParams.get('store');

// Display error message
function displayError(message) {
  errorMessageElement.textContent = message;
  errorMessageElement.style.display = 'block';
  reportButton.style.display = 'none';
  manualInputElement.style.display = 'block'; // QRコードの不備があった場合、手動入力を促す
}

// Display store information
function displayStoreInfo(storeData) {
  storeInfoElement.innerHTML = `
    <p class="mb-2"><strong class="font-semibold">店舗ID:</strong> ${storeId}</p>
    <p class="mb-2"><strong class="font-semibold">店舗名:</strong> ${storeData.name || '未設定'}</p>
    <p><strong class="font-semibold">住所:</strong> ${storeData.address || '未設定'}</p>
  `;
  reportButton.style.display = 'block';
}

// Report inventory
async function reportInventory() {
  try {
    reportButton.disabled = true;
    reportButton.textContent = '報告中...';

    // Save restock report using storeId as the document ID
    await db.collection('restock-reports').doc(storeId).set({
      storeId: storeId,
      reportedAt: firebase.firestore.FieldValue.serverTimestamp() // 現在のサーバータイムスタンプをセット
    });

    reportButton.style.display = 'none';
    thankYouElement.style.display = 'block';
  } catch (error) {
    console.error("Error adding document: ", error);
    displayError(`報告の送信中にエラーが発生しました: ${error.message}`);
  } finally {
    reportButton.disabled = false;
    reportButton.textContent = '最後の1冊を報告する';
  }
}

// Initialize the page
async function init() {
  if (!storeId) {
    displayError('店舗IDが指定されていません。URLに「?store=店舗ID」を追加するか、以下から入力してください。');
    return;
  }

  try {
    const doc = await db.collection('stores').doc(storeId).get();
    if (doc.exists) {
      displayStoreInfo(doc.data());
    } else {
      displayError(`店舗ID: ${storeId} の情報が見つかりません`);
    }
  } catch (error) {
    console.error("Error getting document:", error);
    displayError('店舗情報の取得中にエラーが発生しました。');
  }
}

// 手動で店舗IDを入力してリンクを更新する
manualSubmitButton.addEventListener('click', function() {
  const manualStoreId = manualStoreIdInput.value.trim();
  if (manualStoreId) {
    storeId = manualStoreId; // URLで渡されたstoreIdを手動入力のIDに変更
    window.history.replaceState(null, null, `?store=${storeId}`); // URLを更新
    init(); // 新しい店舗IDで再初期化
  } else {
    displayError('有効な店舗IDを入力してください。');
  }
});

// Event listener for report button
reportButton.addEventListener('click', reportInventory);

// Event listener for closing the tab
closeTabButton.addEventListener('click', function() {
  window.close();
});

// Initialize the page when DOM is fully loaded
document.addEventListener('DOMContentLoaded', init);
