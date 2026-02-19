// restock.gs - V5 (GÜVENLİ BAŞLANGIÇ STOĞU)

function processPendingIntakesForCode_(stokKoduRaw) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const giris = ss.getSheetByName(SHEET_GIRIS);
  const stok = ss.getSheetByName(SHEET_STOK);
  const key = normalizeKey_(stokKoduRaw);
  if (!key) return;

  const last = giris.getLastRow();
  const vals = giris.getRange(2, 1, last - 1, Math.max(G_ONAY, giris.getLastColumn())).getValues();
  const rowsToApprove = [];
  
  for (let i = 0; i < vals.length; i++) {
    if (normalizeKey_(vals[i][G_STOK_KODU - 1]) === key && !vals[i][G_ONAY - 1]) {
      rowsToApprove.push(2 + i);
    }
  }
  if (rowsToApprove.length === 0) return;

  const idxObj = buildStokDualIndexFast_();
  let stokRow = findStokRowByKeysFast_(idxObj, stokKoduRaw, "");
  
  if (stokRow <= 0) {
    stokRow = stok.getLastRow() + 1;
    stok.getRange(stokRow, S_STOK_KODU).setValue(stokKoduRaw);
    stok.getRange(stokRow, S_BASLANGIC).setValue(0); // İlk defa açılıyorsa 0
    yazFormul_(stokRow);
  }

  setBusy_(true);
  try {
    const now = new Date();
    rowsToApprove.forEach(row => {
      giris.getRange(row, G_ONAY).setValue(true);
      giris.getRange(row, G_GIRIS_TARIH).setValue(now).setNumberFormat("dd-mm-yyyy");
      if (typeof ensureLockedNote_ === 'function') ensureLockedNote_(giris, row);
    });
    updateStokGirisTarihiForCode_(stokKoduRaw, now);
    yazFormul_(stokRow); // I sütununa dokunmadan formülü tazeler
    SpreadsheetApp.flush();
  } finally { setBusy_(false); }
}

function recomputeStoreFromApproved_() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert("DİKKAT", "Stok formüllerini tüm satırlar için yeniden bağlayacağım. Başlangıç stoklarına dokunulmayacak. Devam?", ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  const stok = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STOK);
  const last = stok.getLastRow();
  setBusy_(true);
  try {
    for (let r = 2; r <= last; r++) { yazFormul_(r); }
    SpreadsheetApp.flush();
    ui.alert("Tüm stok formülleri güncellendi.");
  } finally { setBusy_(false); }
}

function processAllPendingExits_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cikis = ss.getSheetByName(SHEET_CIKIS);
  const last = cikis.getLastRow();
  const codes = new Set();
  const vals = cikis.getRange(2, 1, last - 1, X_ONAY).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (!vals[i][X_ONAY - 1]) {
      let k = normalizeKey_(vals[i][X_STOK_KODU - 1]);
      if (k) codes.add(k);
    }
  }
  codes.forEach(k => processPendingExitsForCode_(k));
}