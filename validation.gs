/***** ZORUNLU ALAN KONTROLÜ & VURGULAR (opsiyonel) *****/

// GİRİŞ sayfasındaki belirli bir satırın zorunlu alanlarını kontrol eder
// Eksik kod varsa hücre arka planını sarı (#FFF2CC) yapar; tamamsa sıfırlar
function validateGirisRow_(row) {
// GİRİŞ sayfasını al
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sh = ss.getSheetByName(SHEET_GIRIS);

// Satırın stok kodunu oku
const code = sh.getRange(row, G_STOK_KODU).getValue();

// Tüm satır aralığını seç (1. sütundan son sütuna kadar)
const rng = sh.getRange(row, 1, 1, sh.getLastColumn());

// Eğer stok kodu boşsa, arka planı sarıya boya (eksik alan uyarısı)
if (!code) {
rng.setBackground("#FFF2CC"); // açık sarı vurgulama
} else {
rng.setBackground(null); // arka planı sıfırla
}
}