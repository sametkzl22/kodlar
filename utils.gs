/***** GENEL YARDIMCILAR
 * Bu dosya; anahtar normalizasyonu, başlık haritalama, kod üretimi/parçalama,
 * tarih parse etme ve toplu işlem kilidi (reentrancy guard) fonksiyonlarını içerir.
 ********************************************************************************************/

/**
 * normalizeKey_
 * Kullanıcı girdilerini eşleştirme/indeksleme için normalize eder:
 * - null/undefined -> ""
 * - trim
 * - birden fazla boşluğu tek boşluğa indirger
 * - BÜYÜK HARFE çevirir
 * Örn: "  Abc   123  " -> "ABC 123"
 */
function normalizeKey_(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim().replace(/\s+/g, " ").toUpperCase();
}

/**
 * headerMap_
 * Verilen sayfanın 1. satırındaki başlıkları okuyup bir Map döner.
 * Map:  "BAŞLIK" -> kolonNo (1-indeksli)
 * - Boş başlıklar atlanır.
 * - Eşleştirme için başlıklar üst harfe çevrilir.
 */
function headerMap_(sheet) {
  const row = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = new Map();
  for (let c = 0; c < row.length; c++) {
    const key = String(row[c] || "").trim().toUpperCase();
    if (key) map.set(key, c + 1);
  }
  return map;
}

/**
 * headerIndexByName_
 * Tek bir başlık ismine göre kolon numarasını döndürür.
 * Bulunamazsa 0 döner.
 */
function headerIndexByName_(sheet, headerName) {
  const m = headerMap_(sheet);
  return m.get(String(headerName).trim().toUpperCase()) || 0;
}

/**
 * leftPadNumber_
 * Sayıyı soldan belirli genişliğe kadar “pad” karakteri ile doldurur.
 * Örn: leftPadNumber_(12, 4, "0") -> "0012"
 */
function leftPadNumber_(num, width, pad = "0") {
  const s = String(num);
  if (s.length >= width) return s;
  return new Array(width - s.length + 1).join(pad) + s;
}

/**
 * makeStockCode_
 * Prefiks + sayısal kısımdan stok kodu oluşturur.
 * Sayısal kısım STOK_KOD_NUM_PAD kadar sıfır dolgulu yazılır.
 * Örn: prefix="LAZ", number=10, pad=4 => "LAZ0010"
 */
function makeStockCode_(prefix, number) {
  const padded = leftPadNumber_(number, STOK_KOD_NUM_PAD, "0");
  return String(prefix || STOK_KOD_DEFAULT_PREFIX) + padded;
}

/**
 * splitStockCode_
 * Bir stok kodunu “prefix” ve “numeric tail” olarak parçalar.
 * Sondaki sayıyı STOK_KOD_NUM_REGEX ile bulur.
 * Örn: "LAZ0030" -> { prefix: "LAZ", number: 30 }
 * Eğer sonda sayı yoksa number=NaN döner ve tüm ifade prefix kabul edilir.
 */
function splitStockCode_(code) {
  const s = String(code || "");
  const m = s.match(STOK_KOD_NUM_REGEX);
  if (!m) return { prefix: s, number: NaN };
  const num = parseInt(m[1], 10);
  const prefix = s.slice(0, m.index);
  return { prefix, number: isNaN(num) ? NaN : num };
}

/**
 * parseDate_
 * Hücreden gelen farklı tipte tarihleri Date nesnesine çevirir.
 * Destek:
 *  - Date nesnesi → direkt döndürülür
 *  - Seri sayı (Excel/Sheets) → 1899-12-30 epoch’una göre hesaplanır
 *  - String → new Date(string) ile denenir
 * Başarısız olursa null döner.
 */
function parseDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;

  const s = String(v).trim();
  if (!s) return null;

  // Sayısal seri (ör. 45567 gibi)
  if (!isNaN(s)) {
    const n = Number(s);
    const epoch = new Date(1899, 11, 30); // Google Sheets/Excel epoch
    const d = new Date(epoch.getTime() + n * 24 * 60 * 60 * 1000);
    return isNaN(d.getTime()) ? null : d;
  }

  // Serbest metin tarih
  const tryD = new Date(s);
  return isNaN(tryD.getTime()) ? null : tryD;
}

/***** TOPLU İŞLEM KİLİDİ (onEdit reentrancy guard)
 * setBusy_/isBusy_ ile kritik bloklar sırasında onEdit içinde
 * yinelenen tetiklemelerin önüne geçilir.
 *******************************************************************/

/**
 * setBusy_
 * true -> kilidi açar (BUSY), false -> kilidi kapatır.
 */
function setBusy_(flag) {
  PropertiesService.getScriptProperties().setProperty("BUSY_FLAG", flag ? "1" : "");
}

/**
 * isBusy_
 * Şu anda kilit açık mı? (BUSY_FLAG == "1")
 */
function isBusy_() {
  return PropertiesService.getScriptProperties().getProperty("BUSY_FLAG") === "1";
}

function enforceUniqueCodeForNewProduct_(row) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const giris = ss.getSheetByName(SHEET_GIRIS);
  const stok  = ss.getSheetByName(SHEET_STOK);
  if (!giris || !stok || row < 2) return true;

  const codeRaw = giris.getRange(row, G_STOK_KODU).getValue();
  const codeKey = normalizeKey_(codeRaw);
  if (!codeKey) return true; // boşsa kontrol yok

  // Kodun işaret ettiği STOK satırı var mı?
  const idxObj  = buildStokDualIndexFast_();
  const stokRow = findStokRowByKeysFast_(idxObj, codeRaw, "");
  if (stokRow <= 0) return true; // STOK'ta yok => benzersiz, izin ver

  // GİRİŞ satırındaki 3’lü anahtar
  const catKey   = normalizeKey_(giris.getRange(row, G_KATEGORI).getValue());
  const markaKey = normalizeKey_(giris.getRange(row, G_MARKA).getValue());
  const modelKey = normalizeKey_(giris.getRange(row, G_MODEL).getValue());

  // Eğer 3’lü anahtar boşsa (kullanıcı kodla mevcut ürünü çağırmak istiyor olabilir) => izin ver
  if (!catKey || !markaKey || !modelKey) return true;

  // STOK’taki aynı kodun 3’lü anahtarı
  const sCat   = normalizeKey_(stok.getRange(stokRow, S_KATEGORI).getValue());
  const sMarka = normalizeKey_(stok.getRange(stokRow, S_MARKA).getValue());
  const sModel = normalizeKey_(stok.getRange(stokRow, S_MODEL).getValue());

  const sameProduct = (catKey === sCat && markaKey === sMarka && modelKey === sModel);
  if (sameProduct) return true; // aynı ürün için mevcut kodu kullanmak serbest

  // ——— ÇAKIŞMA: Kod başka bir üründe kullanılıyor
  try {
    SpreadsheetApp.getUi().alert(
      "Kod Çakışması",
      "Girdiğiniz kod zaten farklı bir üründe kullanılıyor: " +
      (stok.getRange(stokRow, S_MARKA).getValue() || "") + " / " +
      (stok.getRange(stokRow, S_MODEL).getValue() || "") +
      ". Lütfen farklı bir kod girin.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getActive().toast("Kod çakışması: Kod başka ürüne ait. Lütfen farklı bir kod girin.", "Uyarı", 6);
  }

  // Kodu geri al: önceki değeri bilemiyorsak temizlemek en güvenlisi
  giris.getRange(row, G_STOK_KODU).clearContent();
  return false;
}