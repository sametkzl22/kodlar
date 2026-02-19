/***** ÖNGÖRÜ & OTOMATİK TAMAMLAMA & KOD ÜRETME (üçgensiz serbest giriş) *****/

/**
 * Bir sütunun (başlıktan aşağı) mevcut tüm data validation kurallarını temizler.
 * Amaç: Yeni dropdown/serbest giriş kuralı uygulamadan önce eski “katı” kuralları kaldırmak.
 */
function clearColumnValidation_(range) {
  try {
    const sh = range.getSheet();                 // Aralığın ait olduğu sayfa
    const col = range.getColumn();               // Hangi sütun?
    const startRow = 2;                          // 1. satır başlık; 2’den itibaren temizle
    const rows = Math.max(sh.getMaxRows() - startRow + 1, 0);
    if (rows > 0) sh.getRange(startRow, col, rows, 1).clearDataValidations();
  } catch (e) {}
}

/**
 * Verilen hücre aralığına serbest girişe izin veren bir dropdown uygular.
 * - values doluysa: listedeki seçenekler + liste dışına serbest giriş
 * - values boş/undefined ise: tamamen serbest (validation yok)
 * Not: allowInvalid(true) üçgen uyarı gösterebilir; suppress fonksiyonu ile lokal olarak kaldırıyoruz.
 */
function setDropdownAllowFree_(range, values) {
  try {
    clearColumnValidation_(range); // Sütun genelindeki önceki katı kuralları temizle
    if (values && values.length) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(values, true) // listedeki değerler + dropdown
        .setAllowInvalid(true)            // liste dışına yazmaya da izin ver
        .setHelpText("Listede yoksa yeni bir değer girebilirsiniz.")
        .build();
      range.setDataValidation(rule);
    } else {
      range.setDataValidation(null); // hiçbir doğrulama yok => tamamen serbest
    }
  } catch (e) {}
}

/**
 * Kullanıcı, dropdown listesinde olmayan bir değer girdiyse, yalnız o hücrenin validation'ını kaldır.
 * Böylece Google Sheets’in sarı üçgen uyarısı görünmez (sadece o hücre için).
 */
function suppressWarningIfFreeEntry_(range, listValues) {
  try {
    const cur = String(range.getValue() ?? "");
    if (!cur) return;                             // Boşsa iş yok
    const list = (listValues || []).map(v => String(v));
    const inList = list.includes(cur);
    if (!inList) {
      range.setDataValidation(null);              // Sadece bu hücrede doğrulamayı kapat
    }
  } catch (e) {}
}

/**
 * GİRİŞ satırındaki 3 alan (Kategori+Marka+Model) TAM dolmadan kod üretme/arama yapmayız.
 * Bu yardımcı, 3’lü anahtar eksikken A (STOK KODU)’nu temizleyerek “yapışma”yı engeller.
 */
function resetGirisCodeIfKeysIncomplete_(row) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const giris = ss.getSheetByName(SHEET_GIRIS);

  const kategori = normalizeKey_(giris.getRange(row, G_KATEGORI).getValue());
  const marka    = normalizeKey_(giris.getRange(row, G_MARKA).getValue());
  const model    = normalizeKey_(giris.getRange(row, G_MODEL).getValue());

  if (!kategori || !marka || !model) {
    giris.getRange(row, G_STOK_KODU).clearContent(); // <<< kritik: eksikken kodu temizle
  }
}

/**
 * 3’lü anahtar TAMAMLANDIĞINDA çalışan ana akıl:
 * 1) STOK’ta aynı (Kategori+Marka+Model) 3’lüsü varsa → kodu ve (boş hedef alanlara) detayları getir.
 * 2) Yoksa → ilgili kategorideki en büyük numaralı kodu bul, +CODE_STEP ile yeni ve benzersiz kod üret.
 */
function handleGirisAutocomplete_(row) {
  if (!AUTOCOMPLETE_ENABLED) return; // Genel ayardan kapatılabilir

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const giris = ss.getSheetByName(SHEET_GIRIS);

  // GİRİŞ hücre değerleri (ham)
  const kategori = String(giris.getRange(row, G_KATEGORI).getValue() || "");
  const marka    = String(giris.getRange(row, G_MARKA).getValue() || "");
  const model    = String(giris.getRange(row, G_MODEL).getValue() || "");
  const curCode  = String(giris.getRange(row, G_STOK_KODU).getValue() || "");

  // Normalize edilmiş anahtarlar (büyük harf, tek boşluk vb.)
  const catKey   = normalizeKey_(kategori);
  const markaKey = normalizeKey_(marka);
  const modelKey = normalizeKey_(model);

  // 3 alan TAM dolmadan asla işlem yapma (erken tetik ve yanlış eşleşmenin önüne geçer)
  if (!catKey || !markaKey || !modelKey) return;

  // STOK’tan indeksleri topla ve 3’lü anahtar oluştur
  const maps = buildCatalogMaps_();
  const cbmKey = catKey + "||" + markaKey + "||" + modelKey;

  // 1) Tam eşleşme varsa: kodu GİRİŞ!A’ya yaz ve boş hedef alanları STOK’tan doldur
  const found = maps.rowByCatBrandModel && maps.rowByCatBrandModel.get(cbmKey);
  if (found && found.row > 0) {
    if (curCode !== String(found.code || "")) {
      giris.getRange(row, G_STOK_KODU).setValue(String(found.code || ""));
    }
    try {
      // Hedef: GİRİŞ satırı; Kaynak: STOK’taki found.row; Sadece boş alanları doldur
      const headerMapG = headerMap_(giris);
      fillDetailsFromStokRowFast_(giris, row, found.row, headerMapG, null);
    } catch (e) {}
    return; // Eşleşme bulundu, bitti
  }

  // 2) Eşleşme yoksa: kategoride yeni kod öner (+CODE_STEP) ve global benzersizliğini kontrol et
  const newCode = suggestNewCodeForCategory_(catKey, maps.codesByCat, maps);
  if (newCode && curCode !== newCode) {
    giris.getRange(row, G_STOK_KODU).setValue(newCode);
  }
}

/**
 * Satır bazında 3’lü anahtar kontrolü:
 * - Eksikse: stok kodunu temizler ve biter.
 * - Tamamsa: otomatik tamamlama/üretme akışını (handleGirisAutocomplete_) çalıştırır.
 */
function guardAndTriggerAutocompleteForRow_(row) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const giris = ss.getSheetByName(SHEET_GIRIS);

  const kategori = normalizeKey_(giris.getRange(row, G_KATEGORI).getValue());
  const marka    = normalizeKey_(giris.getRange(row, G_MARKA).getValue());
  const model    = normalizeKey_(giris.getRange(row, G_MODEL).getValue());

  if (!kategori || !marka || !model) {
    // Tamamlanmadan önce girilmiş eski bir kod varsa, yapışmayı engellemek için sil
    const curCode = giris.getRange(row, G_STOK_KODU).getValue();
    if (curCode !== "" && curCode !== null) {
      giris.getRange(row, G_STOK_KODU).clearContent();
    }
    return; // 3’lü tamamlanmadan arama YOK
  }

  // 3’lü tamam: şimdi asıl işlevi çağır
  handleGirisAutocomplete_(row);
}

/**
 * Verilen kodun globalde kullanılıp kullanılmadığını kontrol eder.
 * Amaç: Yeni kod üretirken çakışmayı engellemek.
 */
function isCodeUsed_(code, maps) {
  const key = normalizeKey_(code);
  if (!key) return false;
  return (maps && ((maps.byStok && maps.byStok.has(key)) || (maps.bySirket && maps.bySirket.has(key))));
}

/**
 * Yeni stok kodu önericisi:
 * - İlgili kategorideki tüm kodları inceler, en büyük numarayı bulur.
 * - +CODE_STEP ekleyerek aday kod üretir.
 * - Aday kod globalde kullanılıyorsa +CODE_STEP ile arttırarak ilk boşluğu bulur.
 * - Kategoride hiç kod yoksa default prefix + STOK_KOD_START_FROM’tan başlar.
 */
function suggestNewCodeForCategory_(catKey, codesByCatMap, maps) {
  try {
    const list = (codesByCatMap && codesByCatMap.get(catKey)) ? codesByCatMap.get(catKey) : [];

    // Kategoride hiç kod yoksa: default prefix + start_from, benzersiz olacak şekilde artır
    if (!list || list.length === 0) {
      let num = STOK_KOD_START_FROM;
      let candidate = makeStockCode_(STOK_KOD_DEFAULT_PREFIX, num);
      while (isCodeUsed_(candidate, maps)) {
        num += CODE_STEP;
        candidate = makeStockCode_(STOK_KOD_DEFAULT_PREFIX, num);
      }
      return candidate;
    }

    // Kategorideki en büyük numaralı kodu ve onun prefix’ini bul
    let maxNumber = NaN;
    let prefixOfMax = STOK_KOD_DEFAULT_PREFIX;
    for (let i = 0; i < list.length; i++) {
      const { prefix, number } = splitStockCode_(String(list[i] || ""));
      if (!isNaN(number) && (isNaN(maxNumber) || number > maxNumber)) {
        maxNumber = number;
        prefixOfMax = prefix;
      }
    }

    // Başlangıç: (maxNumber || start_from) + CODE_STEP
    let base = isNaN(maxNumber) ? (STOK_KOD_START_FROM - CODE_STEP) : maxNumber;
    let num = base + CODE_STEP;
    let candidate = makeStockCode_(prefixOfMax, num);

    // Global benzersizlik kontrolü: çakışıyorsa +CODE_STEP ile boşluğu bul
    while (isCodeUsed_(candidate, maps)) {
      num += CODE_STEP;
      candidate = makeStockCode_(prefixOfMax, num);
    }
    return candidate;

  } catch (e) {
    // Her ihtimale karşı: default prefix ile güvenli fallback
    let num = STOK_KOD_START_FROM;
    let candidate = makeStockCode_(STOK_KOD_DEFAULT_PREFIX, num);
    while (isCodeUsed_(candidate, maps)) {
      num += CODE_STEP;
      candidate = makeStockCode_(STOK_KOD_DEFAULT_PREFIX, num);
    }
    return candidate;
  }
}

/**
 * GİRİŞ sayfasında ilgili satırın J (giriş tarihi) hücresini NOW yapar,
 * ve aynı kodu STOK LİSTESİ’nde K sütununa (son giriş tarihi) yansıtır.
 * Not: Mevcut akışta bu fonksiyon onEdit içinde otomatik çağrılmıyor; menü/işleme sırasında kullanılıyor.
 */
function touchGirisDateAndPropagate_(row) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const giris = ss.getSheetByName(SHEET_GIRIS);

  const stokKodu = String(giris.getRange(row, G_STOK_KODU).getValue() || "").trim();
  const now = new Date();

  setBusy_(true);
  try {
    // GİRİŞ > J (G_GIRIS_TARIH) hücresi
    giris.getRange(row, G_GIRIS_TARIH).setValue(now);
    giris.getRange(row, G_GIRIS_TARIH).setNumberFormat("dd-mm-yyyy"); // İstersen saat de eklenebilir

    // STOK LİSTESİ > K (son giriş) sütununa yansıt (kod varsa)
    if (stokKodu) {
      updateStokGirisTarihiForCode_(stokKodu, now);
    }
    SpreadsheetApp.flush();
  } finally {
    setBusy_(false);
  }
}

/**
 * ÇIKIŞ sayfasında ilgili satırın C (tarih) hücresini NOW yapar,
 * ve aynı kodun STOK LİSTESİ L sütunundaki (son çıkış) tarihini günceller.
 */
function touchCikisDateAndPropagate_(row) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const cikis = ss.getSheetByName(SHEET_CIKIS);

  const stokKodu = String(cikis.getRange(row, X_STOK_KODU).getValue() || "").trim();
  if (!stokKodu) return; // Kod yoksa iş yok

  const now = new Date();
  setBusy_(true);
  try {
    cikis.getRange(row, X_TARIH).setValue(now);
    cikis.getRange(row, X_TARIH).setNumberFormat("dd-mm-yyyy hh:mm"); // Saatli format
    updateStokCikisTarihiForCode_(stokKodu, now); // STOK L (S_CIKIS_TARIHI)
    SpreadsheetApp.flush();
  } finally {
    setBusy_(false);
  }
}

/**
 * ÇIKIŞ!C (tarih) hücresini kullanıcı manuel düzenlediyse, o değeri parse edip STOK’a yansıt.
 */
function propagateEditedCikisDate_(row) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const cikis = ss.getSheetByName(SHEET_CIKIS);

  const stokKodu = String(cikis.getRange(row, X_STOK_KODU).getValue() || "").trim();
  if (!stokKodu) return;

  const raw = cikis.getRange(row, X_TARIH).getValue();
  const dt  = parseDate_(raw);
  if (!dt) return; // Geçerli tarih değilse dokunma

  setBusy_(true);
  try {
    cikis.getRange(row, X_TARIH).setValue(dt);
    cikis.getRange(row, X_TARIH).setNumberFormat("dd-mm-yyyy hh:mm");
    updateStokCikisTarihiForCode_(stokKodu, dt);
    SpreadsheetApp.flush();
  } finally {
    setBusy_(false);
  }
}

// Seçili KODA ait TÜM bekleyen (K=false) girişleri işle:
// - Bekleyen miktarların TOPLAMINI bul
// - STOK'ta ilgili satırın I (Başlangıç) sütununa EKLE (depo et)
// - Satırları K=TRUE yap, tarihi yaz ve gerekirse STOK satırını oluştur
// - J formülünü tazele
function processPendingIntakesForCode_(stokKoduRaw) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const giris = ss.getSheetByName(SHEET_GIRIS);
  const stok  = ss.getSheetByName(SHEET_STOK);
  const key   = normalizeKey_(stokKoduRaw);
  if (!key) return;

  const last = giris.getLastRow();
  if (last < 2) return;

  const maxCol = Math.max(G_ONAY, giris.getLastColumn());
  const vals = giris.getRange(2, 1, last - 1, maxCol).getValues(); // A..K

  // ——— K=false olan bekleyen satırlar ve toplam adet
  const rowsToApprove = [];
  let pendingQty = 0;
  for (let i = 0; i < vals.length; i++) {
    const rowIdx   = 2 + i;
    const codeKey  = normalizeKey_(vals[i][G_STOK_KODU - 1]); // A
    const approved = !!vals[i][G_ONAY - 1];                   // K
    const qty      = Number(vals[i][G_ADET - 1]) || 0;        // B
    if (codeKey === key && !approved) {
      rowsToApprove.push(rowIdx);
      pendingQty += qty;
    }
  }
  if (rowsToApprove.length === 0) return;

  // ——— STOK satırını bul/oluştur
  const idxObj  = buildStokDualIndexFast_();
  let stokRow   = findStokRowByKeysFast_(idxObj, stokKoduRaw, "");
  if (stokRow <= 0) {
    stokRow = stok.getLastRow() + 1;
    stok.getRange(stokRow, S_STOK_KODU).setValue(stokKoduRaw);
    const sampleRow = rowsToApprove[0];
    [
      { g: G_KATEGORI, s: S_KATEGORI },
      { g: G_MARKA,    s: S_MARKA },
      { g: G_MODEL,    s: S_MODEL },
      { g: G_OZELLIK,  s: S_OZELLIK },
      { g: G_ACIKLAMA, s: S_ACIKLAMA },
      { g: G_BIRIM,    s: S_BIRIM },
      { g: G_RAF,      s: S_RAF }
    ].forEach(m => {
      const v = giris.getRange(sampleRow, m.g).getValue();
      if (v !== "" && v !== null) stok.getRange(stokRow, m.s).setValue(v);
    });
    yazFormul_(stokRow);
  }

  setBusy_(true);
  try {
    // I (Başlangıç) += bekleyen toplam
    const currentI = Number(stok.getRange(stokRow, S_BASLANGIC).getValue()) || 0;
    stok.getRange(stokRow, S_BASLANGIC).setValue(currentI + pendingQty);

    // Satırları TRUE yap + J tarih yaz + HEMEN KİLİTLE
    // ——— Bekleyen satırları TRUE yap, tarih at ve stok tarihine yansıt
    const now = new Date();
   rowsToApprove.forEach(row => {
    giris.getRange(row, G_ONAY).setValue(true);     // K = TRUE
    giris.getRange(row, G_GIRIS_TARIH).setValue(now);
    giris.getRange(row, G_GIRIS_TARIH).setNumberFormat("dd-mm-yyyy");
    ensureLockedNote_(giris, row);                  // <<< sadece NOT: "LOCKED"
  });

    updateStokGirisTarihiForCode_(stokKoduRaw, now);
    yazFormul_(stokRow);
    SpreadsheetApp.flush();
  } finally {
    setBusy_(false);
  }
}

/**
 * Sayfadaki tüm bekleyen (K=false) girişleri kod bazında gruplayıp sırayla işler.
 * Kullanım: “Tüm bekleyenleri işle (K=false)” menü eylemi.
 */
function processAllPendingIntakes_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const giris = ss.getSheetByName(SHEET_GIRIS);
  const last  = giris.getLastRow();
  if (last < 2) return;

  const maxCol = Math.max(G_ONAY, giris.getLastColumn());
  const vals   = giris.getRange(2, 1, last - 1, maxCol).getValues();

  // Aynı koda ait bekleyen satırları tek seferde işlemek için benzersiz kod kümesi üret
  const codes = new Set();
  for (let i = 0; i < vals.length; i++) {
    const approved = !!vals[i][G_ONAY - 1];
    if (!approved) {
      const key = normalizeKey_(vals[i][G_STOK_KODU - 1]);
      if (key) codes.add(key);
    }
  }
  // Her kod için bekleyenleri işle
  codes.forEach(k => processPendingIntakesForCode_(k));
}
// checkbox locked
function ensureLockedNote_(sh, row) {
  const cell = sh.getRange(row, G_ONAY);
  const note = cell.getNote() || "";
  if (!/LOCKED/i.test(note)) cell.setNote("LOCKED");
}
function ensureLockedNoteAt_(sh, row, col) {
  const cell = sh.getRange(row, col);
  const note = cell.getNote() || "";
  if (!/LOCKED/i.test(note)) cell.setNote("LOCKED");
}
// GİRİŞ’te tek bir satırı, yazılan stok koduna göre STOK’tan doldurur (adet hariç).
function fillGirisDetailsByCode_(row) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const giris = ss.getSheetByName(SHEET_GIRIS);
  const stok  = ss.getSheetByName(SHEET_STOK);
  if (!giris || !stok || row < 2) return;

  const code = String(giris.getRange(row, G_STOK_KODU).getValue() || "").trim();
  if (!code) return;

  // STOK’ta hızlı arama için index
  const idxObj  = buildStokDualIndexFast_();
  const stokRow = findStokRowByKeysFast_(idxObj, code, "");
  if (stokRow <= 0) return;

  // Hedef: GİRİŞ satırı; Kaynak: STOK’ta bulunan satır
  // Not: ADET (G_ADET) zaten FIELDS_TO_COPY_FROM_STOK listesinde yok — üzerine yazılmaz.
  const headerMapG = headerMap_(giris);
  fillDetailsFromStokRowFast_(giris, row, stokRow, headerMapG, idxObj);
}

function SISTEM_RESET() {
  // Meşgul bayrağını zorla indir
  PropertiesService.getScriptProperties().deleteProperty("BUSY_FLAG");
  
  // Kullanıcıya bilgi ver
  SpreadsheetApp.getUi().alert("✅ Sistem kilidi kaldırıldı! Artık sorgulama yapabilirsin.");
}

/**
 * Sağ taraftaki HTML kenar çubuğunu açar.
 */
function openSidebarPanel() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Stok Kontrol Paneli')
    .setWidth(350); // Genişlik sabitlenmiştir
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * --- İSTEMCİ (HTML) API FONKSİYONLARI ---
 * Aşağıdaki fonksiyonlar Sidebar.html içinden google.script.run ile çağrılır.
 */

// 1. Ürün Arama Fonksiyonu
function clientSearchProduct(code) {
  try {
    if (!code) return { success: false, message: "Kod boş olamaz." };
    
    // Mevcut yardımcı fonksiyonlarını kullanıyoruz
    const idxObj = buildStokDualIndexFast_();
    const stokRow = findStokRowByKeysFast_(idxObj, code, "");
    
    if (stokRow <= 0) {
      return { success: false, found: false, message: "Ürün bulunamadı." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stokSh = ss.getSheetByName(SHEET_STOK);
    
    // Stok sayfasından verileri çek
    const data = {
      code: stokSh.getRange(stokRow, S_STOK_KODU).getValue(),
      marka: stokSh.getRange(stokRow, S_MARKA).getValue(),
      model: stokSh.getRange(stokRow, S_MODEL).getValue(),
      ozellik: stokSh.getRange(stokRow, S_OZELLIK).getValue(),
      birim: stokSh.getRange(stokRow, S_BIRIM).getValue(),
      raf: stokSh.getRange(stokRow, S_RAF).getValue(),
      stok: stokSh.getRange(stokRow, S_GUNCEL).getValue()
    };
    
    return { success: true, found: true, data: data };
    
  } catch (e) {
    return { success: false, message: "Hata: " + e.message };
  }
}

// 2. Hızlı İşlem (Giriş/Çıkış) Fonksiyonu



function clientQuickTransaction(type, code, amount, projeAdi) { // <--- projeAdi eklendi
  if (isBusy_()) {
    return { success: false, message: "Sistem şu an meşgul, lütfen bekleyin." };
  }

  try {
    const qty = Number(amount);
    if (!code || qty <= 0) return { success: false, message: "Kod veya miktar geçersiz." };

    if (type === 'CIKIS' || type.indexOf('ÇIKIŞ') !== -1) {
      
      const idxObj = buildStokDualIndexFast_();
      const stokRow = findStokRowByKeysFast_(idxObj, code, "");
      
      if (stokRow <= 0) return { success: false, message: "Bu ürün stok listesinde bulunamadı." };

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const stokSh = ss.getSheetByName(SHEET_STOK);
      const currentStock = Number(stokSh.getRange(stokRow, S_GUNCEL).getValue()) || 0;

      if (qty > currentStock) {
        return { 
          success: false, 
          message: `❌ YETERSİZ STOK!\nMevcut: ${currentStock}\nÇıkılmak İstenen: ${qty}\nİşlem İptal Edildi.` 
        };
      }

      // Proje adını buraya iletiyoruz
      apiOutboundInternal_(code, qty, projeAdi);
      
      return { success: true, message: `${qty} adet ÇIKIŞ yapıldı. Kalan: ${currentStock - qty}` };
    } else {
      // Giriş işleminde proje adı genelde olmaz, boş gönderiyoruz
      apiInboundInternal_(code, qty);
      return { success: true, message: `${qty} adet GİRİŞ yapıldı.` };
    }
  } catch (e) {
    return { success: false, message: "İşlem Hatası: " + e.message };
  }
}

// 3. Yeni Kart Oluşturma Fonksiyonu
function clientCreateCard(form) {
  if (isBusy_()) return { success: false, message: "Sistem meşgul." };

  try {
    // dashboard.gs'deki fonksiyonu kullan
    apiCreateInternal_(form);
    return { success: true, message: `Kart açıldı: ${form.code}` };
  } catch (e) {
    return { success: false, message: "Hata: " + e.message };
  }
}

// 4. Kategoriye Göre Otomatik Kod Önerisi
function clientSuggestCode(category) {
  try {
    const maps = buildCatalogMaps_();
    const catKey = normalizeKey_(category);
    // Kod.gs içindeki mantığı kullanıyoruz
    const newCode = suggestNewCodeForCategory_(catKey, maps.codesByCat, maps);
    return { success: true, code: newCode };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Kategori Listesini Frontend'e gönderir
function clientGetCategories() {
  // dashboard.gs'deki sabitleri kullanıyoruz
  return { 
    kategoriler: KATEGORI_LISTESI, 
    birimler: BIRIM_LISTESI 
  };
}

// kod.gs dosyasının EN ALTINA ekle:

/**
 * 5. Detaylı Arama Fonksiyonu (Marka/Model ile)
 */
function clientSearchByDetails(kategori, marka, model) {
  try {
    // 1. Girdileri normalize et (büyük harf, boşluk temizle)
    const catKey = normalizeKey_(kategori);
    const brandKey = normalizeKey_(marka);
    const modelKey = normalizeKey_(model);

    if (!catKey || !brandKey || !modelKey) {
      return { success: false, message: "Lütfen Kategori, Marka ve Model alanlarını doldurun." };
    }

    // 2. Katalog haritasını oluştur
    const maps = buildCatalogMaps_();
    
    // 3. Üçlü anahtarı oluştur: KAT||MARKA||MODEL
    const key = catKey + "||" + brandKey + "||" + modelKey;

    // 4. Haritada var mı bak
    if (maps.rowByCatBrandModel && maps.rowByCatBrandModel.has(key)) {
      const found = maps.rowByCatBrandModel.get(key);
      // Bulunduysa, o ürünün kodunu alıp mevcut 'clientSearchProduct' fonksiyonuna yönlendiriyoruz.
      // Böylece aynı veri formatını (stok, raf vb.) tekrar yazmamıza gerek kalmıyor.
      return clientSearchProduct(found.code);
    } else {
      return { success: false, found: false, message: "Bu kriterlere uyan ürün bulunamadı." };
    }

  } catch (e) {
    return { success: false, message: "Arama Hatası: " + e.message };
  }
}