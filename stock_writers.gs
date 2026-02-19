/***** STOK YAZICILAR & DETAY KOPYALAYICILAR – HIZLI
 * Bu dosya; STOK sayfasından hedef satırlara alan kopyalama, kod ile satır
 * bulup detay doldurma ve “GÜNCEL (J)” formülünü yazma işlemlerini içerir.
 * Notlar:
 * - FILL_ONLY_IF_EMPTY=true iken yalnız boş hedef hücrelere yazar.
 * - “pushGirisToStok_” mevcut stok satırlarını ASLA güncellemez (politikaya bağlı).
 ********************************************************************************************/

/**
 * STOK sayfasındaki bir satırdan, hedef sayfadaki bir satıra alanları kopyalar.
 * - Hedef: targetSheet/targetRow
 * - Kaynak: STOK/stokRow
 * - Hangi kolonlar?: FIELDS_TO_COPY_FROM_STOK listesindeki başlıklara göre
 * - FILL_ONLY_IF_EMPTY: true ise hedefte boş olan hücrelere yazar; false’ta üzerine yazar.
 */
function fillDetailsFromStokRowFast_(targetSheet, targetRow, stokRow, headerMapTarget, idxObj) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const stok = ss.getSheetByName(SHEET_STOK);

  const lastCol = targetSheet.getLastColumn();
  const rowVals = targetSheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];
  let changed = false;

  FIELDS_TO_COPY_FROM_STOK.forEach(f => {
    const c = headerMapTarget.get(String(f.header).toUpperCase()); // hedefte bu başlığın olduğu kolon
    if (c > 0) {
      const currentVal = rowVals[c - 1];
      const newVal = stok.getRange(stokRow, f.stokCol).getValue(); // STOK’tan oku
      const shouldWrite = FILL_ONLY_IF_EMPTY ? (currentVal === "" || currentVal === null) : true;
      if (shouldWrite && String(currentVal) !== String(newVal)) {
        rowVals[c - 1] = newVal;
        changed = true;
      }
    }
  });

  if (changed) {
    targetSheet.getRange(targetRow, 1, 1, lastCol).setValues([rowVals]);
  }
}

/**
 * Seçili aralıktaki satırları (aktif sayfada) dolaşır; her satır için
 * stok/şirket kodu ile STOK’ta satırı bulur ve detayları (boş hücrelere) kopyalar.
 * Kullanım: Menüden “Seçili aralığı doldur (stok+şirket kodu)”.
 */
function autofillSelectionByKeys_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const sel = sh.getActiveRange();
  if (!sel) return;

  // Hedef sayfanın başlık haritası: “başlık -> kolon no”
  const headerMapT = headerMap_(sh);

  // “STOK KODU” ve “ŞİRKET KODU” kolonlarını başlığa göre esnek bul
  const stokCol   = (ALT_KEY_HEADERS.stokKodu  .map(h => headerMapT.get(h.toUpperCase())).find(x => x > 0)) || 0;
  const sirketCol = (ALT_KEY_HEADERS.sirketKodu.map(h => headerMapT.get(h.toUpperCase())).find(x => x > 0)) || 0;
  if (!stokCol && !sirketCol) return; // Hiçbiri yoksa işlem yapamayız

  // STOK’ta hızlı arama için ikili index
  const idxObj = buildStokDualIndexFast_();

  // Seçimin ilk satırından itibaren tüm satırları oku
  const r1 = sel.getRow();
  const vals = sh.getRange(r1, 1, sel.getNumRows(), sh.getLastColumn()).getValues();

  setBusy_(true);
  try {
    for (let i = 0; i < vals.length; i++) {
      const row = r1 + i;
      const stokKodu   = stokCol   ? vals[i][stokCol - 1]   : "";
      const sirketKodu = sirketCol ? vals[i][sirketCol - 1] : "";
      const stokRow = findStokRowByKeysFast_(idxObj, stokKodu, sirketKodu);
      if (stokRow > 0) fillDetailsFromStokRowFast_(sh, row, stokRow, headerMapT, idxObj);
    }
    SpreadsheetApp.flush();
  } finally {
    setBusy_(false);
  }
}

/**
 * Yalnız aktif satırı doldurur. (Seçim gerekmez.)
 * Kullanım: Menüden “Aktif satırı doldur (stok+şirket kodu)”.
 */
function autofillActiveRowByKeys_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const row = sh.getActiveCell().getRow();
  if (row < 2) return; // başlık satırı yok

  const headerMapT = headerMap_(sh);
  const stokCol   = (ALT_KEY_HEADERS.stokKodu  .map(h => headerMapT.get(h.toUpperCase())).find(x => x > 0)) || 0;
  const sirketCol = (ALT_KEY_HEADERS.sirketKodu.map(h => headerMapT.get(h.toUpperCase())).find(x => x > 0)) || 0;
  if (!stokCol && !sirketCol) return;

  const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const idxObj = buildStokDualIndexFast_();

  setBusy_(true);
  try {
    const stokRow = findStokRowByKeysFast_(idxObj,
      stokCol   ? vals[stokCol - 1]   : "",
      sirketCol ? vals[sirketCol - 1] : ""
    );
    if (stokRow > 0) fillDetailsFromStokRowFast_(sh, row, stokRow, headerMapT, idxObj);
    SpreadsheetApp.flush();
  } finally {
    setBusy_(false);
  }
}

/**
 * Tüm sayfayı (adı verilen) parçalar halinde dolaşır ve her parçayı doldurur.
 * Büyük veri setlerinde zaman aşımına düşmemek için kullanışlıdır.
 * @param {string} sheetName  Hedef sayfa adı
 * @param {number} startRow   Başlangıç satırı (varsayılan 2)
 * @param {number} chunkSize  Parça boyutu (varsayılan 300)
 */
function autofillAllByKeysChunked_(sheetName, startRow, chunkSize) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  const last = sh.getLastRow();

  const headerMapT = headerMap_(sh);
  const stokCol   = (ALT_KEY_HEADERS.stokKodu  .map(h => headerMapT.get(h.toUpperCase())).find(x => x > 0)) || 0;
  const sirketCol = (ALT_KEY_HEADERS.sirketKodu.map(h => headerMapT.get(h.toUpperCase())).find(x => x > 0)) || 0;
  if (!stokCol && !sirketCol) return;

  const idxObj = buildStokDualIndexFast_();
  let r = Math.max(startRow || 2, 2);
  const end = last;

  setBusy_(true);
  try {
    while (r <= end) {
      const n = Math.min(chunkSize || 300, end - r + 1);
      const vals = sh.getRange(r, 1, n, sh.getLastColumn()).getValues();

      for (let i = 0; i < n; i++) {
        const row = r + i;
        const stokKodu   = stokCol   ? vals[i][stokCol - 1]   : "";
        const sirketKodu = sirketCol ? vals[i][sirketCol - 1] : "";
        const stokRow = findStokRowByKeysFast_(idxObj, stokKodu, sirketKodu);
        if (stokRow > 0) fillDetailsFromStokRowFast_(sh, row, stokRow, headerMapT, idxObj);
      }
      SpreadsheetApp.flush();
      r += n; // sıradaki parçaya geç
    }
  } finally {
    setBusy_(false);
  }
}

/*** STOK adet ve tarih (geriye uyum) *********************************************************/

/**
 * Formül argüman ayırıcıyı yerel ayara göre belirler.
 * - İngilizce benzeri yereller: “,” (virgül)
 * - TR gibi yereller: “;” (noktalı virgül)
 */
function getArgSep_() {
  const loc = (SpreadsheetApp.getActive().getSpreadsheetLocale() || "").toLowerCase();
  return (/^(en|ja|zh|ko)/.test(loc)) ? "," : ";";
}

// Dil ayırıcısına duyarlı argüman ayırıcı (en/tr vs.)
function getArgSep_() {
  const loc = (SpreadsheetApp.getActive().getSpreadsheetLocale() || "").toLowerCase();
  return (/^(en|ja|zh|ko)/.test(loc)) ? "," : ";";
}



/**
 * STOK!J (GÜNCEL) formülünü satır bazında yazar:
 * J_r = I_r (Sabit Başlangıç) 
 * + SUMIF(GİRİŞ!A:A, C_r, GİRİŞ!B:B)  <-- Tüm Girişler
 * - SUMIF(ÇIKIŞ!A:A, C_r, ÇIKIŞ!B:B)  <-- Tüm Çıkışlar
 */
function yazFormul_(rowNum) {
  const stok = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STOK);
  const sep  = getArgSep_();

  // J = I (Sabit) + (Tüm GİRİŞ) - (Tüm ÇIKIŞ)
  const f =
    "=I" + rowNum +
    // GİRİŞ: Onay (K) durumuna bakmaksızın o koda ait TÜM adetleri topla
    "+IFERROR(SUMIF('GİRİŞ'!A:A" + sep + "C" + rowNum + sep + "'GİRİŞ'!B:B)" + sep + "0)" +
    // ÇIKIŞ: O koda ait TÜM çıkışları düş
    "-IFERROR(SUMIF('ÇIKIŞ'!A:A" + sep + "C" + rowNum + sep + "'ÇIKIŞ'!B:B)" + sep + "0)";

  const rng = stok.getRange(rowNum, S_GUNCEL);
  rng.setFormula(f);
  rng.setNumberFormat("#,##0");
}

/**
 * ÇIKIŞ sayfasında bir satır için ÜRÜN (Marka / Model) hücresini doldurur.
 * Kaynak: GİRİŞ (A=Stok Kodu, D=Marka, E=Model)
 */
function fillCikisUrunForRow_(row) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const cikis  = ss.getSheetByName(SHEET_CIKIS);
  const stok   = ss.getSheetByName(SHEET_STOK);
  const giris  = ss.getSheetByName(SHEET_GIRIS); // fallback için
  if (!cikis || row < 2) return;

  const codeRaw   = cikis.getRange(row, X_STOK_KODU).getValue(); // ÇIKIŞ!A
  const codeKey   = normalizeKey_(codeRaw);
  const targetCol = (typeof X_URUN !== "undefined" ? X_URUN : 4); // ÇIKIŞ!D (ÜRÜN)

  // Kod boşsa: ürünü temizle ve çık
  if (!codeKey) {
    cikis.getRange(row, targetCol).clearContent();
    return;
  }

  let product = "";

  // 1) ÖNCE STOK LİSTESİ'NDEN ÇEK (asıl kaynak)
  if (stok) {
    const idxObj  = buildStokDualIndexFast_();
    const stokRow = findStokRowByKeysFast_(idxObj, codeRaw, "");
    if (stokRow > 0) {
      const marka = stok.getRange(stokRow, S_MARKA).getValue() || "";
      const model = stok.getRange(stokRow, S_MODEL).getValue() || "";
      product = String(marka).trim() + (marka && model ? " / " : "") + String(model).trim();
    }
  }

  // 2) Fallback: STOK’ta bulunamazsa GİRİŞ’ten ilk eşleşmeyi dene
  if (!product && giris) {
    const last = giris.getLastRow();
    if (last >= 2) {
      const vals = giris.getRange(2, 1, last - 1, 5).getValues(); // A..E
      for (let i = 0; i < vals.length; i++) {
        if (normalizeKey_(vals[i][0]) === codeKey) { // A=kod
          const marka = vals[i][3] || "";           // D=marka
          const model = vals[i][4] || "";           // E=model
          product = String(marka).trim() + (marka && model ? " / " : "") + String(model).trim();
          break;
        }
      }
    }
  }

  cikis.getRange(row, targetCol).setValue(product);
}

/**
 * Tüm STOK satırlarına (2’den son satıra kadar) J formülünü yeniden uygular.
 * Kullanım: Menüden “Tüm stokları güncelle”.
 */
/**
 * Tüm STOK satırlarına (2’den son satıra kadar) J formülünü yeniden uygular.
 * Kullanım: Menüden “Tüm stokları güncelle”.
 * OPTİMİZASYON: Tek tek yazmak yerine formülleri bellekte oluşturup toplu yazar.
 */
function reapplyGuncelToAll_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STOK);
  const last = sh.getLastRow();
  if (last < 2) return;

  const formulas = [];
  const sep = getArgSep_();

  // Formül şablonunu hazırla (satır numarası dinamik)
  // J = I (Sabit) + (GİRİŞ Toplam) - (ÇIKIŞ Toplam)
  for (let r = 2; r <= last; r++) {
    const f = "=I" + r +
      "+IFERROR(SUMIF('GİRİŞ'!A:A" + sep + "C" + r + sep + "'GİRİŞ'!B:B)" + sep + "0)" +
      "-IFERROR(SUMIF('ÇIKIŞ'!A:A" + sep + "C" + r + sep + "'ÇIKIŞ'!B:B)" + sep + "0)";
    formulas.push([f]);
  }

  // Tek seferde yaz (Batch Operation)
  sh.getRange(2, S_GUNCEL, formulas.length, 1).setFormulas(formulas).setNumberFormat("#,##0");
  SpreadsheetApp.flush();
}

/**
 * STOK sayfasında “Son Giriş Tarihi (K)” hücresini, verilen stok kodu için günceller.
 * Not: Koda karşılık gelen satır bulunamazsa sessiz geçer.
 */
function updateStokGirisTarihiForCode_(stokKodu, tarih) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stok = ss.getSheetByName(SHEET_STOK);
  const idx = buildStokIndex_();
  const row = idx.byStokKodu.get(normalizeKey_(stokKodu));
  if (row) stok.getRange(row, S_GIRIS_TARIHI).setValue(tarih);
}

/**
 * STOK sayfasında “Son Çıkış Tarihi (L)” hücresini, verilen stok kodu için günceller.
 */
function updateStokCikisTarihiForCode_(stokKodu, tarih) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stok = ss.getSheetByName(SHEET_STOK);
  const idx = buildStokIndex_();
  const row = idx.byStokKodu.get(normalizeKey_(stokKodu));
  if (row) stok.getRange(row, S_CIKIS_TARIHI).setValue(tarih);
}

/*** GİRİŞ -> STOK push (yalnız YENİ kayıt) ***************************************************/

/**
 * GİRİŞ sayfasındaki bir satırı STOK’a “yansıtma” işlemi.
 * Politika:
 *  - MIRROR_GIRIS_TO_STOK=false ise hiç çalışmaz.
 *  - ALLOW_UPDATE_EXISTING_STOK=false iken, STOK’ta mevcut satır varsa asla güncellemez, sadece formülü tazeler.
 *  - STOK’ta bulunamazsa yeni satır açar ve temel alanları kopyalar.
 */
function pushGirisToStok_(row) {
  if (!MIRROR_GIRIS_TO_STOK) return;

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const giris = ss.getSheetByName(SHEET_GIRIS);
  const stok  = ss.getSheetByName(SHEET_STOK);

  // GİRİŞ başlık haritası (başlığa göre kolon bulmak için)
  const headerMapG = headerMap_(giris);

  // “STOK KODU” ve “ŞİRKET KODU” kolonları (esnek başlık eşleşmesi ile)
  const stokCol   = (ALT_KEY_HEADERS.stokKodu  .map(h => headerMapG.get(h.toUpperCase())).find(x => x > 0)) || G_STOK_KODU;
  const sirketCol = (ALT_KEY_HEADERS.sirketKodu.map(h => headerMapG.get(h.toUpperCase())).find(x => x > 0)) || 0;

  // GİRİŞ satırından kod(lar)ı oku
  const stokKodu   = giris.getRange(row, stokCol).getValue();
  const sirketKodu = sirketCol ? giris.getRange(row, sirketCol).getValue() : "";

  // STOK’ta kodu ara
  const idxObj  = buildStokDualIndexFast_();
  const stokRow = findStokRowByKeysFast_(idxObj, stokKodu, sirketKodu);

  // VAR OLAN STOK SATIRINI GÜNCELLEME: kapalıysa, varsa çık (sadece formülü tazele)
  if (!ALLOW_UPDATE_EXISTING_STOK && stokRow > 0) {
    yazFormul_(stokRow);           // J formülünü tazele (güncel adet hesaplanır)
    SpreadsheetApp.flush();
    return;
  }

  // stokRow <= 0: Bu kod için STOK’ta satır yok → yeni satır aç
  if (stokRow <= 0) {
    const newRow = stok.getLastRow() + 1;
    stok.getRange(newRow, S_STOK_KODU).setValue(stokKodu);

    // GİRİŞ’ten temel alanları STOK’a taşı (varsa)
    const mappings = [
      { g: G_KATEGORI, s: S_KATEGORI },
      { g: G_MARKA,    s: S_MARKA },
      { g: G_MODEL,    s: S_MODEL },
      { g: G_OZELLIK,  s: S_OZELLIK },
      { g: G_ACIKLAMA, s: S_ACIKLAMA },
      { g: G_BIRIM,    s: S_BIRIM },
      { g: G_RAF,      s: S_RAF }
    ];
    mappings.forEach(m => {
      const v = giris.getRange(row, m.g).getValue();
      if (v !== "" && v !== null) stok.getRange(newRow, m.s).setValue(v);
    });

    // GÜNCEL (J) formülünü yeni satıra yaz
    yazFormul_(newRow);
    SpreadsheetApp.flush();
  }
}