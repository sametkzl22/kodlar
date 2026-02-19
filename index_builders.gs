/***** STOK İNDEKS & KATALOGLAR – BATCH
 * Bu dosya, STOK LİSTESİ sayfasından tek seferde indeks/katalog
 * yapılarını üretir. Diğer modüller bu indeksleri hızlı arama ve
 * kod üretimi (kategori içi max + step) için kullanır.
 *
 * Üretilen yapılar:
 *  - byStok:    STOK KODU (normalize) -> satır no
 *  - bySirket:  Şirket Kodu (şu an stok kodu ile aynı) -> satır no
 *  - rowByBrandModel:  "MARKA||MODEL"            -> {row, code}
 *  - rowByCatBrandModel:"KATEGORİ||MARKA||MODEL" -> {row, code}
 *  - codesByCat: "KATEGORİ" -> ["0010", "0020", ...]  (yeni kod üretmek için kaynak)
 ******************************************************************/

// STOK kataloğunu ve indekslerini tek seferde hazırla
function buildCatalogMaps_() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const stok = ss.getSheetByName(SHEET_STOK);
  const last = stok.getLastRow();

  // Çıktı objesi: tüm indeks/katalog yapıları burada toplanır
  const res = {
    byStok: new Map(),             // STOK KODU -> row (hızlı bulma)
    bySirket: new Map(),           // Şirket Kodu -> row (şu an stok kodu ile aynı alan)
    rowByBrandModel: new Map(),    // "MARKA||MODEL" -> { row, code } (geriye uyum)
    rowByCatBrandModel: new Map(), // "KAT||MARKA||MODEL" -> { row, code } (asıl 3’lü eşleşme)
    codesByCat: new Map()          // "KATEGORİ" -> [kodlar] (kategori içi max bulmak için)
  };
  if (last < 2) return res;        // veri yoksa boş haritalar dön

  // STOK’ta B..M (12 sütun) aralığını tek seferde oku (2. satırdan itibaren)
  const values = stok.getRange(2, 2, last - 1, 12).getValues(); // B..M

  // Her satır için indeks/katalogları besle
  for (let i = 0; i < values.length; i++) {
    const row = i + 2; // gerçek satır numarası (başlık + offset)

    // Ham değerler
    const kategori = normalizeKey_(values[i][S_KATEGORI - 2]); // B
    const codeRaw  = values[i][S_STOK_KODU - 2];               // C
    const markaRaw = values[i][S_MARKA - 2];                   // D
    const modelRaw = values[i][S_MODEL - 2];                   // E

    // Kod ve normalize anahtarlar
    const code = String(codeRaw || "");
    const stokKey   = normalizeKey_(codeRaw);
    const sirketKey = normalizeKey_(codeRaw); // Ayrı bir “şirket kodu” alanınız yoksa stok koduyla aynı kullanılır

    // ——— 1) Doğrudan kod -> satır indeksleri
    if (stokKey)   res.byStok.set(stokKey, row);
    if (sirketKey) res.bySirket.set(sirketKey, row);

    // ——— 2) Kategori -> kod listesi (yeni kod üretiminde max’ı bulmak için)
    if (kategori) {
      if (!res.codesByCat.has(kategori)) res.codesByCat.set(kategori, []);
      if (code) res.codesByCat.get(kategori).push(code);
    }

    // ——— 3) 2’li ve 3’lü anahtar indeksleri
    const markaKey = normalizeKey_(markaRaw);
    const modelKey = normalizeKey_(modelRaw);

    // Geriye uyumluluk: sadece marka||model ile de satır bulunabilsin
    if (markaKey || modelKey) {
      res.rowByBrandModel.set(markaKey + "||" + modelKey, { row, code });
    }

    // Asıl hedef: kategori+marka+model 3’lüsüyle deterministik eşleşme
    if (kategori && markaKey && modelKey) {
      res.rowByCatBrandModel.set(
        kategori + "||" + markaKey + "||" + modelKey,
        { row, code }
      );
    }
  }

  return res;
}

/**
 * buildCatalogMaps_()’i çağırıp yalnız stok/şirket kodu index’lerini döndürür.
 * Kullanım: Hızlı kod -> satır aramalarında.
 */
function buildStokDualIndexFast_() {
  const maps = buildCatalogMaps_();
  return { byStok: maps.byStok, bySirket: maps.bySirket };
}

/**
 * Stok kodu ve/veya şirket kodu ile hızlı satır bulma.
 * Öncelik:
 *  1) stok kodu varsa doğrudan byStok’tan
 *  2) yoksa şirket kodundan bySirket’ten
 * Not: Şu an iki alan da aynı değere bakıyor; ileride ayrışırsa hazır.
 */
function findStokRowByKeysFast_(idxObj, stokKoduRaw, sirketKoduRaw) {
  const s1 = normalizeKey_(stokKoduRaw);
  const s2 = normalizeKey_(sirketKoduRaw);
  const { byStok, bySirket } = idxObj;

  // Her iki anahtar da verilmişse: stok kodu öncelikli (daha spesifik)
  if (s1 && s2) {
    const r1 = byStok.get(s1);
    if (r1) return r1;
  }
  // Tek başına stok kodu varsa
  if (s1 && byStok.has(s1))   return byStok.get(s1);
  // Tek başına şirket kodu varsa
  if (s2 && bySirket.has(s2)) return bySirket.get(s2);

  return 0; // bulunamadı
}

/**
 * Eski arayüzler için basitleştirilmiş index objesi döndürür.
 * Şu an sadece byStokKodu map’ini sağlar.
 */
function buildStokIndex_() {
  const idx = buildStokDualIndexFast_();
  return { byStokKodu: idx.byStok };
}