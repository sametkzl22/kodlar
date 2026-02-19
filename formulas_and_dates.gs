/***** FORMÜLLER & TARİHLER – SAĞLAM
 * Bu dosya STOK sayfasındaki tarih sütunlarını (K: Son Giriş, L: Son Çıkış)
 * iki farklı yöntemle günceller:
 *   1) uygulaStokTarihFormulleri_ : ArrayFormula + QUERY/VLOOKUP ile canlı bağlar
 *   2) hesaplaVeYazStokTarihleri_: Formülsüz hesaplar ve tarihlerı hücrelere yazar
 *
 * Notlar:
 * - “uygulaStokTarihFormulleri_” dinamik bir çözüm; GİRİŞ/ÇIKIŞ verisi değiştikçe STOK’ta yansır.
 * - “hesaplaVeYazStokTarihleri_” ise sabit değer yazar (formül bırakmaz), raporlamada tercih edilebilir.
 **************************************************************************************************/

function uygulaStokTarihFormulleri_() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const stok = ss.getSheetByName(SHEET_STOK);
  const r2   = 2; // başlık altı ilk satır

  // Önce mevcut içerikleri temizleyelim (K ve L sütunları)
  const lastRow = stok.getMaxRows();
  if (lastRow >= r2) {
    stok.getRange(r2, S_GIRIS_TARIHI, lastRow - r2 + 1, 1).clearContent(); // K
    stok.getRange(r2, S_CIKIS_TARIHI, lastRow - r2 + 1, 1).clearContent(); // L
  }

  /***** SON GİRİŞ TARİHİ (K)
   * Mantık:
   *  - GİRİŞ sayfasından {Kod, Tarih} kolonlarını al
   *  - Tarihi DATEVALUE ile tarihe dönüştürmeyi dene, olmazsa ham değeri kullan
   *  - Kod bazında MAX(Tarih) al (yani en son giriş)
   *  - STOK!C’deki kodları TO_TEXT ile eşleştirip VLOOKUP ile tarihleri getir
   **********************************************************************/
  const formGiris =
    "=ARRAYFORMULA(IF(LEN($C" + r2 + ":$C)=0,," +                                   // STOK!C boşsa sonuç boş
      "IFERROR(VLOOKUP(TO_TEXT($C" + r2 + ":$C)," +                                // Kod anahtar
        "QUERY({TO_TEXT('GİRİŞ'!$A:$A), IFERROR(DATEVALUE('GİRİŞ'!$J:$J),'GİRİŞ'!$J:$J)}," +
              "\"select Col1, max(Col2) where Col1 is not null group by Col1 label max(Col2) ''\", 0)," +
        "2, FALSE), \"\")))";                                                      // eşleşme yoksa boş döndür


   

  /***** SON ÇIKIŞ TARİHİ (L)
   * Mantık:
   *  - ÇIKIŞ sayfasından {Kod, Tarih} kolonlarını al
   *  - Tarihi DATEVALUE ile tarihe dönüştürmeyi dene, olmazsa ham değeri kullan
   *  - Kod bazında MAX(Tarih) al (yani en son çıkış)
   *  - STOK!C ile VLOOKUP
   * Dikkat: Burada örnekte ÇIKIŞ tarih kolonu S olarak verilmişti (A..S). Kendi sayfanda tarih kolonunun
   * gerçekten S olup olmadığını kontrol et. Değilse $S:$S kısmını doğru sütuna çevir.
   **********************************************************************/
  const formCikis =
    "=ARRAYFORMULA(IF(LEN($C" + r2 + ":$C)=0,," +
      "IFERROR(VLOOKUP(TO_TEXT($C" + r2 + ":$C)," +
        "QUERY({TO_TEXT('ÇIKIŞ'!$A:$A), IFERROR(DATEVALUE('ÇIKIŞ'!$S:$S),'ÇIKIŞ'!$S:$S)}," +
              "\"select Col1, max(Col2) where Col1 is not null group by Col1 label max(Col2) ''\", 0)," +
        "2, FALSE), \"\")))";

  // Formülleri K ve L sütunlarına yaz
  stok.getRange(r2, S_GIRIS_TARIHI).setFormula(formGiris);
  stok.getRange(r2, S_CIKIS_TARIHI).setFormula(formCikis);

  // Görsel format: tarih olarak göster
  if (lastRow >= r2) {
    stok.getRange(r2, S_GIRIS_TARIHI, lastRow - r2 + 1, 1).setNumberFormat("yyyy-mm-dd");
    stok.getRange(r2, S_CIKIS_TARIHI, lastRow - r2 + 1, 1).setNumberFormat("yyyy-mm-dd");
  }
}

/**
 * (Opsiyonel) Formülsüz hesaplayıp K/L’ye doğrudan yazar.
 * Kullanım senaryosu: Formül istemediğiniz/sabit değerle raporlamak istediğiniz durumlar.
 * Not: Kodlar TO_TEXT ile normalize edilmediği için burada normalizeKey_ kullanıyoruz.
 */
function hesaplaVeYazStokTarihleri_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const stok  = ss.getSheetByName(SHEET_STOK);
  const giris = ss.getSheetByName(SHEET_GIRIS);
  const cikis = ss.getSheetByName(SHEET_CIKIS);

  // ——— GİRİŞ için: Kod -> En Son Tarih (Map)
  const mapGiris = new Map();
  {
    const last = giris.getLastRow();
    if (last >= 2) {
      // A..J aralığı (A:Kod, J:Tarih) — J’nin gerçekten 10. kolon olduğuna dikkat.
      const vals = giris.getRange(2, 1, last - 1, Math.max(10, giris.getLastColumn())).getValues(); // A..J
      for (let i = 0; i < vals.length; i++) {
        const codeRaw = vals[i][0];  // A
        const dateRaw = vals[i][9];  // J
        const code = normalizeKey_(codeRaw);
        if (!code) continue;

        const dt = parseDate_(dateRaw); // Date/seri/string’e dayanıklı parse
        if (!dt) continue;

        // Map’te yoksa ekle; varsa en büyüğü (son) tut
        if (!mapGiris.has(code) || dt > mapGiris.get(code)) {
          mapGiris.set(code, dt);
        }
      }
    }
  }

  // ——— ÇIKIŞ için: Kod -> En Son Tarih (Map)
  const mapCikis = new Map();
  {
    const last = cikis.getLastRow();
    if (last >= 2) {
      const maxCol = cikis.getLastColumn();
      const vals = cikis.getRange(2, 1, last - 1, maxCol).getValues();
      for (let i = 0; i < vals.length; i++) {
        const codeRaw = vals[i][0];     // A
        const dateRaw = vals[i][18];    // S (örnek: 19. kolon). Kendi sayfanda tarih kolonunu doğrula!
        const code = normalizeKey_(codeRaw);
        if (!code) continue;

        const dt = parseDate_(dateRaw);
        if (!dt) continue;

        if (!mapCikis.has(code) || dt > mapCikis.get(code)) {
          mapCikis.set(code, dt);
        }
      }
    }
  }

  // ——— STOK sayfasına yaz: K (son giriş), L (son çıkış)
  const last = stok.getLastRow();
  for (let r = 2; r <= last; r++) {
    const code = normalizeKey_(stok.getRange(r, S_STOK_KODU).getValue());

    const g = mapGiris.get(code);
    const c = mapCikis.get(code);

    if (g) stok.getRange(r, S_GIRIS_TARIHI).setValue(g);
    if (c) stok.getRange(r, S_CIKIS_TARIHI).setValue(c);
  }

  // Görsel format: tarih
  if (last >= 2) {
    stok.getRange(2, S_GIRIS_TARIHI, last - 1, 1).setNumberFormat("yyyy-mm-dd");
    stok.getRange(2, S_CIKIS_TARIHI, last - 1, 1).setNumberFormat("yyyy-mm-dd");
  }
}