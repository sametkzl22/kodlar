// dashboard.gs - V17 (GLOBAL ÇAKIŞMA KONTROLÜ: Duplicate Code Fix)

// --- ÖZEL AYARLAR ---
const KATEGORI_LISTESI = ["ELEKTRİK", "PNÖMATİK", "OTOMASYON", "DEMİRBAŞ", "ROBOT", "ŞİRKET STOK", "LAZER", "MEKANİK"];
const BIRIM_LISTESI    = ["ADET", "UZUNLUK (M)", "KG", "LİTRE", "PAKET", "SET"]; 

// Menü Kurulumu
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = "KONTROL PANELİ";
  let sh = ss.getSheetByName(name);
  
  if (sh) { sh.clear(); } 
  else { sh = ss.insertSheet(name, 0); }

  sh.getRange("B2").setValue("🔍 ÜRÜN SORGULA").setFontWeight("bold").setFontSize(14).setFontColor("#1a73e8");
  sh.getRange("B3").setValue("Stok Kodu:");
  sh.getRange("C3").setBackground("#fff2cc").setBorder(true, true, true, true, null, null).setFontWeight("bold").setHorizontalAlignment("center");
  
  sh.getRange("B5").setValue("Marka / Model:");
  sh.getRange("B6").setValue("Mevcut Stok:");
  sh.getRange("B7").setValue("Raf Yeri:");
  sh.getRange("B8").setValue("Özellikler:");
  sh.getRange("B5:B8").setFontWeight("bold").setFontColor("gray");
  sh.getRange("C5:C8").setValue("-").setHorizontalAlignment("left");

  sh.getRange("B10").setValue("🆕 YENİ KART AÇ").setFontWeight("bold").setFontSize(12).setFontColor("#e37400");
  sh.getRange("B10:C10").setBackground("#fce8b2").merge().setHorizontalAlignment("center");

  const labels = [["Kategori (Seç):"], ["Marka:"], ["Model:"], ["Özellikler:"], ["Açıklama:"], ["Birim (Seç):"], ["Başlangıç Adeti:"], ["KARTI OLUŞTUR:"]];
  sh.getRange("B11:B18").setValues(labels);
  sh.getRange("C11:C17").setBackground("#f3f3f3");

  sh.getRange("C11").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(KATEGORI_LISTESI).build()).setBackground("#e6f4ea");
  sh.getRange("C16").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(BIRIM_LISTESI).build()).setValue("ADET").setBackground("#e6f4ea");
  sh.getRange("C17").setValue(0);
  sh.getRange("C18").insertCheckboxes();

  sh.getRange("E2").setValue("⚡ HIZLI İŞLEM").setFontWeight("bold").setFontSize(14).setFontColor("#d93025");
  sh.getRange("E4").setValue("İşlem Türü");
  sh.getRange("F4").setValue("Stok Kodu");
  sh.getRange("G4").setValue("Adet");
  sh.getRange("H4").setValue("UYGULA");
  sh.getRange("I4").setValue("TEMİZLE");

  sh.getRange("E4:I4").setBackground("#f1f3f4").setFontWeight("bold").setHorizontalAlignment("center");

  sh.getRange("E5").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["GİRİŞ YAP (+)", "ÇIKIŞ YAP (-)"]).build()).setValue("GİRİŞ YAP (+)").setBackground("#e6f4ea");
  sh.getRange("F5").setBackground("#fff2cc");
  sh.getRange("G5").setBackground("#fff2cc");
  sh.getRange("H5").insertCheckboxes();
  sh.getRange("I5").insertCheckboxes().setBackground("#fce8b2");

  sh.setColumnWidth(1, 20); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 220);
  sh.setColumnWidth(4, 40); sh.setColumnWidth(5, 120); sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 80); sh.setColumnWidth(8, 60); sh.setColumnWidth(9, 70);
  sh.setHiddenGridlines(true);
  
  SpreadsheetApp.getUi().alert("✅ Panel Güncellendi!");
}

// Edit Trigger
function handleDashboardEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== "KONTROL PANELİ") return;

  const row = range.getRow();
  const col = range.getColumn();
  const val = range.getValue();

  // 1. SORGULAMA
  if (row === 3 && col === 3) {
    sheet.getRange("C5:C8").setValue("-").setFontColor("black").setFontWeight("normal");
    if (!val) return;

    const idxObj = buildStokDualIndexFast_();
    const stokRow = findStokRowByKeysFast_(idxObj, val, "");

    if (stokRow > 0) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const stokSh = ss.getSheetByName(SHEET_STOK);
      const marka = stokSh.getRange(stokRow, S_MARKA).getValue();
      const model = stokSh.getRange(stokRow, S_MODEL).getValue();
      const adet  = stokSh.getRange(stokRow, S_GUNCEL).getValue();
      const raf   = stokSh.getRange(stokRow, S_RAF).getValue();
      const ozellik = stokSh.getRange(stokRow, S_OZELLIK).getValue();
      const birim = stokSh.getRange(stokRow, S_BIRIM).getValue() || "";

      sheet.getRange("C5").setValue(marka + " " + model);
      sheet.getRange("C6").setValue(adet + " " + birim).setFontWeight("bold").setFontColor(adet > 0 ? "green" : "red");
      sheet.getRange("C7").setValue(raf);
      sheet.getRange("C8").setValue(ozellik);
    } else {
      sheet.getRange("C5").setValue("❌ KAYIT YOK!").setFontColor("red");
    }
    return;
  }

  // 2. YENİ KART (Sheet üzerinden manuel)
  if (row === 18 && col === 3 && val === true) {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(500)) return; 

    try {
      range.setValue(false);
      SpreadsheetApp.getActive().toast("Kart oluşturuluyor...", "Bekleyiniz");

      const kat = sheet.getRange("C11").getValue();
      const mar = sheet.getRange("C12").getValue();
      const mod = sheet.getRange("C13").getValue();
      const oze = sheet.getRange("C14").getValue();
      const aci = sheet.getRange("C15").getValue();
      const bir = sheet.getRange("C16").getValue();
      const bas = sheet.getRange("C17").getValue();

      if (!kat || !mar || !mod || !bir) {
        SpreadsheetApp.getUi().alert("HATA: Kategori, Marka, Model ve Birim zorunludur.");
        return;
      }
      
      // Kod Üret
      const generatedResult = clientGetNextCode(kat); 
      const generatedCode = generatedResult.code;

      const ui = SpreadsheetApp.getUi();
      const confirm = ui.alert("Onay", `Ürün: ${mar} - ${mod}\nKod: ${generatedCode}\nBaşlangıç: ${bas} ${bir}\n\nOluşturulsun mu?`, ui.ButtonSet.YES_NO);

      if (confirm !== ui.Button.YES) return;

      apiCreateInternal_({
          code: generatedCode, kategori: kat, marka: mar, model: mod,
          ozellik: oze, aciklama: aci, birim: bir, baslangic: bas
      });
      
      SpreadsheetApp.getActive().toast(`✅ Kart Açıldı: ${generatedCode}`, "Başarılı");
      
      sheet.getRange("C11:C15").clearContent();
      sheet.getRange("C17").setValue(0);
      sheet.getRange("C3").setValue(generatedCode);
      sheet.getRange("C5").setValue(mar + " " + mod);
      sheet.getRange("C6").setValue(bas + " " + bir).setFontColor("green");

    } catch (err) {
      SpreadsheetApp.getUi().alert("Hata: " + err.message);
    } finally {
      lock.releaseLock();
    }
  }

  // 3. HIZLI İŞLEM
  if (row === 5 && col === 8 && val === true) {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(2000)) { range.setValue(false); return; }

    try {
      range.setValue(false);
      SpreadsheetApp.flush();
      SpreadsheetApp.getActive().toast("İşleniyor...", "Bekleyiniz");

      const islem = sheet.getRange("E5").getValue();
      const kod   = sheet.getRange("F5").getValue();
      const adet  = sheet.getRange("G5").getValue();

      if (!kod || !adet || adet <= 0) {
        SpreadsheetApp.getActive().toast("Hata: Kod ve Adet eksik.");
        return;
      }

      const idxObj = buildStokDualIndexFast_();
      const stokRow = findStokRowByKeysFast_(idxObj, kod, "");

      if (stokRow <= 0) {
         SpreadsheetApp.getUi().alert("⚠️ Ürün Yok");
         return;
      }

      if (islem.indexOf("GİRİŞ") !== -1) {
        apiInboundInternal_(kod, adet);
        SpreadsheetApp.getActive().toast(`✅ ${adet} adet GİRİŞ yapıldı.`);
      } else {
        apiOutboundInternal_(kod, adet);
        SpreadsheetApp.getActive().toast(`✅ ${adet} adet ÇIKIŞ yapıldı.`);
      }
      
      sheet.getRange("F5").clearContent();
      sheet.getRange("G5").clearContent();
      
      const sorgulanan = sheet.getRange("C3").getValue();
      if (String(sorgulanan) === String(kod)) {
         const ss = SpreadsheetApp.getActiveSpreadsheet();
         const stokSh = ss.getSheetByName(SHEET_STOK);
         const guncelAdet = stokSh.getRange(stokRow, S_GUNCEL).getValue();
         const birim = stokSh.getRange(stokRow, S_BIRIM).getValue() || "";
         sheet.getRange("C6").setValue(guncelAdet + " " + birim).setFontColor(guncelAdet > 0 ? "green" : "red");
      }

    } catch (err) {
      SpreadsheetApp.getUi().alert("Hata: " + err.message);
    } finally {
      lock.releaseLock();
    }
  }

  // 4. TEMİZLE
  if (row === 5 && col === 9 && val === true) {
    range.setValue(false);
    sheet.getRange("F5").clearContent(); 
    sheet.getRange("G5").clearContent(); 
    sheet.getRange("E5").setValue("GİRİŞ YAP (+)");
    sheet.getRange("C3").clearContent(); 
    sheet.getRange("C5:C8").setValue("-").setFontColor("black").setFontWeight("normal");
    sheet.getRange("C11:C15").clearContent(); 
    sheet.getRange("C16").setValue("ADET"); 
    sheet.getRange("C17").setValue(0);
    SpreadsheetApp.getActive().toast("Tüm panel temizlendi.", "Hazır 🧹");
  }
}

// ----------------------------------------------------
// İÇ FONKSİYONLAR
// ----------------------------------------------------

function getNextRealRow_(sheet, colIndex) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  var values = sheet.getRange(1, colIndex, lastRow, 1).getValues();
  for (var i = lastRow - 1; i >= 0; i--) {
    var cellValue = values[i][0];
    if (cellValue !== "" && cellValue !== null && String(cellValue).trim() !== "") {
      return (i + 1) + 1; 
    }
  }
  return 2;
}

function yazFormul_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_STOK);
  
  const getColLetter = (colIdx) => {
    let temp, letter = '';
    while (colIdx > 0) {
      temp = (colIdx - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      colIdx = (colIdx - temp - 1) / 26;
    }
    return letter;
  };

  const L_KOD      = getColLetter(typeof S_STOK_KODU !== 'undefined' ? S_STOK_KODU : 1);
  const L_BASLANGIC= getColLetter(typeof S_BASLANGIC !== 'undefined' ? S_BASLANGIC : 12);
  const L_G_KOD    = getColLetter(typeof G_STOK_KODU !== 'undefined' ? G_STOK_KODU : 1);
  const L_G_ADET   = getColLetter(typeof G_ADET !== 'undefined' ? G_ADET : 2);
  const L_C_KOD    = getColLetter(typeof X_STOK_KODU !== 'undefined' ? X_STOK_KODU : 1);
  const L_C_ADET   = getColLetter(typeof X_ADET !== 'undefined' ? X_ADET : 2);

  const hucreKod = L_KOD + row;
  const hucreBas = L_BASLANGIC + row;
  
  // NOKTALI VİRGÜL (;) FORMÜL
  const formul = `=${hucreBas} + SUMIF('${SHEET_GIRIS}'!${L_G_KOD}:${L_G_KOD}; ${hucreKod}; '${SHEET_GIRIS}'!${L_G_ADET}:${L_G_ADET}) - SUMIF('${SHEET_CIKIS}'!${L_C_KOD}:${L_C_KOD}; ${hucreKod}; '${SHEET_CIKIS}'!${L_C_ADET}:${L_C_ADET})`;

  const colGuncel = (typeof S_GUNCEL !== 'undefined' ? S_GUNCEL : 10);
  sh.getRange(row, colGuncel).setFormula(formul);
}

// ----------------------------------------------------
// İŞLEM API'LERİ
// ----------------------------------------------------

function apiCreateInternal_(form) {
  const code = form.code;
  const cat  = form.kategori;
  const brand= form.marka;
  const model= form.model;
  const oze  = form.ozellik;
  const aci  = form.aciklama;
  const bir  = form.birim;
  const bas  = Number(form.baslangic) || 0; 
  const raf  = form.raf; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stok = ss.getSheetByName(SHEET_STOK);
  
  const idxObj = buildStokDualIndexFast_();
  if (findStokRowByKeysFast_(idxObj, code, "") > 0) throw new Error("Bu stok kodu zaten var!");

  const newRow = getNextRealRow_(stok, S_STOK_KODU);
  
  stok.getRange(newRow, S_STOK_KODU).setNumberFormat("@").setValue(String(code));
  stok.getRange(newRow, S_KATEGORI).setValue(cat);
  stok.getRange(newRow, S_MARKA).setValue(brand);
  stok.getRange(newRow, S_MODEL).setValue(model);
  
  if (oze) stok.getRange(newRow, S_OZELLIK).setValue(oze);
  if (aci) stok.getRange(newRow, S_ACIKLAMA).setValue(aci);
  if (bir) stok.getRange(newRow, S_BIRIM).setValue(bir);
  if (raf) stok.getRange(newRow, S_RAF).setValue(raf);
  
  stok.getRange(newRow, S_BASLANGIC).setValue(bas);
  
  yazFormul_(newRow);
  stok.getRange(newRow, S_GIRIS_TARIHI).setValue(new Date()).setNumberFormat("yyyy-mm-dd");
  
  SpreadsheetApp.flush(); 

  updateShortHistory_('YENI', code, `${brand} ${model}`, bas, ""); 
  
  return { success: true, message: `Kart açıldı. Kod: ${code}` };
}

function apiInboundInternal_(code, qty) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const giris = ss.getSheetByName(SHEET_GIRIS);
  const stok = ss.getSheetByName(SHEET_STOK);
  
  const idxObj = buildStokDualIndexFast_();
  const stokRow = findStokRowByKeysFast_(idxObj, code, "");
  let realCode = code;
  let urunIsmi = "Yeni Giriş";

  if (stokRow > 0) { 
    realCode = stok.getRange(stokRow, S_STOK_KODU).getValue(); 
    const m = stok.getRange(stokRow, S_MARKA).getValue();
    const mo = stok.getRange(stokRow, S_MODEL).getValue();
    urunIsmi = m + " " + mo;
  }

  const last = getNextRealRow_(giris, G_STOK_KODU);
  
  giris.getRange(last, G_STOK_KODU).setNumberFormat("@").setValue(realCode);
  giris.getRange(last, G_ADET).setValue(qty);
  giris.getRange(last, G_ACIKLAMA).setValue("Panel Hızlı Giriş");
  giris.getRange(last, G_GIRIS_TARIH).setValue(new Date()).setNumberFormat("dd-mm-yyyy");
  
  const onayHucre = giris.getRange(last, G_ONAY);
  onayHucre.setValue(true);
  onayHucre.setNote("LOCKED");
  
  if (typeof fillGirisDetailsByCode_ === 'function') fillGirisDetailsByCode_(last);
  
  if (stokRow > 0) { yazFormul_(stokRow); }
  SpreadsheetApp.flush();

  updateShortHistory_('GIRIS', realCode, urunIsmi, qty, "");
}

function apiOutboundInternal_(code, qty, projeAdi) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cikis = ss.getSheetByName(SHEET_CIKIS);
  const stok = ss.getSheetByName(SHEET_STOK);
  
  const idxObj = buildStokDualIndexFast_();
  const stokRow = findStokRowByKeysFast_(idxObj, code, "");
  let realCode = code;
  let urunIsmi = "Stok Çıkışı";

  if (stokRow > 0) { 
    realCode = stok.getRange(stokRow, S_STOK_KODU).getValue(); 
    const m = stok.getRange(stokRow, S_MARKA).getValue();
    const mo = stok.getRange(stokRow, S_MODEL).getValue();
    urunIsmi = m + " " + mo;
  }

  const last = getNextRealRow_(cikis, X_STOK_KODU);
  
  cikis.getRange(last, X_STOK_KODU).setNumberFormat("@").setValue(realCode);
  cikis.getRange(last, X_ADET).setValue(qty);
  if (projeAdi) { cikis.getRange(last, X_PROJE).setValue(projeAdi); }
  
  const colTarih = (typeof X_TARIH !== "undefined" ? X_TARIH : 3);
  const colOnay = (typeof X_ONAY !== "undefined" ? X_ONAY : 7);
  
  cikis.getRange(last, colTarih).setValue(new Date()).setNumberFormat("dd-mm-yyyy");
  
  const onayHucre = cikis.getRange(last, colOnay);
  onayHucre.setValue(true);
  onayHucre.setNote("LOCKED");
  
  if (typeof fillCikisUrunForRow_ === 'function') fillCikisUrunForRow_(last);
  
  if (stokRow > 0) { yazFormul_(stokRow); }
  SpreadsheetApp.flush();

  updateShortHistory_('CIKIS', realCode, urunIsmi, qty, projeAdi);
}

// ----------------------------------------------------
// WEB APP YARDIMCILARI (GÜNCELLENMİŞ KOD ÜRETİMİ)
// ----------------------------------------------------

function clientCreateCard(form) {
    return apiCreateInternal_(form);
}

// --- FİNAL KOD ÜRETİM MANTIĞI ---
function clientGetNextCode(kategori) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stok = ss.getSheetByName(SHEET_STOK);
  
  const colCatIdx = (typeof S_KATEGORI !== 'undefined' ? S_KATEGORI : 2) - 1;
  const colKodIdx = (typeof S_STOK_KODU !== 'undefined' ? S_STOK_KODU : 1) - 1;

  const data = stok.getDataRange().getValues();
  
  // 1. Kategorideki EN BÜYÜK kodun yanı sıra,
  // TÜM saydadaki mevcut kodların listesini de alıyoruz (Çakışma kontrolü için)
  let allExistingCodes = new Set();
  let maxKodInCategory = 0;
  let foundInCategory = false;

  for (let i = 1; i < data.length; i++) {
    const rowKat = data[i][colCatIdx];
    const val = data[i][colKodIdx];
    const rowKod = parseInt(val, 10);

    if (!isNaN(rowKod)) {
      allExistingCodes.add(rowKod); // Tüm kodları kaydet

      if (rowKat === kategori) {
        if (rowKod > maxKodInCategory) {
          maxKodInCategory = rowKod;
          foundInCategory = true;
        }
      }
    }
  }

  // 2. Kategoriye göre başlangıç kodunu belirle
  let candidate = 0;
  if (!foundInCategory || maxKodInCategory === 0) {
     const baslangiclar = {
       "ELEKTRİK": 10000,
       "PNÖMATİK": 20000,
       "OTOMASYON": 30000,
       "DEMİRBAŞ": 40000,
       "ROBOT": 50000,
       "ŞİRKET STOK": 60000,
       "LAZER": 70000,
       "MEKANİK": 80000
     };
     candidate = baslangiclar[kategori] || 90000;
  } else {
     candidate = maxKodInCategory + 10;
  }

  // 3. GLOBAL ÇAKIŞMA KONTROLÜ (Collision Check)
  // Eğer hesaplanan aday kod (candidate) listede ZATEN VARSA (başka kategoride bile olsa),
  // boş bir yer bulana kadar 10 eklemeye devam et.
  while (allExistingCodes.has(candidate)) {
    candidate += 10;
  }

  return { success: true, code: candidate };
}

function clientQuickTransaction(type, code, amount, projeAdi) {
  if (typeof isBusy_ !== 'undefined' && isBusy_()) {
    return { success: false, message: "⚠️ Sistem meşgul." };
  }
  try {
    const cleanType = String(type).trim().toUpperCase();
    const cleanCode = String(code).trim();
    const cleanQty  = Number(amount);

    if (!cleanCode) return { success: false, message: "❌ Stok kodu boş." };
    if (!cleanQty || cleanQty <= 0) return { success: false, message: "❌ Miktar > 0 olmalı." };

    if (cleanType === 'CIKIS' || cleanType === 'ÇIKIŞ') {
      const idxObj = buildStokDualIndexFast_();
      const stokRow = findStokRowByKeysFast_(idxObj, cleanCode, "");
      if (stokRow <= 0) return { success: false, message: "❌ Ürün bulunamadı!" };
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const stokSh = ss.getSheetByName(SHEET_STOK);
      const currentStock = Number(stokSh.getRange(stokRow, S_GUNCEL).getValue()) || 0;
      
      if (cleanQty > currentStock) return { success: false, message: `🛑 Yetersiz Stok! (Mevcut: ${currentStock})` };
      
      apiOutboundInternal_(cleanCode, cleanQty, projeAdi);
      return { success: true, message: `📉 ÇIKIŞ BAŞARILI: -${cleanQty}` };
    } 
    else {
      apiInboundInternal_(cleanCode, cleanQty);
      return { success: true, message: `📈 GİRİŞ BAŞARILI: +${cleanQty}` };
    }
  } catch (e) {
    return { success: false, message: "Hata: " + e.message };
  }
}

function clientSearchProduct(code) {
  const idxObj = buildStokDualIndexFast_();
  const row = findStokRowByKeysFast_(idxObj, code, "");
  if (row <= 0) return { success: false, message: "Ürün bulunamadı" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_STOK);
  
  return {
    success: true, found: true,
    data: {
      code: sh.getRange(row, S_STOK_KODU).getValue(),
      marka: sh.getRange(row, S_MARKA).getValue(),
      model: sh.getRange(row, S_MODEL).getValue(),
      stok: sh.getRange(row, S_GUNCEL).getValue(),
      raf: sh.getRange(row, S_RAF).getValue(),
      ozellik: sh.getRange(row, S_OZELLIK).getValue(),
      birim: sh.getRange(row, S_BIRIM).getValue()
    }
  };
}

function clientGetProductsByBrand(brandName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stok = ss.getSheetByName(SHEET_STOK);
  const data = stok.getDataRange().getValues();
  let results = [];
  const searchMarka = String(brandName).toUpperCase().trim();

  const idxMarka   = (typeof S_MARKA !== 'undefined' ? S_MARKA : 3) - 1;
  const idxKod     = (typeof S_STOK_KODU !== 'undefined' ? S_STOK_KODU : 1) - 1;
  const idxModel   = (typeof S_MODEL !== 'undefined' ? S_MODEL : 4) - 1;
  const idxOzellik = (typeof S_OZELLIK !== 'undefined' ? S_OZELLIK : 5) - 1;
  const idxStok    = (typeof S_GUNCEL !== 'undefined' ? S_GUNCEL : 10) - 1;
  const idxRaf     = (typeof S_RAF !== 'undefined' ? S_RAF : 13) - 1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[idxKod]) continue; 
    if (String(row[idxMarka]).toUpperCase().indexOf(searchMarka) > -1) {
      results.push({
        code: row[idxKod], marka: row[idxMarka], model: row[idxModel],
        ozellik: row[idxOzellik], stok: row[idxStok], raf: row[idxRaf]
      });
    }
  }
  return { success: true, data: results };
}

// Unicode NFC + normalizeKey_ — Türkçe İ/Ş/Ö/Ü/Ç farklı byte temsillerini eşitler
function normalizeSearch_(v) {
  if (v === null || v === undefined) return "";
  return String(v).normalize('NFC').trim().replace(/\s+/g, " ").toUpperCase();
}

function clientGetProductsByFilters(kat, mar, mod) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stok = ss.getSheetByName(SHEET_STOK);
  const lastRow = stok.getLastRow();
  if (lastRow < 2) return { success: true, data: [] };

  // A1'den başlayarak al — getDataRange() boş A sütununu atlayabilir
  const data = stok.getRange(1, 1, lastRow, S_RAF).getValues();

  const searchKat = normalizeSearch_(kat);
  const searchMar = normalizeSearch_(mar);
  const searchMod = normalizeSearch_(mod);

  const idxKat     = S_KATEGORI - 1;   // B=1
  const idxMarka   = S_MARKA - 1;      // D=3
  const idxModel   = S_MODEL - 1;      // E=4
  const idxKod     = S_STOK_KODU - 1;  // C=2
  const idxOzellik = S_OZELLIK - 1;    // F=5
  const idxStok    = S_GUNCEL - 1;     // J=9
  const idxRaf     = S_RAF - 1;        // M=12

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[idxKod]) continue;

    // Kategori: tam eşleşme (===) — NFC ile normalize edilmiş
    if (searchKat && normalizeSearch_(row[idxKat]) !== searchKat) continue;
    // Marka: kısmi eşleşme (indexOf / contains)
    if (searchMar && normalizeSearch_(row[idxMarka]).indexOf(searchMar) === -1) continue;
    // Model: kısmi eşleşme (indexOf / contains)
    if (searchMod && normalizeSearch_(row[idxModel]).indexOf(searchMod) === -1) continue;

    results.push({
      code: String(row[idxKod] || ""),
      marka: String(row[idxMarka] || ""),
      model: String(row[idxModel] || ""),
      ozellik: String(row[idxOzellik] || ""),
      stok: (typeof row[idxStok] === 'number' ? row[idxStok] : (Number(row[idxStok]) || 0)),
      raf: String(row[idxRaf] || "")
    });
  }

  return { success: true, data: results };
}


function updateShortHistory_(type, code, urunAdi, adet, projeAdi) {
  try {
    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty('LAST_5_TX');
    let list = stored ? JSON.parse(stored) : [];
    
    const d = new Date();
    const dateStr = ("0" + d.getDate()).slice(-2) + "." + ("0" + (d.getMonth() + 1)).slice(-2) + "." + d.getFullYear();

    list.unshift({ type, code, urunAdi: urunAdi || "Ürün", adet: adet || 0, proje: projeAdi || "", dateStr });
    if (list.length > 5) list = list.slice(0, 5);

    props.setProperty('LAST_5_TX', JSON.stringify(list));
  } catch (e) { console.error(e); }
}

function clientGetShortHistory() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty('LAST_5_TX');
  return stored ? JSON.parse(stored) : [];
}

// Son hareketler geçmişini sıfırla
function clientClearHistory() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_5_TX');
  Logger.log("✅ Son hareketler geçmişi temizlendi.");
}