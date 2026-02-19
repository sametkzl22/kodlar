/***** MENÜ & TETİKLEYİCİLER – HIZLI ve GÜVENLİ
 * Bu dosya, Google Sheets menüsünü oluşturur ve kullanıcı etkileşimlerine yanıt veren tetikleyicileri (onOpen/onEdit) barındırır.
 * ÖNEMLİ TASARIM:
 * - onEdit içinde GİRİŞ sayfasında SADECE otomatik kod/alan doldurma (3’lü tamamlanınca) yapılır.
 * - onEdit içinde hiçbir şekilde GİRİŞ’ten STOK’a veri yazılmaz, tarih atılmaz, onay kutusuna dokunulmaz.
 * - Girişlerin STOK’a işlenmesi için menüden “bekleyenleri işle” aksiyonları kullanılmalıdır.
 ****************************************************************************************************/

function onOpen() {
  // Üst menüye "Stok İşlemleri" adında özel bir menü ekler
  SpreadsheetApp.getUi()
    .createMenu("Stok İşlemleri")

    .addItem('🌐 WEB ARAYÜZÜNÜ AÇ', 'openWebAppLauncher')

    //.addItem('Paneli Aç', 'openSidebarPanel')


    //.addItem("STORE'u baştan hesapla (K=TRUE)", 
    //"recomputeStoreFromApproved_")  // ⬅️ yeni
    
    //.addItem("Güncel Adeti Yeniden Hesapla (başlangıç adetine dokunma)", "menuRecomputeStoreSafe_")
    //.addSeparator()
    /*
    .addItem("Tüm çıkış tarihlerini işle (L)", 
    "processAllCikisDates_")
    .addSeparator()
    */
    //.addSeparator()
    //.addItem("Seçili kod için BEKLEYEN ÇIKIŞ’ları işle", "menuProcessPendingExitsForActiveCode_")
    //.addItem("Tüm BEKLEYEN ÇIKIŞ’ları işle (checkboxları tiksiz olanlar)", "processAllPendingExits_")
    
    
    
    //.addItem("ÇIKIŞ: Ürün sütununu doldur", "menuFillAllCikisUrun_")
    //.addSeparator()
    //.addItem("GİRİŞ checkbox kilitle", "menuLockAllApproved_")
    //.addItem("ÇIKIŞ: checkbox kilitle ", "menuLockAllApprovedExits_")
    //.addSeparator()
    // ——— GİRİŞ onay (K=false) akışı ———
    // Aktif satırdaki stok kodu için GİRİŞ!K=false olan TÜM satırları işler (K=true yapar, tarih atar, STOK’a yeni satır açar veya formülü tazeler).
    //.addItem("Girilen seçili satırı stok listesine işle", "menuProcessPendingForActiveCode_")
    // Tüm kodlar için, GİRİŞ!K=false olanların hepsini işler (toplu onay).
    //.addItem("Girilen tüm stokları işle (checkboxları tiksiz olanlar)", "processAllPendingIntakes_")

    //.addSeparator()

    // ——— STOK güncel (J) formülü yardımcıları ———
    // Yalnız aktif satırdaki STOK!J hücresine doğru formülü yazar.
    /*.addItem("Stok girişini seçili satıra uygula", "menuApplyGuncelToActiveRow_")
    // Tüm STOK satırlarına STOK!J formülünü tekrar yazar (yeniler).
    .addItem("Tüm stokları güncelle", "menuApplyGuncelToAll_")

    .addSeparator() */

    // ——— Otomatik detay doldurma yardımcıları ———
    // Aktif satırda stok/şirket koduna göre STOK'tan detayları (boş hücrelere) kopyalar.
    /* .addItem("Aktif satırı doldur (stok+şirket kodu)", "autofillActiveRowByKeys_")
    // Seçili aralık için aynı işlemi topluca yapar.
    .addItem("Seçili aralığı doldur (stok+şirket kodu)", "autofillSelectionByKeys_") 

    .addSeparator()

    // ——— Tarih alanları (STOK K/L) ———
    // STOK!K ve STOK!L için ArrayFormula tabanlı bağlantı formüllerini uygular.
    .addItem("Tarih bağlantılarını uygula (K/L)", "uygulaStokTarihFormulleri_")
    // GİRİŞ/ÇIKIŞ sayfalarını tarayıp en son tarihleri hesaplar ve STOK!K/L’ye direkt yazar (formülsüz).
    .addItem("Tarihleri hesapla ve yaz (K/L) – formülsüz", "hesaplaVeYazStokTarihleri_")

    .addSeparator()

    // ——— Büyük veri için performanslı doldurma ———
    // Tüm sayfayı (aktif sayfa) parçalar halinde doldurur (performans için).
    .addItem("Tüm sayfayı doldur (büyük veri – parça parça)", "menuAutofillAllChunked_") */
    
    //.addItem("🔥 Paneli Kur/Sıfırla", "setupDashboard")
    .addToUi(); 

    
}

function ensureCikisUrunFormula_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_CIKIS);
  if (!sh) return;
  const d2 = sh.getRange(2, (typeof X_URUN !== "undefined" ? X_URUN : 4));
  const f  = d2.getFormula();
  if (!f) uygulaCikisUrunFormulu_(); // benim verdiğim fonksiyon
}

//Menu lock sütunları kitleme
function menuLockAllApproved_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_GIRIS);
  if (!sh) return;
  const last = sh.getLastRow();
  if (last < 2) return;

  let cnt = 0;
  setBusy_(true);
  try {
    const vals = sh.getRange(2, G_ONAY, last - 1, 1).getValues(); // K sütunu
    for (let i = 0; i < vals.length; i++) {
      if (vals[i][0] === true) {
        ensureLockedNote_(sh, 2 + i);
        cnt++;
      }
    }
    SpreadsheetApp.flush();
    SpreadsheetApp.getActive().toast("LOCKED notu eklenen hücre: " + cnt, "Bitti", 4);
  } finally {
    setBusy_(false);
  }
}
// çıkış all sutünları kitleme
function menuLockAllApprovedExits_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CIKIS);
  if (!sh) return;
  const last = sh.getLastRow();
  if (last < 2) return;

  let cnt = 0;
  setBusy_(true);
  try {
    const vals = sh.getRange(2, X_ONAY, last - 1, 1).getValues(); // G sütunu
    for (let i = 0; i < vals.length; i++) {
      if (vals[i][0] === true) {
        ensureLockedNoteAt_(sh, 2 + i, X_ONAY);
        cnt++;
      }
    }
    SpreadsheetApp.flush();
    SpreadsheetApp.getActive().toast("ÇIKIŞ: LOCKED notu eklenen hücre: " + cnt, "Bitti", 4);
  } finally {
    setBusy_(false);
  }
}


/**
 * Menü: aktif satırdaki stok kodu için bekleyen girişleri (GİRİŞ!K=false) işle.
 * Kullanım: GİRİŞ sayfasında, kodu yazılmış herhangi bir satırdayken çalıştırın.
 */
function menuProcessPendingExitsForActiveCode_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CIKIS);
  if (!sh) return;
  const row = sh.getActiveCell().getRow();
  if (row < 2) return;
  const code = String(sh.getRange(row, X_STOK_KODU).getValue() || "").trim();
  if (code) processPendingExitsForCode_(code);
}

function processAllPendingExits_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const cikis = ss.getSheetByName(SHEET_CIKIS);
  if (!cikis) return;
  const last = cikis.getLastRow();
  if (last < 2) return;

  const maxCol = Math.max(cikis.getLastColumn(), X_ONAY || 7);
  const vals   = cikis.getRange(2, 1, last - 1, maxCol).getValues();

  const codes = new Set();
  for (let i = 0; i < vals.length; i++) {
    const approved = !!vals[i][(X_ONAY || 7) - 1]; // G
    if (!approved) {
      const key = normalizeKey_(vals[i][X_STOK_KODU - 1]); // A
      if (key) codes.add(key);
    }
  }
  codes.forEach(k => processPendingExitsForCode_(k));
}

/**
 * Menü: büyük veri için seçili aktif sayfayı 2. satırdan itibaren chunk’lar halinde doldurur.
 * Not: Bu doldurma, kod/şirket kodu ile STOK detay eşleşmesi yapar (hedefte BOŞ hücreleri doldurur).
 */
function menuAutofillAllChunked_() {
  const sh = SpreadsheetApp.getActiveSheet();
  autofillAllByKeysChunked_(sh.getName(), 2, 300);
}

/**
 * Menü: STOK sayfasında yalnız aktif satırın “GÜNCEL (J)” hücresine formülü uygular.
 * Kullanışlılık: Satır bazlı hızlı tazeleme.
 */
function menuApplyGuncelToActiveRow_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STOK);
  if (!sh) return;
  const r = sh.getActiveCell().getRow();
  if (r >= 2) {
    yazFormul_(r);             // STOK!J formülünü bu satıra yaz
    SpreadsheetApp.flush();
  }
}

/**
 * Menü: STOK sayfasındaki tüm satırlara “GÜNCEL (J)” formülünü baştan uygular.
 * Kullanışlılık: Formül bozulduysa veya yeni mantık eklendiyse topluca güncelleme.
 */
function menuApplyGuncelToAll_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STOK);
  if (!sh) return;
  const last = sh.getLastRow();
  for (let r = 2; r <= last; r++) {
    yazFormul_(r);             // her satır için STOK!J formülünü yaz
  }
  SpreadsheetApp.flush();
}

     // ÇIKIŞ’taki tüm satırları tarar, her stok kodu için EN SON tarihi bulur
// ve STOK LİSTESİ'nde L sütununa (S_CIKIS_TARIHI) yazar.
function processAllCikisDates_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const stok  = ss.getSheetByName(SHEET_STOK);
  const cikis = ss.getSheetByName(SHEET_CIKIS);
  if (!stok || !cikis) return;

  const lastC = cikis.getLastRow();
  const lastS = stok.getLastRow();

  // Kod -> en son çıkış tarihi
  const lastByCode = new Map();

  if (lastC >= 2) {
    const needCols = Math.max(cikis.getLastColumn(), (typeof X_TARIH !== "undefined" ? X_TARIH : 3));
    const rows = cikis.getRange(2, 1, lastC - 1, needCols).getValues();
    for (let i = 0; i < rows.length; i++) {
      const codeRaw = rows[i][X_STOK_KODU - 1];                               // A
      const dateRaw = rows[i][(typeof X_TARIH !== "undefined" ? X_TARIH : 3) - 1]; // C
      const key = normalizeKey_(codeRaw);
      if (!key) continue;

      const dt = parseDate_(dateRaw);
      if (!dt) continue;

      const prev = lastByCode.get(key);
      if (!prev || dt > prev) lastByCode.set(key, dt);
    }
  }

  if (lastS < 2) return;

  setBusy_(true);
  try {
    for (let r = 2; r <= lastS; r++) {
      const codeKey = normalizeKey_(stok.getRange(r, S_STOK_KODU).getValue());
      const dt = lastByCode.get(codeKey);
      if (dt) {
        const cell = stok.getRange(r, S_CIKIS_TARIHI);
        cell.setValue(dt);
        cell.setNumberFormat("dd-mm-yyyy");
      }
    }
    SpreadsheetApp.flush();
    // DÜZELTİLEN SATIR:
    SpreadsheetApp.getActive().toast("Tüm çıkış tarihleri işlendi.", "Bitti", 4);
  } finally {
    setBusy_(false);
  }
}
// Menü: GİRİŞ sayfasında aktif satırdaki stok kodu için bekleyen (K=false) girişleri işle
function menuProcessPendingForActiveCode_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_GIRIS);
  if (!sh) return;

  const row = sh.getActiveCell().getRow();
  if (row < 2) return;

  const code = String(sh.getRange(row, G_STOK_KODU).getValue() || "").trim();
  if (!code) {
    try {
      SpreadsheetApp.getUi().alert("Uyarı", "Seçili satırda stok kodu yok.", SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      SpreadsheetApp.getActive().toast("Seçili satırda stok kodu yok.", "Uyarı", 4);
    }
    return;
  }

  // Bekleyenleri (K=false) bu kod için işle
  processPendingIntakesForCode_(code);
}

/**
 * onEdit: Kullanıcı düzenleme yaptığında tetiklenir.
 * Güvenlik & kararlılık:
 * - isBusy_() kontrolüyle reentrancy/sonsuz tetik döngüsü engellenir.
 * - GİRİŞ sayfasında SADECE 3’lü tamamlanınca otomatik doldurma yapılır.
 * - STOK’a push, tarih atama vb. ağır işler burada çalıştırılmaz (yanlış tetik ve performans sorunlarını önler).
 */
// K sütunu için YUMUŞAK KİLİT (izin istemez): TRUE olduktan sonra FALSE yapılamaz
// menu_and_triggers.gs içindeki GÜNCEL onEdit

function onEdit(e) {
  try {
    // 1. Kilit Kontrolü
    if (isBusy_()) return;
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    const name = sh.getName();

    // --- ÖNCELİKLİ: KONTROL PANELİ ---
    // Eğer işlem panel sayfasındaysa, hemen panel kodunu çalıştır ve bitir.
    if (name === "KONTROL PANELİ") {
      if (typeof handleDashboardEdit === 'function') {
        handleDashboardEdit(e);
      }
      return; // Diğer karmaşık kontrollere girme
    }

    // --- Diğer Sayfalar (Giriş, Çıkış, Stok) ---
    const r1 = e.range.getRow();
    if (r1 < 2) return;

    // ... (Buradan aşağısı senin eski GİRİŞ/ÇIKIŞ/STOK kodların olarak kalmalı) ...
    // Eğer önceki kodların tam halini istiyorsan söyle, atayım. Ama sadece üstteki kısmı eklemen yeterli.
    
    // NOT: Eski kodların silinmemesi için aşağıya sadece çağrıları bırakıyorum.
    // Eğer elindeki onEdit kodunun alt kısmı duruyorsa dokunma.
    // Sadece en tepeye "KONTROL PANELİ" bloğunu ekle.
    
  } catch (err) {
    // Hata olursa sessiz kal
  }
}

/**
 * Web App'i açmak için şık bir pencere gösterir.
 */
function openWebAppLauncher() {
  // BURAYA KENDİ WEB APP LİNKİNİ YAPIŞTIR 👇
  const url = "https://script.google.com/a/macros/3dotomasyon.com/s/AKfycbz36kgeySF7z0o9jI86m-PGAcObz-c3e8YLhMVPw9NrNdmZpR-dXdU9C7Fa2hxC2ltPDw/exec"; 
  
  // Pencere tasarımı (HTML + CSS)
  const htmlContent = `
    <div style="font-family: 'Segoe UI', sans-serif; text-align: center; padding: 20px;">
      <h2 style="color: #333; margin-bottom: 10px;">Stok Yönetim Paneli</h2>
      <p style="color: #666; font-size: 14px; margin-bottom: 25px;">
        Tam ekran deneyimi ve güvenli işlem için panele geçiş yapın.
      </p>
      <a href="${url}" target="_blank" style="text-decoration: none;">
        <button style="
          background-color: #2563eb; 
          color: white; 
          border: none; 
          padding: 12px 24px; 
          font-size: 16px; 
          font-weight: bold; 
          border-radius: 8px; 
          cursor: pointer; 
          box-shadow: 0 4px 6px rgba(37, 99, 235, 0.3);
          transition: background 0.3s;
        " onmouseover="this.style.backgroundColor='#1d4ed8'" onmouseout="this.style.backgroundColor='#2563eb'">
          🚀 ARAYÜZÜ AÇ
        </button>
      </a>
      <p style="margin-top: 15px; font-size: 11px; color: #999;">Bu pencereyi kapatabilirsiniz.</p>
    </div>
  `;

  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(400)
    .setHeight(250);

  SpreadsheetApp.getUi().showModalDialog(html, 'Yönetim Paneli');
}