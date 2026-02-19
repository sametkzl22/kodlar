/***** KONFİG & SÜTUN HARİTASI *****
 * Bu dosya tüm sayfa adlarını, kolon numaralarını ve davranış bayraklarını merkezî olarak tanımlar.
 * Not: Google Sheets API sütunları 1-indekslidir (A=1, B=2, ...).
 ******************************************/

// ——— Sayfa adları ———
const SHEET_GIRIS = "GİRİŞ";
const SHEET_STOK  = "STOK LİSTESİ";
const SHEET_CIKIS = "ÇIKIŞ";

/***** GİRİŞ Sayfası (A–K) *****
 * A: STOK KODU
 * B: ADET
 * C: KATEGORİ
 * D: ÜRÜN MARKASI
 * E: ÜRÜN MODELİ
 * F: ÜRÜN ÖZELLİKLERİ
 * G: AÇIKLAMA
 * H: BİRİM
 * I: RAF SIRASI
 * J: GİRİŞ TARİHİ (işlenince NOW yazılır)
 * K: ONAY (Checkbox; FALSE=bekliyor, TRUE=işlendi)
 ******************************************/
const G_STOK_KODU   = 1;  // A
const G_ADET        = 2;  // B
const G_KATEGORI    = 3;  // C
const G_MARKA       = 4;  // D
const G_MODEL       = 5;  // E
const G_OZELLIK     = 6;  // F
const G_ACIKLAMA    = 7;  // G
const G_BIRIM       = 8;  // H
const G_RAF         = 9;  // I
const G_GIRIS_TARIH = 10; // J
const G_ONAY        = 11; // K (Checkbox)

/***** STOK LİSTESİ (B–M) *****
 * B: KATEGORİ
 * C: STOK KODU (şirket kodu ile aynı kullanılıyor)
 * D: ÜRÜN MARKASI
 * E: ÜRÜN MODELİ
 * F: ÜRÜN ÖZELLİKLERİ
 * G: AÇIKLAMA
 * H: BİRİM
 * I: BAŞLANGIÇ MİKTARI
 * J: GÜNCEL (formülle: I + onaylı girişler − çıkışlar)
 * K: SON GİRİŞ TARİHİ
 * L: SON ÇIKIŞ TARİHİ
 * M: RAF SIRASI
 ******************************************/
const S_KATEGORI     = 2;   // B
const S_STOK_KODU    = 3;   // C
const S_SIRKET_KODU  = S_STOK_KODU; // alias (ayrı bir sütun yoksa aynı referans)
const S_MARKA        = 4;   // D
const S_MODEL        = 5;   // E
const S_OZELLIK      = 6;   // F
const S_ACIKLAMA     = 7;   // G
const S_BIRIM        = 8;   // H
const S_BASLANGIC    = 9;   // I
const S_GUNCEL       = 10;  // J
const S_GIRIS_TARIHI = 11;  // K
const S_CIKIS_TARIHI = 12;  // L
const S_RAF          = 13;  // M

/***** ÇIKIŞ Sayfası (A–C) *****
 * A: STOK KODU
 * B: ADET
 * C: TARİH (manuel veya otomatik NOW)
 ******************************************/
const X_STOK_KODU = 1; // A
const X_ADET      = 2; // B
const X_TARIH     = 3; // C
const X_URUN      = 4; // D sütunu: "ÜRÜN"
const X_ONAY      = 7; // G checkbox  
const X_PROJE     = 6; // F Sütunu (Kullanılan Proje)

/***** STOK KODU ÜRETİM PARAMETRELERİ *****
 * Amaç: Kategori içindeki en yüksek kodu bulup +CODE_STEP ile yeni kod oluşturmak.
 * Kod formatı: <prefix><padded number>  (örn: "" + 0010 -> "0010")
 ******************************************/
const STOK_KOD_STRATEGY       = "AUTO_INCREMENT"; // şimdilik tek strateji
const STOK_KOD_NUM_REGEX      = /(\d+)$/;         // kod sonundaki sayıyı yakalar
const STOK_KOD_NUM_PAD        = 4;                // 0001, 0011, 0123 gibi padding
const STOK_KOD_START_FROM     = 1;                // kategoride hiç kod yoksa başlangıç
const STOK_KOD_DEFAULT_PREFIX = "";               // istenirse kategori bazlı prefix’e genişletilebilir

/***** ALAN KOPYALAMA (STOK -> hedef) *****
 * STOK sayfasındaki bu kolonlar, hedef satırda aynı başlığı taşıyan sütunlara kopyalanır.
 * Dikkat: FILL_ONLY_IF_EMPTY=true ise yalnızca hedef hücre boşsa yazılır.
 ******************************************/
const FIELDS_TO_COPY_FROM_STOK = [
  { header: "KATEGORİ",         stokCol: S_KATEGORI },
  { header: "ÜRÜN MARKASI",     stokCol: S_MARKA },
  { header: "ÜRÜN MODELİ",      stokCol: S_MODEL },
  { header: "ÜRÜN ÖZELLİKLERİ", stokCol: S_OZELLIK },
  { header: "AÇIKLAMA",         stokCol: S_ACIKLAMA },
  { header: "BİRİM",            stokCol: S_BIRIM },
  { header: "RAF SIRASI",       stokCol: S_RAF }
];

/***** Alternatif başlıklar (esnek eşleşme) *****
 * Kullanıcı farklı yazım kullanırsa yine doğru sütunu bulabilelim diye.
 ******************************************/
const ALT_KEY_HEADERS = {
  stokKodu:   ["STOK KODU", "Stok Kodu", "Ürün Kodu"],
  sirketKodu: ["ŞİRKET KODU", "Sirket Kodu", "Şirket Kodu"]
};

/***** SENKRON POLİTİKASI *****
 * GİRİŞ’ten STOK’a yazma yalnızca yeni satır oluştururken yapılır.
 * Mevcut STOK satırlarının üzerine otomatik yazılmaz (güvenli mod).
 ******************************************/
const MIRROR_GIRIS_TO_STOK = true;
const ALLOW_UPDATE_EXISTING_STOK = false; // mevcut satırı asla güncelleme

// STOK -> Hedef doldurma davranışı: yalnız BOŞ hedef hücreleri doldur (true) / zorla yaz (false)
const FILL_ONLY_IF_EMPTY = true;

/***** OTOMATİK TAMAMLAMA *****
 * 3’lü anahtar (Kategori+Marka+Model) tamamlanınca:
 * - aynı ürün STOK’ta varsa: kod ve detayları getir
 * - yoksa: kategorideki en yüksek kod + CODE_STEP ile benzersiz yeni kod üret
 ******************************************/
const AUTOCOMPLETE_ENABLED = true;
const AUTOCOMPLETE_MAX_SUGGESTIONS = 50; // ileriye dönük: dropdown öneri listesi sınırı
const CODE_STEP = 10;                     // yeni kod artışı (örn: 0010, 0020, 0030)