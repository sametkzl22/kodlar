// webapp.gs

/**
 * Web Uygulaması olarak çalıştırıldığında (URL'e gidildiğinde) bu fonksiyon devreye girer.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Sidebar') // Mevcut HTML dosyamızı kullanıyoruz
    .setTitle('Stok Yönetim Sistemi') // Sekme adı
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // Mobilde düzgün görünmesi için
}

/**
 * Web App içerisinden gelen verileri işlemek için gerekli yetkiyi sağlar.
 */
function getUrl() {
  return ScriptApp.getService().getUrl();
}