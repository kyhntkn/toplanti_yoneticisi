# Toplantı Yöneticisi – Windows (EXE + Microsoft Store)

Electron tabanlı masaüstü sürüm. Uygulama kodu `app/index.html` içindedir (tarayıcı sürümüyle aynı dosya).

## 1) GitHub'da derleme
1. Bu klasörün **içeriğini** yeni bir GitHub deposunun köküne yükleyin (`package.json` kökte olmalı).
2. Depoda **Actions → "Derle (EXE + Store)" → Run workflow** deyin (ya da `v1.0.0` gibi bir etiket gönderin).
3. İş bitince sayfanın altındaki **Artifacts → ToplantiYoneticisi-windows** dosyasını indirin:
   - `ToplantiYoneticisi-Kurulum-1.0.0.exe` → kurulum sihirbazı
   - `ToplantiYoneticisi-Tasinabilir-1.0.0.exe` → kurulum gerektirmez
   - `ToplantiYoneticisi-1.0.0.appx` → Microsoft Store paketi

Yerelde denemek için (Node 20+): `npm install` → `npm start`. Derleme: `npm run dist:exe` / `npm run dist:store`.

## 2) Microsoft Store'a yükleme
1. Partner Center'da uygulama adını rezerve edin: **Ürün yönetimi → Ürün kimliği** sayfasından şu 3 değeri alın:
   `Package/Identity/Name`, `Package/Identity/Publisher` (CN=...), `PublisherDisplayName`.
2. `package.json` → `build.appx` bölümündeki `identityName`, `publisher`, `publisherDisplayName` alanlarını bu değerlerle değiştirin (ayrıca `author` ve `copyright`).
3. Workflow'u tekrar çalıştırıp `.appx` dosyasını indirin ve Partner Center'da **Paketler** adımına yükleyin. İmzasız paket kabul edilir, Store paketi kendisi imzalar.
4. Mağaza sayfası için `store-listing/logo-300x300.png` kullanabilirsiniz. Ekran görüntülerini siz eklersiniz.
> Yer tutucu kimlikle üretilen `.appx` yalnızca denemedir, Store kabul etmez. Sürümü yükseltirken `package.json` içindeki `version` değerini artırın (Store için son hane 0 kalır).

## 3) Özellikler (masaüstüne özel)
- **Tam otomatik e-posta:** Tutanak sayfası → ✉ E-posta → alıcıyı yazın → **Gönder** (hemen yollar) veya **Taslak olarak aç**. PDF/Word ek olarak eklenir. Klasik masaüstü Outlook gerekir; yeni (web tabanlı) Outlook COM'u desteklemez. Outlook "bir program sizin adınıza e-posta gönderiyor" uyarısı verebilir.
- **Sistem tepsisi:** Pencereyi kapatınca uygulama arka planda çalışır, toplantı/termin bildirimleri gelmeye devam eder. Tamamen kapatmak için tepsi simgesi → Çıkış. Tepsi menüsünden "Windows ile başlat" seçilebilir (Store/MSIX sürümünde çalışmayabilir).
- Veriler bilgisayarda yerel saklanır, internete gönderilmez.

## 4) Mevcut verileri taşıma
Tarayıcı sürümünde **💾 Yedekle** ile yedek alın, masaüstü uygulamada **📂 Geri Yükle** ile yükleyin.

## 5) Bilinen noktalar
- İmzasız `.exe` ilk açılışta Windows SmartScreen uyarısı verebilir ("Yine de çalıştır"). Uyarıyı kaldırmak için kod imzalama sertifikası gerekir. Store sürümünde bu sorun yoktur.
- Takvime ".ics" ekleme ve Google Takvim bağlantısı varsayılan uygulamayla açılır.
