# DABS Night 2.0 – Test Sürümü Kurulumu

Bu paket **test sürümü**. Canlı uygulamaya ve oradaki verilere dokunmaz:
- Kendi verilerini Firebase'de ayrı `test_...` bölümlerinde tutar.
- Ekranın üstünde turuncu **TEST VERSION** şeridi görünür.
- Telefonda **DABS Night TEST** adıyla, canlı uygulamadan ayrı kurulur.

Süre: yaklaşık 10–15 dakika.

## 1. Firebase kurallarını güncelle (bir kez)

1. Firebase → **Firestore Database → Rules** sekmesini aç.
2. İçindekini tamamen sil.
3. Bu paketteki `firestore.rules` dosyasının **tamamını** yapıştır.
4. **Publish** de.

Bu kurallar canlı uygulamayı bozmaz. Canlı uygulama için geçerli olan bölüm aynı kalıyor; altına test bölümü ekleniyor. İki e-postan da canlıda şimdilik sahip olarak duruyor. Test sürümünde ise sahip sadece **fatihyilmaz.en@gmail.com**.

## 2. GitHub'da test deposu aç

1. GitHub → **+ → New repository**.
   - Ad: `dabs-night-test`
   - **Public** seç.
   - **Create repository** de.
2. **uploading an existing file** linkine tıkla. Bu klasördeki **tüm dosyaları ve `icons` klasörünü** sürükle. **Commit changes** de.
3. **Settings → Pages** bölümünde:
   - Source: **Deploy from a branch**
   - Branch: **main**, klasör: **/(root)**
   - **Save** de.
4. 1–2 dakika sonra adres hazır olur: `https://KULLANICIADIN.github.io/dabs-night-test/`

Firebase'e yeni bir alan adı eklemene gerek yok; `KULLANICIADIN.github.io` zaten tanıtılı.

## 3. Test sürümüne gir

1. Adresi telefonda aç ve **fatihyilmaz.en@gmail.com** ile giriş yap.
   - Google ile girersen doğrulama gerekmez.
   - E-posta/şifre ile girersen doğrulama linkine tıklaman gerekir.
2. **Menu → Copy live nights into test** de. Canlıdaki son 30 gece testin içine kopyalanır. Canlı veri sadece okunur, değişmez.
3. İstersen **Menu → People** bölümünden isim listelerini yapıştır.
4. Bir iki ekip arkadaşına test adresini gönder. Testte ayrı onay gerekir: **Menu → Team & approvals**.

## 4. Denenecekler

- [ ] Başkasının girdiği bir işi ya da DONE bir işi değiştir → uyarı çıkmalı. Kartta **Change history** görünmeli.
- [ ] Aynı günün DABS'ini ikinci kez yükle (başka biri de denesin) → "already loaded" uyarısı çıkmalı.
- [ ] **Menu → Replace tonight's DABS** ile düzeltilmiş bir DABS yükle → girilen bilgiler korunmalı.
- [ ] **Concern list**: ekle, düzenle, Resolved yap, sil. Her biri uyarı vermeli. Ertesi gece açık maddeler görünmeli.
- [ ] **Export Excel** → önce concern hatırlatması çıkmalı. Excel'de sağ üstte rapor saati, altta kırmızı concern listesi olmalı.
- [ ] Android'de **Share file** → paylaşma çalışmazsa dosya İndirilenler'e kaydedilip mesaj çıkmalı.
- [ ] Üstte "Synced · last edit 05:18" görünmeli.
- [ ] Bir task'ın kartında önceki gecelerin geçmişi görünmeli.
- [ ] **Menu → People**: toplu yapıştır, ekle, çıkar; trade sayıları görünmeli.
- [ ] **Menu → Daily Activity Briefing**: IC, TMIC ve First Aider'ı yaz, trade seç, **Word** indir, yazdır.
- [ ] **Menu → Settings** (sadece admin): sabit bilgiler ve imza listesi eşleşmesi.
- [ ] **Menu → Archive**: **Show older nights** ve gece silme.

## 5. Her şey tamamsa canlıya alma

Bana "canlıya alalım" de. Sana canlı paketi ve son kuralları göndereceğim. Sen sadece asıl `dabs-night` deposuna yüklersin. Canlıdaki mevcut geceler aynen kalır. Testteki veriler istersen sonra silinebilir.
