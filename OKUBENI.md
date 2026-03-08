# Sled Test Analyzer 🚀

Sled Test Analyzer, araç çarpışma testi (Sled) verilerinin analiz edilmesi, grafiklere dönüştürülmesi ve docx (Word) raporu halinde sunulması için özel olarak geliştirilmiş profesyonel bir veri analitik uygulamasıdır. Efe Nakcı tarafından tasarlanmıştır.

## Gereksinimler

Programı çalışıyorken arka planda Python'ın ihtiyacı olan kütüphaneler:
- `pandas` (Veri okuma)
- `numpy` (Matematiksel işlemleri yapabilmek için)
- `PyQt5` (Kullanıcı arayüzünü çalıştırmak için)
- `matplotlib` (O profesyonel grafikleri çizebilmek için)
- `openpyxl` (Excel dosyalarını işleyebilmek için)
- `docxtpl` ve `python-docx` (Grafiklerinizi şablon Word raporlarına dönüştürebilmek için)

## Nasıl Kullanılır? (Adım Adım Kılavuz)

### 1- Verilerin Yüklenmesi (Sol Üst Köşe)
- **Actual Data Yükle:** Sisteminize dışarıdan aldığınız "gerçekleşen" test verisini (mesela `velocity.xlsx`) uygulamaya tanıtın. Bu excel dosyasında *Time*, *Velocity* ve *Acceleration* gibi kolonların bulunması büyük önem arz eder.
- **Target Data Yükle:** Eğer varsa, hedeflenen çarpışma verilerini (mesela `target.xlsx`) bu butonla uygulamaya tanıtın. İçerisinde *Target Velocity* ve *Target Acceleration* gibi istenen değerlerin bulunmasını gerektirir. (Girilmesi şart değildir, girilmezse o grafikler tek çizgi halinde hedefsiz çıkar.)

### 2- Grafikleri Görme 
- Her iki veriyi de seçtikten sonra alttaki yeşil renkli **"Oluştur / Güncelle"** butonuna basın. Uygulama verileri birbirine oturtup anında grafikleri ana ekrana çizecektir.
- Sol alttaki **⬅** ve **➡** (Önceki/Sonraki) ok tuşlarıyla 3 farklı grafiğiniz (*Spul*, *Acceleration vs Velocity*, *Actual vs Target Acceleration*) arasında serbestçe geçiş yapabilirsiniz. Her grafiğin altında otomatik olarak size sunduğu profesyonel minimum/maksimum okuma tabloları da değişecektir.

### 3- Offset (Kayıklık) Ayarları (Sağ Üst Köşe)
Gelen verilerde zaman farklılıkları (kayması/gecikmesi) varsa bunları grafik üzerinde sola doğru kaydırarak giderebilirsiniz:
- Sağ üstteki tabloda (**Grafik Offset Ayarları**) her bir grafik için "Mevcut Değer" sütununa bir saniye kayması (milisaniye / `ms` cinsinden) girebilirsiniz. O rakamı (örneğin 15.0) girdiğiniz an sadece o grafik anında sola doğru (-15 ms) kayacaktır.
- **Evrensel Offset:** Eğer tüm grafiklerin 15 ms kaymasını istiyorsanız en alttaki **"Tüm grafiklere aynı anda evrensel offset uygula"** kutucuğuna 15 yazın. Tüm tablonun anında 15'e dönüşeceğini ve 3 grafiğin de topluca kaydığını göreceksiniz!
- **Not:** Offset ayarları, girildikleri an grafiğe anlık olarak yansır, "Oluştur" tuşuna dahi basmanıza gerek kalmaz.

### 4- Kaydetme ve Çıktı Alma (Alt Taraf)
- **Kayıt Dizini:** Çıktıların hangi klasöre atılacağını gösterir. **"Gözat..."** butonunu kullanarak klasörünüzü (örneğin Masaüstünüzü) seçebilirsiniz.
- **Tüm Grafikleri Kaydet (.png):** O an ekranda ne varsa arkaplanda 3 sekmedeki 3 yüksek çözünürlüklü grafiği (`Spul.png`, `Acc_vs_Vel.png` vb.) otomatik çizer ve tek tıkla seçtiğiniz klasöre resim formatında kaydeder.
- **Rapor Oluştur (Word):** Seçtiğiniz dizinde bir *`Template.docx`* bulunmak zorundadır! (Uygulama bu şablonun içindeki `{{SPUL}}` vs komutları kullanacaktır.)
  1. Mavi butona tıklayın.
  2. Karşınıza **Rapor Bilgileri** adında bir kutu çıkacak. İstediğiniz "Test No", "Tarih" ve "Proje Adı" değerlerini yazıp Ok deyin.
  3. Uygulama saniyeler içerisinde sizin girdiğiniz numaraya özel, örneğin `graphs_096.docx` adında çok profesyonel bir Word dokümanı üretip o dosyayı sizin için hedef seçtiğiniz klasörün içine koyacaktır!

Bol raporlamalar!
- *Created by Efe Nakcı* 🚀
