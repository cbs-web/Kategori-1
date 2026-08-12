# Çanakkale Jeoloji Kütüphanesi – klasör ve harita kullanımı

## Klasörden rapor ekleme

1. **Araçlar → Çanakkale Jeoloji Kütüphanesi** ekranını açın.
2. **Klasörden Rapor Ekle** düğmesine basıp proje klasörünü seçin.
3. Birden fazla Word veya KML adayı varsa ana raporu ve parsel KML'sini seçin.
4. Ada/parsel bilgileriyle KML önizlemesini kontrol edin ve **Forma Aktar** deyin.
5. Metinleri ve onay durumunu kontrol ettikten sonra **Kaydet** düğmesine basın.

Program alt klasörleri de tarar. Geçici Word dosyaları, gizli dosyalar, bozuk DOCX'ler,
poligon içermeyen veya geçersiz KML'ler otomatik kayıt olarak kullanılmaz.

## İlçe klasöründen toplu aktarma

1. Kütüphane ekranındaki **İlçe Klasöründen Toplu Aktar** düğmesine basın.
2. `AYVACIK` gibi ilçe klasörünü seçin; program bütün alt klasörleri tarar.
3. Her proje klasöründeki ana rapor Word'ü ile aynı proje sınırındaki parsel KML'si eşleştirilir.
4. Kesin eşleşmeler yeşil ve seçili, eksik veya belirsiz kayıtlar uyarılı gösterilir.
5. **Seçili Hazır Kayıtları Aktar** ile bütün kesin eşleşmeleri tek onayla kaydedin.

`JEOFİZİK`, `lab`, `evraklar` ve fotoğraf klasörleri ayrı proje olarak kabul edilmez.
Kardeş proje klasörleri arasında Word veya KML eşleştirilmez. Aynı kayıt ikinci kez
taranırsa atlanır; aynı künye için farklı dosya görülürse olası revizyon olarak işaretlenir.
Belirsiz satırlar **Sorunlu Kaydı Tek Proje Olarak İncele** ile ayrıca düzeltilebilir.

Toplu eşleştirmede yalnız Word belgesinin içindeki künye esas alınmaz. Program ilçe ve
proje klasörünü, Word dosya adını, Word içeriğini ve KML adını/Placemark bilgisini birlikte
değerlendirir. Örneğin `0/673` projesindeki `673-parsel.kml`, KML'de ada yazmasa da diğer
kaynaklar ada `0` bilgisini doğruluyorsa otomatik eşleşir. Word içeriğinde eski bir ada/parsel
veya yerleşim kalmış fakat klasör, Word dosya adı ve KML aynı güncel bilgiyi destekliyorsa
kayıt düzeltilmiş olarak sarı uyarıyla hazırlanır; kullanılan kaynaklar ve düzeltme notu
kayıtla birlikte saklanır. Kaynaklarda açık ada/parsel çelişkisi varsa otomatik aktarım yapılmaz.

## Haritada rapor parsellerini görme

1. Çalıştığınız projenin parsel KML'sini **KML Yükle** ile açın.
2. Haritalar sekmesindeki **Jeoloji kütüphanesi parsellerini göster** seçeneğini açık bırakın.
3. Haritayı kaydırdığınızda veya yakınlaştırdığınızda görünür alandaki kayıtlar yenilenir.
4. Eski rapor parseline tıklayarak jeoloji metinlerini inceleyebilir ve kontrollü biçimde projeye uygulayabilirsiniz.

Renkler:

- Mavi: üzerinde çalışılan proje parseli
- Yeşil: onaylı kütüphane raporu
- Turuncu: taslak kayıt (yalnız “Taslakları da göster” açıksa)
- Sarı: haritada seçilen kütüphane parseli

Harita bir raporu otomatik olarak uygun ilan etmez ve uzaklık göstermez. Jeoloji içeriğinin
uygunluğunu kullanıcı parsel konumu, formasyon ve rapor içeriğini birlikte değerlendirerek belirler.

## Parsel merkezli rapor görüntüsü

1. Çalışılan parsele ait KML'yi **KML Yükle** ile açın.
2. Haritalar sekmesinde **Parsel Haritası Hazırla** düğmesine basın.
3. Program parseli ortalar, tamamını gösterecek yakınlaştırmayı seçer ve hibrit uydu
   altlığı üzerine kırmızı sınır, yarı saydam dolgu ve ada/parsel etiketini işler.
4. Önizlemeyi onaylayıp `Parsel_Haritasi.png` dosyasının kaydedileceği klasörü seçin.
5. Ana rapor şablonunda `[PARSEL_HARITASI]` etiketini tek başına bir paragrafa yazın.

Görsel yolu ve üretildiği geometri özeti proje kaydında saklanır. KML veya ada/parsel daha
sonra değişirse rapor ön kontrolü eski görüntünün kullanılmasını engeller ve yeniden hazırlama
ister. Görsel altındaki kaynak açıklaması bunun taslak parsel gösterimi olduğunu belirtir.

## Seçilen jeolojiyi çıktı Word'e aktarma

Kütüphaneden doğrudan bir kayıt projeye uygulandığında, kaydın saklanan JEOLOJİ Word paketi
projeye bağlanır. Ana rapor şablonunda `[JEOLOJI_BOLUMU]` etiketi varsa çıktı hazırlanırken
bu paket metinleri, tabloları, resimleri, şekilleri ve biçimlendirmeleriyle birlikte etikete
yerleştirilir. Kütüphane paketi seçiliyken ilçe/köy adına göre başka bir jeoloji şablonu aranmaz.

Otomatik önerinin genel jeoloji ve inceleme alanı metinleri farklı kütüphane kayıtlarından
geliyorsa yanlış raporun şekillerini taşımamak için tam Word paketi bağlanmaz. Bağlı paket daha
sonra silinir veya bulunamazsa program sessizce başka bir jeoloji şablonuna geçmez; rapor ön
kontrolünde kaydın yeniden uygulanmasını ister. Bağlantı, kayıt kimliği ve içerik özetiyle proje
dosyasında saklanır.

## KML'siz eski kayıtlar

Kütüphanede kaydı seçin, **KML Bağla** ile parsel KML'sini seçin ve kaydı yeniden kaydedin.
KML'siz kayıtlar kütüphane listesinde korunur ancak haritada gösterilmez.
