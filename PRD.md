# K-1 Ürün Gereksinimleri

## 1. Ürün özeti

K-1, Kategori-1 zemin ve temel etüt işlerinin arazi öncesinden nihai Word/PDF teslimine
kadar tek proje kaydı üzerinden yürütüldüğü, Windows üzerinde çalışan tek kullanıcılı bir
mühendislik masaüstü uygulamasıdır. Program veri toplamayı, kontrollü hesaplamayı,
haritalamayı ve rapor üretimini hızlandırır; mühendislik kararının yerine geçmez.

## 2. Kullanıcı ve çalışma ortamı

- Birincil kullanıcı: jeoloji/jeofizik ve zemin etüt raporu hazırlayan mühendis.
- Ortam: Windows, yerel veya ağ klasörlerindeki proje dosyaları, Microsoft Word/PDF çıktısı.
- Temel veri: proje JSON'u; bağlı KML/KMZ/JPEG/DOCX/XLSX/PDF ve üretilmiş haritalar.
- Dil: Türkçe; mühendislik birimleri ve mevzuat terimleri korunur.

## 3. Başarı ölçütleri

- Proje aşaması ve kimliği her an ilk bakışta anlaşılır olmalı.
- Kullanıcı aynı bilgiyi farklı sekmelerde yeniden girmemeli.
- Rapor/harita üretimi sırasında eksikler açıkça gösterilmeli, veri kaybı olmamalı.
- AI ve eski rapor önerileri yalnız danışman niteliğinde olmalı; son seçim kullanıcıda kalmalı.
- Sık etkileşimler 100 ms civarında geri bildirim vermeli; uzun işler arayüzü dondurmamalı.
- Rapor statik etiket dönüşümü gerçek TASLAK üzerinde 500 ms altında olmalı.
- Başlangıç bağımlılık denetimi 250 ms, `import k1` medyanı 1 saniye altında hedeflenir.

## 4. Kapsam dışı ve güvenlik sınırları

- Program otomatik jeolojik uygunluk veya nihai teknik onay vermez.
- AI sonucu güven, kanıt ve kaynakla gösterilir; onaysız biçimde rapora uygulanmaz.
- Haritadaki eski rapor parselleri yalnız karşılaştırma/referans içindir.
- Eksik ekler ve taahhütname nihai üretimi zorla engellemez; açık uyarı verir.
- `CEBE_ATILANLAR.md` içindeki canlı Harita/Rapor Tablosu Tasarım Editörü bu sürümün
  kapsamında değildir.

## 5. Proje yaşam döngüsü

1. Yeni iş doğrudan Yazım Aşamasında veya isteğe bağlı Ön Değer akışıyla başlar.
2. Ön değerde kullanıcı yaklaşık `qt`, `ks`, zemin sınıfı ve TDTH PDF bilgisini kaydeder.
3. Yazım aşamasında proje, arazi, bina, jeofizik, laboratuvar, AÇ/YN, taşıma ve ekler işlenir.
4. Biten proje yeniden açıldığında salt okunur İzleme veya kontrollü Düzeltme seçilir.
5. Düzeltmeler revizyon olarak izlenir; kaydedilmemiş değişiklikte geçiş/çıkış korunur.
6. Nihai Word, PDF ve sıralı ekler kullanıcı denetimiyle üretilir.

## 6. İşlevsel gereksinimler

### Proje ve durum

- `0. Özet`, yalnız mevcut aşamayı, proje sahibi/adı, konum, ada/parsel ve geçerli parsel
  haritasını göstermelidir.
- Aşama değişiklikleri ve revizyonlar proje kaydında tutulmalıdır.
- Salt okunur mod tüm düzenleme noktalarında uygulanmalıdır.

### Haritalar ve parsel

- Kullanıcı KML yükleyebilmeli veya TKGM servisinden parsel geometrisi alabilmelidir.
- Yerbulduru, mühendislik jeolojisi, jeoloji/jeofizik lokasyon ve parsel haritası üretilmelidir.
- Rapor haritalarında jeoloji kütüphanesi parselleri görünmemelidir.
- Çalışılan parsel ile eski rapor parselleri yalnız etkileşimli Haritalar sekmesinde ayrışmalıdır.
- Harita işaret ve çizgi renklerinin alan anlamı UI sadeleştirmesinden etkilenmemelidir.

### Jeoloji kütüphanesi ve 1/100.000 paftalar

- İlçe klasörünün altındaki Word/KML çiftleri kaynak kanıtlarıyla eşleştirilmelidir.
- Seçilen jeoloji Word'ü metin, tablo, resim, şekil ve biçimiyle 2. JEOLOJİ altına eklenmelidir.
- Genel jeoloji haritası parsel merkezli hazırlanmalı; birimler yaşlıdan gence sıralanmalıdır.
- 2.1 metni program/kütüphane veya seçili eski rapordan gelebilmeli; 2.1.1 seçili eski Word'den
  kontrollü biçimde alınmalıdır.
- Gemini ve OpenAI denetimleri yerel sonuçtan ayrı çalıştırılmalı, ortak sonuç son karar olarak
  otomatik kabul edilmemelidir.
- Ayrıntılı kullanım ve aktarım sözleşmesi `CANAKKALE_JEOLOJI_KUTUPHANESI.md` içindedir.

### Arazi, laboratuvar ve jeofizik

- AÇ/YN kayıtları harita, Excel/pano laboratuvar verisi ve rapor tablolarında aynı kimliği
  kullanmalıdır; `AÇ1` için ikinci `AÇ-1` kaydı açılmamalıdır.
- Laboratuvar Excel/CSV ve çalışma sayfası biçimindeki pano verisi eşlenebilmelidir.
- Jeofizik Word ekleri gerektiğinde PDF'e otomatik dönüştürülebilmelidir.

### Taşıma gücü

- YASS var/yok seçimi ve varsa derinliği açık alanlarla girilmelidir.
- Kohezyon ve içsel sürtünme açısı güvenli tarafta 2/3'e indirilebilmeli, ardından kullanıcı
  tarafından düzenlenebilmelidir.
- Kullanılan azaltma rapor metninde tek, onaylı mühendislik cümlesiyle açıklanmalıdır.
- `qk`, `qt` ve sabit 40 katsayılı `ks` formülleri doğrulanabilir değer/tablo olarak üretilmelidir.
- Kullanıcı rapor metnini düzenleyebilmeli; tablo yer işareti korunmalıdır.

### Rapor, taahhütname ve ekler

- Statik etiketler biçimi bozmadan topluca değiştirilmelidir.
- Jeoloji paketi ve stratigrafik kesit normal paragraf stiliyle, doğru başlık altına eklenmelidir.
- Şekil/tablo numaraları ve alanları nihai Word/PDF öncesi güncellenmelidir.
- Taahhütname düzeni RaporPro uyumlu olmalı; mühendis bilgisi eksikliği anlaşılır olmalıdır.
- Eksik ekler uyarı olmalı, zorunlu blok olmamalıdır; taahhütnameler ekler PDF'ine alınmamalıdır.

## 7. Kullanıcı deneyimi ilkeleri

- Her görünümde bir dolu ana işlem; ikincil işlemler nötr/çerçeveli olmalıdır.
- Kırmızı yalnız silme/hata, amber yalnız çözülmemiş risk için kullanılmalıdır.
- Renkli rozet ve tam satır boyama yerine metin, ağırlık ve boşlukla hiyerarşi kurulmalıdır.
- Ana etiketler kullanıcı diliyle yazılmalı; `[KOD]` gibi şablon ayrıntıları ana formdan
  kaldırılmalıdır.
- Tek satırlı tablolar 32 px satır yüksekliğiyle yoğun ama okunur olmalıdır.
- Başarılı olağan işlemler modal kutu yerine alt durum çubuğunda bildirilmelidir.

## 8. Görsel sistem ve hareket

- Yazı tipi: Windows'a doğal Segoe UI.
- Zemin: açık nötr; yüzey: beyaz; ana vurgu: uygulamada önceden kullanılan `#2563eb`.
- Jeolojik/harita semantik renkleri aynen korunur.
- Kullanıcı pencereleri 170 ms ease-out giriş (8 px yukarı yerleşme + opaklık) ve 110 ms
  çıkış kullanır. Ölçek/sıçrama/parlama yoktur.
- Animasyonlar `Araçlar > Pencere Animasyonları` ile kapatılabilir; test/headless ve
  `K1_REDUCED_MOTION=1` durumunda otomatik devre dışıdır.
- Harita dışa aktarımı sırasında içerik/tablo animasyonu yapılmaz.

## 9. Performans ve ölçüm

Benchmark komutları `benchmarks/README.md` içinde kayıtlıdır.

| Ölçüm | Başlangıç | Güncel | Sonuç |
|---|---:|---:|---:|
| Statik Word etiketleri, gerçek TASLAK | 4,440 sn | 0,051 sn | %98,8 azalma, 86,5× |
| `import k1` medyanı | 1,279 sn | 0,866 sn | %32 azalma |
| Bağımlılık kontrolü medyanı | 1,257 sn | 0,122 sn | %90 azalma |
| Etkileşimde tam proje toplama | yenileme başına 1 | 0 | yalnız kayıt/geçiş sınırında |

Sonraki performans adayları:

- Harita karo beklemesini `sleep`/iç içe `update()` yerine iptal edilebilir `after()` akışına almak.
- Büyük iş klasörü taramasını worker'a taşımak.
- Jeoloji kütüphanesinde sayfalama ve Treeview fark güncellemesi.
- Harita poligonlarını kimlik bazlı farkla yenilemek.
- Büyük Word şablonunun ön kontrol/üretim arasında tekrar açılmasını azaltmak.

## 10. Veri, güvenlik ve kurtarma

- Proje yolları taşınabilir/yeniden çözülebilir biçimde saklanmalıdır.
- JSON yazımı atomik olmalı; uygulama kaydedilmemiş değişiklikte kullanıcıyı uyarmalıdır.
- API anahtarları proje JSON'una yazılmamalı; Windows Kimlik Bilgileri Yöneticisi kullanılmalıdır.
- Kaynak Word/KML/KMZ/JPEG kanıtları ve seçim notları izlenebilir olmalıdır.

## 11. Kabul ölçütleri

- Tam otomatik test paketi ve eski unittest paketleri geçer.
- Gerçek Tk kurulumu olan bilgisayarda başlangıç/aç-kapat, salt okunur ve animasyon smoke testi geçer.
- 100/125/150% Windows DPI'da ana ekranlar taşmadan okunur.
- Rapor benchmarkı 500 ms, başlangıç benchmarkı 1 sn hedefini karşılar.
- Harita ve Word çıktılarında mevcut mühendislik sembolleri/biçimleri regresyona uğramaz.

## 12. Durum ve ilerleme — 13 Ağustos 2026

Tamamlananlar:

- Başlangıç Git kaydı: `7ff5c91`.
- Kill AI Slop taraması ve Tk/ttkbootstrap için elle görsel denetim.
- Merkezi sade stil sistemi, 32 px tablo satırları, tek ana işlem hiyerarşisi.
- Özet/Raporlama/Ekler/Taşıma/Harita ana işlem sadeleştirmesi.
- Merkezi pencere giriş/çıkış animasyonu ve azaltılmış hareket yolu.
- Kirli durum bayrağıyla tuş başına tüm proje serileştirmesinin kaldırılması.
- Ağır pandas/rapor yüklerinin geciktirilmesi ve hızlı bağımlılık metadata kontrolü.
- Toplu Word etiketi değiştirme ve tekrar üretilebilir benchmarklar.

Doğrulama bekleyenler:

- Bu makinedeki eksik Tcl/Tk (`init.tcl`) nedeniyle gerçek pencere/DPI görsel smoke testi.
- Harita ağ/karo bekleme akışının asenkronlaştırılması (yüksek riskli, ayrı sürüm).

Ertelenenler:

- Canlı Harita ve Word Tablo Tasarım Editörü (`CEBE_ATILANLAR.md`).
