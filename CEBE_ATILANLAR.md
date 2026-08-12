# Cebe Atılan Geliştirmeler

## Harita ve Rapor Tablosu Tasarım Editörü

- Kayıt tarihi: 10 Ağustos 2026
- Durum: Ertelendi; şu anda uygulanmayacak.
- Hatırlatma anahtarı: Kullanıcı “Ne yapmayı unuttuk?” veya “Ne yapacaktık?” dediğinde bu geliştirmeyi hatırlat.

### Amaç

Program içine, harita çizimlerindeki nesneleri ve rapor içindeki Word tablolarını biçimlendirmek için ortak bir **Tasarım Editörü** eklemek.

### Harita/pafta tarafı

- AÇ ve SS yazılarının metni, yazı tipi, boyutu, rengi, hizası ve konumu
- Araştırma çukuru, serim hattı ve diğer sembollerin tipi, boyutu ve rengi
- Parsel sınırı ile diğer çizgilerin rengi, kalınlığı ve çizgi tipi
- Başlık, açıklama alanı, koordinat tablosu ve pafta çerçevesinin düzenlenmesi
- Nesneleri fareyle taşıma ve yeniden boyutlandırma
- Mühendislik Jeolojisi, Jeoloji Lokasyon ve Jeofizik Lokasyon haritaları için ayrı şablonlar

### Rapor tablosu tarafı

- Yazı tipi, boyutu, kalınlığı, rengi ve yatay/dikey hizalama
- Satır yüksekliği, sütun genişliği ve hücre iç boşlukları
- Kenarlık tipi, rengi ve kalınlığı
- Hücre arka planı, başlık satırı ve hücre birleştirme
- Tablo genişliği, sayfa hizası, üst/alt boşluklar
- Satırların sayfada bölünmesini engelleme ve başlık satırını yeni sayfalarda tekrarlama
- Hesaplanan hücreleri varsayılan olarak kilitli tutma

### Canlı önizleme

- Ayarlar solda, seçilen harita veya tablo sağda yan yana gösterilecek.
- Değişiklikler önizlemeye anında yansıyacak; kalıcı kayıt yalnız **Uygula/Kaydet** ile yapılacak.
- Seçilen nesne, hücre, satır veya sütun önizlemede vurgulanacak.
- Geri Al, İleri Al, İptal ve Varsayılana Dön seçenekleri olacak.
- Harita önizlemesi nihai çıktıyla birebir üretilecek.
- Tablolar için hızlı canlı önizlemeye ek olarak kesin sayfa düzenini gösteren **Word Görünümünü Yenile** seçeneği bulunacak.

### Ayar kapsamı

- Yalnız mevcut projeye uygula
- İlgili harita veya tablo türünün varsayılanı yap
- Bütün yeni projeler için genel şablon olarak kaydet
- Varsayılan şablona sıfırla

### Teknik yaklaşım

Oluşturulmuş JPG veya DOCX dosyasının pikselleri/son hali düzenlenmeyecek. Harita nesneleri ve tablo stilleri ayrı, kimlikli şablon ayarları olarak saklanacak; rapor ve haritalar bu ayarlardan yeniden üretilecek.
