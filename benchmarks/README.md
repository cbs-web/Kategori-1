# K-1 performans benchmarkları

Çalışma dizininde:

```powershell
python benchmarks/benchmark_baslangic.py --repeat 7
python benchmarks/benchmark_rapor_etiketleri.py --repeat 3
```

Başlangıç ölçümü her tekrarı ayrı Python sürecinde çalıştırır ve kullanıcı profilinden
bağımsız, geçici bir uygulama veri klasörü kullanır. Rapor etiketi ölçümü gerçek
`ornek_sablonlar/rapor/TASLAK.docx` dosyasını kullanır.
