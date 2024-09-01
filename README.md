# İzin Formu Uygulaması

Bu Python uygulaması, çalışanların izin taleplerini yönetmek ve belgelemek için geliştirilmiş bir masaüstü uygulamasıdır. `tkinter` ve `openpyxl` gibi kütüphaneler kullanılarak oluşturulmuştur ve çeşitli işlevleri desteklemektedir.

## Özellikler

- **Tarih Seçimi:** Kullanıcı, izin başlangıç ve bitiş tarihlerini seçebilir. Bu tarihleri temel alarak izin gün sayısı otomatik olarak hesaplanır.
- **Kişi Bilgileri Yönetimi:** Ad, soyad ve görev gibi bilgileri hızlıca girmenizi sağlar. Kişi bilgileri bir dosyadan yüklenebilir ve düzenlenebilir.
- **Yarım Gün ve Ucu Açık İzin Seçenekleri:** Kullanıcılar yarım gün izin veya ucu açık (belirsiz) izin talebinde bulunabilir.
- **Masaüstüne Kopyalama:** Oluşturulan izin formu, istenirse otomatik olarak masaüstüne kopyalanabilir.
- **Dosya Kaydetme ve Yazdırma:** Oluşturulan izin formu, Excel dosyası olarak kaydedilir ve yazdırılabilir.
- **Ayar Yönetimi:** Uygulama ayarları, kullanıcı tarafından düzenlenebilir ve bu ayarlar `settings.json` dosyasında saklanır.
- **Tema Desteği:** Uygulama, `azure-dark` temasını kullanır ve modern bir kullanıcı arayüzü sunar.
- **Excel Entegrasyonu:** İzin formları Excel formatında kaydedilir ve farklı konumlara kopyalanabilir.

## Gereksinimler

- Python 3.7+
- Aşağıdaki Python kütüphaneleri:
  - `tkinter`
  - `tkcalendar`
  - `openpyxl`
  - `pywin32`
  - `json`

## Kurulum

1. Gerekli kütüphaneleri yüklemek için aşağıdaki komutu kullanın:

   pip install tkcalendar openpyxl pywin32
   
Kullanım
Uygulamayı başlattığınızda, ana form ekranı açılır. İlgili bilgileri doldurduktan sonra "Kaydet" veya "Kaydet ve Yazdır" düğmelerini kullanarak formu oluşturabilir ve yazdırabilirsiniz.

Ayrıca, uygulamanın ayarlarını menüden düzenleyebilir ve kişi bilgilerini güncelleyebilirsiniz.

Ayarlar
Ayarlar settings.json dosyasında saklanır ve kullanıcı tarafından düzenlenebilir. Örneğin, form üzerindeki hücrelerin yerlerini değiştirmek veya bir alanı devre dışı bırakmak mümkündür.
