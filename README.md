# ExchangeServer-PST-Exporter-for-AD-OU
PST Exporter for Active Directory
Açıklama
PST-Exporter-for-Active-Directory bir PowerShell betiği olup, belirli bir Organizational Unit (OU) içindeki tüm posta kutularının PST dosyalarını dışa aktarır.

Özellikler
Belirli bir OU'dan posta kutularını alır.
Her posta kutusu için bir PST dışa aktarma talebi oluşturur.
Hedef klasörü otomatik olarak oluşturur eğer mevcut değilse.
Dışa aktarma taleplerinin durumunu kontrol eder ve zaten tamamlanmış talepleri atlar.
Hata mesajlarına karşı hata yakalama sağlar.
Kullanım
PowerShell betiğini indirin.
Betik içindeki parametreleri ihtiyacınıza göre düzenleyin:
$ouPath: Dışa aktarmak istediğiniz posta kutularının bulunduğu OU yolu.
$targetFolder: PST dosyalarının dışa aktarılacağı klasör yolu.
PowerShell betiğini çalıştırın.
Lisans
Bu projenin lisansı MIT lisansı ile korunmaktadır.

Bu öneriyi kullanarak GitHub'da bir repo oluşturabilir ve bu Readme'yi kullanabilirsiniz. Bu şekilde, repo hızla kurulabilir ve kullanıcılara nasıl kullanılacağı konusunda yönergeler sağlar.
