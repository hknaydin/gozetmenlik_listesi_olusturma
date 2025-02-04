# coding=utf-8
import re
from datetime import datetime, date

import pandas as pd
#import sys # This import is not needed anymore
#from importlib import reload # This import is not needed anymore

#reload(sys) # This line is not needed anymore
#sys.setdefaultencoding('utf-8') # This line is not needed anymore and causing the error
# Dosyayı okur ve bir dataframe'e dönüştürür.
df = pd.read_excel('/content/data/sinav_programi_ek2.xlsx', engine='openpyxl') # Use 'openpyxl' for .xlsx files or 'xlrd' for .xls files

def format_date(date_str):
    if date_str == "ara sınav yok" or date_str == "sınav yok":
        return datetime.max.date()
    else:
        try:
            tarih = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S').date()
        except ValueError:
            try:
                tarih = datetime.strptime(date_str, '%d.%m.%Y').date()
            except ValueError:
                tarih = datetime.max.date()  # Eğer tarih okunamazsa, max değeri döndür.

        return tarih
# Sınıfı tanımlar.
class Sinav(object):
    """
    Bir sınavı temsil eden sınıf.

    Özellikleri:
      - sinif: Sınıf.
      - sinav_adi: Sınav adı.
      - ders_kodu: Ders kodu.
      - ogretim_uyesi: Öğretim üyesi.
      - tarih: Sınav tarihi.
      - saat: Sınav saati.
      - yer: Sınav yeri.
    """
    def __init__(self, sinif, sinav_adi, ders_kodu, ogretim_uyesi, tarih, saat, yer):
        self.sinif = sinif
        self.sinav_adi = sinav_adi
        self.ders_kodu = ders_kodu
        self.ogretim_uyesi = ogretim_uyesi
        self.tarih = tarih
        self.saat = saat
        self.yer = yer

sinavlar = []
for i in range(len(df)):
    sinif = str(df.loc[i, 'Sınıf'])
    sinav_adi = df.loc[i, 'Dersin Adi'].strip()
    ders_kodu = df.loc[i, 'Ders Kodu'].strip()
    ogretim_uyesi = df.loc[i, 'Ogretim Uyesi'].strip()
    tarih = str(df.loc[i, 'Tarih'])
    saat = str(df.loc[i, 'Saat'])
    yer = str(df.loc[i, 'Yer'])

    sinav = Sinav(sinif, sinav_adi, ders_kodu, ogretim_uyesi, tarih, saat, yer)
    sinavlar.append(sinav)

# Tarih ve saat bilgisine göre sınavları sıralar
sinavlar_sirali = sorted(sinavlar, key=lambda x: (format_date(x.tarih), x.saat))
# Sıralanmış sınavları yazdırır.
for sinav in sinavlar_sirali:
    print(u"Sınav Adı: " + sinav.sinav_adi)
    print("Ders Kodu: " + sinav.ders_kodu)
    print(u"Öğretim Üyesi: " + sinav.ogretim_uyesi)
    print("Tarih: " + sinav.tarih)
    print("Saat: " + sinav.saat)
    print("Yer: " + sinav.yer)
    print("-" * 50)
# Sıralanmış sınavları Excel dosyasına yazar.
df_sirali = pd.DataFrame(columns=['Sınıf', 'Sınav Adı', 'Ders Kodu', 'Öğretim Üyesi', 'Tarih', 'Saat', 'Yer'])

all_sinavlar = []
for sinav in sinavlar_sirali:
    all_sinavlar.append({
        'Sınıf': sinav.sinif,
        'Sınav Adı': sinav.sinav_adi,
        'Ders Kodu': sinav.ders_kodu,
        'Öğretim Üyesi': sinav.ogretim_uyesi,
        'Tarih': format_date(sinav.tarih),
        'Saat': sinav.saat,
        'Yer': sinav.yer
    })

# Create the DataFrame from the list of dictionaries
df_sirali = pd.DataFrame(all_sinavlar)

# Sonuçları yeni bir Excel dosyasına yaz.
df_sirali.to_excel('sinav_programi_ek2_sirali.xlsx', index=False)
