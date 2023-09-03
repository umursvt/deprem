import pandas as pd

# Excel dosyalarını okur
source_zurich = pd.read_excel('SOURCE ZURICH.xlsx', skiprows=3)
source_howden = pd.read_excel('howden.xlsx', sheet_name='SOM',  skiprows=1)


# ZURICH verilerini sözlüğe dönüştürür
dict_zurich = {}
for index, row in source_zurich.iterrows():
    if row['SHEET'] == "SOM":
        sigortalanan = row['SİGORTA BEDELİ']
        bedel = row['BİLDİRİLEN GÜNCEL BEDELLER 2023-2024']
        bolge = row['DEPREM BÖLGESİ']
        dict_zurich[(sigortalanan, bedel)] = bolge

# HOWDEN verilerini sözlüğe dönüştürür
dict_howden = {}
for index, row in source_howden.iterrows():
    sigortalanan = row['SİGORTA BEDELİ']
    bedel = row['2023-2024 SİGORTA BEDELLERİ']
    bolge = row['DEPREM BÖLGESİ']
    dict_howden[(sigortalanan, bedel)] = bolge

# Eşleşmeyi kontrol et ve KONTROL sütununu günceller
for index_z, row_z in source_zurich.iterrows():
    key_z = (row_z['SİGORTA BEDELİ'], row_z['BİLDİRİLEN GÜNCEL BEDELLER 2023-2024'])
    if key_z in dict_howden:
        value_z = row_z['DEPREM BÖLGESİ']
        value_h = dict_howden[key_z]
        index_h = source_howden[(source_howden['SİGORTA BEDELİ'] == key_z[0]) & (source_howden['2023-2024 SİGORTA BEDELLERİ'] == key_z[1])].index[0]
        if value_z == value_h:
            source_howden.at[index_h, 'KONTROL'] = 'EŞİT'
        else:
            source_howden.at[index_h,'KONTROL'] = 'EŞİT DEĞİL'




    # ANADOLU
excel_file = 'SOM.xlsx'
with pd.ExcelWriter(excel_file) as writer:

    source_howden.to_excel(writer, sheet_name='SOM', index=False)
  
  
print("Güncellenmiş veriler Excel dosyasına kaydedildi.")
