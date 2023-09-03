import pandas as pd

# Zurich verilerini oku
file_zurich = "SOURCE ZURICH.xlsx"
page_name = "GÜNCEL BEDELLER 2023-2024"
starting_row = 4
read_zurich = pd.read_excel(file_zurich, sheet_name=page_name, skiprows=starting_row - 1)

# Howden verilerini oku
read_anadolu = pd.read_excel('howden.xlsx', sheet_name='ANADOLU', skiprows=1)

# Zurich ve Howden verilerini karşılaştır
for i, row in read_zurich.iterrows():
    if row['SHEET'] == "ANADOLU":
        zurich_ucret = row['BİLDİRİLEN GÜNCEL BEDELLER 2023-2024']
        if zurich_ucret <= 0:
            continue
        else:
            zurich_deprem = row['DEPREM BÖLGESİ']
            
            # Howden verisini bul
            howden_row = read_anadolu[read_anadolu['DEPREM BÖLGESİ'] == zurich_deprem]
            if not howden_row.empty or howden_row=='':
                howden_ucret = howden_row.iloc[i]['2023-2024 SİGORTA BEDELLERİ']
                
                # Karşılaştırma yap
                if zurich_ucret == howden_ucret:
                    print(f"ANADOLU - Deprem Bölgesi: {zurich_deprem}, Zurich Ücret: {zurich_ucret}, Howden Ücret: {howden_ucret} - Eşleşiyor")
                else:
                    print(f"ANADOLU - Deprem Bölgesi: {zurich_deprem}, Zurich Ücret: {zurich_ucret}, Howden Ücret: {howden_ucret} - Eşleşmiyor")
            else:
                print(f"Howden verisi bulunamadı - Deprem Bölgesi: {zurich_deprem}")

