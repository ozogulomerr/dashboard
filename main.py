import requests
import json
import pandas as pd
import os
import io
import math
from datetime import datetime

# --- AYARLAR ---
# API Key'i GitHub Secrets'tan alacağız, eğer bulamazsa (lokalde çalışırken) senin yazdığını kullanır.
API_KEY = os.environ.get("EVDS_API_KEY", "eEojQT7PgD") 
START_DATE = "01-01-2021"
# Bitiş tarihini dinamik yapmak daha iyidir, her gün güncelleneceği için bugünün tarihini alabiliriz.
# Ama senin formatın sabitse böyle de kalabilir.
END_DATE = datetime.now().strftime("%d-%m-%Y")

# EXCEL DOSYA YOLLARI (Direkt ana dizinden okuyacak)
PATH_BUTCE = "butce.xlsx"
PATH_NAKIT = "Nakit Dengesi.xlsx"
PATH_ATIL = "atılisgucu.xlsx"
PATH_PMI = "imalat sanayi pmi.xlsx"
PATH_GSYH_ONCU = "GSYH_Oncu.xlsx"

headers = {"key": API_KEY, "User-Agent": "Mozilla/5.0"}
requests.packages.urllib3.disable_warnings()

# --- YARDIMCI FONKSİYONLAR ---
def clean_nan(value):
    """NaN değerleri None (null) yapar, böylece JSON bozulmaz."""
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
    return value

def veri_cek_evds(series_code, aylik_yap=False):
    url = f"https://evds2.tcmb.gov.tr/service/evds/series={series_code}&startDate={START_DATE}&endDate={END_DATE}&type=json"
    if aylik_yap: url += "&frequency=5&aggregationTypes=avg"
    try:
        resp = requests.get(url, headers=headers, verify=False)
        data = resp.json()
        return data["items"] if "items" in data else []
    except: return []

def veri_cek_fred(series_id):
    try:
        url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}"
        resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, verify=False)
        if resp.status_code == 200:
            df = pd.read_csv(io.StringIO(resp.text))
            df.columns = [c.strip() for c in df.columns]
            
            # Tarih sütunu düzeltme
            if 'observation_date' in df.columns:
                df.rename(columns={'observation_date': 'DATE'}, inplace=True)
            
            if 'DATE' in df.columns:
                df['DATE'] = pd.to_datetime(df['DATE'])
                return df
    except: pass
    return pd.DataFrame()

print("📡 Veriler çekiliyor ve işleniyor...")

# ==========================================
# 1. MAKROEKONOMİK BİRİNCİL GÖSTERGELER
# ==========================================

# --- A) GSYH ---
gsyh_list = []
gsyh_raw = veri_cek_evds("TP.GSYIH26.IFK.ZH")
if gsyh_raw:
    temp_vals = [float(x["TP_GSYIH26_IFK_ZH"]) for x in gsyh_raw if x.get("TP_GSYIH26_IFK_ZH")]
    temp_dates = [x["Tarih"] for x in gsyh_raw if x.get("TP_GSYIH26_IFK_ZH")]
    for i in range(len(temp_vals)):
        if "2023" in temp_dates[i] or "2024" in temp_dates[i] or "2025" in temp_dates[i]:
            yillik = ((temp_vals[i] - temp_vals[i-4])/temp_vals[i-4])*100 if i>=4 else 0
            gsyh_list.append({"tarih": temp_dates[i], "yillik": round(yillik, 1)})

# --- B) TÜFE ---
tufe_list = []
tufe_raw = veri_cek_evds("TP.FG.J0")
if tufe_raw:
    vals = [float(x["TP_FG_J0"]) for x in tufe_raw if x.get("TP_FG_J0")]
    dates = [x["Tarih"] for x in tufe_raw if x.get("TP_FG_J0")]
    for i in range(len(vals)):
        if "2023" in dates[i] or "2024" in dates[i] or "2025" in dates[i]:
            aylik = ((vals[i] - vals[i-1])/vals[i-1])*100 if i>0 else 0
            yillik = ((vals[i] - vals[i-12])/vals[i-12])*100 if i>=12 else 0
            tufe_list.append({"tarih": dates[i], "aylik": round(aylik, 2), "yillik": round(yillik, 2)})

# --- C) Yİ-ÜFE ---
ufe_list = []
ufe_raw = veri_cek_evds("TP.TUFE1YI.T1")
if ufe_raw:
    vals = [float(x["TP_TUFE1YI_T1"]) for x in ufe_raw if x.get("TP_TUFE1YI_T1")]
    dates = [x["Tarih"] for x in ufe_raw if x.get("TP_TUFE1YI_T1")]
    for i in range(len(vals)):
        if "2023" in dates[i] or "2024" in dates[i] or "2025" in dates[i]:
            aylik = ((vals[i] - vals[i-1])/vals[i-1])*100 if i>0 else 0
            yillik = ((vals[i] - vals[i-12])/vals[i-12])*100 if i>=12 else 0
            ufe_list.append({"tarih": dates[i], "aylik": round(aylik, 2), "yillik": round(yillik, 2)})

# --- D) CARİ DENGE ---
cari_list = []
cari_raw = veri_cek_evds("TP.HARICCARIACIK.K1-TP.HARICCARIACIK.K10")
if cari_raw:
    df_cari = pd.DataFrame(cari_raw)
    df_cari['TP_HARICCARIACIK_K1'] = pd.to_numeric(df_cari['TP_HARICCARIACIK_K1'])
    df_cari['TP_HARICCARIACIK_K10'] = pd.to_numeric(df_cari['TP_HARICCARIACIK_K10'])
    df_cari['Cari_Yillik'] = df_cari['TP_HARICCARIACIK_K1'].rolling(window=12).sum()
    df_cari['Cekirdek_Yillik'] = df_cari['TP_HARICCARIACIK_K10'].rolling(window=12).sum()
    for _, row in df_cari.iterrows():
        yil = row["Tarih"].split("-")[0]
        if int(yil) >= 2023 and pd.notnull(row['Cari_Yillik']):
            cari_list.append({
                "tarih": row["Tarih"],
                "cari_aylik": int(row["TP_HARICCARIACIK_K1"]),
                "cari_yillik": int(row["Cari_Yillik"]),
                "cekirdek_aylik": int(row["TP_HARICCARIACIK_K10"]),
                "cekirdek_yillik": int(row["Cekirdek_Yillik"])
            })

# --- E) BÜTÇE DENGESİ ---
butce_list = []
if os.path.exists(PATH_BUTCE):
    try:
        df = pd.read_excel(PATH_BUTCE)
        df = df.dropna(subset=["Tarih"])
        for _, row in df.iterrows():
            tarih = str(row["Tarih"])
            if isinstance(row["Tarih"], pd.Timestamp): tarih = row["Tarih"].strftime('%d-%m-%Y')
            if "2023" in tarih or "2024" in tarih or "2025" in tarih:
                butce_list.append({
                    "tarih": tarih,
                    "butce_aylik": clean_nan(float(row["Bütçe Dengesi Aylık"])),
                    "butce_yillik": clean_nan(float(row["Bütçe Dengesi Yıllık"])),
                    "faizdisi_aylik": clean_nan(float(row["Faiz Dışı Denge Aylık"])),
                    "faizdisi_yillik": clean_nan(float(row["Faiz Dışı Denge Yıllık"]))
                })
    except: pass

# --- F) HAZİNE NAKİT ---
nakit_list = []
if os.path.exists(PATH_NAKIT):
    try:
        df = pd.read_excel(PATH_NAKIT)
        df = df.dropna(subset=["Tarih"])
        for _, row in df.iterrows():
            tarih = str(row["Tarih"])
            if isinstance(row["Tarih"], pd.Timestamp): tarih = row["Tarih"].strftime('%d-%m-%Y')
            if "2023" in tarih or "2024" in tarih or "2025" in tarih:
                nakit_list.append({
                    "tarih": tarih,
                    "nakit_aylik": clean_nan(float(row["Nakit Dengesi Aylık"])),
                    "nakit_yillik": clean_nan(float(row["Nakit Dengesi Yıllık"])),
                    "faizdisi_aylik": clean_nan(float(row["Faiz Dışı Nakit Denge Aylık"])),
                    "faizdisi_yillik": clean_nan(float(row["Faiz Dışı Nakit Denge Yıllık"]))
                })
    except: pass

# --- G) İŞGÜCÜ ---
atıl_dict = {}
if os.path.exists(PATH_ATIL):
    try:
        df = pd.read_excel(PATH_ATIL)
        df.columns = [c.strip() for c in df.columns] # Column isimlerini temizle
        if "Tarih" in df.columns and "Atıl İşgücü" in df.columns:
            df = df.dropna(subset=["Tarih", "Atıl İşgücü"])
            for _, row in df.iterrows():
                # Tarih formatlama (YYYY-MM-DD olarak standardize ediyoruz)
                try:
                    ts = pd.to_datetime(row["Tarih"], dayfirst=True)
                    t_str = ts.strftime('%Y-%m-%d')
                    
                    val = float(row["Atıl İşgücü"])
                    # Eğer oran olarak girildiyse (örn: 0.23 -> %23), 100 ile çarpıp yüzdelik yapıyoruz
                    if val < 1.0 and val > 0:
                        val = val * 100
                        
                    atıl_dict[t_str] = val
                except: pass
    except Exception as e:
        print(f"Excel hatası (Atıl İşgücü): {e}")

isgucu_list = []
isgucu_raw = veri_cek_evds("TP.TIG08-TP.TIG06")

if isgucu_raw:
    for item in isgucu_raw:
        raw_tarih = item["Tarih"]
        
        # EVDS tarihini parse edip formatlayalım
        try:
            dt = pd.to_datetime(raw_tarih, dayfirst=True)
            lookup_date = dt.strftime('%Y-%m-%d')
            year = dt.year
            
            if year >= 2023:
                # Sözlükten veriyi çek (Yoksa None döner)
                atil_val = atıl_dict.get(lookup_date)
                
                # Sadece veri varsa ekle (İsteğe bağlı, EVDS verisi varsa ekleyelim, atıl boş gelebilir)
                if item.get("TP_TIG08"):
                    isgucu_list.append({
                        "tarih": raw_tarih, 
                        "issizlik": float(item["TP_TIG08"]),
                        "katilim": float(item["TP_TIG06"]),
                        "atil": atil_val
                    })
        except: pass

# --- H) TCMB FONLAMA ---
fon_list = []
# Serileri parametre olarak tanımlayalım
series = "TP.APIFON4-TP.BISTTLREF.ORAN-TP.APIFON3"
fon_raw = veri_cek_evds(series)

# Hata ayıklama için ilk veriyi kontrol etmekte fayda var (print ile)
# print(fon_raw[0]) 

for item in fon_raw:
    tarih = item["Tarih"]
    
    # DÜZELTME 1: Yılı almak için son parçayı alıyoruz (GG-AA-YYYY varsayımıyla)
    # Eğer veri YYYY-MM-DD geliyorsa senin kodun doğruydu, 
    # ama boş dönüyorsa %99 ihtimalle format GG-AA-YYYY'dir.
    yil_str = tarih.split("-")[-1] 
    
    # Güvenlik önlemi: Yıl verisi bazen boş gelebilir, kontrol edelim
    if not yil_str.isdigit():
        continue
        
    yil = int(yil_str)
    
    # TP_APIFON4 verisinin None olmadığından emin olalım
    if yil >= 2023 and item.get("TP_APIFON4") is not None:
        try:
            fon_list.append({
                "tarih": tarih,
                "aofm": float(item["TP_APIFON4"]),
                # get kullanırken default değer atamak iyidir
                "tlref": float(item.get("TP_BISTTLREF_ORAN") or 0), 
                "net_fonlama": int(float(item.get("TP_APIFON3") or 0))
            })
        except ValueError:
            # Veri bazen boş string "" gelebilir, float'a çevrilemez
            continue
            
# ==========================================
# I. GSYH ÖNCÜ GÖSTERGELERİ
# ==========================================
oncu_gostergeler_list = []

# --- 1. SANAYİ (EVDS'den çekilmeye devam edecek) ---

# --- 2. DİĞERLERİ (Excel'den çekilecek) ---
# Hizmet, Ticaret, Perakende, İnşaat, Sanayi

excel_mapping = {
    "Hizmet": "hizmet",
    "Ticaret": "ticaret",
    "Perakende": "perakende",
    "İnşaat": "insaat",
    "Sanayi": "sanayi"
}

if os.path.exists(PATH_GSYH_ONCU):
    try:
        df_oncu = pd.read_excel(PATH_GSYH_ONCU)
        
        # Tarih formatını garantiye al
        df_oncu["Tarih"] = pd.to_datetime(df_oncu["Tarih"])
        df_oncu = df_oncu.sort_values("Tarih").reset_index(drop=True)
        
        # Her bir tür için işlem yap
        for col_name, json_key in excel_mapping.items():
            if col_name in df_oncu.columns:
                temp_series = []
                vals = df_oncu[col_name].values
                dates = df_oncu["Tarih"].values
                
                for i in range(len(vals)):
                    current_date = pd.to_datetime(dates[i])
                    
                    if current_date.year >= 2023:
                        # Yıllık değişim hesabı (12 ay öncesine göre)
                        if i >= 12:
                            val_now = vals[i]
                            val_prev = vals[i-12]
                            if val_prev != 0:
                                yillik_degisim = ((val_now - val_prev) / val_prev) * 100
                            else:
                                yillik_degisim = 0
                        else:
                            yillik_degisim = 0
                        
                        # Tarih formatını EVDS gibi yapalım (İsteğe bağlı, burada dd-mm-yyyy kullanıyorum)
                        tarih_str = current_date.strftime('%d-%m-%Y')
                        
                        temp_series.append({
                            "tarih": tarih_str,
                            "deger": float(vals[i]),
                            "yillik": round(yillik_degisim, 1)
                        })
                
                oncu_gostergeler_list.append({
                    "tur": json_key,
                    "data": temp_series
                })
            else:
                print(f"Uyarı: Excel dosyasında '{col_name}' sütunu bulunamadı.")
                
    except Exception as e:
        print(f"GSYH Öncü Excel okuma hatası: {e}")
else:
    print(f"Hata: GSYH Öncü Excel dosyası bulunamadı: {PATH_GSYH_ONCU}")
# ==========================================
# 2. İKİNCİL GÖSTERGELER
# ==========================================

# --- A) İMALAT SANAYİ PMI VERİSİNİ HAZIRLAMA ---
pmi_dict = {}

if os.path.exists(PATH_PMI):
    try:
        df = pd.read_excel(PATH_PMI)
        df = df.dropna(subset=["Tarih", "İmalat Sanayi PMI"])
        
        # Excel'deki tarihi datetime objesine çevir
        df["Tarih"] = pd.to_datetime(df["Tarih"])
        
        for _, row in df.iterrows():
            # Standardize et: YYYY-MM-DD (Örn: 2023-01-01)
            t_str = row["Tarih"].strftime('%Y-%m-%d')
            pmi_dict[t_str] = float(row["İmalat Sanayi PMI"])
        
    except Exception as e:
        print(f"Excel okuma hatası: {e}")
else:
    print("Excel dosyası bulunamadı.")

# --- B) EVDS VERİSİ İLE BİRLEŞTİRME (DÜZELTİLMİŞ) ---
imalat_list = []

# EVDS'den KKO verisini çekiyoruz
kko_raw = veri_cek_evds("TP.KKO2.IS.TOP") 

if 'kko_raw' in locals() and kko_raw:
    for item in kko_raw:
        raw_tarih = item["Tarih"] 
        
        try:
            dt_object = pd.to_datetime(raw_tarih)
            lookup_key = dt_object.strftime('%Y-%m-%d')
            
            if dt_object.year >= 2023:
                # 1. PMI Verisini Sözlükten Çek (Yoksa None gelir)
                pmi_degeri = pmi_dict.get(lookup_key)
                
                # 2. KKO Verisini Al
                val_kko_raw = item.get("TP_KKO2_IS_TOP")
                
                # 3. KONTROL: Her iki veri de MEVCUTSA ve 0'dan BÜYÜKSE listeye ekle.
                # Böylece PMI henüz açıklanmadıysa o ayı listeye hiç almayız,
                # grafik ve panel bir önceki (tam olan) ayda durur.
                if pmi_degeri is not None and val_kko_raw is not None:
                    val_kko = float(val_kko_raw)
                    
                    if pmi_degeri > 0 and val_kko > 0:
                        imalat_list.append({
                            "tarih": lookup_key,
                            "kko": val_kko,
                            "pmi": pmi_degeri
                        })
                        
        except Exception as err:
            print(f"Hata ({raw_tarih}): {err}")
else:
    print("kko_raw verisi yok.")

# --- B) GÜVEN ENDEKSLERİ ---
guven_list = []
# İki seriyi tek istekte çekiyoruz: Tüketici Güven ve Reel Kesim Güven
guven_raw = veri_cek_evds("TP.TG2.Y01-TP.GY1.N2.MA")

for item in guven_raw:
    tarih = item["Tarih"]
    yil = tarih.split("-")[0]
    
    if int(yil) >= 2023:
        # EVDS, JSON anahtarlarında noktaları alt çizgiye çevirir.
        # Bu yüzden "TP.GY1.N2.MA" yerine "TP_GY1_N2_MA" kullanmalıyız.
        
        # Veri bazen boş string "" veya None gelebilir, güvenli dönüşüm yapalım:
        val_tuketici = item.get("TP_TG2_Y01")
        val_reel = item.get("TP_GY1_N2_MA") # <--- Düzeltilen kısım (Nokta yerine Alt Çizgi)

        guven_list.append({
            "tarih": tarih,
            # Değer varsa float'a çevir, yoksa 0 ata
            "tuketici": int(float(val_tuketici)) if val_tuketici else 0,
            "reel": float(val_reel) if val_reel else 0
        })

# ==========================================
# 3. ABD & EURO (FRED)
# ==========================================

# --- FED FAİZ ---
fed_list = []
df_fed = veri_cek_fred("DFEDTARU") 
df_effr = veri_cek_fred("EFFR")     
df_lower = veri_cek_fred("DFEDTARL")

if not df_fed.empty and not df_effr.empty and 'DATE' in df_fed.columns and 'DATE' in df_effr.columns:
    m = pd.merge(df_fed, df_effr, on="DATE", how="inner", suffixes=('_U', '_E'))
    if not df_lower.empty and 'DATE' in df_lower.columns:
        m = pd.merge(m, df_lower, on="DATE", how="inner")
    
    m = m[m['DATE'].dt.year >= 2023]
    for _, r in m.iterrows():
        if len(r) >= 4:
            fed_list.append({
                "tarih": r["DATE"].strftime('%d-%m-%Y'),
                "ust": clean_nan(float(r.iloc[1])), 
                "efektif": clean_nan(float(r.iloc[2])), 
                "alt": clean_nan(float(r.iloc[3]))
            })

# --- ABD TÜFE ---
uscpi_list = []
df_uscpi = veri_cek_fred("CPIAUCSL")
if not df_uscpi.empty and 'DATE' in df_uscpi.columns:
    col_name = df_uscpi.columns[1]
    df_uscpi['MoM'] = df_uscpi[col_name].pct_change(1)*100
    df_uscpi['YoY'] = df_uscpi[col_name].pct_change(12)*100
    df_uscpi = df_uscpi[df_uscpi['DATE'].dt.year >= 2023]
    for _, r in df_uscpi.iterrows():
        if pd.notnull(r['MoM']):
            uscpi_list.append({
                "tarih": r["DATE"].strftime('%d-%m-%Y'), 
                "aylik": round(r["MoM"],1), 
                "yillik": round(r["YoY"],1)
            })

# --- ECB FAİZ ---
ecb_list = []
df1 = veri_cek_fred("ECBDFR")
df2 = veri_cek_fred("ECBMRRFR")
df3 = veri_cek_fred("ECBMLFR")
if not df1.empty and not df2.empty and not df3.empty:
    m = pd.merge(df1, df2, on="DATE", how="outer")
    m = pd.merge(m, df3, on="DATE", how="outer")
    m.fillna(method='ffill', inplace=True)
    m = m[m['DATE'].dt.year >= 2023]
    for _, r in m.iterrows():
        if len(r) >= 4:
            ecb_list.append({
                "tarih": r["DATE"].strftime('%d-%m-%Y'),
                "mevduat": clean_nan(float(r.iloc[1])),
                "refinans": clean_nan(float(r.iloc[2])),
                "marjinal": clean_nan(float(r.iloc[3]))
            })

# --- EURO TÜFE ---
eurocpi_list = []
df_eu = veri_cek_fred("CP0000EZ19M086NEST")
if not df_eu.empty and 'DATE' in df_eu.columns:
    col_name = df_eu.columns[1]
    df_eu['MoM'] = df_eu[col_name].pct_change(1)*100
    df_eu['YoY'] = df_eu[col_name].pct_change(12)*100
    df_eu = df_eu[df_eu['DATE'].dt.year >= 2023]
    for _, r in df_eu.iterrows():
        if pd.notnull(r['MoM']):
            eurocpi_list.append({
                "tarih": r["DATE"].strftime('%d-%m-%Y'), 
                "aylik": round(r["MoM"],1), 
                "yillik": round(r["YoY"],1)
            })

# ==========================================
# KAYDET VE TEMİZLE
# ==========================================

final_data = {
    "gsyh": gsyh_list, 
    "tufe": tufe_list, 
    "ufe": ufe_list, 
    "cari": cari_list,
    "butce": butce_list, 
    "nakit": nakit_list, 
    "isgucu": isgucu_list, 
    "fonlama": fon_list,
    "imalat": imalat_list, 
    "guven": guven_list, 
    "fed": fed_list, 
    "uscpi": uscpi_list,
    "ecb": ecb_list, 
    "eurocpi": eurocpi_list,
    "gsyh_oncu": oncu_gostergeler_list  # <--- BURAYI EKLE
}

# Recursively clean NaNs in the final structure
def sanitize_json(obj):
    if isinstance(obj, dict):
        return {k: sanitize_json(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [sanitize_json(i) for i in obj]
    elif isinstance(obj, float):
        if math.isnan(obj) or math.isinf(obj):
            return None
    return obj

final_clean = sanitize_json(final_data)

with open('veri.json', 'w', encoding='utf-8') as f:
    json.dump(final_clean, f, ensure_ascii=False, indent=4)

print("✅ GÜNCELLENDİ!")
