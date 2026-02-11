import requests
import json
import pandas as pd
import os
import io
import math
from datetime import datetime

# --- AYARLAR ---
API_KEY = os.environ.get("EVDS_API_KEY", "eEojQT7PgD") 
START_DATE = "01-01-2021"
END_DATE = datetime.now().strftime("%d-%m-%Y")

# EXCEL DOSYA YOLLARI
PATH_BUTCE = "butce.xlsx"
PATH_NAKIT = "Nakit Dengesi.xlsx"
PATH_ATIL = "atılisgucu.xlsx"
PATH_PMI = "imalat sanayi pmi.xlsx"
PATH_GSYH_ONCU = "GSYH_Oncu.xlsx"

headers = {"key": API_KEY, "User-Agent": "Mozilla/5.0"}
requests.packages.urllib3.disable_warnings()

# --- YARDIMCI FONKSİYONLAR ---
def clean_nan(value):
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
    return value

def get_year_from_date(date_str):
    """DD-MM-YYYY veya YYYY-MM-DD formatından yılı integer olarak döner."""
    try:
        if "-" in date_str:
            parts = date_str.split("-")
            if len(parts[0]) == 4: return int(parts[0]) # YYYY-MM-DD
            if len(parts[2]) == 4: return int(parts[2]) # DD-MM-YYYY
    except:
        return 0
    return 0

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
        # DÜZELTME: Sabit yıl kontrolü yerine >= 2023 kontrolü
        yil = get_year_from_date(temp_dates[i])
        if yil >= 2023:
            yillik = ((temp_vals[i] - temp_vals[i-4])/temp_vals[i-4])*100 if i>=4 else 0
            gsyh_list.append({"tarih": temp_dates[i], "yillik": round(yillik, 1)})

# --- B) TÜFE ---
tufe_list = []
tufe_raw = veri_cek_evds("TP.FG.J0")
if tufe_raw:
    vals = [float(x["TP_FG_J0"]) for x in tufe_raw if x.get("TP_FG_J0")]
    dates = [x["Tarih"] for x in tufe_raw if x.get("TP_FG_J0")]
    for i in range(len(vals)):
        yil = get_year_from_date(dates[i])
        if yil >= 2023:
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
        yil = get_year_from_date(dates[i])
        if yil >= 2023:
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
        yil = get_year_from_date(row["Tarih"])
        if yil >= 2023 and pd.notnull(row['Cari_Yillik']):
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
            yil = get_year_from_date(tarih)
            if yil >= 2023:
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
            yil = get_year_from_date(tarih)
            if yil >= 2023:
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
        df.columns = [c.strip() for c in df.columns]
        if "Tarih" in df.columns and "Atıl İşgücü" in df.columns:
            df = df.dropna(subset=["Tarih", "Atıl İşgücü"])
            for _, row in df.iterrows():
                try:
                    ts = pd.to_datetime(row["Tarih"], dayfirst=True)
                    t_str = ts.strftime('%Y-%m-%d')
                    val = float(row["Atıl İşgücü"])
                    if val < 1.0 and val > 0: val = val * 100
                    atıl_dict[t_str] = val
                except: pass
    except: pass

isgucu_list = []
isgucu_raw = veri_cek_evds("TP.TIG08-TP.TIG06")

if isgucu_raw:
    for item in isgucu_raw:
        raw_tarih = item["Tarih"]
        try:
            dt = pd.to_datetime(raw_tarih, dayfirst=True)
            lookup_date = dt.strftime('%Y-%m-%d')
            if dt.year >= 2023:
                atil_val = atıl_dict.get(lookup_date)
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
fon_raw = veri_cek_evds("TP.APIFON4-TP.BISTTLREF.ORAN-TP.APIFON3")

if fon_raw:
    df_fon = pd.DataFrame(fon_raw)
    df_fon.rename(columns={"Tarih": "tarih", "TP_APIFON4": "aofm", "TP_BISTTLREF_ORAN": "tlref", "TP_APIFON3": "net_fonlama"}, inplace=True)
    cols_to_convert = ["aofm", "tlref", "net_fonlama"]
    for col in cols_to_convert:
        df_fon[col] = pd.to_numeric(df_fon[col], errors='coerce')
    
    df_fon["tarih_dt"] = pd.to_datetime(df_fon["tarih"], dayfirst=True)
    df_fon = df_fon.sort_values("tarih_dt")
    # Düzeltme: Yıl >= 2023
    df_fon = df_fon[df_fon["tarih_dt"].dt.year >= 2023]
    df_fon = df_fon.dropna(subset=["aofm", "tlref"], how="all")

    for _, row in df_fon.iterrows():
        aofm_val = row["aofm"] if pd.notnull(row["aofm"]) else None
        tlref_val = row["tlref"] if pd.notnull(row["tlref"]) else None
        net_val = row["net_fonlama"] if pd.notnull(row["net_fonlama"]) else 0
        fon_list.append({"tarih": row["tarih"], "aofm": aofm_val, "tlref": tlref_val, "net_fonlama": net_val})

# ==========================================
# I. GSYH ÖNCÜ GÖSTERGELERİ (REVIZE EDİLDİ)
# ==========================================
oncu_gostergeler_list = []

# Yeni Sütun İsimleri Eşleşmesi
# JSON Key -> [Excel Başlık Prefix]
# Excel'de "Hizmet takvim ar." ve "Hizmet mevsim ar." şeklinde olduğu varsayılmıştır.
excel_sectors = {
    "hizmet": "Hizmet",
    "ticaret": "Ticaret",
    "perakende": "Perakende",
    "insaat": "İnşaat",
    "sanayi": "Sanayi"
}

if os.path.exists(PATH_GSYH_ONCU):
    try:
        df_oncu = pd.read_excel(PATH_GSYH_ONCU)
        # Sütun isimlerindeki boşlukları temizleyelim (garanti olsun)
        df_oncu.columns = [c.strip() for c in df_oncu.columns]
        
        df_oncu["Tarih"] = pd.to_datetime(df_oncu["Tarih"])
        df_oncu = df_oncu.sort_values("Tarih").reset_index(drop=True)
        
        for json_key, excel_prefix in excel_sectors.items():
            col_takvim = f"{excel_prefix} takvim ar."
            col_mevsim = f"{excel_prefix} mevsim ar."
            
            # Her iki sütun da var mı kontrol et
            if col_takvim in df_oncu.columns and col_mevsim in df_oncu.columns:
                temp_series = []
                vals_takvim = df_oncu[col_takvim].values
                vals_mevsim = df_oncu[col_mevsim].values
                dates = df_oncu["Tarih"].values
                
                for i in range(len(dates)):
                    current_date = pd.to_datetime(dates[i])
                    
                    if current_date.year >= 2023:
                        # 1. YILLIK DEĞİŞİM (Takvim ar. sütunundan, 12 ay öncesine göre)
                        if i >= 12:
                            v_now = vals_takvim[i]
                            v_prev = vals_takvim[i-12]
                            if v_prev != 0 and pd.notnull(v_now) and pd.notnull(v_prev):
                                yillik_degisim = ((v_now - v_prev) / v_prev) * 100
                            else: yillik_degisim = None
                        else: yillik_degisim = None
                        
                        # 2. AYLIK DEĞİŞİM (Mevsim ar. sütunundan, 1 ay öncesine göre)
                        if i >= 1:
                            v_now_m = vals_mevsim[i]
                            v_prev_m = vals_mevsim[i-1]
                            if v_prev_m != 0 and pd.notnull(v_now_m) and pd.notnull(v_prev_m):
                                aylik_degisim = ((v_now_m - v_prev_m) / v_prev_m) * 100
                            else: aylik_degisim = None
                        else: aylik_degisim = None
                        
                        # Veri Ekleme
                        tarih_str = current_date.strftime('%d-%m-%Y')
                        
                        # Eğer veriler NaN değilse listeye alalım
                        if yillik_degisim is not None or aylik_degisim is not None:
                            temp_series.append({
                                "tarih": tarih_str,
                                "yillik": round(yillik_degisim, 1) if yillik_degisim is not None else None,
                                "aylik": round(aylik_degisim, 1) if aylik_degisim is not None else None
                            })
                
                oncu_gostergeler_list.append({
                    "tur": json_key,
                    "data": temp_series
                })
            else:
                print(f"Uyarı: '{excel_prefix}' için gerekli sütunlar bulunamadı.")
                
    except Exception as e:
        print(f"GSYH Öncü Excel hatası: {e}")
else:
    print(f"Hata: Dosya yok -> {PATH_GSYH_ONCU}")

# ==========================================
# 2. İKİNCİL GÖSTERGELER
# ==========================================

# --- A) İMALAT SANAYİ PMI ---
pmi_dict = {}
if os.path.exists(PATH_PMI):
    try:
        df = pd.read_excel(PATH_PMI)
        df = df.dropna(subset=["Tarih", "İmalat Sanayi PMI"])
        df["Tarih"] = pd.to_datetime(df["Tarih"])
        for _, row in df.iterrows():
            t_str = row["Tarih"].strftime('%Y-%m-%d')
            pmi_dict[t_str] = float(row["İmalat Sanayi PMI"])
    except: pass

# --- B) KKO ve Birleştirme ---
imalat_list = []
kko_raw = veri_cek_evds("TP.KKO2.IS.TOP") 
if kko_raw:
    for item in kko_raw:
        raw_tarih = item["Tarih"] 
        try:
            dt_object = pd.to_datetime(raw_tarih)
            lookup_key = dt_object.strftime('%Y-%m-%d')
            if dt_object.year >= 2023:
                pmi_degeri = pmi_dict.get(lookup_key)
                val_kko_raw = item.get("TP_KKO2_IS_TOP")
                if pmi_degeri is not None and val_kko_raw is not None:
                    if pmi_degeri > 0 and float(val_kko_raw) > 0:
                        imalat_list.append({
                            "tarih": lookup_key,
                            "kko": float(val_kko_raw),
                            "pmi": pmi_degeri
                        })
        except: pass

# --- C) GÜVEN ENDEKSLERİ ---
guven_list = []
guven_raw = veri_cek_evds("TP.TG2.Y01-TP.GY1.N2.MA")
for item in guven_raw:
    tarih = item["Tarih"]
    yil = get_year_from_date(tarih)
    if yil >= 2023:
        val_tuketici = item.get("TP_TG2_Y01")
        val_reel = item.get("TP_GY1_N2_MA")
        guven_list.append({
            "tarih": tarih,
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
if not df_fed.empty and not df_effr.empty:
    m = pd.merge(df_fed, df_effr, on="DATE", how="inner", suffixes=('_U', '_E'))
    if not df_lower.empty: m = pd.merge(m, df_lower, on="DATE", how="inner")
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
if not df_uscpi.empty:
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
if not df_eu.empty:
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
# 4. PARA VE BANKA
# ==========================================
banka_list = []
banka_raw = veri_cek_evds("TP.KTF10-TP.KTF101-TP.KTF11-TP.KTF12-TP.KTF17-TP.KTF18-TP.TRY.MT01-TP.TRY.MT02-TP.TRY.MT06")
if banka_raw:
    for item in banka_raw:
        tarih = item["Tarih"]
        yil = get_year_from_date(tarih)
        if yil >= 2023:
            try:
                ihtiyac_net = float(item.get("TP_KTF10") or 0)
                ihtiyac_kmh_dahil = float(item.get("TP_KTF101") or 0)
                tasit = float(item.get("TP_KTF11") or 0)
                konut = float(item.get("TP_KTF12") or 0)
                ticari_genel = float(item.get("TP_KTF17") or 0)
                ticari_net = float(item.get("TP_KTF18") or 0)
                mev_1ay = float(item.get("TP_TRY_MT01") or 0)
                mev_3ay = float(item.get("TP_TRY_MT02") or 0)
                mev_toplam = float(item.get("TP_TRY_MT06") or 0)
                if (ihtiyac_net + ticari_genel + mev_toplam) > 0:
                    banka_list.append({
                        "tarih": tarih,
                        "ihtiyac_net": ihtiyac_net, "ihtiyac_toplam": ihtiyac_kmh_dahil,
                        "tasit": tasit, "konut": konut,
                        "ticari_toplam": ticari_genel, "ticari_net": ticari_net,
                        "mev_1ay": mev_1ay, "mev_3ay": mev_3ay, "mev_toplam": mev_toplam
                    })
            except: continue

yp_list = []
yp_raw = veri_cek_evds("TP.KTF17.EUR-TP.KTF17.USD-TP.EUR.MT06-TP.USD.MT06")
if yp_raw:
    for item in yp_raw:
        tarih = item["Tarih"]
        yil = get_year_from_date(tarih)
        if yil >= 2023:
            try:
                ticari_eur = float(item.get("TP_KTF17_EUR") or 0)
                ticari_usd = float(item.get("TP_KTF17_USD") or 0)
                mev_eur = float(item.get("TP_EUR_MT06") or 0)
                mev_usd = float(item.get("TP_USD_MT06") or 0)
                if (ticari_eur + ticari_usd + mev_eur + mev_usd) > 0:
                    yp_list.append({
                        "tarih": tarih, "ticari_eur": ticari_eur, "ticari_usd": ticari_usd,
                        "mev_eur": mev_eur, "mev_usd": mev_usd
                    })
            except: continue

para_arzi_list = []
m3_raw = veri_cek_evds("TP.HPBITABLO1.18-TP.HPBITABLO1.4-TP.HPBITABLO1.12-TP.HPBITABLO1.20")
if m3_raw:
    for item in m3_raw:
        tarih = item["Tarih"]
        yil = get_year_from_date(tarih)
        if yil >= 2021:
            try:
                m3 = float(item.get("TP_HPBITABLO1_18") or 0)
                vadesiz_tl = float(item.get("TP_HPBITABLO1_4") or 0)
                vadeli_tl = float(item.get("TP_HPBITABLO1_12") or 0)
                ppf = float(item.get("TP_HPBITABLO1_20") or 0)
                if m3 > 0:
                    para_arzi_list.append({
                        "tarih": tarih, "m3": m3, "vadesiz_tl": vadesiz_tl,
                        "vadeli_tl": vadeli_tl, "ppf": ppf
                    })
            except: continue

# ==========================================
# KAYDET
# ==========================================
eski_veri = {}
eski_meta = {}
if os.path.exists('veri.json'):
    try:
        with open('veri.json', 'r', encoding='utf-8') as f:
            eski_veri = json.load(f)
            if "meta" in eski_veri: eski_meta = eski_veri["meta"]
    except: pass

bugun_str = datetime.now().strftime("%Y-%m-%d")

def get_update_date(key_name, new_list):
    if not new_list: return eski_meta.get(key_name, "2000-01-01")
    try:
        new_last_date = new_list[-1]["tarih"]
        old_list = eski_veri.get(key_name, [])
        old_last_date = old_list[-1]["tarih"] if old_list else None
        if new_last_date != old_last_date: return bugun_str
        else: return eski_meta.get(key_name, bugun_str)
    except: return bugun_str

meta_data = {
    "gsyh": get_update_date("gsyh", gsyh_list),
    "tufe": get_update_date("tufe", tufe_list),
    "ufe": get_update_date("ufe", ufe_list),
    "cari": get_update_date("cari", cari_list),
    "butce": get_update_date("butce", butce_list),
    "nakit": get_update_date("nakit", nakit_list),
    "isgucu": get_update_date("isgucu", isgucu_list),
    "fonlama": get_update_date("fonlama", fon_list),
    "imalat": get_update_date("imalat", imalat_list),
    "guven": get_update_date("guven", guven_list),
    "fed": get_update_date("fed", fed_list),
    "uscpi": get_update_date("uscpi", uscpi_list),
    "ecb": get_update_date("ecb", ecb_list),
    "eurocpi": get_update_date("eurocpi", eurocpi_list),
    "gsyh_oncu": get_update_date("gsyh_oncu", oncu_gostergeler_list),
    "banka": get_update_date("banka", banka_list),
    "yp": get_update_date("yp", yp_list),
    "para_arzi": get_update_date("para_arzi", para_arzi_list)
}

final_data = {
    "meta": meta_data,
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
    "gsyh_oncu": oncu_gostergeler_list,
    "banka": banka_list,
    "yp": yp_list,
    "para_arzi": para_arzi_list
}

def sanitize_json(obj):
    if isinstance(obj, dict): return {k: sanitize_json(v) for k, v in obj.items()}
    elif isinstance(obj, list): return [sanitize_json(i) for i in obj]
    elif isinstance(obj, float):
        if math.isnan(obj) or math.isinf(obj): return None
    return obj

final_clean = sanitize_json(final_data)

with open('veri.json', 'w', encoding='utf-8') as f:
    json.dump(final_clean, f, ensure_ascii=False, indent=4)

print("✅ VERİLER GÜNCELLENDİ!")
