import streamlit as st
import pandas as pd
from datetime import datetime
from github import Github
import io

st.set_page_config(page_title="Evidence Prohlídek Online", page_icon="🚂", layout="wide")
st.title("🚂 Online Evidence a správa prohlídek")

# --- 1. GITHUB INTEGRACE ---
try:
    g = Github(st.secrets["GITHUB_TOKEN"])
    repo = g.get_repo(st.secrets["REPO_NAME"])
except Exception as e:
    st.error("Není nastaven GitHub Token nebo REPO_NAME v Secrets na Streamlit Cloud!")
    st.stop()

# --- 2. NAČÍTÁNÍ EXCELU Z GITHUB ---
def nacist_data_z_excelu():
    try:
        content = repo.get_contents("prohlidky.xlsx")
        excel_data = io.BytesIO(content.decoded_content)
        
        xls = pd.ExcelFile(excel_data)
        dostupne_listy = xls.sheet_names
        
        st.caption(f"ℹ️ Načtené listy v souboru: **{', '.join(dostupne_listy)}**")

        # Načtení podle pořadí záložek (0 = 1. list, 1 = 2. list, 2 = 3. list)
        kbs_df = pd.read_excel(xls, sheet_name=0, dtype=str) if len(dostupne_listy) > 0 else pd.DataFrame()
        ls06_df = pd.read_excel(xls, sheet_name=1, dtype=str) if len(dostupne_listy) > 1 else pd.DataFrame()
        radio_df = pd.read_excel(xls, sheet_name=2, dtype=str) if len(dostupne_listy) > 2 else pd.DataFrame()
        
        return kbs_df.fillna(""), ls06_df.fillna(""), radio_df.fillna("")
    except Exception as e:
        st.error(f"Chyba při načítání souboru prohlidky.xlsx z GitHubu: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# --- 3. UKLÁDÁNÍ VŠECH LISTŮ DO EXCELU ---
def ulozit_vse_do_excelu(kbs_df, ls06_df, radio_df):
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            kbs_df.to_excel(writer, sheet_name='KBS', index=False)
            ls06_df.to_excel(writer, sheet_name='LS06', index=False)
            radio_df.to_excel(writer, sheet_name='Radiostanice', index=False)
        
        excel_bytes = output.getvalue()
        content = repo.get_contents("prohlidky.xlsx")
        repo.update_file(content.path, "Aktualizace dat z webové aplikace", excel_bytes, content.sha)
        st.toast("✅ Všechny 3 listy byly úspěšně uloženy do prohlidky.xlsx na GitHub!", icon="💾")
    except Exception as e:
        st.error(f"Chyba při ukládání na GitHub: {e}")

# --- 4. NAČTENÍ A PŘEJMENOVÁNÍ DAT ---
kbs_df, ls06_df, radio_df = nacist_data_z_excelu()

# Sem doplň názvy sloupců, které chceš změnit:
kbs_df = kbs_df.rename(columns={
    "Unnamed: 0": "Číslo mašiny",
    "Datum provedení prohlídky a rozsah": "Datum provedení",
    "Unnamed: 2": "Provedena prohlídka",
    " Přístí prohlídka a rozsah": "Příští prohlídka",
    "Unnamed: 4": "Budoucí prohlídka",
    "Starý název 2": "Nový název 2"
})

ls06_df = ls06_df.rename(columns={
    "Unnamed: 0": "Číslo mašiny",
    "Datum provedení prohlídky a rozsah": "Datum prohlídky",
    "Unnamed: 2": "Provedena prohlídka",
    "Přístí prohlídka a rozsah": "Příští prohlídka",
    "Unnamed: 4": "Budoucí prohlídka",
    "2026-09-06 00:00:00": "Typ",
    "AKTUÁLNÍ PROHLÍDKA": "Výrobní číslo",
})

radio_df = radio_df.rename(columns={
    "Starý název 1": "Nový název 1"
})

# --- 5. VÝPOČET TERMÍNŮ A ZVÝRAZNĚNÍ ---
dnes = datetime.now()
zacatek_aktualniho_mesice = dnes.replace(day=1, hour=0, minute=0, second=0, microsecond=0)

aktualni_mesic = dnes.month
aktualni_rok = dnes.year

pristi_mesic_datum = (zacatek_aktualniho_mesice + pd.Timedelta(days=32)).replace(day=1)
pristi_mesic = pristi_mesic_datum.month
pristi_rok = pristi_mesic_datum.year

# 1. Propadlé prohlídky (před tímto měsícem)
def ziskej_propadle_masiny(df, sloupec_data):
    if df.empty or sloupec_data not in df.columns:
        return []
    df_temp = df.copy()
    df_temp["_dt"] = pd.to_datetime(df_temp[sloupec_data], errors="coerce")
    filtrovane = df_temp[df_temp["_dt"] < zacatek_aktualniho_mesice]
    prvni_sloupec = df.columns[0]
    return filtrovane[prvni_sloupec].dropna().unique().tolist()

# 2. Prohlídky v tomto měsíci
def ziskej_masiny_tento_mesic(df, sloupec_data):
    if df.empty or sloupec_data not in df.columns:
        return []
    df_temp = df.copy()
    df_temp["_dt"] = pd.to_datetime(df_temp[sloupec_data], errors="coerce")
    filtrovane = df_temp[
        (df_temp["_dt"].dt.month == aktualni_mesic) & 
        (df_temp["_dt"].dt.year == aktualni_rok)
    ]
    prvni_sloupec = df.columns[0]
    return filtrovane[prvni_sloupec].dropna().unique().tolist()

# 3. Prohlídky příští měsíc
def ziskej_masiny_pristi_mesic(df, sloupec_data):
    if df.empty or sloupec_data not in df.columns:
        return []
    df_temp = df.copy()
    df_temp["_dt"] = pd.to_datetime(df_temp[sloupec_data], errors="coerce")
    filtrovane = df_temp[
        (df_temp["_dt"].dt.month == pristi_mesic) & 
        (df_temp["_dt"].dt.year == pristi_rok)
    ]
    prvni_sloupec = df.columns[0]
    return filtrovane[prvni_sloupec].dropna().unique().tolist()

# Zvýraznění řádků (Červená = Propadlé, Oranžová = Tento měsíc, Žlutá = Příští měsíc)
def zvyrazni_terminy(row, sloupec_data):
    try:
        dt = pd.to_datetime(row[sloupec_data])
        if pd.notnull(dt):
            if dt < zacatek_aktualniho_mesice:
                return ['background-color: #ffcdd2; color: #b71c1c; font-weight: bold'] * len(row)  # Červená
            elif dt.month == aktualni_mesic and dt.year == aktualni_rok:
                return ['background-color: #ffe0b2; color: #e65100; font-weight: bold'] * len(row)  # Oranžová
            elif dt.month == pristi_mesic and dt.year == pristi_rok:
                return ['background-color: #fff9c4; color: #f57f17; font-weight: bold'] * len(row)  # Žlutá
    except:
        pass
    return [''] * len(row)
# --- 6. UI A TABULKY ---
tab_kbs, tab_ls06, tab_radio = st.tabs(["📋 KBS prohlídky", "📟 LS06", "📻 Radiostanice"])

# --- 6. UI A TABULKY ---
tab_kbs, tab_ls06, tab_radio = st.tabs(["📋 KBS prohlídky", "📟 LS06", "📻 Radiostanice"])

# ==================== KBS List ====================
with tab_kbs:
    st.subheader("KBS Prohlídky")
    
    # 1. Vyhledání datových sloupců
    sloupce_s_datem_kbs = []
    mozne_nazvy_kbs = ["Datum provedení", "Příští prohlídka", "Další V1"]
    for col in kbs_df.columns:
        if col in mozne_nazvy_kbs or "datum" in col.lower():
            sloupce_s_datem_kbs.append(col)

    # 2. Převod na datetime a formátování na psaný měsíc a rok
    kbs_config = {}
    for col in sloupce_s_datem_kbs:
        kbs_df[col] = pd.to_datetime(kbs_df[col], errors="coerce")
        kbs_config[col] = st.column_config.DateColumn(
            col,
            format="MMMM YYYY",
            step=1
        )

    # 3. Kontrola termínů
    sloupec_pro_upozorneni = "Příští prohlídka" if "Příští prohlídka" in kbs_df.columns else (sloupce_s_datem_kbs[-1] if sloupce_s_datem_kbs else None)
    
    if sloupec_pro_upozorneni:
        propadle = ziskej_propadle_masiny(kbs_df, sloupec_pro_upozorneni)
        tento = ziskej_masiny_tento_mesic(kbs_df, sloupec_pro_upozorneni)
        pristi = ziskej_masiny_pristi_mesic(kbs_df, sloupec_pro_upozorneni)
        
        if propadle:
            st.error(f"🚨 **PROPADLÁ PROHLÍDKA:** {', '.join(propadle)}")
        if tento:
            st.warning(f"🟧 **PROHLÍDKA TENTO MĚSÍC ({aktualni_mesic}/{aktualni_rok}):** {', '.join(tento)}")
        if pristi:
            st.info(f"🟨 **Pozor na příští měsíc ({pristi_mesic}/{pristi_rok}):** {', '.join(pristi)}")
            
        styled_kbs = kbs_df.style.apply(zvyrazni_terminy, sloupec_data=sloupec_pro_upozorneni, axis=1)
    else:
        styled_kbs = kbs_df

    # 4. Vykreslení tabulky
    edited_kbs = st.data_editor(
        styled_kbs, 
        use_container_width=True, 
        num_rows="dynamic", 
        hide_index=True, 
        column_config=kbs_config,
        key="kbs_editor"
    )

# ==================== LS06 List ====================
with tab_ls06:
    st.subheader("LS06")
    
    # 1. Vyhledání datových sloupců
    sloupce_s_datem_ls06 = []
    mozne_nazvy_ls06 = ["Příští datum", "Datum prohlídky", "Datum vykonání", "Datum"]
    for col in ls06_df.columns:
        if col in mozne_nazvy_ls06 or "datum" in col.lower():
            sloupce_s_datem_ls06.append(col)

    # 2. Převod na datetime a formátování na psaný měsíc a rok
    ls06_config = {}
    for col in sloupce_s_datem_ls06:
        ls06_df[col] = pd.to_datetime(ls06_df[col], errors="coerce")
        ls06_config[col] = st.column_config.DateColumn(
            col,
            format="MMMM YYYY",
            step=1
        )

    # 3. Kontrola termínů
    sloupec_pro_upozorneni = sloupce_s_datem_ls06[-1] if sloupce_s_datem_ls06 else None
    if sloupec_pro_upozorneni:
        propadle = ziskej_propadle_masiny(ls06_df, sloupec_pro_upozorneni)
        tento = ziskej_masiny_tento_mesic(ls06_df, sloupec_pro_upozorneni)
        pristi = ziskej_masiny_pristi_mesic(ls06_df, sloupec_pro_upozorneni)
        
        if propadle:
            st.error(f"🚨 **PROPADLÁ PROHLÍDKA:** {', '.join(propadle)}")
        if tento:
            st.warning(f"🟧 **PROHLÍDKA TENTO MĚSÍC ({aktualni_mesic}/{aktualni_rok}):** {', '.join(tento)}")
        if pristi:
            st.info(f"🟨 **Pozor na příští měsíc ({pristi_mesic}/{pristi_rok}):** {', '.join(pristi)}")
            
        styled_ls06 = ls06_df.style.apply(zvyrazni_terminy, sloupec_data=sloupec_pro_upozorneni, axis=1)
    else:
        styled_ls06 = ls06_df

    # 4. Vykreslení tabulky
    edited_ls06 = st.data_editor(
        styled_ls06, 
        use_container_width=True, 
        num_rows="dynamic", 
        hide_index=True, 
        column_config=ls06_config,
        key="ls06_editor"
    )

# ==================== Radiostanice List ====================
with tab_radio:
    st.subheader("Radiostanice")
    
    # 1. Vyhledání datových sloupců
    sloupce_s_datem_radio = []
    mozne_nazvy_radio = ["Datum prohlídky", "Příští prohlídka", "Datum"]
    for col in radio_df.columns:
        if col in mozne_nazvy_radio or "datum" in col.lower():
            sloupce_s_datem_radio.append(col)

    # 2. Převod na datetime a formátování na psaný měsíc a rok
    radio_config = {}
    for col in sloupce_s_datem_radio:
        radio_df[col] = pd.to_datetime(radio_df[col], errors="coerce")
        radio_config[col] = st.column_config.DateColumn(
            col,
            format="MMMM YYYY",
            step=1
        )

    # 3. Kontrola termínů
    sloupec_pro_upozorneni = sloupce_s_datem_radio[-1] if sloupce_s_datem_radio else None
    if sloupec_pro_upozorneni:
        propadle = ziskej_propadle_masiny(radio_df, sloupec_pro_upozorneni)
        tento = ziskej_masiny_tento_mesic(radio_df, sloupec_pro_upozorneni)
        pristi = ziskej_masiny_pristi_mesic(radio_df, sloupec_pro_upozorneni)
        
        if propadle:
            st.error(f"🚨 **PROPADLÁ PROHLÍDKA:** {', '.join(propadle)}")
        if tento:
            st.warning(f"🟧 **PROHLÍDKA TENTO MĚSÍC ({aktualni_mesic}/{aktualni_rok}):** {', '.join(tento)}")
        if pristi:
            st.info(f"🟨 **Pozor na příští měsíc ({pristi_mesic}/{pristi_rok}):** {', '.join(pristi)}")
            
        styled_radio = radio_df.style.apply(zvyrazni_terminy, sloupec_data=sloupec_pro_upozorneni, axis=1)
    else:
        styled_radio = radio_df

    # 4. Vykreslení tabulky
    edited_radio = st.data_editor(
        styled_radio, 
        use_container_width=True, 
        num_rows="dynamic", 
        hide_index=True, 
        column_config=radio_config,
        key="radio_editor"
    )

# --- 7. GLOBÁLNÍ TLAČÍTKO ULOŽIT ---
st.divider()
if st.button("💾 Uložit všechny změny do Excelu na GitHub", type="primary", use_container_width=True):
    ulozit_vse_do_excelu(edited_kbs, edited_ls06, edited_radio)
