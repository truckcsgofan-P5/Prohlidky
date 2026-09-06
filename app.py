import streamlit as st
import pandas as pd
from datetime import datetime
from github import Github
import io

st.set_page_config(page_title="Evidence Prohlídek Online", page_icon="🚂", layout="wide")
st.title("🚂 Online Evidence a správa prohlídek")

# --- GITHUB INTEGRACE ---
try:
    g = Github(st.secrets["GITHUB_TOKEN"])
    repo = g.get_repo(st.secrets["REPO_NAME"])
except Exception as e:
    st.error("Není nastaven GitHub Token v Secrets! Zkontrolujte nastavení Streamlit Cloud.")
    st.stop()

def nacist_csv(file_path, vychozi_data):
    try:
        content = repo.get_contents(file_path)
        return pd.read_csv(io.StringIO(content.decoded_content.decode('utf-8')), dtype=str)
    except Exception:
        df = pd.DataFrame(vychozi_data)
        repo.create_file(file_path, f"Inicializace {file_path}", df.to_csv(index=False))
        return df

def ulozit_csv(file_path, df):
    csv_data = df.to_csv(index=False)
    try:
        content = repo.get_contents(file_path)
        repo.update_file(content.path, "Aktualizace dat z webu", csv_data, content.sha)
        st.toast("✅ Úspěšně uloženo na GitHub!", icon="💾")
    except Exception as e:
        st.error(f"Chyba při ukládání na GitHub: {e}")

# --- PŘÍPRAVA TERMÍNŮ ---
dnes = datetime.now()
pristi_mesic_datum = (dnes.replace(day=1) + pd.Timedelta(days=32)).replace(day=1)
pristi_mesic = pristi_mesic_datum.month
pristi_rok = pristi_mesic_datum.year

def ziskej_masiny_pristi_mesic(df, sloupec_data):
    if df.empty or sloupec_data not in df.columns:
        return []
    df_temp = df.copy()
    df_temp["_dt"] = pd.to_datetime(df_temp[sloupec_data], errors="coerce")
    filtrovane = df_temp[
        (df_temp["_dt"].dt.month == pristi_mesic) & 
        (df_temp["_dt"].dt.year == pristi_rok)
    ]
    return filtrovane["Číslo mašiny"].dropna().unique().tolist()

def zvyrazni_pristi_mesic(row, sloupec_data):
    try:
        dt = pd.to_datetime(row[sloupec_data])
        if dt.month == pristi_mesic and dt.year == pristi_rok:
            return ['background-color: #ffeba0; color: #000000; font-weight: bold'] * len(row)
    except:
        pass
    return [''] * len(row)

# Funkce pro načtení konkrétních listů z Excelu uloženého na GitHubu
def nacist_data_z_excelu():
    try:
        content = repo.get_contents("prohlidky.xlsx")  # Název tvého Excel souboru na GitHubu
        excel_data = io.BytesIO(content.decoded_content)
        
        # Načtení podle přesných názvů listů v Excelu
        kbs_df = pd.read_excel(excel_data, sheet_name="KBS", dtype=str)
        ls06_df = pd.read_excel(excel_data, sheet_name="LS06", dtype=str)
        radio_df = pd.read_excel(excel_data, sheet_name="Radiostanice", dtype=str)
        
        return kbs_df, ls06_df, radio_df
    except Exception as e:
        st.error(f"Soubor prohlidky.xlsx nebyl na GitHubu nalezen nebo se nepodařilo načíst listy: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# Načtení reálných dat z Excelu
kbs_df, ls06_df, radio_df = nacist_data_z_excelu()

# --- HLAVNÍ NAVIGACE A LISTY ---
tab_kbs, tab_ls06, tab_radio = st.tabs(["📋 KBS prohlídky", "📟 LS06", "📻 Radiostanice"])

with tab_kbs:
    st.subheader("KBS Prohlídky")
    upozorneni = ziskej_masiny_pristi_mesic(kbs_df, "Datum další prohlídky")
    if upozorneni:
        st.warning(f"⚠️ **Pozor na prohlídku příští měsíc ({pristi_mesic}/{pristi_rok}):** {', '.join(upozorneni)}")
    
    edited_kbs = st.data_editor(
        kbs_df.style.apply(zvyrazni_pristi_mesic, sloupec_data="Datum další prohlídky", axis=1),
        use_container_width=True, num_rows="dynamic", hide_index=True, key="kbs_editor"
    )
    if st.button("💾 Uložit KBS na GitHub", type="primary", key="save_kbs"):
        ulozit_csv("kbs_data.csv", edited_kbs)

with tab_ls06:
    st.subheader("LS06")
    upozorneni = ziskej_masiny_pristi_mesic(ls06_df, "Příští datum")
    if upozorneni:
        st.warning(f"⚠️ **Pozor na prohlídku příští měsíc ({pristi_mesic}/{pristi_rok}):** {', '.join(upozorneni)}")
    
    edited_ls06 = st.data_editor(
        ls06_df.style.apply(zvyrazni_pristi_mesic, sloupec_data="Příští datum", axis=1),
        use_container_width=True, num_rows="dynamic", hide_index=True, key="ls06_editor"
    )
    if st.button("💾 Uložit LS06 na GitHub", type="primary", key="save_ls06"):
        ulozit_csv("ls06_data.csv", edited_ls06)

with tab_radio:
    st.subheader("Radiostanice")
    upozorneni = ziskej_masiny_pristi_mesic(radio_df, "Datum prohlídky")
    if upozorneni:
        st.warning(f"⚠️ **Pozor na prohlídku příští měsíc ({pristi_mesic}/{pristi_rok}):** {', '.join(upozorneni)}")
    
    edited_radio = st.data_editor(
        radio_df.style.apply(zvyrazni_pristi_mesic, sloupec_data="Datum prohlídky", axis=1),
        use_container_width=True, num_rows="dynamic", hide_index=True, key="radio_editor"
    )
    if st.button("💾 Uložit Radiostanice na GitHub", type="primary", key="save_radio"):
        ulozit_csv("radio_data.csv", edited_radio)
