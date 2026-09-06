import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from github import Github
import io
import extra_streamlit_components as stx
import time
import hashlib

st.set_page_config(page_title="Evidence Prohlídek Online", page_icon="🚂", layout="wide")


# ==================== INICIALIZACE COOKIES ====================
cookie_manager = stx.CookieManager(key="muj_spravce_cookies")


# ==================== 0. PŘIHLAŠOVACÍ SYSTÉM ====================
if "users" in st.secrets:
    UZIVATELE = dict(st.secrets["users"])
else:
    UZIVATELE = {}
    st.error("⚠️ Nebyli načteni žádní uživatelé. Nastav prosím sekci [users] v Secrets.")

def vytvorit_bezpecny_token(username):
    """Vytvoří bezpečný hash (otisk) z hesla, aby se cookie nedala zfalšovat."""
    heslo = UZIVATELE.get(username, "")
    text_k_zasifrovani = f"{username}_tajny_klic_{heslo}"
    return hashlib.sha256(text_k_zasifrovani.encode()).hexdigest()

def overit_prihlaseni():
    # 1. Nejprve zkusíme načíst uživatele ze zachráněné Cookie
    cookie_val = cookie_manager.get(cookie="auth_token")
    
    if cookie_val and "::" in str(cookie_val):
        cookie_user, cookie_token = cookie_val.split("::", 1)
        # Ověříme, zda souhlasí token (obrana proti podvržení cookie)
        if cookie_user in UZIVATELE and vytvorit_bezpecny_token(cookie_user) == cookie_token:
            st.session_state["logged_in"] = True
            st.session_state["user"] = cookie_user
            return True

    # 2. Klasická session kontrola
    if st.session_state.get("logged_in"):
        return True

    # Zobrazení přihlašovacího okna
    st.markdown("<h2 style='text-align: center;'>🔐 Přihlášení do aplikace</h2>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("Uživatelské jméno")
            password = st.text_input("Heslo", type="password")
            # Přidáno zaškrtávací políčko
            remember_me = st.checkbox("Pamatovat si mě na tomto zařízení (30 dní)", value=True)
            submit = st.form_submit_button("Přihlásit se", use_container_width=True)
            
            if submit:
                if username in UZIVATELE and UZIVATELE[username] == password:
                    st.session_state["logged_in"] = True
                    st.session_state["user"] = username
                    
                    if remember_me:
                        # Vytvoříme token ve formátu: jmeno::hash
                        bezpecny_token = f"{username}::{vytvorit_bezpecny_token(username)}"
                        # Uložíme do cookies na 30 dní
                        cookie_manager.set(
                            "auth_token", 
                            bezpecny_token, 
                            expires_at=datetime.now() + timedelta(days=30)
                        )
                        time.sleep(0.5) # Krátké zdržení, aby se cookie stihla uložit do prohlížeče
                        
                    st.success("Přihlášení úspěšné!")
                    st.rerun()
                else:
                    st.error("❌ Nesprávné uživatelské jméno nebo heslo.")
        
    return False

# Zastaví vykonávání skriptu, pokud uživatel NENÍ přihlášen
if not overit_prihlaseni():
    st.stop()

# ==================== POSTRANNÍ PANEL (ODHLÁŠENÍ) ====================
with st.sidebar:
    st.write(f"👤 Přihlášen: **{st.session_state.get('user', '')}**")
    if st.button("Odhlásit se", use_container_width=True):
        # Smazání cookie a session
        cookie_manager.delete("auth_token")
        st.session_state["logged_in"] = False
        time.sleep(0.5) # Zdržení pro smazání z prohlížeče
        st.rerun()

# ==================== HLAVNÍ APLIKACE ====================
st.title("🚂 Prohlídky VZ a Radiostanic")

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

# --- 4. NAČTENÍ A INICIALIZACE SESSION STATE ---
if "kbs_df" not in st.session_state or "ls06_df" not in st.session_state or "radio_df" not in st.session_state:
    kbs_df, ls06_df, radio_df = nacist_data_z_excelu()

    st.session_state["kbs_df"] = kbs_df.rename(columns={
        "Unnamed: 0": "Číslo mašiny",
        "Datum provedení prohlídky a rozsah": "Datum provedení",
        "Unnamed: 2": "Provedena prohlídka",
        " Přístí prohlídka a rozsah": "Příští prohlídka",
        "Unnamed: 4": "Budoucí prohlídka",
        "Starý název 2": "Nový název 2"
    })

    st.session_state["ls06_df"] = ls06_df.rename(columns={
        "Unnamed: 0": "Číslo mašiny",
        "Datum provedení prohlídky a rozsah": "Datum prohlídky",
        "Unnamed: 2": "Rozsah",
        "Přístí prohlídka a rozsah": "Příští prohlídka",
        "Unnamed: 4": "Rozsah",
        "2026-09-06 00:00:00": "Typ",
        "AKTUÁLNÍ PROHLÍDKA": "Výrobní číslo",
    })

    st.session_state["radio_df"] = radio_df.rename(columns={
        "Starý název 1": "Nový název 1"
    })

kbs_df = st.session_state["kbs_df"]
ls06_df = st.session_state["ls06_df"]
radio_df = st.session_state["radio_df"]

# --- 5. VÝPOČET TERMÍNŮ A ZVÝRAZNĚNÍ ---
dnes = datetime.now()
zacatek_aktualniho_mesice = dnes.replace(day=1, hour=0, minute=0, second=0, microsecond=0)

aktualni_mesic = dnes.month
aktualni_rok = dnes.year

pristi_mesic_datum = (zacatek_aktualniho_mesice + pd.Timedelta(days=32)).replace(day=1)
pristi_mesic = pristi_mesic_datum.month
pristi_rok = pristi_mesic_datum.year

def ziskej_propadle_masiny(df, sloupec_data):
    if df.empty or sloupec_data not in df.columns:
        return []
    df_temp = df.copy()
    df_temp["_dt"] = pd.to_datetime(df_temp[sloupec_data], errors="coerce")
    filtrovane = df_temp[df_temp["_dt"] < zacatek_aktualniho_mesice]
    prvni_sloupec = df.columns[0]
    return filtrovane[prvni_sloupec].dropna().unique().tolist()

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

def zvyrazni_terminy(row, sloupec_data):
    try:
        dt = pd.to_datetime(row[sloupec_data])
        if pd.notnull(dt):
            if dt < zacatek_aktualniho_mesice:
                return ['background-color: #ffcdd2; color: #b71c1c; font-weight: bold'] * len(row)
            elif dt.month == aktualni_mesic and dt.year == aktualni_rok:
                return ['background-color: #ffe0b2; color: #e65100; font-weight: bold'] * len(row)
            elif dt.month == pristi_mesic and dt.year == pristi_rok:
                return ['background-color: #fff9c4; color: #f57f17; font-weight: bold'] * len(row)
    except:
        pass
    return [''] * len(row)

# --- 6. UI A TABULKY ---
tab_kbs, tab_ls06, tab_radio = st.tabs(["📋 KBS prohlídky", "📟 LS06", "📻 Radiostanice"])

# ==================== KBS List ====================
with tab_kbs:
    st.subheader("KBS Prohlídky")
    
    sloupce_s_datem_kbs = []
    mozne_nazvy_kbs = ["Datum provedení", "Datum prohlídky", "Příští prohlídka", "Další V1"]
    for col in kbs_df.columns:
        if col in mozne_nazvy_kbs or "datum" in col.lower():
            sloupce_s_datem_kbs.append(col)

    sloupec_provedeni = "Datum provedení" if "Datum provedení" in kbs_df.columns else ("Datum prohlídky" if "Datum prohlídky" in kbs_df.columns else (sloupce_s_datem_kbs[0] if sloupce_s_datem_kbs else None))
    sloupec_pristi = "Příští prohlídka" if "Příští prohlídka" in kbs_df.columns else (sloupce_s_datem_kbs[-1] if len(sloupce_s_datem_kbs) > 1 else None)

    kbs_config = {}
    for col in sloupce_s_datem_kbs:
        kbs_df[col] = pd.to_datetime(kbs_df[col], errors="coerce")
        kbs_config[col] = st.column_config.DateColumn(
            col,
            format="DD.MM.YYYY",
            step=1
        )

    # Vlastní šířky vybraných sloupců
    if "Číslo mašiny" in kbs_df.columns:
        kbs_config["Číslo mašiny"] = st.column_config.TextColumn("Číslo mašiny", width="small")
    if "Rozsah" in kbs_df.columns:
        kbs_config["Rozsah"] = st.column_config.TextColumn("Rozsah", width="small")
    #if "Číslo mašiny" in kbs_df.columns:
        #kbs_config["Číslo mašiny"] = st.column_config.TextColumn("Číslo mašiny", width="small")

    sloupec_pro_upozorneni = sloupec_pristi if sloupec_pristi else (sloupce_s_datem_kbs[-1] if sloupce_s_datem_kbs else None)
    
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

    edited_kbs = st.data_editor(
        styled_kbs, 
        use_container_width=True, 
        num_rows="dynamic", 
        hide_index=True, 
        column_config=kbs_config,
        key="kbs_editor"
    )

    if sloupec_provedeni and sloupec_pristi and sloupec_provedeni in edited_kbs.columns and sloupec_pristi in edited_kbs.columns:
        spocitane_pristi = pd.to_datetime(edited_kbs[sloupec_provedeni]).apply(
            lambda x: x + pd.DateOffset(months=3) if pd.notnull(x) else pd.NaT
        )
        
        if not edited_kbs[sloupec_pristi].equals(spocitane_pristi):
            edited_kbs[sloupec_pristi] = spocitane_pristi
            st.session_state["kbs_df"] = edited_kbs
            st.rerun()


# ==================== LS06 List ====================
with tab_ls06:
    st.subheader("LS06")
    
    sloupce_s_datem_ls06 = []
    mozne_nazvy_ls06 = ["Datum vykonání", "Datum prohlídky", "Příští datum", "Příští prohlídka", "Datum"]
    for col in ls06_df.columns:
        if col in mozne_nazvy_ls06 or "datum" in col.lower():
            sloupce_s_datem_ls06.append(col)

    sloupec_provedeni_ls = "Datum vykonání" if "Datum vykonání" in ls06_df.columns else ("Datum prohlídky" if "Datum prohlídky" in ls06_df.columns else (sloupce_s_datem_ls06[0] if sloupce_s_datem_ls06 else None))
    sloupec_pristi_ls = "Příští datum" if "Příští datum" in ls06_df.columns else ("Příští prohlídka" if "Příští prohlídka" in ls06_df.columns else (sloupce_s_datem_ls06[-1] if len(sloupce_s_datem_ls06) > 1 else None))

    ls06_config = {}
    for col in sloupce_s_datem_ls06:
        ls06_df[col] = pd.to_datetime(ls06_df[col], errors="coerce")
        ls06_config[col] = st.column_config.DateColumn(
            col,
            format="DD.MM.YYYY",
            step=1
        )

    sloupec_pro_upozorneni_ls = sloupec_pristi_ls if sloupec_pristi_ls else (sloupce_s_datem_ls06[-1] if sloupce_s_datem_ls06 else None)
    
    if sloupec_pro_upozorneni_ls:
        propadle = ziskej_propadle_masiny(ls06_df, sloupec_pro_upozorneni_ls)
        tento = ziskej_masiny_tento_mesic(ls06_df, sloupec_pro_upozorneni_ls)
        pristi = ziskej_masiny_pristi_mesic(ls06_df, sloupec_pro_upozorneni_ls)
        
        if propadle:
            st.error(f"🚨 **PROPADLÁ PROHLÍDKA:** {', '.join(propadle)}")
        if tento:
            st.warning(f"🟧 **PROHLÍDKA TENTO MĚSÍC ({aktualni_mesic}/{aktualni_rok}):** {', '.join(tento)}")
        if pristi:
            st.info(f"🟨 **Pozor na příští měsíc ({pristi_mesic}/{pristi_rok}):** {', '.join(pristi)}")
            
        styled_ls06 = ls06_df.style.apply(zvyrazni_terminy, sloupec_data=sloupec_pro_upozorneni_ls, axis=1)
    else:
        styled_ls06 = ls06_df

    edited_ls06 = st.data_editor(
        styled_ls06, 
        use_container_width=True, 
        num_rows="dynamic", 
        hide_index=True, 
        column_config=ls06_config,
        key="ls06_editor"
    )

    if sloupec_provedeni_ls and sloupec_pristi_ls and sloupec_provedeni_ls in edited_ls06.columns and sloupec_pristi_ls in edited_ls06.columns:
        spocitane_pristi = pd.to_datetime(edited_ls06[sloupec_provedeni_ls]).apply(
            lambda x: x + pd.DateOffset(months=12) if pd.notnull(x) else pd.NaT
        )
        
        if not edited_ls06[sloupec_pristi_ls].equals(spocitane_pristi):
            edited_ls06[sloupec_pristi_ls] = spocitane_pristi
            st.session_state["ls06_df"] = edited_ls06
            st.rerun()

# ==================== Radiostanice List ====================
with tab_radio:
    st.subheader("Radiostanice")
    
    sloupce_s_datem_radio = []
    mozne_nazvy_radio = ["Datum prohlídky", "Příští prohlídka", "Datum vykonání", "Příští datum", "Datum"]
    for col in radio_df.columns:
        if col in mozne_nazvy_radio or "datum" in col.lower():
            sloupce_s_datem_radio.append(col)

    # Určení sloupců pro provedenou a příští prohlídku
    sloupec_provedeni_rad = "Datum prohlídky" if "Datum prohlídky" in radio_df.columns else ("Datum vykonání" if "Datum vykonání" in radio_df.columns else (sloupce_s_datem_radio[0] if sloupce_s_datem_radio else None))
    sloupec_pristi_rad = "Příští prohlídka" if "Příští prohlídka" in radio_df.columns else ("Příští datum" if "Příští datum" in radio_df.columns else (sloupce_s_datem_radio[-1] if len(sloupce_s_datem_radio) > 1 else None))

    radio_config = {}
    for col in sloupce_s_datem_radio:
        radio_df[col] = pd.to_datetime(radio_df[col], errors="coerce")
        radio_config[col] = st.column_config.DateColumn(
            col,
            format="DD.MM.YYYY",
            step=1
        )

    sloupec_pro_upozorneni_rad = sloupec_pristi_rad if sloupec_pristi_rad else (sloupce_s_datem_radio[-1] if sloupce_s_datem_radio else None)
    
    if sloupec_pro_upozorneni_rad:
        propadle = ziskej_propadle_masiny(radio_df, sloupec_pro_upozorneni_rad)
        tento = ziskej_masiny_tento_mesic(radio_df, sloupec_pro_upozorneni_rad)
        pristi = ziskej_masiny_pristi_mesic(radio_df, sloupec_pro_upozorneni_rad)
        
        if propadle:
            st.error(f"🚨 **PROPADLÁ PROHLÍDKA:** {', '.join(propadle)}")
        if tento:
            st.warning(f"🟧 **PROHLÍDKA TENTO MĚSÍC ({aktualni_mesic}/{aktualni_rok}):** {', '.join(tento)}")
        if pristi:
            st.info(f"🟨 **Pozor na příští měsíc ({pristi_mesic}/{pristi_rok}):** {', '.join(pristi)}")
            
        styled_radio = radio_df.style.apply(zvyrazni_terminy, sloupec_data=sloupec_pro_upozorneni_rad, axis=1)
    else:
        styled_radio = radio_df

    edited_radio = st.data_editor(
        styled_radio, 
        use_container_width=True, 
        num_rows="dynamic", 
        hide_index=True, 
        column_config=radio_config,
        key="radio_editor"
    )

    # --- AUTOMATICKÝ VÝPOČET PŘÍŠTÍ PROHLÍDKY (CO 4 ROKY / 48 MĚSÍCŮ) ---
    if sloupec_provedeni_rad and sloupec_pristi_rad and sloupec_provedeni_rad in edited_radio.columns and sloupec_pristi_rad in edited_radio.columns:
        spocitane_pristi = pd.to_datetime(edited_radio[sloupec_provedeni_rad]).apply(
            lambda x: x + pd.DateOffset(years=4) if pd.notnull(x) else pd.NaT
        )
        
        if not edited_radio[sloupec_pristi_rad].equals(spocitane_pristi):
            edited_radio[sloupec_pristi_rad] = spocitane_pristi
            st.session_state["radio_df"] = edited_radio
            st.rerun()


# --- 7. GLOBÁLNÍ TLAČÍTKO ULOŽIT ---
st.divider()
if st.button("💾 Uložit všechny změny do Excelu na GitHub", type="primary", use_container_width=True):
    ulozit_vse_do_excelu(edited_kbs, edited_ls06, edited_radio)
