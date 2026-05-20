import hashlib
import hmac
from datetime import date
from pathlib import Path

import gspread
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2.service_account import Credentials

# =========================================================
# CONFIG
# =========================================================

st.set_page_config(
    page_title="SIGE-CTAR",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

LOGO_PATH = "assets/LOGO.jpg"

REQUIRED_SHEETS = [
    "FACT_CTAR_SEGUIMIENTO",
    "DIM_EQUIPO",
    "DIM_SERVICIO",
    "DIM_TIPO_PROCESO",
    "DIM_ESTADO",
    "DIM_RESPONSABLE",
    "DIM_PRIORIDAD",
    "FACT_BAJAS",
    "FACT_REPOSICIONES",
    "FACT_ADQUISICIONES",
    "FACT_ALERTAS",
]

# =========================================================
# ESTILO
# =========================================================

st.markdown("""
<style>

.stApp {
    background: #f4f6fb;
}

.block-container {
    padding-top: 1rem;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1a2b4a 0%, #0f1d33 100%);
}

section[data-testid="stSidebar"] * {
    color: #eaf2ff !important;
}

.main-header {
    background: linear-gradient(90deg, #1a2b4a 0%, #2563eb 100%);
    padding: 24px;
    border-radius: 18px;
    color: white;
    margin-bottom: 20px;
    box-shadow: 0 8px 22px rgba(15,23,42,.15);
}

.main-header h1 {
    margin: 0;
    color: white;
}

.main-header p {
    margin-top: 5px;
    color: #dbeafe;
}

.metric-card {
    background: white;
    border-radius: 16px;
    padding: 18px;
    border: 1px solid #e5e7eb;
    box-shadow: 0 4px 12px rgba(15,23,42,.06);
}

.metric-title {
    font-size: 12px;
    color: #64748b;
    text-transform: uppercase;
    font-weight: 700;
}

.metric-value {
    font-size: 34px;
    font-weight: 900;
    color: #1f2937;
    margin-top: 8px;
}

.login-box {
    background: white;
    padding: 24px;
    border-radius: 18px;
    border: 1px solid #e5e7eb;
    box-shadow: 0 8px 20px rgba(15,23,42,.08);
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# AUTH
# =========================================================

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


def get_users():
    return {
        "admin": {
            "name": "Administrador CTAR",
            "password_hash": hash_password("admin123"),
            "role": "admin",
        },
        "ctar": {
            "name": "Usuario CTAR",
            "password_hash": hash_password("ctar123"),
            "role": "ctar",
        },
        "hospital": {
            "name": "Usuario Hospital",
            "password_hash": hash_password("hospital123"),
            "role": "hospital",
        },
        "ifiscal": {
            "name": "Inspector Fiscal",
            "password_hash": hash_password("if123"),
            "role": "if",
        },
    }


def authenticate(username, password):
    users = get_users()

    if username not in users:
        return None

    user = users[username]

    if hmac.compare_digest(
        user["password_hash"],
        hash_password(password)
    ):
        return user

    return None


def login():

    col1, col2 = st.columns([1, 4])

    with col1:
        try:
            st.image(LOGO_PATH, width=160)
        except:
            st.markdown("## SIGE-CTAR")

    with col2:
        st.markdown("# SIGE-CTAR")
        st.caption("Sistema de Gestión y Trazabilidad CTAR")

    st.markdown("""
    <div class="login-box">
    Plataforma institucional para gestión de:
    <ul>
        <li>Bajas</li>
        <li>Reposiciones</li>
        <li>Adquisiciones</li>
        <li>Seguimiento SIC</li>
        <li>Gestión CTAR</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)

    with st.form("login"):

        username = st.text_input("Usuario")
        password = st.text_input("Clave", type="password")

        submit = st.form_submit_button("Ingresar")

    if submit:

        user = authenticate(username.strip(), password)

        if user:
            st.session_state["user"] = user
            st.rerun()

        else:
            st.error("Usuario o clave incorrecta")


if "user" not in st.session_state:
    login()
    st.stop()

# =========================================================
# ROLES
# =========================================================

def role():
    return st.session_state["user"]["role"]


def is_hospital():
    return role() == "hospital"


def can_edit():
    return role() in ["admin", "ctar", "if"]

# =========================================================
# GOOGLE SHEETS
# =========================================================

@st.cache_data(ttl=30)
def load_google_sheets():

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )

    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(
        st.secrets["google_sheet"]["spreadsheet_id"]
    )

    tables = {}

    for ws in spreadsheet.worksheets():

        data = ws.get_all_records()

        df = pd.DataFrame(data)

        df.columns = [
            str(c).strip().replace(" ", "_")
            for c in df.columns
        ]

        tables[ws.title] = df

    return tables


try:
    tables = load_google_sheets()

except Exception as e:

    st.error("No se pudo conectar con Google Sheets")
    st.code(str(e))
    st.stop()

# =========================================================
# MODELO
# =========================================================

df = tables.get("FACT_CTAR_SEGUIMIENTO", pd.DataFrame())

if not df.empty:

    for col in ["Fecha_Ingreso", "Fecha_Compromiso"]:

        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    today = pd.Timestamp(date.today())

    if "Fecha_Compromiso" in df.columns:

        df["Dias_Atraso"] = (
            today - df["Fecha_Compromiso"]
        ).dt.days

        df["Vencido"] = df["Dias_Atraso"] > 0

# =========================================================
# HELPERS
# =========================================================

def header(title, subtitle):

    st.markdown(f"""
    <div class="main-header">
        <h1>{title}</h1>
        <p>{subtitle}</p>
    </div>
    """, unsafe_allow_html=True)


def metric_card(title, value):

    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">{title}</div>
        <div class="metric-value">{value}</div>
    </div>
    """, unsafe_allow_html=True)

# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    try:
        st.image(LOGO_PATH, use_container_width=True)
    except:
        st.markdown("## SIGE-CTAR")

    st.markdown("## SIGE-CTAR")
    st.caption("Sistema CTAR")

    st.markdown("---")

    st.caption(
        f"Usuario: {st.session_state['user']['name']}"
    )

    st.caption(
        f"Rol: {st.session_state['user']['role']}"
    )

    if is_hospital():

        pages = [
            "Resumen Ejecutivo",
            "Seguimiento",
            "Bajas",
            "Reposiciones",
            "Adquisiciones",
        ]

    else:

        pages = [
            "Resumen Ejecutivo",
            "Seguimiento",
            "Bajas",
            "Reposiciones",
            "Adquisiciones",
            "Alertas",
        ]

    page = st.radio("Menú", pages)

    st.markdown("---")

    if st.button("Cerrar sesión"):

        st.session_state.clear()
        st.rerun()

# =========================================================
# RESUMEN
# =========================================================

if page == "Resumen Ejecutivo":

    header(
        "SIGE-CTAR · Resumen Ejecutivo",
        "Seguimiento operacional CTAR"
    )

    if df.empty:

        st.info("No hay datos cargados")

        st.stop()

    total = len(df)

    revision = df["Estado"].astype(str).str.lower().str.contains(
        "revision|revisión",
        na=False
    ).sum()

    aprobadas = df["Estado"].astype(str).str.lower().str.contains(
        "aprob",
        na=False
    ).sum()

    compra = df["Estado"].astype(str).str.lower().str.contains(
        "compra",
        na=False
    ).sum()

    altas = df["Prioridad"].astype(str).str.lower().eq(
        "alta"
    ).sum()

    vencidas = int(df["Vencido"].sum())

    if is_hospital():

        c1, c2, c3, c4, c5 = st.columns(5)

        with c1:
            metric_card("Total", total)

        with c2:
            metric_card("En revisión", revision)

        with c3:
            metric_card("Aprobadas", aprobadas)

        with c4:
            metric_card("En compra", compra)

        with c5:
            metric_card("Prioridad alta", altas)

    else:

        c1, c2, c3, c4, c5, c6 = st.columns(6)

        with c1:
            metric_card("Total", total)

        with c2:
            metric_card("En revisión", revision)

        with c3:
            metric_card("Aprobadas", aprobadas)

        with c4:
            metric_card("En compra", compra)

        with c5:
            metric_card("Prioridad alta", altas)

        with c6:
            metric_card("Vencidas", vencidas)

    st.markdown("---")

    g1, g2, g3 = st.columns(3)

    with g1:

        st.markdown("### Estados")

        estado_df = (
            df.groupby("Estado")
            .size()
            .reset_index(name="Cantidad")
        )

        fig = px.bar(
            estado_df,
            x="Estado",
            y="Cantidad",
            text="Cantidad"
        )

        st.plotly_chart(fig, use_container_width=True)

    with g2:

        st.markdown("### Tipo Proceso")

        fig = px.pie(
            df,
            names="Tipo_Proceso",
            hole=0.45
        )

        st.plotly_chart(fig, use_container_width=True)

    with g3:

        st.markdown("### Prioridad")

        fig = px.pie(
            df,
            names="Prioridad",
            hole=0.45
        )

        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Seguimiento principal")

    cols = [
        "ID_CTAR",
        "SIC",
        "Equipo",
        "Servicio",
        "Tipo_Proceso",
        "Estado",
        "Responsable",
        "Prioridad",
        "Fecha_Compromiso",
        "Proxima_Accion",
    ]

    st.dataframe(
        df[[c for c in cols if c in df.columns]],
        use_container_width=True,
        hide_index=True
    )

# =========================================================
# SEGUIMIENTO
# =========================================================

elif page == "Seguimiento":

    header(
        "Seguimiento CTAR",
        "Consulta de solicitudes"
    )

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

# =========================================================
# BAJAS
# =========================================================

elif page == "Bajas":

    header(
        "Bajas y Extravíos",
        "Control de bajas y extravíos"
    )

    view = df[
        df["Tipo_Proceso"]
        .astype(str)
        .str.lower()
        .str.contains("baja|extrav", na=False)
    ]

    st.dataframe(view, use_container_width=True)

# =========================================================
# REPOSICIONES
# =========================================================

elif page == "Reposiciones":

    header(
        "Reposiciones",
        "Seguimiento reposiciones"
    )

    view = df[
        df["Tipo_Proceso"]
        .astype(str)
        .str.lower()
        .str.contains("repos", na=False)
    ]

    st.dataframe(view, use_container_width=True)

# =========================================================
# ADQUISICIONES
# =========================================================

elif page == "Adquisiciones":

    header(
        "Adquisiciones",
        "Seguimiento compras"
    )

    view = df[
        df["Tipo_Proceso"]
        .astype(str)
        .str.lower()
        .str.contains("adquis|compra", na=False)
    ]

    st.dataframe(view, use_container_width=True)

# =========================================================
# ALERTAS
# =========================================================

elif page == "Alertas":

    header(
        "Alertas",
        "Procesos críticos"
    )

    view = df[
        (df["Prioridad"]
        .astype(str)
        .str.lower() == "alta")
        |
        (df["Vencido"] == True)
    ]

    st.dataframe(view, use_container_width=True)
