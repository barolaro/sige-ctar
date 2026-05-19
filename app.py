
import io
import os
import re
from datetime import date, datetime
from typing import Dict, Optional, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None


# =========================================================
# SIGE-CTAR
# Sistema de Gestión y Trazabilidad CTAR
# Hospital / CTAR / Inspector Fiscal / Concesionaria
# =========================================================

st.set_page_config(
    page_title="SIGE-CTAR",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_NAME = "SIGE-CTAR"
DEFAULT_EXCEL = "data/CTAR_RelationalModel.xlsx"


# =========================================================
# AUTENTICACIÓN SIMPLE POR USUARIO / CLAVE
# =========================================================

import hashlib
import hmac

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def get_users_config():
    try:
        return dict(st.secrets["auth"]["users"])
    except Exception:
        # Usuarios demo. Cambiar obligatoriamente antes de usar con Hospital.
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

def authenticate(username: str, password: str):
    users = get_users_config()
    if username not in users:
        return None

    user = users[username]
    expected_hash = str(user.get("password_hash", ""))
    provided_hash = hash_password(password)

    if hmac.compare_digest(expected_hash, provided_hash):
        return {
            "username": username,
            "name": user.get("name", username),
            "role": user.get("role", "viewer"),
        }
    return None

def login_screen():
    st.markdown(
        """
        <style>
        .login-box {
            max-width: 460px;
            margin: 8vh auto 0 auto;
            background: white;
            padding: 34px;
            border-radius: 18px;
            border: 1px solid #e5e7eb;
            box-shadow: 0 12px 34px rgba(15, 23, 42, .12);
        }
        .login-title {
            font-size: 28px;
            font-weight: 800;
            color: #1a2b4a;
            margin-bottom: 4px;
        }
        .login-subtitle {
            color: #64748b;
            font-size: 14px;
            margin-bottom: 22px;
        }
        </style>
        <div class="login-box">
            <div class="login-title">🏥 SIGE-CTAR</div>
            <div class="login-subtitle">Sistema de Gestión y Trazabilidad CTAR</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.form("login_form"):
        username = st.text_input("Usuario")
        password = st.text_input("Clave", type="password")
        submit = st.form_submit_button("Ingresar")

    if submit:
        user = authenticate(username.strip(), password)
        if user:
            st.session_state["authenticated"] = True
            st.session_state["user"] = user
            st.rerun()
        else:
            st.error("Usuario o clave incorrecta.")

def require_login():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if not st.session_state["authenticated"]:
        login_screen()
        st.stop()

def logout_button():
    user = st.session_state.get("user", {})
    st.sidebar.markdown("---")
    st.sidebar.caption(f"Usuario: {user.get('name','')}")
    st.sidebar.caption(f"Rol: {user.get('role','')}")
    if st.sidebar.button("Cerrar sesión"):
        st.session_state.clear()
        st.rerun()

def user_role():
    return st.session_state.get("user", {}).get("role", "viewer")

def can_edit():
    return user_role() in ["admin", "ctar", "if"]

def can_config():
    return user_role() == "admin"

def filter_pages_by_role(pages):
    role = user_role()
    if role == "hospital":
        return [
            "Resumen Ejecutivo",
            "Seguimiento",
            "Bajas",
            "Reposiciones",
            "Adquisiciones",
            "Alertas",
        ]
    if role in ["if", "ctar"]:
        return [
            "Resumen Ejecutivo",
            "Seguimiento",
            "Bajas",
            "Reposiciones",
            "Adquisiciones",
            "Alertas",
            "Registro",
        ]
    return pages


REQUIRED_TABLES = [
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
# CSS
# =========================================================

st.markdown(
    """
    <style>
    :root {
        --ctar-blue: #1a2b4a;
        --ctar-blue2: #2563eb;
        --ctar-bg: #f4f6fb;
        --ctar-card: #ffffff;
        --ctar-text: #1f2937;
        --ctar-muted: #64748b;
        --ctar-red: #dc2626;
        --ctar-orange: #f59e0b;
        --ctar-green: #16a34a;
        --ctar-border: #e5e7eb;
    }
    .main {
        background: var(--ctar-bg);
    }
    h1, h2, h3 {
        color: var(--ctar-text);
    }
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 3rem;
    }
    .ctar-header {
        background: linear-gradient(90deg, #1a2b4a, #2563eb);
        border-radius: 18px;
        padding: 22px 26px;
        color: white;
        margin-bottom: 18px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, .14);
    }
    .ctar-header h1 {
        color: white;
        font-size: 28px;
        margin: 0;
        padding: 0;
    }
    .ctar-header p {
        color: #dbeafe;
        margin: 4px 0 0 0;
        font-size: 14px;
    }
    .metric-card {
        background: white;
        border-radius: 16px;
        padding: 16px 18px;
        border: 1px solid var(--ctar-border);
        box-shadow: 0 4px 14px rgba(15, 23, 42, .06);
    }
    .metric-title {
        font-size: 12px;
        text-transform: uppercase;
        color: var(--ctar-muted);
        letter-spacing: .04em;
        font-weight: 700;
    }
    .metric-value {
        font-size: 30px;
        color: var(--ctar-text);
        font-weight: 800;
        line-height: 1.2;
        margin-top: 6px;
    }
    .metric-sub {
        font-size: 12px;
        color: var(--ctar-muted);
        margin-top: 4px;
    }
    .section-card {
        background: white;
        border-radius: 16px;
        padding: 18px;
        border: 1px solid var(--ctar-border);
        box-shadow: 0 4px 14px rgba(15, 23, 42, .05);
        margin-bottom: 14px;
    }
    .badge {
        display: inline-block;
        border-radius: 999px;
        padding: 4px 10px;
        font-size: 12px;
        font-weight: 700;
    }
    .badge-red { background: #fee2e2; color: #991b1b; }
    .badge-orange { background: #fef3c7; color: #92400e; }
    .badge-green { background: #dcfce7; color: #166534; }
    .badge-blue { background: #dbeafe; color: #1e40af; }
    .small-muted {
        color: var(--ctar-muted);
        font-size: 12px;
    }
    div[data-testid="stSidebar"] {
        background: #1a2b4a;
    }
    div[data-testid="stSidebar"] * {
        color: #e5edff;
    }
    div[data-testid="stSidebar"] .stRadio label {
        color: #e5edff !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# Helpers
# =========================================================

def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().replace(" ", "_") for c in df.columns]
    return df


def to_date(series):
    return pd.to_datetime(series, errors="coerce", dayfirst=False)


def normalize_text(s):
    if pd.isna(s):
        return ""
    return str(s).strip()


def safe_count(df: pd.DataFrame) -> int:
    if df is None or df.empty:
        return 0
    return len(df)


def status_color(estado: str) -> str:
    e = normalize_text(estado).lower()
    if "cerr" in e or "instal" in e:
        return "green"
    if "observ" in e or "atras" in e:
        return "red"
    if "compra" in e or "revisión" in e or "revision" in e:
        return "orange"
    return "blue"


def priority_color(prioridad: str) -> str:
    p = normalize_text(prioridad).lower()
    if p == "alta":
        return "red"
    if p == "media":
        return "orange"
    return "green"


def make_badge(text: str, color: str) -> str:
    return f'<span class="badge badge-{color}">{text}</span>'


def validate_tables(tables: Dict[str, pd.DataFrame]) -> Tuple[bool, list]:
    missing = [t for t in REQUIRED_TABLES if t not in tables]
    return len(missing) == 0, missing


# =========================================================
# Loaders
# =========================================================

@st.cache_data(ttl=60)
def load_from_excel_bytes(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    tables = {}
    for sheet in xls.sheet_names:
        tables[sheet] = clean_columns(pd.read_excel(xls, sheet_name=sheet))
    return tables


@st.cache_data(ttl=60)
def load_from_excel_path(path: str) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(path)
    tables = {}
    for sheet in xls.sheet_names:
        tables[sheet] = clean_columns(pd.read_excel(xls, sheet_name=sheet))
    return tables


def load_from_gspread(spreadsheet_id: str) -> Dict[str, pd.DataFrame]:
    """
    Lee Google Sheets con credenciales de service account.
    En Streamlit Cloud, agregar secrets:
    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
    client_email = "..."
    client_id = "..."
    token_uri = "https://oauth2.googleapis.com/token"

    [google_sheet]
    spreadsheet_id = "..."
    """
    if gspread is None or Credentials is None:
        st.error("Falta instalar gspread y google-auth. Revise requirements.txt.")
        return {}

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(spreadsheet_id)

    tables = {}
    for ws in spreadsheet.worksheets():
        records = ws.get_all_records()
        tables[ws.title] = clean_columns(pd.DataFrame(records))
    return tables


def update_gsheet_row(spreadsheet_id: str, worksheet_name: str, row_values: list):
    """
    Agrega una fila al Google Sheet.
    Requiere secrets de service account.
    """
    if gspread is None or Credentials is None:
        raise RuntimeError("gspread no está instalado.")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    client = gspread.authorize(creds)
    ws = client.open_by_key(spreadsheet_id).worksheet(worksheet_name)
    ws.append_row(row_values, value_input_option="USER_ENTERED")


# =========================================================
# Data Model
# =========================================================

def build_model(tables: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    """
    Construye una tabla enriquecida para visualizaciones.
    Mantiene el modelo relacional y crea vista de trabajo.
    """
    fact = tables.get("FACT_CTAR_SEGUIMIENTO", pd.DataFrame()).copy()

    if fact.empty:
        return {"seguimiento": fact, **tables}

    for col in ["Fecha_Ingreso", "Fecha_Compromiso"]:
        if col in fact.columns:
            fact[col] = to_date(fact[col])

    joins = [
        ("DIM_EQUIPO", "ID_Equipo", ["Equipo", "Nro_Inventario", "Familia"]),
        ("DIM_SERVICIO", "ID_Servicio", ["Servicio", "Area"]),
        ("DIM_TIPO_PROCESO", "ID_Tipo_Proceso", ["Tipo_Proceso"]),
        ("DIM_ESTADO", "ID_Estado", ["Estado", "Orden"]),
        ("DIM_RESPONSABLE", "ID_Responsable", ["Responsable"]),
        ("DIM_PRIORIDAD", "ID_Prioridad", ["Prioridad"]),
    ]

    df = fact.copy()
    for table_name, key, keep_cols in joins:
        dim = tables.get(table_name, pd.DataFrame()).copy()
        if not dim.empty and key in df.columns and key in dim.columns:
            cols = [key] + [c for c in keep_cols if c in dim.columns]
            df = df.merge(dim[cols], on=key, how="left")

    today = pd.Timestamp(date.today())
    if "Fecha_Ingreso" in df.columns:
        df["Dias_Desde_Ingreso"] = (today - df["Fecha_Ingreso"]).dt.days
    else:
        df["Dias_Desde_Ingreso"] = 0

    if "Fecha_Compromiso" in df.columns:
        df["Dias_Atraso"] = (today - df["Fecha_Compromiso"]).dt.days
        df["Vencido"] = (df["Dias_Atraso"] > 0) & (~df.get("Estado", "").astype(str).str.lower().str.contains("cerr|instal", regex=True, na=False))
    else:
        df["Dias_Atraso"] = 0
        df["Vencido"] = False

    if "Prioridad" not in df.columns:
        df["Prioridad"] = ""
    if "Estado" not in df.columns:
        df["Estado"] = ""
    if "Tipo_Proceso" not in df.columns:
        df["Tipo_Proceso"] = ""
    if "Responsable" not in df.columns:
        df["Responsable"] = ""
    if "Servicio" not in df.columns:
        df["Servicio"] = ""
    if "Equipo" not in df.columns:
        df["Equipo"] = ""
    if "Etapa" not in df.columns:
        df["Etapa"] = ""

    tables["seguimiento"] = df
    return tables


# =========================================================
# Login requerido
# =========================================================

require_login()

# =========================================================
# Sidebar
# =========================================================

with st.sidebar:
    st.markdown("## 🏥 SIGE-CTAR")
    st.caption("Sistema de Gestión y Trazabilidad CTAR")
    st.markdown("---")

    available_pages = filter_pages_by_role([
        "Resumen Ejecutivo",
        "Seguimiento",
        "Bajas",
        "Reposiciones",
        "Adquisiciones",
        "Alertas",
        "Registro",
        "Configuración",
    ])

    page = st.radio("Menú", available_pages)

    st.markdown("---")
    data_source = st.selectbox(
        "Fuente de datos",
        ["Excel local / cargado", "Google Sheets"],
    )

    uploaded_file = None
    spreadsheet_id = None

    if data_source == "Excel local / cargado":
        uploaded_file = st.file_uploader("Cargar Excel CTAR", type=["xlsx"])
        st.caption("Si no carga archivo, se usará el Excel de ejemplo incluido.")
    else:
        default_id = ""
        try:
            default_id = st.secrets["google_sheet"]["spreadsheet_id"]
        except Exception:
            default_id = ""
        spreadsheet_id = st.text_input("Google Sheet ID", value=default_id)
        st.caption("Debe compartir el Sheet con el correo del service account.")

    st.markdown("---")
    st.caption("Versión piloto robusto · CTAR")
    logout_button()


# =========================================================
# Load data
# =========================================================

tables = {}
try:
    if data_source == "Google Sheets" and spreadsheet_id:
        tables = load_from_gspread(spreadsheet_id)
    else:
        if uploaded_file is not None:
            tables = load_from_excel_bytes(uploaded_file.read())
        elif os.path.exists(DEFAULT_EXCEL):
            tables = load_from_excel_path(DEFAULT_EXCEL)
        else:
            st.warning("No se encontró archivo de datos. Cargue un Excel CTAR.")
            tables = {}
except Exception as e:
    st.error(f"No fue posible cargar datos: {e}")
    tables = {}

ok, missing = validate_tables(tables) if tables else (False, REQUIRED_TABLES)
if not ok and page != "Configuración":
    st.warning("Faltan algunas tablas del modelo. Revise Configuración.")
    if missing:
        st.caption("Tablas faltantes: " + ", ".join(missing))

model = build_model(tables) if tables else {}
df = model.get("seguimiento", pd.DataFrame())


# =========================================================
# Header
# =========================================================

def header(title: str, subtitle: str):
    st.markdown(
        f"""
        <div class="ctar-header">
            <h1>{title}</h1>
            <p>{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# Global Filters
# =========================================================

def filtered_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    st.markdown("### Filtros")
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        servicios = ["Todos"] + sorted([x for x in df["Servicio"].dropna().unique().tolist() if str(x).strip()])
        f_servicio = st.selectbox("Servicio", servicios)

    with c2:
        tipos = ["Todos"] + sorted([x for x in df["Tipo_Proceso"].dropna().unique().tolist() if str(x).strip()])
        f_tipo = st.selectbox("Tipo proceso", tipos)

    with c3:
        estados = ["Todos"] + sorted([x for x in df["Estado"].dropna().unique().tolist() if str(x).strip()])
        f_estado = st.selectbox("Estado", estados)

    with c4:
        responsables = ["Todos"] + sorted([x for x in df["Responsable"].dropna().unique().tolist() if str(x).strip()])
        f_resp = st.selectbox("Responsable", responsables)

    out = df.copy()
    if f_servicio != "Todos":
        out = out[out["Servicio"] == f_servicio]
    if f_tipo != "Todos":
        out = out[out["Tipo_Proceso"] == f_tipo]
    if f_estado != "Todos":
        out = out[out["Estado"] == f_estado]
    if f_resp != "Todos":
        out = out[out["Responsable"] == f_resp]
    return out


# =========================================================
# KPI Cards
# =========================================================

def kpi_card(title, value, sub="", color="blue"):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">{title}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# Pages
# =========================================================

if page == "Resumen Ejecutivo":
    header(
        "SIGE-CTAR · Resumen Ejecutivo",
        "Control de bajas, reposiciones, adquisiciones, SIC, responsables y alertas.",
    )

    if df.empty:
        st.info("Cargue datos para visualizar el tablero.")
    else:
        view = filtered_df(df)

        total = len(view)
        en_revision = view["Estado"].astype(str).str.lower().str.contains("revisión|revision", regex=True, na=False).sum()
        aprobadas = view["Estado"].astype(str).str.lower().str.contains("aprob", regex=True, na=False).sum()
        en_compra = view["Estado"].astype(str).str.lower().str.contains("compra", regex=True, na=False).sum()
        vencidas = int(view["Vencido"].sum()) if "Vencido" in view.columns else 0
        criticas = view["Prioridad"].astype(str).str.lower().eq("alta").sum()

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1: kpi_card("Total solicitudes", total, "Registros CTAR")
        with c2: kpi_card("En revisión", en_revision, "Etapa CTAR / IF")
        with c3: kpi_card("Aprobadas", aprobadas, "Con acuerdo")
        with c4: kpi_card("En compra", en_compra, "BACO / OC / Proveedor")
        with c5: kpi_card("Vencidas", vencidas, "Fecha compromiso vencida")
        with c6: kpi_card("Prioridad alta", criticas, "Riesgo operacional")

        st.markdown("---")
        g1, g2, g3 = st.columns([1.2, 1, 1])

        with g1:
            st.markdown("#### Flujo por estado")
            if "Estado" in view.columns:
                estado_counts = view.groupby(["Estado"]).size().reset_index(name="Cantidad")
                if "Orden" in view.columns:
                    orden = view.groupby("Estado")["Orden"].min().reset_index()
                    estado_counts = estado_counts.merge(orden, on="Estado", how="left").sort_values("Orden")
                fig = px.bar(
                    estado_counts,
                    x="Estado",
                    y="Cantidad",
                    text="Cantidad",
                    title=None,
                )
                fig.update_layout(height=330, margin=dict(l=10, r=10, t=20, b=10))
                st.plotly_chart(fig, use_container_width=True)

        with g2:
            st.markdown("#### Tipo de proceso")
            if "Tipo_Proceso" in view.columns:
                fig = px.pie(view, names="Tipo_Proceso", hole=0.45)
                fig.update_layout(height=330, margin=dict(l=10, r=10, t=20, b=10))
                st.plotly_chart(fig, use_container_width=True)

        with g3:
            st.markdown("#### Prioridad")
            if "Prioridad" in view.columns:
                fig = px.pie(view, names="Prioridad", hole=0.45)
                fig.update_layout(height=330, margin=dict(l=10, r=10, t=20, b=10))
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Alertas principales")
        alert_cols = [
            "ID_CTAR", "SIC", "Equipo", "Servicio", "Tipo_Proceso", "Estado",
            "Responsable", "Prioridad", "Dias_Desde_Ingreso", "Dias_Atraso",
            "Riesgo_Clinico", "Proxima_Accion"
        ]
        alert_view = view.copy()
        if "Vencido" in alert_view.columns:
            alert_view = alert_view[(alert_view["Vencido"]) | (alert_view["Prioridad"].astype(str).str.lower() == "alta")]
        st.dataframe(alert_view[[c for c in alert_cols if c in alert_view.columns]], use_container_width=True, hide_index=True)


elif page == "Seguimiento":
    header(
        "Seguimiento Operacional CTAR",
        "Consulta rápida por SIC, equipo, inventario, servicio, estado y responsable.",
    )

    if df.empty:
        st.info("Cargue datos para visualizar seguimiento.")
    else:
        view = filtered_df(df)

        q = st.text_input("Buscar por SIC, equipo, inventario, servicio o motivo")
        if q:
            ql = q.lower()
            mask = pd.Series(False, index=view.index)
            for col in ["SIC", "Equipo", "Nro_Inventario", "Servicio", "Motivo", "Riesgo_Clinico"]:
                if col in view.columns:
                    mask = mask | view[col].astype(str).str.lower().str.contains(re.escape(ql), na=False)
            view = view[mask]

        cols = [
            "ID_CTAR", "SIC", "Equipo", "Nro_Inventario", "Servicio", "Tipo_Proceso",
            "Estado", "Etapa", "Responsable", "Prioridad", "Fecha_Ingreso",
            "Fecha_Compromiso", "Dias_Desde_Ingreso", "Dias_Atraso",
            "Motivo", "Riesgo_Clinico", "Ultima_Gestion", "Proxima_Accion", "Link_Documento"
        ]
        st.dataframe(view[[c for c in cols if c in view.columns]], use_container_width=True, hide_index=True)

        st.download_button(
            "Descargar seguimiento filtrado CSV",
            data=view.to_csv(index=False).encode("utf-8-sig"),
            file_name="seguimiento_ctar_filtrado.csv",
            mime="text/csv",
        )


elif page == "Bajas":
    header("Bajas y Extravíos", "Gestión específica de bajas solicitadas, aprobadas, observadas y pendientes.")

    if df.empty:
        st.info("Cargue datos para visualizar bajas.")
    else:
        view = df[df["Tipo_Proceso"].astype(str).str.lower().str.contains("baja|extrav", regex=True, na=False)]
        view = filtered_df(view) if not view.empty else view

        c1, c2, c3 = st.columns(3)
        with c1: kpi_card("Total bajas/extravíos", len(view), "Solicitudes")
        with c2: kpi_card("Prioridad alta", (view["Prioridad"].astype(str).str.lower() == "alta").sum(), "Críticas")
        with c3: kpi_card("Vencidas", int(view["Vencido"].sum()) if "Vencido" in view else 0, "Con atraso")

        st.markdown("#### Detalle")
        st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Reposiciones":
    header("Reposiciones", "Seguimiento de reposición desde evaluación CTAR hasta instalación o cierre.")

    if df.empty:
        st.info("Cargue datos para visualizar reposiciones.")
    else:
        view = df[df["Tipo_Proceso"].astype(str).str.lower().str.contains("repos", regex=True, na=False)]
        view = filtered_df(view) if not view.empty else view

        c1, c2, c3 = st.columns(3)
        with c1: kpi_card("Reposiciones", len(view), "Total")
        with c2: kpi_card("En compra", view["Estado"].astype(str).str.lower().str.contains("compra", na=False).sum(), "Proceso compra")
        with c3: kpi_card("Pendientes", view["Estado"].astype(str).str.lower().str.contains("revisión|revision|observ", regex=True, na=False).sum(), "Por resolver")

        st.markdown("#### Estado de reposiciones")
        if not view.empty:
            fig = px.bar(view.groupby("Estado").size().reset_index(name="Cantidad"), x="Estado", y="Cantidad", text="Cantidad")
            fig.update_layout(height=320)
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Adquisiciones":
    header("Adquisiciones", "Control de BACO, OC, proveedor, compra, recepción e instalación.")

    adq = model.get("FACT_ADQUISICIONES", pd.DataFrame())
    if df.empty:
        st.info("Cargue datos para visualizar adquisiciones.")
    else:
        view = df[df["Tipo_Proceso"].astype(str).str.lower().str.contains("adquis|compra", regex=True, na=False)]
        view = filtered_df(view) if not view.empty else view

        if not adq.empty and "ID_CTAR" in adq.columns:
            view = view.merge(adq, on="ID_CTAR", how="left", suffixes=("", "_Adq"))

        c1, c2, c3 = st.columns(3)
        with c1: kpi_card("Adquisiciones", len(view), "Total")
        with c2: kpi_card("En proceso", view["Estado"].astype(str).str.lower().str.contains("compra|curso", regex=True, na=False).sum(), "Compra")
        with c3: kpi_card("Con proveedor", view.get("Proveedor", pd.Series(dtype=str)).notna().sum() if "Proveedor" in view else 0, "Registradas")

        st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Alertas":
    header("Alertas y Observaciones", "Procesos críticos, atrasos, SIC cerrados sin respuesta y observaciones pendientes.")

    if df.empty:
        st.info("Cargue datos para visualizar alertas.")
    else:
        alertas = model.get("FACT_ALERTAS", pd.DataFrame()).copy()

        automaticas = df[
            (df["Prioridad"].astype(str).str.lower() == "alta") |
            (df["Vencido"] == True) |
            (df["Estado"].astype(str).str.lower().str.contains("observ", regex=True, na=False))
        ].copy()

        st.markdown("#### Alertas automáticas del sistema")
        cols = ["ID_CTAR", "SIC", "Equipo", "Servicio", "Estado", "Responsable", "Prioridad", "Dias_Atraso", "Riesgo_Clinico", "Proxima_Accion"]
        st.dataframe(automaticas[[c for c in cols if c in automaticas.columns]], use_container_width=True, hide_index=True)

        st.markdown("#### Alertas registradas")
        if not alertas.empty:
            st.dataframe(alertas, use_container_width=True, hide_index=True)
        else:
            st.caption("No hay alertas registradas en FACT_ALERTAS.")


elif page == "Registro":
    if not can_edit():
        st.error("No tiene permisos para registrar o modificar solicitudes.")
        st.stop()
    header("Registro de Solicitudes", "Formulario simple para ingresar nuevas solicitudes CTAR.")

    st.info("En modo Google Sheets, este formulario puede agregar una fila a FACT_CTAR_SEGUIMIENTO. En modo Excel, genera una fila descargable.")

    with st.form("nueva_solicitud"):
        c1, c2, c3 = st.columns(3)
        with c1:
            id_ctar = st.text_input("ID_CTAR", value=f"CTAR-{date.today().year}-XXX")
            sic = st.text_input("N° SIC")
            equipo = st.text_input("Equipo")
            inventario = st.text_input("N° inventario")
        with c2:
            servicio = st.text_input("Servicio")
            tipo = st.selectbox("Tipo proceso", ["Baja", "Reposición", "Adquisición", "Extravío", "Observación"])
            estado = st.selectbox("Estado", ["Ingresado", "En revisión CTAR", "Observado", "Aprobado", "En compra", "Recepción / instalación", "Cerrado"])
            responsable = st.selectbox("Responsable", ["Hospital", "CTAR", "Inspector Fiscal", "SCMS", "MINSAL"])
        with c3:
            prioridad = st.selectbox("Prioridad", ["Alta", "Media", "Baja"])
            fecha_ingreso = st.date_input("Fecha ingreso", value=date.today())
            fecha_compromiso = st.date_input("Fecha compromiso", value=date.today())
            link = st.text_input("Link documento Drive")

        motivo = st.text_area("Motivo")
        riesgo = st.text_area("Riesgo clínico / operacional")
        proxima = st.text_area("Próxima acción")

        submitted = st.form_submit_button("Registrar solicitud")

    if submitted:
        new_row = {
            "ID_CTAR": id_ctar,
            "SIC": sic,
            "Equipo": equipo,
            "Nro_Inventario": inventario,
            "Servicio": servicio,
            "Tipo_Proceso": tipo,
            "Estado": estado,
            "Responsable": responsable,
            "Prioridad": prioridad,
            "Fecha_Ingreso": fecha_ingreso.isoformat(),
            "Fecha_Compromiso": fecha_compromiso.isoformat(),
            "Motivo": motivo,
            "Riesgo_Clinico": riesgo,
            "Proxima_Accion": proxima,
            "Link_Documento": link,
        }

        if data_source == "Google Sheets" and spreadsheet_id:
            st.warning("Para escribir en Google Sheets, la hoja FACT_CTAR_SEGUIMIENTO debe tener las columnas exactas del modelo relacional.")
            try:
                # Orden esperado por FACT_CTAR_SEGUIMIENTO:
                row_values = [
                    id_ctar, sic, "", "", "", "", "", "", fecha_ingreso.isoformat(),
                    fecha_compromiso.isoformat(), "", motivo, riesgo, "Registro manual",
                    proxima, link
                ]
                update_gsheet_row(spreadsheet_id, "FACT_CTAR_SEGUIMIENTO", row_values)
                st.success("Solicitud registrada en Google Sheets.")
                st.cache_data.clear()
            except Exception as e:
                st.error(f"No se pudo escribir en Google Sheets: {e}")
        else:
            st.success("Solicitud generada. Descargue CSV para incorporar a su base.")
            st.download_button(
                "Descargar fila nueva CSV",
                data=pd.DataFrame([new_row]).to_csv(index=False).encode("utf-8-sig"),
                file_name="nueva_solicitud_ctar.csv",
                mime="text/csv",
            )


elif page == "Configuración":
    if not can_config():
        st.error("Solo el administrador puede acceder a configuración.")
        st.stop()
    header("Configuración del Sistema", "Validación del modelo, estructura de Google Sheets y guía de uso.")

    st.markdown("### Estado del modelo de datos")
    if tables:
        ok, missing = validate_tables(tables)
        if ok:
            st.success("Modelo cargado correctamente.")
        else:
            st.error("Faltan tablas requeridas.")
            st.write(missing)

        st.markdown("### Tablas cargadas")
        st.write({k: len(v) for k, v in tables.items()})
    else:
        st.warning("No hay datos cargados.")

    st.markdown("### Estructura mínima de Google Sheets")
    st.code(
        """
SIGE_CTAR_MASTER
├── FACT_CTAR_SEGUIMIENTO
├── DIM_EQUIPO
├── DIM_SERVICIO
├── DIM_TIPO_PROCESO
├── DIM_ESTADO
├── DIM_RESPONSABLE
├── DIM_PRIORIDAD
├── FACT_BAJAS
├── FACT_REPOSICIONES
├── FACT_ADQUISICIONES
└── FACT_ALERTAS
        """,
        language="text",
    )

    st.markdown("### Secrets para conexión Google Sheets")
    st.code(
        """
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
client_email = "..."
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"

[google_sheet]
spreadsheet_id = "ID_DEL_GOOGLE_SHEET"
        """,
        language="toml",
    )

    st.markdown("### Comando local")
    st.code("streamlit run app.py", language="bash")
