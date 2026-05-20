import hashlib
import hmac
from datetime import date
from pathlib import Path

import gspread
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2.service_account import Credentials

st.set_page_config(
    page_title="SIGE-CTAR",
    page_icon="🏥",
    layout="wide",
)

LOGO_PATH = "assets/LOGO.jpg"

# =========================
# CSS
# =========================

st.markdown("""
<style>
.stApp { background:#f4f6fb; }
.block-container { padding-top:1rem; }

section[data-testid="stSidebar"] {
    background:#1a2b4a;
}

section[data-testid="stSidebar"] * {
    color:white !important;
}

.header {
    background:linear-gradient(90deg,#1a2b4a,#2563eb);
    color:white;
    padding:24px;
    border-radius:18px;
    margin-bottom:20px;
}

.header h1 {
    color:white;
    margin:0;
}

.header p {
    color:#dbeafe;
    margin:4px 0 0 0;
}

.card {
    background:white;
    padding:18px;
    border-radius:16px;
    border:1px solid #e5e7eb;
    box-shadow:0 4px 12px rgba(0,0,0,.06);
}

.metric-title {
    color:#64748b;
    font-size:12px;
    font-weight:700;
    text-transform:uppercase;
}

.metric-value {
    font-size:34px;
    font-weight:900;
    color:#1f2937;
}
</style>
""", unsafe_allow_html=True)


# =========================
# AUTH
# =========================

def hash_password(password):
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def get_users():
    try:
        return dict(st.secrets["auth"]["users"])
    except Exception:
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
    saved_hash = str(user.get("password_hash", ""))
    input_hash = hash_password(password)

    if hmac.compare_digest(saved_hash, input_hash):
        return {
            "username": username,
            "name": user.get("name", username),
            "role": user.get("role", "viewer"),
        }

    return None


def show_logo(width=160):
    try:
        if Path(LOGO_PATH).exists():
            st.image(LOGO_PATH, width=width)
        else:
            st.markdown("### SIGE-CTAR")
    except Exception:
        st.markdown("### SIGE-CTAR")


def login_page():
    c1, c2 = st.columns([1, 4])

    with c1:
        show_logo(150)

    with c2:
        st.title("SIGE-CTAR")
        st.caption("Sistema de Gestión y Trazabilidad CTAR")

    st.markdown("### Ingreso al sistema")

    with st.form("login_form"):
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
    login_page()
    st.stop()


def role():
    return st.session_state["user"]["role"]


def is_hospital():
    return role() == "hospital"


def can_edit():
    return role() in ["admin", "ctar", "if"]


# =========================
# GOOGLE SHEETS
# =========================

@st.cache_data(ttl=30)
def load_google_sheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes,
    )

    client = gspread.authorize(creds)
    spreadsheet_id = st.secrets["google_sheet"]["spreadsheet_id"]
    spreadsheet = client.open_by_key(spreadsheet_id)

    tables = {}

    for ws in spreadsheet.worksheets():
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        df.columns = [str(c).strip().replace(" ", "_") for c in df.columns]
        tables[ws.title] = df

    return tables


def append_to_sheet(sheet_name, row):
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes,
    )

    client = gspread.authorize(creds)
    spreadsheet_id = st.secrets["google_sheet"]["spreadsheet_id"]
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.worksheet(sheet_name)

    worksheet.append_row(row, value_input_option="USER_ENTERED")
    st.cache_data.clear()


try:
    tables = load_google_sheets()
except Exception as e:
    st.error("No se pudo conectar con Google Sheets.")
    st.code(str(e))
    st.stop()


# =========================
# MODELO DE DATOS
# =========================

def build_model(tables):
    df = tables.get("FACT_CTAR_SEGUIMIENTO", pd.DataFrame()).copy()

    if df.empty:
        return df

    joins = [
        ("DIM_EQUIPO", "ID_Equipo"),
        ("DIM_SERVICIO", "ID_Servicio"),
        ("DIM_TIPO_PROCESO", "ID_Tipo_Proceso"),
        ("DIM_ESTADO", "ID_Estado"),
        ("DIM_RESPONSABLE", "ID_Responsable"),
        ("DIM_PRIORIDAD", "ID_Prioridad"),
    ]

    for sheet_name, key in joins:
        dim = tables.get(sheet_name, pd.DataFrame()).copy()

        if not dim.empty and key in df.columns and key in dim.columns:
            df = df.merge(dim, on=key, how="left")

    for col in ["Fecha_Ingreso", "Fecha_Compromiso"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    today = pd.Timestamp(date.today())

    if "Fecha_Ingreso" in df.columns:
        df["Dias_Desde_Ingreso"] = (today - df["Fecha_Ingreso"]).dt.days
    else:
        df["Dias_Desde_Ingreso"] = 0

    if "Fecha_Compromiso" in df.columns:
        df["Dias_Atraso"] = (today - df["Fecha_Compromiso"]).dt.days
        df["Vencido"] = df["Dias_Atraso"] > 0
    else:
        df["Dias_Atraso"] = 0
        df["Vencido"] = False

    columnas_necesarias = [
        "ID_CTAR",
        "SIC",
        "Equipo",
        "Nro_Inventario",
        "Servicio",
        "Tipo_Proceso",
        "Estado",
        "Responsable",
        "Prioridad",
        "Motivo",
        "Riesgo_Clinico",
        "Proxima_Accion",
        "Link_Documento",
    ]

    for col in columnas_necesarias:
        if col not in df.columns:
            df[col] = ""

    return df


df = build_model(tables)


# =========================
# UI
# =========================

def header(title, subtitle):
    st.markdown(
        f"""
        <div class="header">
            <h1>{title}</h1>
            <p>{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_card(title, value):
    st.markdown(
        f"""
        <div class="card">
            <div class="metric-title">{title}</div>
            <div class="metric-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def apply_filters(data):
    if data.empty:
        return data

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        servicio = st.selectbox(
            "Servicio",
            ["Todos"] + sorted(data["Servicio"].dropna().astype(str).unique().tolist()),
        )

    with c2:
        tipo = st.selectbox(
            "Tipo proceso",
            ["Todos"] + sorted(data["Tipo_Proceso"].dropna().astype(str).unique().tolist()),
        )

    with c3:
        estado = st.selectbox(
            "Estado",
            ["Todos"] + sorted(data["Estado"].dropna().astype(str).unique().tolist()),
        )

    with c4:
        responsable = st.selectbox(
            "Responsable",
            ["Todos"] + sorted(data["Responsable"].dropna().astype(str).unique().tolist()),
        )

    result = data.copy()

    if servicio != "Todos":
        result = result[result["Servicio"] == servicio]

    if tipo != "Todos":
        result = result[result["Tipo_Proceso"] == tipo]

    if estado != "Todos":
        result = result[result["Estado"] == estado]

    if responsable != "Todos":
        result = result[result["Responsable"] == responsable]

    return result


# =========================
# SIDEBAR
# =========================

with st.sidebar:
    show_logo(180)

    st.markdown("## SIGE-CTAR")
    st.caption("Sistema CTAR")

    st.markdown("---")
    st.caption(f"Usuario: {st.session_state['user']['name']}")
    st.caption(f"Rol: {st.session_state['user']['role']}")

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

    if can_edit():
        pages.append("Registro")

    if role() == "admin":
        pages.append("Configuración")

    page = st.radio("Menú", pages)

    st.markdown("---")

    if st.button("Actualizar datos"):
        st.cache_data.clear()
        st.rerun()

    if st.button("Cerrar sesión"):
        st.session_state.clear()
        st.rerun()


# =========================
# PÁGINAS
# =========================

if page == "Resumen Ejecutivo":
    header(
        "SIGE-CTAR · Resumen Ejecutivo",
        "Seguimiento de bajas, reposiciones, adquisiciones y gestión CTAR.",
    )

    if df.empty:
        st.info("No hay datos cargados.")
        st.stop()

    view = apply_filters(df)

    total = len(view)
    revision = view["Estado"].astype(str).str.lower().str.contains("revision|revisión", na=False).sum()
    aprobadas = view["Estado"].astype(str).str.lower().str.contains("aprob", na=False).sum()
    compra = view["Estado"].astype(str).str.lower().str.contains("compra", na=False).sum()
    altas = view["Prioridad"].astype(str).str.lower().eq("alta").sum()
    vencidas = int(view["Vencido"].sum())

    if is_hospital():
        c1, c2, c3, c4, c5 = st.columns(5)
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

    if not is_hospital():
        with c6:
            metric_card("Vencidas", vencidas)

    st.markdown("---")

    g1, g2, g3 = st.columns(3)

    with g1:
        st.markdown("### Estados")
        chart = view.groupby("Estado").size().reset_index(name="Cantidad")
        fig = px.bar(chart, x="Estado", y="Cantidad", text="Cantidad")
        st.plotly_chart(fig, use_container_width=True)

    with g2:
        st.markdown("### Tipo proceso")
        fig = px.pie(view, names="Tipo_Proceso", hole=0.45)
        st.plotly_chart(fig, use_container_width=True)

    with g3:
        st.markdown("### Prioridad")
        fig = px.pie(view, names="Prioridad", hole=0.45)
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
        view[[c for c in cols if c in view.columns]],
        use_container_width=True,
        hide_index=True,
    )


elif page == "Seguimiento":
    header("Seguimiento CTAR", "Consulta por SIC, equipo, servicio, estado y responsable.")

    view = apply_filters(df)

    search = st.text_input("Buscar por SIC, equipo, inventario, servicio o motivo")

    if search:
        s = search.lower()
        mask = pd.Series(False, index=view.index)

        for col in ["SIC", "Equipo", "Nro_Inventario", "Servicio", "Motivo"]:
            if col in view.columns:
                mask = mask | view[col].astype(str).str.lower().str.contains(s, na=False)

        view = view[mask]

    cols = [
        "ID_CTAR",
        "SIC",
        "Equipo",
        "Nro_Inventario",
        "Servicio",
        "Tipo_Proceso",
        "Estado",
        "Responsable",
        "Prioridad",
        "Fecha_Ingreso",
        "Fecha_Compromiso",
        "Dias_Desde_Ingreso",
        "Motivo",
        "Riesgo_Clinico",
        "Proxima_Accion",
        "Link_Documento",
    ]

    if not is_hospital():
        cols.insert(12, "Dias_Atraso")

    st.dataframe(
        view[[c for c in cols if c in view.columns]],
        use_container_width=True,
        hide_index=True,
    )


elif page == "Bajas":
    header("Bajas y Extravíos", "Control de bajas, extravíos y solicitudes asociadas.")

    view = df[
        df["Tipo_Proceso"].astype(str).str.lower().str.contains("baja|extrav", na=False)
    ]

    st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Reposiciones":
    header("Reposiciones", "Seguimiento desde solicitud hasta compra, recepción o cierre.")

    view = df[
        df["Tipo_Proceso"].astype(str).str.lower().str.contains("repos", na=False)
    ]

    st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Adquisiciones":
    header("Adquisiciones", "Control de procesos de compra, BACO, OC y proveedor.")

    view = df[
        df["Tipo_Proceso"].astype(str).str.lower().str.contains("adquis|compra", na=False)
    ]

    st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Alertas":
    header("Alertas", "Procesos críticos, atrasados u observados.")

    view = df[
        (df["Prioridad"].astype(str).str.lower() == "alta")
        | (df["Vencido"] == True)
        | (df["Estado"].astype(str).str.lower().str.contains("observ", na=False))
    ]

    st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Registro":
    header("Registro de Solicitud", "Ingreso directo a Google Sheets.")

    if not can_edit():
        st.error("No tiene permisos para registrar.")
        st.stop()

    with st.form("registro"):
        id_ctar = st.text_input("ID CTAR", value=f"CTAR-{date.today().year}-XXX")
        sic = st.text_input("SIC")
        id_equipo = st.text_input("ID Equipo")
        id_servicio = st.text_input("ID Servicio")
        id_tipo = st.text_input("ID Tipo Proceso")
        id_estado = st.text_input("ID Estado")
        id_resp = st.text_input("ID Responsable")
        id_prioridad = st.text_input("ID Prioridad")
        fecha_ingreso = st.date_input("Fecha ingreso", value=date.today())
        fecha_compromiso = st.date_input("Fecha compromiso", value=date.today())
        motivo = st.text_area("Motivo")
        riesgo = st.text_area("Riesgo clínico")
        ultima = st.text_area("Última gestión")
        proxima = st.text_area("Próxima acción")
        link = st.text_input("Link documento")

        submit = st.form_submit_button("Guardar")

    if submit:
        row = [
            id_ctar,
            sic,
            id_equipo,
            id_servicio,
            id_tipo,
            id_estado,
            id_resp,
            id_prioridad,
            str(fecha_ingreso),
            str(fecha_compromiso),
            "",
            motivo,
            riesgo,
            ultima,
            proxima,
            link,
        ]

        try:
            append_to_sheet("FACT_CTAR_SEGUIMIENTO", row)
            st.success("Solicitud guardada correctamente.")
            st.rerun()
        except Exception as e:
            st.error("No se pudo guardar.")
            st.code(str(e))


elif page == "Configuración":
    header("Configuración", "Validación de conexión y estructura Google Sheets.")

    required = [
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

    missing = [s for s in required if s not in tables]

    st.markdown("### Hojas detectadas")
    st.write(list(tables.keys()))

    st.markdown("### Hojas faltantes")
    st.write(missing if missing else "No faltan hojas.")

    st.markdown("### Cantidad de registros")
    st.write({k: len(v) for k, v in tables.items()})
