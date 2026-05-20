
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

st.markdown("""
<style>
.stApp { background: #f4f6fb; }
.block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
section[data-testid="stSidebar"] { background: linear-gradient(180deg, #1a2b4a 0%, #0f1d33 100%); }
section[data-testid="stSidebar"] * { color: #eaf2ff !important; }
.main-header {
    background: linear-gradient(90deg, #1a2b4a 0%, #2563eb 100%);
    color: white;
    padding: 22px 26px;
    border-radius: 18px;
    margin-bottom: 18px;
    box-shadow: 0 8px 22px rgba(15,23,42,.14);
}
.main-header h1 { color: white; margin: 0; padding: 0; font-size: 30px; letter-spacing: .02em; }
.main-header p { margin: 5px 0 0 0; color: #dbeafe; font-size: 14px; }
.metric-card {
    background: white;
    border-radius: 16px;
    padding: 18px;
    border: 1px solid #e5e7eb;
    box-shadow: 0 4px 16px rgba(15,23,42,.06);
    min-height: 112px;
}
.metric-title { font-size: 12px; color: #64748b; text-transform: uppercase; font-weight: 800; letter-spacing: .04em; }
.metric-value { font-size: 34px; color: #1f2937; font-weight: 900; margin-top: 8px; }
.metric-sub { font-size: 12px; color: #64748b; margin-top: 4px; }
.login-card {
    background: white;
    border-radius: 18px;
    padding: 26px;
    border: 1px solid #e5e7eb;
    box-shadow: 0 8px 24px rgba(15,23,42,.08);
    margin-top: 12px;
}
</style>
""", unsafe_allow_html=True)


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def get_users():
    try:
        return dict(st.secrets["auth"]["users"])
    except Exception:
        return {
            "admin": {"name": "Administrador CTAR", "password_hash": hash_password("admin123"), "role": "admin"},
            "ctar": {"name": "Usuario CTAR", "password_hash": hash_password("ctar123"), "role": "ctar"},
            "hospital": {"name": "Usuario Hospital", "password_hash": hash_password("hospital123"), "role": "hospital"},
            "ifiscal": {"name": "Inspector Fiscal", "password_hash": hash_password("if123"), "role": "if"},
        }


def authenticate(username, password):
    users = get_users()
    if username not in users:
        return None
    user = users[username]
    if hmac.compare_digest(user["password_hash"], hash_password(password)):
        return {"username": username, "name": user["name"], "role": user["role"]}
    return None


def login():
    col_logo, col_title = st.columns([1, 5])
    with col_logo:
        if Path(LOGO_PATH).exists():
            st.image(LOGO_PATH, width=170)
    with col_title:
        st.markdown("# SIGE-CTAR")
        st.caption("Sistema de Gestión y Trazabilidad CTAR")

    st.markdown("""
    <div class="login-card">
        <b>Ingrese con su usuario institucional.</b><br>
        <span style="color:#64748b;">Acceso por perfil: Hospital, CTAR, Inspector Fiscal o Administrador.</span>
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
            st.error("Usuario o clave incorrecta.")


if "user" not in st.session_state:
    login()
    st.stop()


def role():
    return st.session_state["user"]["role"]


def can_edit():
    return role() in ["admin", "ctar", "if"]


def is_hospital():
    return role() == "hospital"


@st.cache_data(ttl=30)
def load_google_sheets():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(st.secrets["google_sheet"]["spreadsheet_id"])

    tables = {}
    for ws in spreadsheet.worksheets():
        df = pd.DataFrame(ws.get_all_records())
        df.columns = [str(c).strip().replace(" ", "_") for c in df.columns]
        tables[ws.title] = df
    return tables


def append_to_sheet(sheet_name, row_values):
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(st.secrets["google_sheet"]["spreadsheet_id"])
    spreadsheet.worksheet(sheet_name).append_row(row_values, value_input_option="USER_ENTERED")
    st.cache_data.clear()


def build_model(tables):
    fact = tables.get("FACT_CTAR_SEGUIMIENTO", pd.DataFrame()).copy()
    if fact.empty:
        return fact

    for col in ["Fecha_Ingreso", "Fecha_Compromiso"]:
        if col in fact.columns:
            fact[col] = pd.to_datetime(fact[col], errors="coerce")

    for sheet, key in [
        ("DIM_EQUIPO", "ID_Equipo"),
        ("DIM_SERVICIO", "ID_Servicio"),
        ("DIM_TIPO_PROCESO", "ID_Tipo_Proceso"),
        ("DIM_ESTADO", "ID_Estado"),
        ("DIM_RESPONSABLE", "ID_Responsable"),
        ("DIM_PRIORIDAD", "ID_Prioridad"),
    ]:
        dim = tables.get(sheet, pd.DataFrame()).copy()
        if not dim.empty and key in fact.columns and key in dim.columns:
            fact = fact.merge(dim, on=key, how="left")

    today = pd.Timestamp(date.today())

    if "Fecha_Ingreso" in fact.columns:
        fact["Dias_Desde_Ingreso"] = (today - fact["Fecha_Ingreso"]).dt.days
    else:
        fact["Dias_Desde_Ingreso"] = 0

    if "Fecha_Compromiso" in fact.columns:
        fact["Dias_Atraso"] = (today - fact["Fecha_Compromiso"]).dt.days
        fact["Vencido"] = fact["Dias_Atraso"] > 0
    else:
        fact["Dias_Atraso"] = 0
        fact["Vencido"] = False

    for col in ["Equipo", "Servicio", "Tipo_Proceso", "Estado", "Responsable", "Prioridad",
                "Riesgo_Clinico", "Proxima_Accion", "Motivo", "Nro_Inventario", "Link_Documento"]:
        if col not in fact.columns:
            fact[col] = ""

    return fact


def header(title, subtitle):
    st.markdown(f"""
    <div class="main-header">
        <h1>{title}</h1>
        <p>{subtitle}</p>
    </div>
    """, unsafe_allow_html=True)


def metric_card(title, value, sub=""):
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">{title}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)


def filters(df):
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        servicio = st.selectbox("Servicio", ["Todos"] + sorted(df["Servicio"].dropna().astype(str).unique().tolist()))
    with c2:
        tipo = st.selectbox("Tipo proceso", ["Todos"] + sorted(df["Tipo_Proceso"].dropna().astype(str).unique().tolist()))
    with c3:
        estado = st.selectbox("Estado", ["Todos"] + sorted(df["Estado"].dropna().astype(str).unique().tolist()))
    with c4:
        responsable = st.selectbox("Responsable", ["Todos"] + sorted(df["Responsable"].dropna().astype(str).unique().tolist()))

    out = df.copy()
    if servicio != "Todos":
        out = out[out["Servicio"] == servicio]
    if tipo != "Todos":
        out = out[out["Tipo_Proceso"] == tipo]
    if estado != "Todos":
        out = out[out["Estado"] == estado]
    if responsable != "Todos":
        out = out[out["Responsable"] == responsable]
    return out


try:
    tables = load_google_sheets()
except Exception as e:
    st.error("No se pudo conectar con Google Sheets.")
    st.code(str(e))
    st.stop()

missing = [s for s in REQUIRED_SHEETS if s not in tables]
df = build_model(tables)


with st.sidebar:
    if Path(LOGO_PATH).exists():
        st.image(LOGO_PATH, use_container_width=True)

    st.markdown("## SIGE-CTAR")
    st.caption("Sistema CTAR")
    st.markdown("---")

    user = st.session_state["user"]
    st.caption(f"Usuario: {user['name']}")
    st.caption(f"Rol: {user['role']}")

    if is_hospital():
        pages = ["Resumen Ejecutivo", "Seguimiento", "Bajas", "Reposiciones", "Adquisiciones"]
    else:
        pages = ["Resumen Ejecutivo", "Seguimiento", "Bajas", "Reposiciones", "Adquisiciones", "Alertas"]

    if can_edit():
        pages.append("Registro")

    if role() == "admin":
        pages.append("Configuración")

    page = st.radio("Menú", pages)
    st.markdown("---")

    if st.button("Cerrar sesión"):
        st.session_state.clear()
        st.rerun()


if page == "Resumen Ejecutivo":
    header("SIGE-CTAR · Resumen Ejecutivo", "Seguimiento de bajas, reposiciones, adquisiciones, SIC y gestión CTAR.")

    if missing:
        st.warning("Faltan hojas en Google Sheets.")
        st.write(missing)

    if df.empty:
        st.info("No hay datos cargados en Google Sheets.")
        st.stop()

    view = filters(df)

    total = len(view)
    revision = view["Estado"].astype(str).str.lower().str.contains("revisión|revision", na=False).sum()
    aprobadas = view["Estado"].astype(str).str.lower().str.contains("aprob", na=False).sum()
    compra = view["Estado"].astype(str).str.lower().str.contains("compra", na=False).sum()
    altas = view["Prioridad"].astype(str).str.lower().eq("alta").sum()
    vencidas = int(view["Vencido"].sum())

    if is_hospital():
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: metric_card("Total solicitudes", total)
        with c2: metric_card("En revisión", revision)
        with c3: metric_card("Aprobadas", aprobadas)
        with c4: metric_card("En compra", compra)
        with c5: metric_card("Prioridad alta", altas)
    else:
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1: metric_card("Total solicitudes", total)
        with c2: metric_card("En revisión", revision)
        with c3: metric_card("Aprobadas", aprobadas)
        with c4: metric_card("En compra", compra)
        with c5: metric_card("Prioridad alta", altas)
        with c6: metric_card("Vencidas", vencidas)

    st.markdown("---")
    g1, g2, g3 = st.columns(3)

    with g1:
        st.markdown("### Estados")
        estado_df = view.groupby("Estado").size().reset_index(name="Cantidad")
        fig = px.bar(estado_df, x="Estado", y="Cantidad", text="Cantidad")
        fig.update_layout(height=330, margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with g2:
        st.markdown("### Tipo de proceso")
        fig = px.pie(view, names="Tipo_Proceso", hole=0.45)
        fig.update_layout(height=330, margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with g3:
        st.markdown("### Prioridad")
        fig = px.pie(view, names="Prioridad", hole=0.45)
        fig.update_layout(height=330, margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Seguimiento principal")
    cols = ["ID_CTAR", "SIC", "Equipo", "Servicio", "Tipo_Proceso", "Estado",
            "Responsable", "Prioridad", "Fecha_Compromiso", "Proxima_Accion"]
    st.dataframe(view[[c for c in cols if c in view.columns]], use_container_width=True, hide_index=True)


elif page == "Seguimiento":
    header("Seguimiento CTAR", "Consulta por SIC, equipo, servicio, estado, responsable y prioridad.")

    if df.empty:
        st.info("No hay datos.")
        st.stop()

    view = filters(df)
    search = st.text_input("Buscar SIC, equipo, servicio o inventario")

    if search:
        s = search.lower()
        mask = pd.Series(False, index=view.index)
        for col in ["SIC", "Equipo", "Servicio", "Nro_Inventario", "Motivo"]:
            if col in view.columns:
                mask = mask | view[col].astype(str).str.lower().str.contains(s, na=False)
        view = view[mask]

    cols = ["ID_CTAR", "SIC", "Equipo", "Nro_Inventario", "Servicio", "Tipo_Proceso",
            "Estado", "Responsable", "Prioridad", "Fecha_Ingreso", "Fecha_Compromiso",
            "Dias_Desde_Ingreso", "Motivo", "Riesgo_Clinico", "Proxima_Accion", "Link_Documento"]

    if not is_hospital():
        cols.insert(12, "Dias_Atraso")

    st.dataframe(view[[c for c in cols if c in view.columns]], use_container_width=True, hide_index=True)


elif page == "Bajas":
    header("Bajas y Extravíos", "Control de bajas, extravíos y solicitudes asociadas.")
    view = df[df["Tipo_Proceso"].astype(str).str.lower().str.contains("baja|extrav", na=False)]
    st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Reposiciones":
    header("Reposiciones", "Seguimiento desde solicitud hasta compra, recepción o cierre.")
    view = df[df["Tipo_Proceso"].astype(str).str.lower().str.contains("repos", na=False)]
    st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Adquisiciones":
    header("Adquisiciones", "Control de procesos de compra, BACO, OC y proveedor.")
    view = df[df["Tipo_Proceso"].astype(str).str.lower().str.contains("adquis|compra", na=False)]
    adq = tables.get("FACT_ADQUISICIONES", pd.DataFrame())
    if not adq.empty and "ID_CTAR" in adq.columns:
        view = view.merge(adq, on="ID_CTAR", how="left")
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

        submit = st.form_submit_button("Guardar en Google Sheets")

    if submit:
        row = [
            id_ctar, sic, id_equipo, id_servicio, id_tipo, id_estado, id_resp,
            id_prioridad, str(fecha_ingreso), str(fecha_compromiso), "",
            motivo, riesgo, ultima, proxima, link,
        ]

        try:
            append_to_sheet("FACT_CTAR_SEGUIMIENTO", row)
            st.success("Solicitud guardada correctamente.")
            st.rerun()
        except Exception as e:
            st.error("No se pudo guardar en Google Sheets.")
            st.code(str(e))


elif page == "Configuración":
    header("Configuración", "Validación de conexión y estructura Google Sheets.")
    st.markdown("### Hojas detectadas")
    st.write(list(tables.keys()))
    st.markdown("### Hojas faltantes")
    st.write(missing if missing else "No faltan hojas.")
    st.markdown("### Cantidad de registros")
    st.write({k: len(v) for k, v in tables.items()})
