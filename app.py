import hashlib
import hmac
from datetime import date
from io import BytesIO
import gspread
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


st.set_page_config(
    page_title="SIGE-CTAR",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================
# CONFIGURACIÓN
# =========================

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

DRIVE_FOLDER_ID = "1wS6MoYjcfZGbZRtbEjPQPQTC0g_A4IGQ"

# ID DEL ARCHIVO XLSX DE RESERVA PRESUPUESTARIA
RESERVA_FILE_ID = "1YU-dxAJlkUW7BLl2DWxssHb1Fo9dmH3m"
RESERVA_WORKSHEET_NAME = "Anexo I f)"


# =========================
# ESTILO
# =========================

st.markdown(
    """
    <style>
    .main { background: #f4f6fb; }
    .block-container { padding-top: 1rem; }

    .header {
        background: linear-gradient(90deg, #1a2b4a, #2563eb);
        color: white;
        padding: 22px;
        border-radius: 16px;
        margin-bottom: 18px;
    }

    .header h1 { color: white; margin: 0; }
    .header p { color: #dbeafe; margin: 4px 0 0 0; }

    .metric-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 16px;
        box-shadow: 0 4px 12px rgba(15,23,42,.06);
        margin-bottom: 10px;
    }

    .metric-title {
        font-size: 12px;
        color: #64748b;
        text-transform: uppercase;
        font-weight: 700;
    }

    .metric-value {
        font-size: 30px;
        font-weight: 800;
        color: #1f2937;
    }

    .doc-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 15px;
        margin-bottom: 12px;
        box-shadow: 0 4px 12px rgba(15,23,42,.06);
    }

    .doc-title {
        font-size: 16px;
        font-weight: 700;
        color: #1f2937;
    }

    .doc-meta {
        font-size: 12px;
        color: #64748b;
        margin-top: 4px;
    }

    .doc-link {
        display: inline-block;
        margin-top: 8px;
        color: #2563eb;
        font-weight: 700;
        text-decoration: none;
    }

    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(5, minmax(160px, 1fr));
        gap: 14px;
        margin-bottom: 22px;
    }

    .kpi-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 18px;
        box-shadow: 0 4px 12px rgba(15,23,42,.08);
        min-height: 135px;
        overflow: hidden;
    }

    .kpi-title {
        font-size: 13px;
        font-weight: 800;
        color: #64748b;
        text-transform: uppercase;
        margin-bottom: 12px;
        line-height: 1.25;
    }

    .kpi-value {
        font-size: clamp(22px, 2.1vw, 30px);
        font-weight: 900;
        color: #1f2937;
        line-height: 1.1;
        white-space: nowrap;
    }

    .kpi-negative {
        color: #b91c1c;
    }

    .kpi-positive {
        color: #047857;
    }

    @media (max-width: 1200px) {
        .kpi-grid {
            grid-template-columns: repeat(2, 1fr);
        }
    }

    @media (max-width: 700px) {
        .kpi-grid {
            grid-template-columns: 1fr;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# LOGIN
# =========================

def hash_password(password: str) -> str:
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

    if hmac.compare_digest(
        user["password_hash"],
        hash_password(password),
    ):
        return {
            "username": username,
            "name": user["name"],
            "role": user["role"],
        }

    return None


def login():
    st.markdown(
        """
        <style>
        .login-title {
            text-align: center;
            font-size: 42px;
            font-weight: 800;
            color: #1f2937;
            margin-bottom: 0px;
        }

        .login-subtitle {
            text-align: center;
            color: #6b7280;
            font-size: 18px;
            margin-bottom: 25px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1.2, 1, 1.2])

    with col2:
        st.markdown(
            """
            <div class="login-title">🏥 SIGE-CTAR</div>
            <div class="login-subtitle">Sistema de Gestión y Trazabilidad CTAR</div>
            """,
            unsafe_allow_html=True,
        )

        with st.form("login"):
            username = st.text_input("Usuario")
            password = st.text_input("Clave", type="password")
            submit = st.form_submit_button(
                "Ingresar",
                use_container_width=True,
            )

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


# =========================
# GOOGLE SHEETS / DRIVE
# =========================

def get_credentials():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes,
    )

    return creds


@st.cache_data(ttl=60)
def load_google_sheets():
    creds = get_credentials()
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


def append_to_sheet(sheet_name, row_values):
    creds = get_credentials()
    client = gspread.authorize(creds)

    spreadsheet_id = st.secrets["google_sheet"]["spreadsheet_id"]
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.worksheet(sheet_name)
    worksheet.append_row(row_values, value_input_option="USER_ENTERED")

    st.cache_data.clear()


@st.cache_data(ttl=60)
def list_drive_files(folder_id):
    creds = get_credentials()

    service = build(
        "drive",
        "v3",
        credentials=creds,
    )

    results = service.files().list(
        q=f"'{folder_id}' in parents and trashed = false",
        fields="files(id, name, mimeType, webViewLink, modifiedTime)",
        orderBy="modifiedTime desc",
        pageSize=100,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    return results.get("files", [])


@st.cache_data(ttl=60)
def load_reserva_presupuestaria():
    creds = get_credentials()

    service = build(
        "drive",
        "v3",
        credentials=creds,
    )

    request = service.files().get_media(
        fileId=RESERVA_FILE_ID,
        supportsAllDrives=True,
    )

    file_buffer = BytesIO()
    downloader = MediaIoBaseDownload(file_buffer, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()

    file_buffer.seek(0)

    df = pd.read_excel(
        file_buffer,
        sheet_name=RESERVA_WORKSHEET_NAME,
        header=None,
        engine="openpyxl",
    )

    return df


def check_sheets(tables):
    missing = [s for s in REQUIRED_SHEETS if s not in tables]
    return missing


# =========================
# MODELO DE DATOS
# =========================

def build_model(tables):
    fact = tables.get("FACT_CTAR_SEGUIMIENTO", pd.DataFrame()).copy()

    if fact.empty:
        return fact

    for col in ["Fecha_Ingreso", "Fecha_Compromiso"]:
        if col in fact.columns:
            fact[col] = pd.to_datetime(fact[col], errors="coerce")

    joins = [
        ("DIM_EQUIPO", "ID_Equipo"),
        ("DIM_SERVICIO", "ID_Servicio"),
        ("DIM_TIPO_PROCESO", "ID_Tipo_Proceso"),
        ("DIM_ESTADO", "ID_Estado"),
        ("DIM_RESPONSABLE", "ID_Responsable"),
        ("DIM_PRIORIDAD", "ID_Prioridad"),
    ]

    df = fact.copy()

    for sheet, key in joins:
        dim = tables.get(sheet, pd.DataFrame()).copy()

        if not dim.empty and key in df.columns and key in dim.columns:
            df = df.merge(dim, on=key, how="left")

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

    for col in [
        "Equipo",
        "Servicio",
        "Tipo_Proceso",
        "Estado",
        "Responsable",
        "Prioridad",
        "Riesgo_Clinico",
        "Proxima_Accion",
    ]:
        if col not in df.columns:
            df[col] = ""

    return df


# =========================
# COMPONENTES
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
        <div class="metric-card">
            <div class="metric-title">{title}</div>
            <div class="metric-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def filters(df):
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        servicio = st.selectbox(
            "Servicio",
            ["Todos"] + sorted(df["Servicio"].dropna().astype(str).unique().tolist()),
        )

    with c2:
        tipo = st.selectbox(
            "Tipo proceso",
            ["Todos"] + sorted(df["Tipo_Proceso"].dropna().astype(str).unique().tolist()),
        )

    with c3:
        estado = st.selectbox(
            "Estado",
            ["Todos"] + sorted(df["Estado"].dropna().astype(str).unique().tolist()),
        )

    with c4:
        responsable = st.selectbox(
            "Responsable",
            ["Todos"] + sorted(df["Responsable"].dropna().astype(str).unique().tolist()),
        )

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


def icon_by_mimetype(mime_type):
    mime_type = str(mime_type).lower()

    if "pdf" in mime_type:
        return "📕"

    if "spreadsheet" in mime_type or "excel" in mime_type:
        return "📊"

    if "document" in mime_type or "word" in mime_type:
        return "📄"

    if "presentation" in mime_type or "powerpoint" in mime_type:
        return "📽️"

    if "folder" in mime_type:
        return "📁"

    return "📎"


def clean_number(value):
    value = str(value).strip()
    value = value.replace(",", "")
    value = value.replace(" ", "")

    return pd.to_numeric(value, errors="coerce")




# =========================
# HOMOLOGACIÓN PLANILLA BAJAS OFICIAL
# =========================

def homologar_fact_bajas(df_bajas):
    """Normaliza los encabezados reales de la planilla de bajas oficial al modelo interno SIGE-CTAR."""
    out = df_bajas.copy()

    out.columns = [
        str(c).strip()
        .replace(" ", "_")
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
        .replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("Ñ", "N")
        .replace("ñ", "n")
        for c in out.columns
    ]

    mapeo = {
        "CARTA": "ID_Baja",
        "FECHA_CA": "Fecha_Baja",
        "SIC": "SIC",
        "NOMBRE_EQUIPO": "Equipo",
        "Serie": "Serie",
        "SERIE": "Serie",
        "Inventario": "Nro_Inventario",
        "INVENTARIO": "Nro_Inventario",
        "Causal": "Motivo_Baja",
        "CAUSAL": "Motivo_Baja",
        "Observacion_AIF": "Observacion_AIF",
        "OBSERVACION_AIF": "Observacion_AIF",
        "CTAR": "GESTION_CTAR",
        "Comentario_SC": "Comentario_SC",
        "COMENTARIO_SC": "Comentario_SC",
        "Estado": "Estado_Baja",
        "ESTADO": "Estado_Baja",
        "Presenta_CTAR": "Presenta_CTAR",
        "PRESENTA_CTAR": "Presenta_CTAR",
    }

    for origen, destino in mapeo.items():
        if origen in out.columns:
            out.rename(columns={origen: destino}, inplace=True)

    columnas_control = [
        "PRIORIDAD_HOSPITAL",
        "JUSTIFICACION_PRIORIDAD",
        "FECHA_ULTIMA_GESTION",
        "ESTADO_CTAR",
        "CERRADO",
        "FECHA_CIERRE",
        "OBSERVACION_CIERRE",
    ]

    for col in columnas_control:
        if col not in out.columns:
            out[col] = ""

    if "ESTADO_CTAR" in out.columns:
        out["ESTADO_CTAR"] = out["ESTADO_CTAR"].replace("", "Pendiente")

    if "CERRADO" in out.columns:
        out["CERRADO"] = out["CERRADO"].replace(
            {
                "": "No",
                "False": "No",
                "false": "No",
                "0": "No",
                "True": "Sí",
                "true": "Sí",
                "1": "Sí",
            }
        )

    return out.fillna("")


# =========================
# CARGA
# =========================

try:
    tables = load_google_sheets()
except Exception as e:
    st.error("No se pudo conectar con Google Sheets.")
    st.code(str(e))
    st.stop()

missing = check_sheets(tables)

if missing:
    st.warning("Faltan hojas en Google Sheets:")
    st.write(missing)

df = build_model(tables)


# =========================
# MENÚ
# =========================

with st.sidebar:
    st.markdown("## 🏥 SIGE-CTAR")
    st.caption("Sistema CTAR")

    user = st.session_state["user"]

    st.markdown("---")
    st.caption(f"Usuario: {user['name']}")
    st.caption(f"Rol: {user['role']}")

    if role() == "hospital":
        pages = [
            "Seguimiento CTAR",
            "Repositorio Documental",
            "Reserva Presupuestaria",
        ]
    else:
        pages = [
            "Resumen Ejecutivo",
            "Seguimiento",
            "Bajas",
            "Priorización Hospital",
            "Gestión CTAR",
            "Histórico CTAR",
            "Reposiciones",
            "Adquisiciones",
            "Repositorio Documental",
            "Reserva Presupuestaria",
            "Alertas",
        ]

        if can_edit():
            pages.append("Registro")

        if role() == "admin":
            pages.append("Configuración")

    page = st.radio("Menú", pages)

    if st.button("Cerrar sesión"):
        st.session_state.clear()
        st.rerun()


# =========================
# PÁGINAS
# =========================

if page == "Resumen Ejecutivo":
    header(
        "SIGE-CTAR · Resumen Ejecutivo",
        "Seguimiento de bajas, reposiciones, adquisiciones, SIC y alertas.",
    )

    if df.empty:
        st.info("No hay datos cargados en Google Sheets.")
        st.stop()

    view = filters(df)

    total = len(view)
    revision = view["Estado"].astype(str).str.lower().str.contains("revisión|revision").sum()
    aprobadas = view["Estado"].astype(str).str.lower().str.contains("aprob").sum()
    compra = view["Estado"].astype(str).str.lower().str.contains("compra").sum()
    vencidas = int(view["Vencido"].sum())
    altas = view["Prioridad"].astype(str).str.lower().eq("alta").sum()

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
        metric_card("Vencidas", vencidas)

    with c6:
        metric_card("Prioridad alta", altas)

    st.markdown("---")

    g1, g2, g3 = st.columns(3)

    with g1:
        st.markdown("### Estados")
        fig = px.bar(
            view.groupby("Estado").size().reset_index(name="Cantidad"),
            x="Estado",
            y="Cantidad",
            text="Cantidad",
        )
        st.plotly_chart(fig, use_container_width=True)

    with g2:
        st.markdown("### Tipo proceso")
        fig = px.pie(view, names="Tipo_Proceso", hole=0.45)
        st.plotly_chart(fig, use_container_width=True)

    with g3:
        st.markdown("### Prioridad")
        fig = px.pie(view, names="Prioridad", hole=0.45)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Alertas principales")

    alertas = view[
        (view["Prioridad"].astype(str).str.lower() == "alta")
        | (view["Vencido"] == True)
        | (view["Estado"].astype(str).str.lower().str.contains("observ", na=False))
    ]

    st.dataframe(alertas, use_container_width=True, hide_index=True)


elif page == "Seguimiento":
    header(
        "Seguimiento CTAR",
        "Consulta por SIC, equipo, servicio, estado, responsable y prioridad.",
    )

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
        "Dias_Atraso",
        "Motivo",
        "Riesgo_Clinico",
        "Proxima_Accion",
        "Link_Documento",
    ]

    st.dataframe(
        view[[c for c in cols if c in view.columns]],
        use_container_width=True,
        hide_index=True,
    )


elif page == "Bajas":
    header(
        "Bajas y Gestión CTAR",
        "Gestión sincronizada con la planilla oficial de bajas. Hospital solo visualiza el seguimiento; CTAR/Admin/IF gestionan.",
    )

    # =========================
    # CONFIGURACIÓN BAJAS OFICIAL
    # =========================

    ESTADOS_CTAR = [
        "Pendiente",
        "En gestión",
        "Requiere antecedentes",
        "Compra iniciada",
        "OC emitida",
        "Recepcionado",
        "Cerrado",
    ]

    PRIORIDADES_HOSPITAL = ["🔴 Roja", "🟠 Naranjo", "🟡 Amarilla", "🟢 Verde"]

    COLUMNAS_BASE_BAJAS = [
        "CARTA_SC",
        "FECHA_CARTA",
        "SIC",
        "NOMBRE_EQUIPO",
        "Serie",
        "Inventario",
        "Causal",
        "Observacion_AIF",
        "CTAR",
        "Comentario_SC",
        "Estado",
        "Presentada_a_CTAR",
    ]

    COLUMNAS_HOSPITAL_BAJAS = [
        "PRIORIDAD_HOSPITAL",
        "JUSTIFICACION_HOSPITAL",
    ]

    COLUMNAS_CTAR_BAJAS = [
        "GESTION_CTAR",
        "FECHA_ULTIMA_GESTION",
        "ESTADO_CTAR",
        "CERRADO",
        "FECHA_CIERRE",
        "OBSERVACION_CIERRE",
    ]

    COLUMNAS_CONTROL_BAJAS = COLUMNAS_HOSPITAL_BAJAS + COLUMNAS_CTAR_BAJAS

    LABELS_BAJAS = {
        "CARTA_SC": "CARTA SC",
        "FECHA_CARTA": "FECHA CARTA",
        "SIC": "SIC",
        "NOMBRE_EQUIPO": "NOMBRE EQUIPO",
        "Serie": "Serie",
        "Inventario": "Inventario",
        "Causal": "Causal",
        "Observacion_AIF": "Observación AIF",
        "CTAR": "CTAR",
        "Comentario_SC": "Comentario SC",
        "Estado": "Estado",
        "Presentada_a_CTAR": "Presentada a CTAR",
        "PRIORIDAD_HOSPITAL": "PRIORIDAD HOSPITAL",
        "JUSTIFICACION_HOSPITAL": "JUSTIFICACIÓN HOSPITAL",
        "GESTION_CTAR": "GESTIÓN CTAR",
        "FECHA_ULTIMA_GESTION": "FECHA ÚLTIMA GESTIÓN",
        "ESTADO_CTAR": "ESTADO CTAR",
        "CERRADO": "CERRADO",
        "FECHA_CIERRE": "FECHA CIERRE",
        "OBSERVACION_CIERRE": "OBSERVACIÓN CIERRE",
    }

    def limpiar_nombre_columna(col):
        col = str(col).strip().replace(" ", "_")
        reemplazos = {
            "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
            "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
            "Ñ": "N", "ñ": "n",
        }
        for a, b in reemplazos.items():
            col = col.replace(a, b)
        return col

    def normalizar_bajas_oficial(dataframe):
        out = dataframe.copy()
        out.columns = [limpiar_nombre_columna(c) for c in out.columns]

        # Compatibilidad con nombres que han aparecido en versiones anteriores.
        alias = {
            # Encabezados oficiales / variantes
            "CARTA": "CARTA_SC",
            "CARTA_SC": "CARTA_SC",
            "FECHA_CA": "FECHA_CARTA",
            "FECHA_CARTA_SC": "FECHA_CARTA",
            "OBSERVACION_AIF": "Observacion_AIF",
            "PRESENTADA_A_CTAR": "Presentada_a_CTAR",
            "PRESENTA_CTAR": "Presentada_a_CTAR",
            "JUSTIFICACION_PRIORIDAD": "JUSTIFICACION_HOSPITAL",
            "JUSTIFICACIÓN_HOSPITAL": "JUSTIFICACION_HOSPITAL",
            "PRIORIDAD": "PRIORIDAD_HOSPITAL",

            # Compatibilidad con el modelo antiguo de FACT_BAJAS
            "ID_Baja": "CARTA_SC",
            "ID_BAJA": "CARTA_SC",
            "ID_CTAR": "CTAR",
            "Fecha_Baja": "FECHA_CARTA",
            "FECHA_BAJA": "FECHA_CARTA",
            "Equipo": "NOMBRE_EQUIPO",
            "EQUIPO": "NOMBRE_EQUIPO",
            "Nro_Inventario": "Inventario",
            "NRO_INVENTARIO": "Inventario",
            "Motivo_Baja": "Causal",
            "MOTIVO_BAJA": "Causal",
            "Estado_Baja": "Estado",
            "ESTADO_BAJA": "Estado",
        }

        for origen, destino in alias.items():
            if origen in out.columns and destino not in out.columns:
                out.rename(columns={origen: destino}, inplace=True)

        columnas_finales = COLUMNAS_BASE_BAJAS + COLUMNAS_CONTROL_BAJAS

        for col in columnas_finales:
            if col not in out.columns:
                out[col] = ""

        for col in columnas_finales:
            out[col] = out[col].fillna("").astype(str).replace({"nan": "", "None": ""})

        out["ESTADO_CTAR"] = out["ESTADO_CTAR"].replace("", "Pendiente")
        out["CERRADO"] = out["CERRADO"].replace(
            {
                "": "No",
                "False": "No",
                "false": "No",
                "0": "No",
                "True": "Sí",
                "true": "Sí",
                "1": "Sí",
            }
        )

        # Mantiene primero las columnas oficiales y luego cualquier columna adicional.
        extras = [c for c in out.columns if c not in columnas_finales]
        return out[columnas_finales + extras].fillna("")

    def df_para_guardar_google_sheets(dataframe):
        out = dataframe.copy().fillna("")
        columnas_orden = COLUMNAS_BASE_BAJAS + COLUMNAS_CONTROL_BAJAS
        extras = [c for c in out.columns if c not in columnas_orden]
        out = out[[c for c in columnas_orden if c in out.columns] + extras]
        out.rename(columns={c: LABELS_BAJAS.get(c, c) for c in out.columns}, inplace=True)
        return out

    def escribir_hoja(sheet_name, df_out):
        creds = get_credentials()
        client = gspread.authorize(creds)
        spreadsheet_id = st.secrets["google_sheet"]["spreadsheet_id"]
        spreadsheet = client.open_by_key(spreadsheet_id)

        try:
            worksheet = spreadsheet.worksheet(sheet_name)
        except Exception:
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=80)

        df_save = df_para_guardar_google_sheets(df_out)
        worksheet.clear()
        worksheet.update(
            [df_save.columns.tolist()] + df_save.astype(str).values.tolist(),
            value_input_option="USER_ENTERED",
        )
        st.cache_data.clear()

    def orden_prioridad(valor):
        v = str(valor).lower()
        if "roja" in v or "rojo" in v or "🔴" in v:
            return 1
        if "naranjo" in v or "naranja" in v or "🟠" in v:
            return 2
        if "amarilla" in v or "amarillo" in v or "🟡" in v:
            return 3
        if "verde" in v or "🟢" in v:
            return 4
        return 9

    def color_fila_prioridad(row):
        p = str(row.get("PRIORIDAD_HOSPITAL", "")).lower()
        cerrado = str(row.get("CERRADO", "")).lower()

        if cerrado in ["sí", "si", "true", "1", "cerrado"]:
            return ["background-color: #dbeafe; color: #1e3a8a"] * len(row)
        if "roja" in p or "rojo" in p or "🔴" in p:
            return ["background-color: #fee2e2; color: #7f1d1d; font-weight: bold"] * len(row)
        if "naranjo" in p or "naranja" in p or "🟠" in p:
            return ["background-color: #ffedd5; color: #7c2d12; font-weight: bold"] * len(row)
        if "amarilla" in p or "amarillo" in p or "🟡" in p:
            return ["background-color: #fef9c3; color: #713f12"] * len(row)
        if "verde" in p or "🟢" in p:
            return ["background-color: #dcfce7; color: #14532d"] * len(row)
        return [""] * len(row)

    def column_config_bajas():
        return {
            "CARTA_SC": st.column_config.TextColumn("CARTA SC"),
            "FECHA_CARTA": st.column_config.TextColumn("FECHA CARTA"),
            "SIC": st.column_config.TextColumn("SIC"),
            "NOMBRE_EQUIPO": st.column_config.TextColumn("NOMBRE EQUIPO"),
            "Serie": st.column_config.TextColumn("Serie"),
            "Inventario": st.column_config.TextColumn("Inventario"),
            "Causal": st.column_config.TextColumn("Causal"),
            "Observacion_AIF": st.column_config.TextColumn("Observación AIF"),
            "CTAR": st.column_config.TextColumn("CTAR"),
            "Comentario_SC": st.column_config.TextColumn("Comentario SC"),
            "Estado": st.column_config.TextColumn("Estado"),
            "Presentada_a_CTAR": st.column_config.TextColumn("Presentada a CTAR"),
            "PRIORIDAD_HOSPITAL": st.column_config.SelectboxColumn("PRIORIDAD HOSPITAL", options=PRIORIDADES_HOSPITAL),
            "JUSTIFICACION_HOSPITAL": st.column_config.TextColumn("JUSTIFICACIÓN HOSPITAL"),
            "GESTION_CTAR": st.column_config.TextColumn("GESTIÓN CTAR"),
            "FECHA_ULTIMA_GESTION": st.column_config.TextColumn("FECHA ÚLTIMA GESTIÓN"),
            "ESTADO_CTAR": st.column_config.SelectboxColumn("ESTADO CTAR", options=ESTADOS_CTAR),
            "CERRADO": st.column_config.SelectboxColumn("CERRADO", options=["No", "Sí"]),
            "FECHA_CIERRE": st.column_config.TextColumn("FECHA CIERRE"),
            "OBSERVACION_CIERRE": st.column_config.TextColumn("OBSERVACIÓN CIERRE"),
        }

    def crear_planilla_hospital(df_bajas):
        base = normalizar_bajas_oficial(df_bajas)
        columnas = COLUMNAS_BASE_BAJAS + COLUMNAS_HOSPITAL_BAJAS
        salida = base[[c for c in columnas if c in base.columns]].copy()
        salida.rename(columns=LABELS_BAJAS, inplace=True)

        matriz = pd.DataFrame(
            {
                "Prioridad": [1, 2, 3, 4],
                "Color": ["🔴 Roja", "🟠 Naranjo", "🟡 Amarilla", "🟢 Verde"],
                "Categoría": [
                    "Crítico / Muy Urgente",
                    "Alta Prioridad / Urgente",
                    "Prioridad Media",
                    "Prioridad Baja",
                ],
                "Criterio de Priorización": [
                    "La ausencia del equipamiento impacta directamente la continuidad de la atención clínica.",
                    "La falta del equipamiento genera limitaciones relevantes en la capacidad operativa del servicio clínico.",
                    "La ausencia del equipamiento no detiene la prestación, pero afecta eficiencia, tiempos o continuidad operativa.",
                    "La falta del equipamiento no genera impacto significativo en la continuidad asistencial.",
                ],
            }
        )

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            salida.to_excel(writer, index=False, sheet_name="Priorizacion_Hospital")
            matriz.to_excel(writer, index=False, sheet_name="Matriz_Prioridad")

            wb = writer.book
            ws = wb["Priorizacion_Hospital"]
            wm = wb["Matriz_Prioridad"]

            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.worksheet.datavalidation import DataValidation

            header_fill = PatternFill("solid", fgColor="1F4E78")
            header_font = Font(color="FFFFFF", bold=True)
            thin = Side(style="thin", color="D9E2F3")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = border

            for row in ws.iter_rows():
                for cell in row:
                    cell.border = border
                    cell.alignment = Alignment(vertical="top", wrap_text=True)

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for col in ws.columns:
                max_len = 0
                letter = col[0].column_letter
                for cell in col:
                    max_len = max(max_len, len(str(cell.value or "")))
                ws.column_dimensions[letter].width = min(max(max_len + 2, 14), 42)

            prioridad_col = None
            for idx, cell in enumerate(ws[1], start=1):
                if cell.value == "PRIORIDAD HOSPITAL":
                    prioridad_col = idx
                    break

            if prioridad_col:
                col_letter = ws.cell(row=1, column=prioridad_col).column_letter
                dv = DataValidation(
                    type="list",
                    formula1='"🔴 Roja,🟠 Naranjo,🟡 Amarilla,🟢 Verde"',
                    allow_blank=True,
                )
                ws.add_data_validation(dv)
                dv.add(f"{col_letter}2:{col_letter}5000")

            colores = {
                "🔴 Roja": "FF0000",
                "🟠 Naranjo": "F4B183",
                "🟡 Amarilla": "FFFF00",
                "🟢 Verde": "00B050",
            }

            for cell in wm[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", wrap_text=True)
                cell.border = border

            for i, row in enumerate(wm.iter_rows(min_row=2), start=2):
                color = wm.cell(i, 2).value
                fill = PatternFill("solid", fgColor=colores.get(color, "FFFFFF"))
                for cell in row:
                    cell.fill = fill
                    cell.border = border
                    cell.alignment = Alignment(vertical="top", wrap_text=True)

            for col in wm.columns:
                max_len = 0
                letter = col[0].column_letter
                for cell in col:
                    max_len = max(max_len, len(str(cell.value or "")))
                wm.column_dimensions[letter].width = min(max(max_len + 2, 14), 75)

        output.seek(0)
        return output.getvalue()

    def actualizar_prioridad(base_df, respuesta_df):
        base = normalizar_bajas_oficial(base_df)
        resp = normalizar_bajas_oficial(respuesta_df)

        key = None
        for candidate in ["CARTA_SC", "SIC", "Inventario", "Serie"]:
            if candidate in base.columns and candidate in resp.columns:
                key = candidate
                break

        if key is None:
            raise ValueError("No se encontró llave común. La planilla debe mantener CARTA SC, SIC, Inventario o Serie.")

        resp_small = resp[[key] + COLUMNAS_HOSPITAL_BAJAS].drop_duplicates(subset=[key], keep="last")
        merged = base.merge(resp_small, on=key, how="left", suffixes=("", "_NUEVO"))

        for col in COLUMNAS_HOSPITAL_BAJAS:
            nuevo = f"{col}_NUEVO"
            if nuevo in merged.columns:
                merged[col] = merged[nuevo].where(
                    merged[nuevo].astype(str).str.strip() != "",
                    merged[col],
                )
                merged.drop(columns=[nuevo], inplace=True)

        return normalizar_bajas_oficial(merged)

    def separar_activos_historico(df_gestion):
        data = normalizar_bajas_oficial(df_gestion)
        cerrado = data["CERRADO"].astype(str).str.lower().isin(
            ["sí", "si", "s", "true", "1", "cerrado"]
        ) | data["ESTADO_CTAR"].astype(str).str.lower().eq("cerrado")

        historico = data[cerrado].copy()
        activos = data[~cerrado].copy()

        if not historico.empty:
            hoy = date.today().isoformat()
            historico["FECHA_CIERRE"] = historico["FECHA_CIERRE"].replace("", hoy)
            historico["FECHA_PASO_HISTORICO"] = hoy

        return activos, historico

    bajas = tables.get("FACT_BAJAS", pd.DataFrame()).copy()

    if bajas.empty:
        st.info("No hay registros en FACT_BAJAS.")
        st.stop()

    bajas = normalizar_bajas_oficial(bajas)

    tab1, tab2, tab3 = st.tabs(
        [
            "📤 Enviar a Hospital",
            "📥 Cargar respuesta Hospital",
            "🛠️ Gestión CTAR",
        ]
    )

    with tab1:
        st.markdown("## 📤 Planilla para enviar al Hospital")
        st.info("Descarga esta planilla y envíala al Hospital. El Hospital solo debe completar PRIORIDAD HOSPITAL y JUSTIFICACIÓN HOSPITAL.")

        archivo_excel = crear_planilla_hospital(bajas)

        st.download_button(
            "📥 Descargar planilla de priorización para Hospital",
            data=archivo_excel,
            file_name=f"Priorizacion_Hospital_CTAR_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("### Vista base enviada")
        columnas_envio = COLUMNAS_BASE_BAJAS + COLUMNAS_HOSPITAL_BAJAS
        vista_envio = bajas[[c for c in columnas_envio if c in bajas.columns]].copy()
        st.dataframe(
            vista_envio,
            use_container_width=True,
            hide_index=True,
            column_config=column_config_bajas(),
        )

    with tab2:
        st.markdown("## 📥 Cargar respuesta del Hospital")
        st.info("Cuando el Hospital devuelva el archivo, súbelo aquí. El sistema actualizará PRIORIDAD HOSPITAL y JUSTIFICACIÓN HOSPITAL en FACT_BAJAS.")

        archivo = st.file_uploader(
            "Subir planilla respondida por Hospital",
            type=["xlsx"],
            key="upload_respuesta_hospital_bajas",
        )

        if archivo is not None:
            try:
                try:
                    respuesta = pd.read_excel(archivo, sheet_name="Priorizacion_Hospital", engine="openpyxl")
                except Exception:
                    archivo.seek(0)
                    respuesta = pd.read_excel(archivo, sheet_name=0, engine="openpyxl")

                actualizada = actualizar_prioridad(bajas, respuesta)
                actualizada["ORDEN_PRIORIDAD"] = actualizada["PRIORIDAD_HOSPITAL"].apply(orden_prioridad)
                actualizada = actualizada.sort_values("ORDEN_PRIORIDAD").drop(columns=["ORDEN_PRIORIDAD"])

                st.success("Planilla procesada correctamente. Revisa la vista previa antes de guardar.")

                st.dataframe(
                    actualizada,
                    use_container_width=True,
                    hide_index=True,
                    column_config=column_config_bajas(),
                )

                if can_edit():
                    if st.button("💾 Guardar priorización en FACT_BAJAS"):
                        escribir_hoja("FACT_BAJAS", actualizada)
                        st.success("FACT_BAJAS actualizado correctamente.")
                        st.rerun()
                else:
                    st.warning("Tu rol no tiene permiso para guardar cambios.")

            except Exception as e:
                st.error("No se pudo procesar la respuesta del Hospital.")
                st.warning("Verifica que la planilla mantenga CARTA SC, SIC, Inventario o Serie.")
                st.code(str(e))

    with tab3:
        st.markdown("## 🛠️ Gestión CTAR")

        activos, historico_nuevo = separar_activos_historico(bajas)
        activos["ORDEN_PRIORIDAD"] = activos["PRIORIDAD_HOSPITAL"].apply(orden_prioridad)
        activos = activos.sort_values("ORDEN_PRIORIDAD").drop(columns=["ORDEN_PRIORIDAD"])

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.metric("Activos", len(activos))
        with c2:
            st.metric("🔴 Rojas", int(activos["PRIORIDAD_HOSPITAL"].str.contains("Roja", na=False).sum()))
        with c3:
            st.metric("🟠 Naranjo", int(activos["PRIORIDAD_HOSPITAL"].str.contains("Naranjo", na=False).sum()))
        with c4:
            st.metric("🟡 Amarilla", int(activos["PRIORIDAD_HOSPITAL"].str.contains("Amarilla|Amarillo", na=False, regex=True).sum()))
        with c5:
            st.metric("🟢 Verde", int(activos["PRIORIDAD_HOSPITAL"].str.contains("Verde", na=False).sum()))

        columnas_gestion = COLUMNAS_BASE_BAJAS + COLUMNAS_CONTROL_BAJAS
        vista = activos[[c for c in columnas_gestion if c in activos.columns]].copy()

        search = st.text_input("Buscar por carta, SIC, equipo, serie, inventario o estado")
        if search:
            s = search.lower()
            mask = pd.Series(False, index=vista.index)
            for col in ["CARTA_SC", "SIC", "NOMBRE_EQUIPO", "Serie", "Inventario", "Causal", "Estado"]:
                if col in vista.columns:
                    mask = mask | vista[col].astype(str).str.lower().str.contains(s, na=False)
            vista = vista[mask]

        disabled_cols = [c for c in vista.columns if c not in COLUMNAS_CONTROL_BAJAS]

        edited = st.data_editor(
            vista,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            disabled=disabled_cols if can_edit() else vista.columns.tolist(),
            column_config=column_config_bajas(),
            key="editor_gestion_bajas_ctar_oficial",
        )

        if can_edit():
            if st.button("💾 Guardar gestión CTAR y mover cerrados a histórico"):
                edited = normalizar_bajas_oficial(edited)
                original = normalizar_bajas_oficial(bajas)

                key = None
                for candidate in ["CARTA_SC", "SIC", "Inventario", "Serie"]:
                    if candidate in original.columns and candidate in edited.columns:
                        key = candidate
                        break

                if key is None:
                    st.error("No se pudo guardar porque no existe CARTA SC, SIC, Inventario o Serie.")
                    st.stop()

                actualizado = original.copy()

                for _, row in edited.iterrows():
                    mask = actualizado[key].astype(str) == str(row[key])
                    for col in COLUMNAS_CONTROL_BAJAS:
                        if col in actualizado.columns and col in edited.columns:
                            actualizado.loc[mask, col] = row[col]

                activos_final, historico_final_nuevo = separar_activos_historico(actualizado)
                historico_existente = tables.get("HISTORICO_CTAR", pd.DataFrame()).copy()

                if not historico_existente.empty:
                    historico_existente = normalizar_bajas_oficial(historico_existente)
                    historico_total = pd.concat(
                        [historico_existente, historico_final_nuevo],
                        ignore_index=True,
                    ).fillna("")
                    if key in historico_total.columns:
                        historico_total = historico_total.drop_duplicates(subset=[key], keep="last")
                else:
                    historico_total = historico_final_nuevo.copy()

                escribir_hoja("FACT_BAJAS", activos_final)
                escribir_hoja("HISTORICO_CTAR", historico_total)

                st.success("Gestión guardada. Los casos cerrados fueron enviados a HISTORICO_CTAR.")
                st.rerun()
        else:
            st.warning("Tu rol solo puede visualizar la gestión CTAR.")

        with st.expander("📚 Ver histórico CTAR actual"):
            historico_existente = tables.get("HISTORICO_CTAR", pd.DataFrame()).copy()
            if historico_existente.empty:
                st.info("Aún no hay histórico CTAR.")
            else:
                historico_existente = normalizar_bajas_oficial(historico_existente)
                st.dataframe(
                    historico_existente,
                    use_container_width=True,
                    hide_index=True,
                    column_config=column_config_bajas(),
                )


elif page == "Seguimiento CTAR":
    header(
        "Seguimiento CTAR",
        "Vista de solo lectura para Hospital: seguimiento de bajas, prioridades y avance de gestión CTAR.",
    )

    ESTADOS_CTAR = [
        "Pendiente",
        "En gestión",
        "Requiere antecedentes",
        "Compra iniciada",
        "OC emitida",
        "Recepcionado",
        "Cerrado",
    ]

    PRIORIDADES_HOSPITAL = ["🔴 Roja", "🟠 Naranjo", "🟡 Amarilla", "🟢 Verde"]

    COLUMNAS_BASE_BAJAS = [
        "CARTA_SC",
        "FECHA_CARTA",
        "SIC",
        "NOMBRE_EQUIPO",
        "Serie",
        "Inventario",
        "Causal",
        "Observacion_AIF",
        "CTAR",
        "Comentario_SC",
        "Estado",
        "Presentada_a_CTAR",
    ]

    COLUMNAS_HOSPITAL_BAJAS = ["PRIORIDAD_HOSPITAL", "JUSTIFICACION_HOSPITAL"]
    COLUMNAS_CTAR_VISIBLE = ["GESTION_CTAR", "FECHA_ULTIMA_GESTION", "ESTADO_CTAR"]

    def limpiar_nombre_columna(col):
        col = str(col).strip().replace(" ", "_")
        reemplazos = {
            "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
            "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
            "Ñ": "N", "ñ": "n",
        }
        for a, b in reemplazos.items():
            col = col.replace(a, b)
        return col

    def normalizar_bajas_hospital(dataframe):
        out = dataframe.copy()
        out.columns = [limpiar_nombre_columna(c) for c in out.columns]

        alias = {
            # Encabezados oficiales / variantes
            "CARTA": "CARTA_SC",
            "CARTA_SC": "CARTA_SC",
            "FECHA_CA": "FECHA_CARTA",
            "FECHA_CARTA_SC": "FECHA_CARTA",
            "OBSERVACION_AIF": "Observacion_AIF",
            "PRESENTADA_A_CTAR": "Presentada_a_CTAR",
            "PRESENTA_CTAR": "Presentada_a_CTAR",
            "JUSTIFICACION_PRIORIDAD": "JUSTIFICACION_HOSPITAL",
            "JUSTIFICACIÓN_HOSPITAL": "JUSTIFICACION_HOSPITAL",

            # Compatibilidad con el modelo antiguo de FACT_BAJAS
            "ID_Baja": "CARTA_SC",
            "ID_BAJA": "CARTA_SC",
            "ID_CTAR": "CTAR",
            "Fecha_Baja": "FECHA_CARTA",
            "FECHA_BAJA": "FECHA_CARTA",
            "Equipo": "NOMBRE_EQUIPO",
            "EQUIPO": "NOMBRE_EQUIPO",
            "Nro_Inventario": "Inventario",
            "NRO_INVENTARIO": "Inventario",
            "Motivo_Baja": "Causal",
            "MOTIVO_BAJA": "Causal",
            "Estado_Baja": "Estado",
            "ESTADO_BAJA": "Estado",
        }

        for origen, destino in alias.items():
            if origen in out.columns and destino not in out.columns:
                out.rename(columns={origen: destino}, inplace=True)

        columnas = COLUMNAS_BASE_BAJAS + COLUMNAS_HOSPITAL_BAJAS + COLUMNAS_CTAR_VISIBLE + ["CERRADO"]
        for col in columnas:
            if col not in out.columns:
                out[col] = ""
            out[col] = out[col].fillna("").astype(str).replace({"nan": "", "None": ""})

        out["ESTADO_CTAR"] = out["ESTADO_CTAR"].replace("", "Pendiente")
        out["CERRADO"] = out["CERRADO"].replace(
            {
                "": "No",
                "False": "No",
                "false": "No",
                "0": "No",
                "True": "Sí",
                "true": "Sí",
                "1": "Sí",
            }
        )
        return out.fillna("")

    def orden_prioridad_hospital(valor):
        v = str(valor).lower()
        if "roja" in v or "rojo" in v or "🔴" in v:
            return 1
        if "naranjo" in v or "naranja" in v or "🟠" in v:
            return 2
        if "amarilla" in v or "amarillo" in v or "🟡" in v:
            return 3
        if "verde" in v or "🟢" in v:
            return 4
        return 9

    def color_fila_hospital(row):
        p = str(row.get("PRIORIDAD_HOSPITAL", "")).lower()
        if "roja" in p or "rojo" in p or "🔴" in p:
            return ["background-color: #fee2e2; color: #7f1d1d; font-weight: bold"] * len(row)
        if "naranjo" in p or "naranja" in p or "🟠" in p:
            return ["background-color: #ffedd5; color: #7c2d12; font-weight: bold"] * len(row)
        if "amarilla" in p or "amarillo" in p or "🟡" in p:
            return ["background-color: #fef9c3; color: #713f12"] * len(row)
        if "verde" in p or "🟢" in p:
            return ["background-color: #dcfce7; color: #14532d"] * len(row)
        return [""] * len(row)

    def column_config_hospital():
        return {
            "CARTA_SC": st.column_config.TextColumn("CARTA SC"),
            "FECHA_CARTA": st.column_config.TextColumn("FECHA CARTA"),
            "SIC": st.column_config.TextColumn("SIC"),
            "NOMBRE_EQUIPO": st.column_config.TextColumn("NOMBRE EQUIPO"),
            "Serie": st.column_config.TextColumn("Serie"),
            "Inventario": st.column_config.TextColumn("Inventario"),
            "Causal": st.column_config.TextColumn("Causal"),
            "Observacion_AIF": st.column_config.TextColumn("Observación AIF"),
            "CTAR": st.column_config.TextColumn("CTAR"),
            "Comentario_SC": st.column_config.TextColumn("Comentario SC"),
            "Estado": st.column_config.TextColumn("Estado"),
            "Presentada_a_CTAR": st.column_config.TextColumn("Presentada a CTAR"),
            "PRIORIDAD_HOSPITAL": st.column_config.TextColumn("PRIORIDAD HOSPITAL"),
            "JUSTIFICACION_HOSPITAL": st.column_config.TextColumn("JUSTIFICACIÓN HOSPITAL"),
            "GESTION_CTAR": st.column_config.TextColumn("GESTIÓN CTAR"),
            "FECHA_ULTIMA_GESTION": st.column_config.TextColumn("FECHA ÚLTIMA GESTIÓN"),
            "ESTADO_CTAR": st.column_config.TextColumn("ESTADO CTAR"),
        }

    bajas = tables.get("FACT_BAJAS", pd.DataFrame()).copy()

    if bajas.empty:
        st.info("No hay registros disponibles para seguimiento CTAR.")
        st.stop()

    bajas = normalizar_bajas_hospital(bajas)

    cerrado = bajas["CERRADO"].astype(str).str.lower().isin(
        ["sí", "si", "s", "true", "1", "cerrado"]
    ) | bajas["ESTADO_CTAR"].astype(str).str.lower().eq("cerrado")

    activos = bajas[~cerrado].copy()

    if activos.empty:
        st.success("No hay casos activos pendientes.")
        st.stop()

    activos["ORDEN_PRIORIDAD"] = activos["PRIORIDAD_HOSPITAL"].apply(orden_prioridad_hospital)
    activos = activos.sort_values("ORDEN_PRIORIDAD").drop(columns=["ORDEN_PRIORIDAD"])

    st.markdown("## 📌 Resumen de seguimiento")

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("Casos activos", len(activos))
    with c2:
        st.metric("🔴 Rojas", int(activos["PRIORIDAD_HOSPITAL"].str.contains("Roja", na=False).sum()))
    with c3:
        st.metric("🟠 Naranjo", int(activos["PRIORIDAD_HOSPITAL"].str.contains("Naranjo", na=False).sum()))
    with c4:
        st.metric("🟡 Amarillo", int(activos["PRIORIDAD_HOSPITAL"].str.contains("Amarilla|Amarillo", na=False, regex=True).sum()))
    with c5:
        st.metric("🟢 Verde", int(activos["PRIORIDAD_HOSPITAL"].str.contains("Verde", na=False).sum()))

    st.markdown("---")
    st.markdown("## 👀 Avance de gestión CTAR")

    columnas_hospital = COLUMNAS_BASE_BAJAS + COLUMNAS_HOSPITAL_BAJAS + COLUMNAS_CTAR_VISIBLE
    vista_hospital = activos[[c for c in columnas_hospital if c in activos.columns]].copy()

    search = st.text_input("Buscar por carta, SIC, equipo, serie, inventario o estado")
    if search:
        s = search.lower()
        mask = pd.Series(False, index=vista_hospital.index)
        for col in ["CARTA_SC", "SIC", "NOMBRE_EQUIPO", "Serie", "Inventario", "Causal", "Estado"]:
            if col in vista_hospital.columns:
                mask = mask | vista_hospital[col].astype(str).str.lower().str.contains(s, na=False)
        vista_hospital = vista_hospital[mask]

    st.dataframe(
        vista_hospital,
        use_container_width=True,
        hide_index=True,
        column_config=column_config_hospital(),
    )

    st.info("Esta vista es solo de consulta. El Hospital no puede modificar información desde este perfil.")


elif page == "Reposiciones":
    header(
        "Reposiciones",
        "Seguimiento desde solicitud hasta compra, recepción o cierre.",
    )

    view = df[
        df["Tipo_Proceso"].astype(str).str.lower().str.contains("repos", na=False)
    ]

    st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Adquisiciones":
    header(
        "Adquisiciones",
        "Control de procesos de compra, BACO, OC y proveedor.",
    )

    view = df[
        df["Tipo_Proceso"].astype(str).str.lower().str.contains("adquis|compra", na=False)
    ]

    adq = tables.get("FACT_ADQUISICIONES", pd.DataFrame())

    if not adq.empty and "ID_CTAR" in adq.columns:
        view = view.merge(adq, on="ID_CTAR", how="left")

    st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Repositorio Documental":
    header(
        "Repositorio Documental CTAR",
        "Consulta de documentos almacenados en Google Drive.",
    )

    st.markdown("### 📂 Documentos disponibles")

    try:
        files = list_drive_files(DRIVE_FOLDER_ID)

        if not files:
            st.info("No se encontraron documentos en la carpeta de Google Drive.")
        else:
            st.success(f"Se encontraron {len(files)} documentos en el repositorio.")

            search_doc = st.text_input("Buscar documento por nombre")

            if search_doc:
                files = [
                    f for f in files
                    if search_doc.lower() in f.get("name", "").lower()
                ]

            if not files:
                st.warning("No se encontraron documentos con ese criterio de búsqueda.")

            for f in files:
                icon = icon_by_mimetype(f.get("mimeType", ""))
                name = f.get("name", "Sin nombre")
                link = f.get("webViewLink", "#")
                modified = f.get("modifiedTime", "")

                st.markdown(
                    f"""
                    <div class="doc-card">
                        <div class="doc-title">{icon} {name}</div>
                        <div class="doc-meta">Última modificación: {modified}</div>
                        <a class="doc-link" href="{link}" target="_blank">Abrir documento</a>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    except Exception as e:
        st.error("No se pudo acceder a la carpeta de Google Drive.")
        st.warning(
            "Verifica que la Google Drive API esté habilitada y que la cuenta de servicio tenga permiso de lector en la carpeta."
        )
        st.code(str(e))


elif page == "Reserva Presupuestaria":
    header(
        "Reserva Presupuestaria",
        "Dashboard ejecutivo del Anexo I f) · Fondo de reserva y gastos durante la explotación.",
    )

    try:
        df_raw = load_reserva_presupuestaria()

        if df_raw.empty:
            st.info("No se encontraron datos en la hoja Anexo I f).")
            st.stop()

        tabla = df_raw.iloc[22:30, 1:9].copy()

        tabla.columns = [
            "Año",
            "Suplemento",
            "VMA",
            "Total general",
            "VMA año explotación",
            "Desembolsos explotación",
            "En revisión desembolsos",
            "Diferencia proyectada",
        ]

        tabla = tabla.dropna(how="all")
        tabla = tabla[tabla["Año"].astype(str).str.strip() != ""]

        columnas_uf = [
            "Suplemento",
            "VMA",
            "Total general",
            "VMA año explotación",
            "Desembolsos explotación",
            "En revisión desembolsos",
            "Diferencia proyectada",
        ]

        for col in columnas_uf:
            tabla[col] = tabla[col].apply(clean_number).fillna(0)

        fila_total = tabla[
            tabla["Año"].astype(str).str.contains("Total", case=False, na=False)
        ]

        tabla_anios = tabla[
            ~tabla["Año"].astype(str).str.contains("Total", case=False, na=False)
        ].copy()

        total_general = fila_total["Total general"].sum()
        total_vma = fila_total["VMA año explotación"].sum()
        total_desembolsos = fila_total["Desembolsos explotación"].sum()
        total_revision = fila_total["En revisión desembolsos"].sum()
        diferencia = fila_total["Diferencia proyectada"].sum()

        porcentaje_usado = 0
        if total_vma > 0:
            porcentaje_usado = ((total_desembolsos + total_revision) / total_vma) * 100

        saldo_disponible = total_vma - total_desembolsos - total_revision

        st.markdown("## 📌 Resumen ejecutivo financiero")

        c1, c2, c3 = st.columns(3)

        with c1:
            st.metric(
                "Fondo VMA explotación",
                f"{total_vma:,.2f} UF"
            )

        with c2:
            st.metric(
                "Total desembolsado",
                f"{total_desembolsos:,.2f} UF"
            )

        with c3:
            st.metric(
                "Saldo disponible estimado",
                f"{saldo_disponible:,.2f} UF"
            )

        c4, c5, c6 = st.columns(3)

        with c4:
            st.metric(
                "En revisión",
                f"{total_revision:,.2f} UF"
            )

        with c5:
            st.metric(
                "Diferencia proyectada",
                f"{diferencia:,.2f} UF"
            )

        with c6:
            st.metric(
                "Uso estimado del fondo",
                f"{porcentaje_usado:,.1f}%"
            )

        st.progress(min(max(porcentaje_usado / 100, 0), 1))

        if diferencia < 0:
            st.error(
                "El análisis muestra una diferencia proyectada negativa, por lo que corresponde mantener seguimiento y control financiero del fondo de reserva."
            )
        else:
            st.success(
                "El análisis muestra una diferencia proyectada positiva o controlada respecto del fondo de reserva."
            )

        st.markdown("---")

        st.markdown("## 📋 Tabla ejecutiva consolidada")

        st.dataframe(
            tabla.style.format({
                col: "{:,.2f}" for col in columnas_uf
            }),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("---")

        st.markdown("## 📊 Comparativo financiero por año")

        grafico_base = tabla_anios.melt(
            id_vars="Año",
            value_vars=[
                "VMA año explotación",
                "Desembolsos explotación",
                "En revisión desembolsos",
            ],
            var_name="Concepto",
            value_name="UF",
        )

        fig = px.bar(
            grafico_base,
            x="Año",
            y="UF",
            color="Concepto",
            barmode="group",
            text_auto=",.0f",
            title="VMA, desembolsos y montos en revisión por año de explotación",
        )

        fig.update_layout(
            height=520,
            xaxis_title="Año de explotación",
            yaxis_title="Monto UF",
            legend_title="Concepto",
            title_x=0.02,
            bargap=0.25,
        )

        st.plotly_chart(fig, use_container_width=True)

        st.markdown("## 🔎 Diferencia proyectada por año")

        fig2 = px.bar(
            tabla_anios,
            x="Año",
            y="Diferencia proyectada",
            text="Diferencia proyectada",
            title="Diferencia proyectada respecto del fondo de reserva",
        )

        fig2.update_traces(
            texttemplate="%{text:,.0f}",
            textposition="outside",
        )

        fig2.update_layout(
            height=500,
            xaxis_title="Año de explotación",
            yaxis_title="Diferencia proyectada UF",
            title_x=0.02,
        )

        st.plotly_chart(fig2, use_container_width=True)

        st.markdown("---")

        st.markdown("## 🧾 Lectura ejecutiva")

        st.info(
            f"""
            De acuerdo con el Anexo I f), el VMA total del año de explotación asciende a **{total_vma:,.2f} UF**. 
            A la fecha, se registran desembolsos por **{total_desembolsos:,.2f} UF** y montos en revisión por 
            **{total_revision:,.2f} UF**, lo que representa un uso estimado del fondo de **{porcentaje_usado:,.1f}%**.

            La diferencia proyectada consolidada corresponde a **{diferencia:,.2f} UF**.
            """
        )

        with st.expander("📄 Ver planilla original completa"):
            st.dataframe(
                df_raw,
                use_container_width=True,
                hide_index=True,
            )

    except Exception as e:
        st.error("No se pudo cargar la Reserva Presupuestaria.")
        st.warning(
            "Verifica que la Google Drive API esté habilitada, que el archivo sea .xlsx y que la cuenta de servicio tenga permiso de lector."
        )
        st.code(str(e))


elif page == "Alertas":
    header(
        "Alertas",
        "Procesos críticos, atrasados u observados.",
    )

    view = df[
        (df["Prioridad"].astype(str).str.lower() == "alta")
        | (df["Vencido"] == True)
        | (df["Estado"].astype(str).str.lower().str.contains("observ", na=False))
    ]

    st.dataframe(view, use_container_width=True, hide_index=True)


elif page == "Registro":
    header(
        "Registro de Solicitud",
        "Ingreso directo a Google Sheets.",
    )

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
            st.error("No se pudo guardar en Google Sheets.")
            st.code(str(e))


elif page == "Configuración":
    header(
        "Configuración",
        "Validación de conexión y estructura Google Sheets.",
    )

    st.markdown("### Hojas detectadas")
    st.write(list(tables.keys()))

    st.markdown("### Hojas faltantes")
    st.write(missing if missing else "No faltan hojas.")

    st.markdown("### Cantidad de registros")
    st.write({k: len(v) for k, v in tables.items()})
