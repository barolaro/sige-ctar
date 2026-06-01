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

    pages = [
        "Resumen Ejecutivo",
        "Seguimiento",
        "Bajas",
        "Reposiciones",
        "Adquisiciones",
        "Repositorio Documental",
        "Reserva Presupuestaria",
    ]

    if role() != "hospital":
        pages.append("Alertas")

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
        "Bajas y Priorización Hospital",
        "Carga simple de planilla hospitalaria, priorización por colores, gestión CTAR y traspaso a histórico.",
    )

    PRIORIDADES = ["🔴 Roja", "🟠 Naranjo", "🟡 Amarilla", "🟢 Verde"]
    ESTADOS_CTAR = ["Pendiente", "En gestión", "En revisión", "Resuelto", "Cerrado"]

    def normalizar_columnas_bajas(dataframe):
        dataframe = dataframe.copy()
        dataframe.columns = [str(c).strip() for c in dataframe.columns]

        columnas_necesarias = {
            "PRIORIDAD_HOSPITAL": "",
            "JUSTIFICACION_PRIORIDAD": "",
            "GESTION_CTAR": "",
            "FECHA_ULTIMA_GESTION": "",
            "ESTADO_CTAR": "Pendiente",
            "CERRADO": "No",
            "FECHA_CIERRE": "",
        }

        for col, default in columnas_necesarias.items():
            if col not in dataframe.columns:
                dataframe[col] = default

        dataframe["PRIORIDAD_HOSPITAL"] = dataframe["PRIORIDAD_HOSPITAL"].replace("", "🟡 Amarilla")
        dataframe["ESTADO_CTAR"] = dataframe["ESTADO_CTAR"].replace("", "Pendiente")
        dataframe["CERRADO"] = dataframe["CERRADO"].replace("", "No")

        return dataframe

    def orden_prioridad(valor):
        valor = str(valor)
        if "🔴" in valor or "Roja" in valor:
            return 1
        if "🟠" in valor or "Naranjo" in valor:
            return 2
        if "🟡" in valor or "Amarilla" in valor:
            return 3
        if "🟢" in valor or "Verde" in valor:
            return 4
        return 5

    def color_prioridad_bajas(row):
        prioridad = str(row.get("PRIORIDAD_HOSPITAL", ""))
        cerrado = str(row.get("CERRADO", "")).lower()

        if cerrado in ["sí", "si", "true", "1", "cerrado"]:
            return ["background-color: #D9E1F2; color: #1F4E78"] * len(row)

        if "🔴" in prioridad or "Roja" in prioridad:
            return ["background-color: #F4CCCC; color: #7F1D1D; font-weight: bold"] * len(row)

        if "🟠" in prioridad or "Naranjo" in prioridad:
            return ["background-color: #FCE4D6; color: #7C2D12"] * len(row)

        if "🟡" in prioridad or "Amarilla" in prioridad:
            return ["background-color: #FFF2CC; color: #78350F"] * len(row)

        if "🟢" in prioridad or "Verde" in prioridad:
            return ["background-color: #D9EAD3; color: #14532D"] * len(row)

        return [""] * len(row)

    def generar_excel_salida(df_activos, df_historico):
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_activos.to_excel(writer, sheet_name="Casos_Activos_CTAR", index=False)
            df_historico.to_excel(writer, sheet_name="Historico_CTAR", index=False)

        output.seek(0)
        return output

    st.markdown("## 1️⃣ Cargar planilla priorizada por Hospital")

    st.info(
        "El Hospital completa solamente las columnas PRIORIDAD_HOSPITAL y JUSTIFICACION_PRIORIDAD. "
        "Luego CTAR carga el archivo, trabaja la gestión y marca CERRADO = Sí cuando el caso termina."
    )

    archivo_hospital = st.file_uploader(
        "Subir planilla Excel del Hospital",
        type=["xlsx"],
        help="Debe corresponder a la planilla simple de priorización de bajas CTAR.",
    )

    if archivo_hospital is None:
        st.warning("Sube la planilla del Hospital para iniciar la gestión CTAR.")
        st.stop()

    try:
        try:
            df_hospital = pd.read_excel(
                archivo_hospital,
                sheet_name="Priorizacion_Hospital",
                header=2,
                engine="openpyxl",
            )
        except Exception:
            archivo_hospital.seek(0)
            df_hospital = pd.read_excel(
                archivo_hospital,
                sheet_name=0,
                header=2,
                engine="openpyxl",
            )

        df_hospital = df_hospital.dropna(how="all")
        df_hospital = normalizar_columnas_bajas(df_hospital)

        # Eliminar filas vacías reales, manteniendo solo casos con CTAR, SIC, equipo o inventario.
        columnas_base = [
            c for c in ["CTAR", "SIC", "NOMBRE EQUIPO", "N° INVENTARIO", "DETALLE"]
            if c in df_hospital.columns
        ]

        if columnas_base:
            mask_no_vacio = df_hospital[columnas_base].astype(str).apply(
                lambda x: x.str.strip().replace("nan", "").ne("")
            ).any(axis=1)
            df_hospital = df_hospital[mask_no_vacio].copy()

        df_hospital["ORDEN_PRIORIDAD"] = df_hospital["PRIORIDAD_HOSPITAL"].apply(orden_prioridad)
        df_hospital = df_hospital.sort_values(
            by=["ORDEN_PRIORIDAD", "CTAR"] if "CTAR" in df_hospital.columns else ["ORDEN_PRIORIDAD"],
            ascending=True,
        ).drop(columns=["ORDEN_PRIORIDAD"])

        st.success(f"Planilla cargada correctamente: {len(df_hospital)} registros detectados.")

        total = len(df_hospital)
        rojos = int(df_hospital["PRIORIDAD_HOSPITAL"].astype(str).str.contains("🔴|Roja", regex=True).sum())
        naranjos = int(df_hospital["PRIORIDAD_HOSPITAL"].astype(str).str.contains("🟠|Naranjo", regex=True).sum())
        amarillos = int(df_hospital["PRIORIDAD_HOSPITAL"].astype(str).str.contains("🟡|Amarilla", regex=True).sum())
        verdes = int(df_hospital["PRIORIDAD_HOSPITAL"].astype(str).str.contains("🟢|Verde", regex=True).sum())
        cerrados = int(df_hospital["CERRADO"].astype(str).str.lower().isin(["sí", "si", "true", "1", "cerrado"]).sum())

        st.markdown("## 2️⃣ Resumen ejecutivo")

        k1, k2, k3, k4, k5, k6 = st.columns(6)

        with k1:
            st.metric("Total casos", total)

        with k2:
            st.metric("🔴 Rojos", rojos)

        with k3:
            st.metric("🟠 Naranjos", naranjos)

        with k4:
            st.metric("🟡 Amarillos", amarillos)

        with k5:
            st.metric("🟢 Verdes", verdes)

        with k6:
            st.metric("✅ Cerrados", cerrados)

        st.markdown("---")

        resumen_prioridad = df_hospital.groupby("PRIORIDAD_HOSPITAL").size().reset_index(name="Cantidad")
        resumen_prioridad["Orden"] = resumen_prioridad["PRIORIDAD_HOSPITAL"].apply(orden_prioridad)
        resumen_prioridad = resumen_prioridad.sort_values("Orden")

        fig = px.bar(
            resumen_prioridad,
            x="PRIORIDAD_HOSPITAL",
            y="Cantidad",
            text="Cantidad",
            title="Casos por prioridad informada por Hospital",
        )

        fig.update_layout(
            height=420,
            xaxis_title="Prioridad Hospital",
            yaxis_title="Cantidad de casos",
            title_x=0.02,
        )

        st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.markdown("## 3️⃣ Gestión CTAR")

        columnas_preferidas = [
            "CTAR",
            "OFICIO",
            "FECHA",
            "SIC",
            "Estado SIC",
            "UNIDAD",
            "NOMBRE EQUIPO",
            "MARCA",
            "MODELO",
            "N° SERIE",
            "N° INVENTARIO",
            "CAUSAL RECLAMADA",
            "CAUSAL APROBADA",
            "DETALLE",
            "VALOR (NETO)",
            "Observaciones",
            "PRIORIDAD_HOSPITAL",
            "JUSTIFICACION_PRIORIDAD",
            "GESTION_CTAR",
            "FECHA_ULTIMA_GESTION",
            "ESTADO_CTAR",
            "CERRADO",
            "FECHA_CIERRE",
        ]

        columnas_vista = [c for c in columnas_preferidas if c in df_hospital.columns]

        if not columnas_vista:
            columnas_vista = df_hospital.columns.tolist()

        vista = df_hospital[columnas_vista].copy()

        # Normalización de tipos para evitar errores en st.data_editor.
        # Las columnas editables deben venir como texto, no como float.
        columnas_texto_editor = [
            "PRIORIDAD_HOSPITAL",
            "JUSTIFICACION_PRIORIDAD",
            "GESTION_CTAR",
            "FECHA_ULTIMA_GESTION",
            "ESTADO_CTAR",
            "CERRADO",
            "FECHA_CIERRE",
            "OBSERVACION_CIERRE",
            "RESPONSABLE_CTAR",
        ]

        for col in columnas_texto_editor:
            if col in vista.columns:
                vista[col] = vista[col].fillna("").astype(str)

        if "PRIORIDAD_HOSPITAL" in vista.columns:
            vista["PRIORIDAD_HOSPITAL"] = vista["PRIORIDAD_HOSPITAL"].replace(
                {"": "🟡 Amarilla", "nan": "🟡 Amarilla", "None": "🟡 Amarilla"}
            )

        if "ESTADO_CTAR" in vista.columns:
            vista["ESTADO_CTAR"] = vista["ESTADO_CTAR"].replace(
                {"": "Pendiente", "nan": "Pendiente", "None": "Pendiente"}
            )

        if "CERRADO" in vista.columns:
            vista["CERRADO"] = vista["CERRADO"].replace(
                {"": "No", "nan": "No", "None": "No", "False": "No", "false": "No", "0": "No", "True": "Sí", "true": "Sí", "1": "Sí"}
            )

        disabled_cols = [c for c in vista.columns]

        if role() == "hospital":
            # Hospital solo propone prioridad y justificación.
            disabled_cols = [
                c for c in vista.columns
                if c not in ["PRIORIDAD_HOSPITAL", "JUSTIFICACION_PRIORIDAD"]
            ]
        elif can_edit():
            # CTAR/Admin/IF trabajan la gestión; se bloquean antecedentes base.
            editable_ctar = [
                "PRIORIDAD_HOSPITAL",
                "JUSTIFICACION_PRIORIDAD",
                "GESTION_CTAR",
                "FECHA_ULTIMA_GESTION",
                "ESTADO_CTAR",
                "CERRADO",
                "FECHA_CIERRE",
            ]
            disabled_cols = [
                c for c in vista.columns
                if c not in editable_ctar
            ]

        st.caption(
            "Edita la gestión CTAR y marca CERRADO = Sí cuando el caso esté terminado. "
            "Al descargar, los cerrados pasan a la hoja Histórico_CTAR."
        )

        edited = st.data_editor(
            vista,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            disabled=disabled_cols,
            column_config={
                "PRIORIDAD_HOSPITAL": st.column_config.SelectboxColumn(
                    "PRIORIDAD_HOSPITAL",
                    options=PRIORIDADES,
                    required=False,
                ),
                "ESTADO_CTAR": st.column_config.SelectboxColumn(
                    "ESTADO_CTAR",
                    options=ESTADOS_CTAR,
                    required=False,
                ),
                "CERRADO": st.column_config.SelectboxColumn(
                    "CERRADO",
                    options=["No", "Sí"],
                    required=False,
                ),
                "GESTION_CTAR": st.column_config.TextColumn(
                    "GESTION_CTAR",
                    help="Registrar avance de la mesa CTAR o gestión semanal.",
                    max_chars=500,
                ),
                "JUSTIFICACION_PRIORIDAD": st.column_config.TextColumn(
                    "JUSTIFICACION_PRIORIDAD",
                    help="Justificación entregada por Hospital.",
                    max_chars=500,
                ),
            },
            key="editor_bajas_ctar",
        )

        for col in columnas_texto_editor:
            if col in edited.columns:
                edited[col] = edited[col].fillna("").astype(str)

        if "CERRADO" in edited.columns:
            edited["CERRADO"] = edited["CERRADO"].replace(
                {"": "No", "nan": "No", "None": "No", "False": "No", "false": "No", "0": "No", "True": "Sí", "true": "Sí", "1": "Sí"}
            )

        edited["ORDEN_PRIORIDAD"] = edited["PRIORIDAD_HOSPITAL"].apply(orden_prioridad)
        edited = edited.sort_values(
            by=["CERRADO", "ORDEN_PRIORIDAD"],
            ascending=[True, True],
        ).drop(columns=["ORDEN_PRIORIDAD"])

        mask_cerrado = edited["CERRADO"].astype(str).str.lower().isin(["sí", "si", "true", "1", "cerrado"])

        casos_activos = edited[~mask_cerrado].copy()
        historico = edited[mask_cerrado].copy()

        if not historico.empty:
            historico["FECHA_REGISTRO_HISTORICO"] = pd.Timestamp.today().strftime("%Y-%m-%d")
            historico["USUARIO_REGISTRO"] = st.session_state["user"]["name"]

        st.markdown("---")
        st.markdown("## 4️⃣ Casos activos priorizados")

        st.dataframe(
            casos_activos.style.apply(color_prioridad_bajas, axis=1),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("## 5️⃣ Histórico de casos cerrados")

        if historico.empty:
            st.info("Aún no hay casos cerrados.")
        else:
            st.dataframe(
                historico.style.apply(color_prioridad_bajas, axis=1),
                use_container_width=True,
                hide_index=True,
            )

        archivo_salida = generar_excel_salida(casos_activos, historico)

        st.download_button(
            label="⬇️ Descargar archivo actualizado CTAR",
            data=archivo_salida,
            file_name="Bajas_CTAR_Actualizado_con_Historico.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        st.error("No se pudo procesar la planilla del Hospital.")
        st.warning(
            "Verifica que el archivo corresponda a la plantilla simple y que la hoja se llame Priorizacion_Hospital."
        )
        st.code(str(e))


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
