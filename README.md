
# SIGE-CTAR

Sistema de Gestión y Trazabilidad CTAR para seguimiento de bajas, reposiciones, adquisiciones, SIC, responsables, alertas y documentos.

## 1. Objetivo

Este sistema permite centralizar el estado operacional de cada solicitud asociada al CTAR.

Sirve para responder preguntas como:

- ¿En qué estado está este SIC?
- ¿Quién tiene la acción pendiente?
- ¿El equipo está en baja, reposición o adquisición?
- ¿Está atrasado?
- ¿Cuál es la prioridad clínica?
- ¿Dónde está el documento de respaldo?

## 2. Estructura recomendada

El sistema trabaja con un Google Sheet o Excel con estas hojas:

- FACT_CTAR_SEGUIMIENTO
- DIM_EQUIPO
- DIM_SERVICIO
- DIM_TIPO_PROCESO
- DIM_ESTADO
- DIM_RESPONSABLE
- DIM_PRIORIDAD
- FACT_BAJAS
- FACT_REPOSICIONES
- FACT_ADQUISICIONES
- FACT_ALERTAS

## 3. Instalación local

Crear entorno:

```bash
python -m venv .venv
```

Activar entorno:

```bash
# Windows
.venv\Scripts\activate

# Mac/Linux
source .venv/bin/activate
```

Instalar dependencias:

```bash
pip install -r requirements.txt
```

Ejecutar:

```bash
streamlit run app.py
```

## 4. Uso con Excel

Deje el archivo `CTAR_RelationalModel.xlsx` dentro de la carpeta `data`.

También puede cargar el Excel manualmente desde el menú lateral.

## 5. Uso con Google Sheets

Debe crear un Google Sheet con las mismas hojas del modelo.

Luego debe crear una credencial de Service Account en Google Cloud y compartir el Sheet con el correo del Service Account.

En Streamlit Cloud o local, crear `.streamlit/secrets.toml` con:

```toml
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "..."
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"

[google_sheet]
spreadsheet_id = "ID_DEL_GOOGLE_SHEET"
```

## 6. Flujo operacional

Hospital solicita o consulta equipo.

Luego el CTAR registra y actualiza:

1. Solicitud recibida.
2. Revisión CTAR.
3. Baja, reposición o adquisición.
4. Observaciones.
5. Responsable actual.
6. Compra, BACO u OC.
7. Recepción o instalación.
8. Cierre.

## 7. Recomendación de implementación

Fase 1:
- Google Sheet
- Streamlit
- Registro y seguimiento

Fase 2:
- Alertas
- Historial
- Documentos Drive

Fase 3:
- Login
- PostgreSQL
- permisos por perfil
