from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError
import os
import requests
import json

# === CONFIGURACIÓN ===
CARPETA_PADRE_ID = '1IqWQkzIEksw6F7bxqRA_od-1f0BHaKXJ'
SPREADSHEET_ID = '1dKEnfq7RxCt8Slk73oGhKrBmYyVayt8vP3gHJJ24wjQ'
WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzGf9eb4pchk9c9k-5i-DBMxNXw8-Vo-iJq3QGJpKWwUs_fDiUwCF4NAY2EUXGjeYM_/exec'
CREDENTIALS_FILE = 'astute-dreamer-463903-s9-bcbff0079541.json'

# === AUTENTICACIÓN ===
def autenticar():
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE,
        scopes=[
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/spreadsheets.readonly'
        ]
    )
    return build('drive', 'v3', credentials=creds), creds

# === SUBIR IMÁGENES ===
def subir_imagenes(service, carpeta_id):
    imagenes = [
        ("Logo1.png", "Logo1.png"),
        ("AVANCE%.png", "AVANCE%.png"),
        ("GANTT.png", "GANTT.png")
    ]
    for nombre_drive, ruta_local in imagenes:
        if not os.path.exists(ruta_local):
            print(f"⚠️ Imagen no encontrada: {ruta_local}")
            continue

        query = f"'{carpeta_id}' in parents and name = '{nombre_drive}' and trashed = false"
        archivos = service.files().list(q=query, fields="files(id)").execute().get('files', [])
        for archivo in archivos:
            service.files().delete(fileId=archivo['id']).execute()
            print(f"🗑️ Imagen anterior eliminada: {nombre_drive}")

        metadata = {'name': nombre_drive, 'parents': [carpeta_id]}
        media = MediaFileUpload(ruta_local, resumable=True)
        service.files().create(body=metadata, media_body=media, fields='id').execute()
        print(f"✅ Imagen subida: {nombre_drive}")

# === LEER DATOS FIJOS ===
def obtener_datos_fijos(creds, nombre_proyecto):
    sheets = build('sheets', 'v4', credentials=creds)
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='Proyectos!A2:G'
    ).execute()

    filas = result.get('values', [])
    for fila in filas:
        if fila[0] == nombre_proyecto:
            return {
                "TITULO1": fila[1],
                "TITULO2": fila[2],
                "BUOWNER": fila[3],
                "PM": fila[4],
                "FECHAI": fila[5],
                "FECHAF": fila[6]
            }
    return None

# === OBTENER O CREAR CARPETA ===
def obtener_o_crear_carpeta(service, nombre_proyecto):
    query = f"'{CARPETA_PADRE_ID}' in parents and name = '{nombre_proyecto}' and mimeType = 'application/vnd.google-apps.folder'"
    resultado = service.files().list(q=query, fields="files(id)").execute()
    folders = resultado.get('files', [])
    if folders:
        return folders[0]['id']
    else:
        metadata = {'name': nombre_proyecto, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [CARPETA_PADRE_ID]}
        folder = service.files().create(body=metadata, fields='id').execute()
        return folder['id']

# === LLAMAR AL WEB APP ===
def llamar_web_app(nombre_proyecto, datos_combinados):
    response = requests.post(
        WEB_APP_URL,
        headers={'Content-Type': 'application/json'},
        data=json.dumps({
            'nombreProyecto': nombre_proyecto,
            'datos': datos_combinados
        })
    )
    if response.status_code == 200:
        print("✅ Presentación actualizada")
        print("🔗 URL:", response.json().get("url"))
    else:
        print(f"❌ Error Web App: {response.status_code} - {response.text}")

# === FUNCIÓN PRINCIPAL ===
def actualizar_presentacion(nombre_proyecto: str, datos_variables: dict):
    drive_service, creds = autenticar()

    # Verificar si el proyecto ya existe
    datos_fijos = obtener_datos_fijos(creds, nombre_proyecto)
    if not datos_fijos:
        raise Exception("⚠️ El proyecto no existe en la hoja 'Proyectos'. No se pueden recuperar datos fijos.")

    carpeta_id = obtener_o_crear_carpeta(drive_service, nombre_proyecto)

    subir_imagenes(drive_service, carpeta_id)

    # Combinar los datos
    datos_finales = {**datos_fijos, **datos_variables}

    # Llamar a Web App
    llamar_web_app(nombre_proyecto, datos_finales)

# === PRUEBA LOCAL ===
if __name__ == '__main__':
    datos_variables = {
        "DESCRIPTION": "Actualización semanal desde el agente",
        "AVANCE1": "Agente OpenAI operativo",
        "AVANCE2": "Subida de imágenes automática",
        "AVANCE3": "Sin intervención manual",
        "PROX1": "Probar nuevo endpoint",
        "PROX2": "Preparar demo",
        "PROX3": "Feedback del equipo",
        "RIESGO1": "Errores de red",
        "RIESGO2": "Pérdida de autenticación",
        "RIESGO3": "Falta de permisos",
        "RESP1": "Marcos",
        "RESP2": "Bot",
        "RESP3": "QA",
        "ESTAD1": "Completado",
        "ESTAD2": "En curso",
        "ESTAD3": "Pendiente",
        "PASO1": "Revisión final",
        "PASO2": "Entrega",
        "PASO3": "Seguimiento"
    }

    actualizar_presentacion("PythonTest", datos_variables)