import os
from flask import Flask, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from werkzeug.utils import secure_filename
import requests
import json

# === CONFIGURACIÓN ===
CARPETA_PADRE_ID = '1IqWQkzIEksw6F7bxqRA_od-1f0BHaKXJ'
SPREADSHEET_ID = '1dKEnfq7RxCt8Slk73oGhKrBmYyVayt8vP3gHJJ24wjQ'
WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzGf9eb4pchk9c9k-5i-DBMxNXw8-Vo-iJq3QGJpKWwUs_fDiUwCF4NAY2EUXGjeYM_/exec'

app = Flask(__name__)

# === AUTENTICACIÓN USANDO VARIABLE DE ENTORNO JSON ===
def autenticar():
    creds_json_str = os.getenv('GOOGLE_CREDS_JSON')
    if not creds_json_str:
        raise Exception("No se encontró la variable de entorno GOOGLE_CREDS_JSON")

    creds_info = json.loads(creds_json_str)
    creds = service_account.Credentials.from_service_account_info(
        creds_info,
        scopes=['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
    )
    drive = build('drive', 'v3', credentials=creds)
    sheets = build('sheets', 'v4', credentials=creds)
    return drive, sheets

# === OBTENER O CREAR CARPETA ===
def obtener_o_crear_carpeta(drive, nombre_proyecto):
    query = f"'{CARPETA_PADRE_ID}' in parents and name = '{nombre_proyecto}' and mimeType = 'application/vnd.google-apps.folder'"
    result = drive.files().list(q=query, fields="files(id)").execute()
    items = result.get('files', [])
    if items:
        return items[0]['id']
    
    carpeta_metadata = {
        'name': nombre_proyecto,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [CARPETA_PADRE_ID]
    }
    carpeta = drive.files().create(body=carpeta_metadata, fields='id').execute()
    return carpeta.get('id')

# === OBTENER DATOS FIJOS EN SHEETS (solo primeras 8 columnas) ===
def obtener_datos_fijos(sheets, nombre_proyecto):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='Proyectos'
    ).execute()

    for row in result.get('values', []):
        if row and row[0] == nombre_proyecto:
            if len(row) < 8:
                raise Exception(f"Fila incompleta para el proyecto '{nombre_proyecto}'. Se requieren al menos 8 columnas.")
            return {
                "TITULO1": row[1],
                "TITULO2": row[2],
                "BUOWNER": row[3],
                "PM": row[4],
                "FECHAI": row[5],
                "FECHAF": row[6],
                "DESCRIPTION": row[7]
            }

    return {}

# === BUSCAR FILA DE PROYECTO EXISTENTE ===
def buscar_fila_proyecto(sheets, nombre_proyecto):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='Proyectos'
    ).execute()
    values = result.get('values', [])
    for i, row in enumerate(values):
        if row and row[0] == nombre_proyecto:
            return i + 1  # Las filas en Sheets son 1-indexadas
    return -1

# === AGREGAR O ACTUALIZAR PROYECTO SIN CarpetaID NI PresentacionBaseID ===
def agregar_o_actualizar_proyecto(sheets, nombre_proyecto, datos_variables):
    fila_proyecto = [
        nombre_proyecto,
        datos_variables.get("TITULO1", ""),
        datos_variables.get("TITULO2", ""),
        datos_variables.get("BUOWNER", ""),
        datos_variables.get("PM", ""),
        datos_variables.get("FECHAI", ""),
        datos_variables.get("FECHAF", ""),
        datos_variables.get("DESCRIPTION", "")
    ]
    fila = buscar_fila_proyecto(sheets, nombre_proyecto)
    if fila == -1:
        # Agregar nueva fila (solo 8 columnas)
        sheets.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range='Proyectos',
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body={'values': [fila_proyecto]}
        ).execute()
    else:
        # Actualizar fila existente y limpiar columnas 9 y 10 con vacíos
        fila_proyecto.extend(["", ""])  # Para limpiar columnas I y J
        range_to_update = f'Proyectos!A{fila}:J{fila}'
        sheets.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_to_update,
            valueInputOption='RAW',
            body={'values': [fila_proyecto]}
        ).execute()

# === AGREGAR REGISTRO DE ACTUALIZACIÓN EN HOJA "Actualizaciones" ===
def agregar_actualizacion(sheets, nombre_proyecto, datos_variables):
    fila_actualizacion = [
        nombre_proyecto,
        datos_variables.get("FECHA1", ""),
        datos_variables.get("AVANCE1", ""),
        datos_variables.get("AVANCE2", ""),
        datos_variables.get("AVANCE3", ""),
        datos_variables.get("PROX1", ""),
        datos_variables.get("PROX2", ""),
        datos_variables.get("PROX3", ""),
        datos_variables.get("RIESGO1", ""),
        datos_variables.get("RIESGO2", ""),
        datos_variables.get("RIESGO3", ""),
        datos_variables.get("RESP1", ""),
        datos_variables.get("RESP2", ""),
        datos_variables.get("RESP3", ""),
        datos_variables.get("ESTAD1", ""),
        datos_variables.get("ESTAD2", ""),
        datos_variables.get("ESTAD3", ""),
        datos_variables.get("PASO1", ""),
        datos_variables.get("PASO2", ""),
        datos_variables.get("PASO3", "")
    ]
    sheets.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range='Actualizaciones',
        valueInputOption='RAW',
        insertDataOption='INSERT_ROWS',
        body={'values': [fila_actualizacion]}
    ).execute()

# === SUBIR IMAGEN DIRECTAMENTE DESDE AGENTE ===
@app.route('/subir-imagen', methods=['POST'])
def subir_imagen_directa():
    nombre_proyecto = request.form.get('nombre_proyecto')
    nombre_archivo = request.form.get('nombre_archivo')
    archivo = request.files.get('archivo')

    if not nombre_proyecto or not nombre_archivo or not archivo:
        return jsonify({"error": "Faltan datos: nombre_proyecto, nombre_archivo o archivo"}), 400

    drive, _ = autenticar()
    carpeta_id = obtener_o_crear_carpeta(drive, nombre_proyecto)

    query = f"'{carpeta_id}' in parents and name = '{nombre_archivo}'"
    archivos = drive.files().list(q=query, fields="files(id)").execute().get('files', [])
    for archivo_existente in archivos:
        drive.files().delete(fileId=archivo_existente['id']).execute()

    if not os.path.exists('temp'):
        os.makedirs('temp')
    temp_path = os.path.join("temp", secure_filename(nombre_archivo))
    archivo.save(temp_path)

    media = MediaFileUpload(temp_path, resumable=True)
    metadata = {'name': nombre_archivo, 'parents': [carpeta_id]}
    drive.files().create(body=metadata, media_body=media, fields='id').execute()

    os.remove(temp_path)

    return jsonify({"mensaje": f"✅ Imagen '{nombre_archivo}' subida correctamente"}), 200

# === LLAMAR A WEB APP ===
def llamar_web_app(nombre_proyecto, datos_finales):
    payload = {
        'nombreProyecto': nombre_proyecto,
        'datos': datos_finales
    }
    headers = {'Content-Type': 'application/json'}
    try:
        response = requests.post(WEB_APP_URL, data=json.dumps(payload), headers=headers, timeout=40)
        response.raise_for_status()
        return response.json().get('url')
    except requests.exceptions.Timeout:
        raise Exception("⏰ Timeout al llamar al WebApp (más de 15 segundos).")
    except requests.exceptions.RequestException as e:
        raise Exception(f"❌ Error al llamar al WebApp: {e}")

# === ENDPOINT PRINCIPAL ===
@app.route('/actualizar-presentacion', methods=['POST'])
def actualizar_presentacion():
    body = request.get_json()
    nombre_proyecto = body.get("nombre_proyecto")
    datos_variables = body.get("datos_variables", {})
    datos_fijos_entrantes = body.get("datos_fijos", {})

    if not nombre_proyecto or not datos_variables:
        return jsonify({"error": "Faltan campos: 'nombre_proyecto' y/o 'datos_variables'"}), 400

    drive, sheets = autenticar()

    # Intentar obtener datos fijos guardados (solo primeras 8 columnas)
    datos_fijos = obtener_datos_fijos(sheets, nombre_proyecto)

    if not datos_fijos:
        if not datos_fijos_entrantes:
            return jsonify({"error": "El proyecto no existe y no se enviaron datos fijos para crearlo"}), 400
        agregar_o_actualizar_proyecto(sheets, nombre_proyecto, datos_fijos_entrantes)
        datos_fijos = datos_fijos_entrantes
    else:
        # Actualizar datos fijos sin CarpetaID ni PresentacionBaseID si se envían nuevos datos fijos
        if datos_fijos_entrantes:
            agregar_o_actualizar_proyecto(sheets, nombre_proyecto, datos_fijos_entrantes)
            datos_fijos = datos_fijos_entrantes

    datos_finales = {**datos_fijos, **datos_variables}

    url = llamar_web_app(nombre_proyecto, datos_finales)

    # Guardar registro de actualización en hoja "Actualizaciones"
    agregar_actualizacion(sheets, nombre_proyecto, datos_variables)

    return jsonify({"mensaje": "✅ Presentación creada o actualizada", "url": url})

# === RUN SERVER ===
if __name__ == '__main__':
    if not os.path.exists('temp'):
        os.makedirs('temp')
    port = int(os.environ.get('PORT', 5002))
    app.run(host='0.0.0.0', port=port, debug=True)
