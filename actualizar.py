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

# === AUTENTICACIÓN ===
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

    metadata = {
        'name': nombre_proyecto,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [CARPETA_PADRE_ID]
    }
    carpeta = drive.files().create(body=metadata, fields='id').execute()
    return carpeta['id']

# === OBTENER DATOS FIJOS DESDE SHEETS ===
def obtener_datos_fijos(sheets, nombre_proyecto):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='Proyectos'
    ).execute()

    for row in result.get('values', []):
        if row[0] == nombre_proyecto:
            return {
                "TITULO1": row[1] if len(row) > 1 else "",
                "TITULO2": row[2] if len(row) > 2 else "",
                "BUOWNER": row[3] if len(row) > 3 else "",
                "PM": row[4] if len(row) > 4 else "",
                "FECHAI": row[5] if len(row) > 5 else "",
                "FECHAF": row[6] if len(row) > 6 else ""
            }
    return {}

# === AGREGAR PROYECTO Y ACTUALIZACIÓN NUEVA ===
def agregar_proyecto_y_actualizacion(sheets, nombre_proyecto, datos_fijos, datos_variables, carpeta_id, presentacion_id):
    fila_proyecto = [
        nombre_proyecto,
        datos_fijos.get("TITULO1", ""),
        datos_fijos.get("TITULO2", ""),
        datos_fijos.get("BUOWNER", ""),
        datos_fijos.get("PM", ""),
        datos_fijos.get("FECHAI", ""),
        datos_fijos.get("FECHAF", ""),
        carpeta_id,
        presentacion_id
    ]
    sheets.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range='Proyectos',
        valueInputOption='RAW',
        insertDataOption='INSERT_ROWS',
        body={'values': [fila_proyecto]}
    ).execute()

    fila_actualizacion = [
        nombre_proyecto,
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
        datos_variables.get("ESTAD3", "")
    ]
    sheets.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range='Actualizaciones',
        valueInputOption='RAW',
        insertDataOption='INSERT_ROWS',
        body={'values': [fila_actualizacion]}
    ).execute()

# === LLAMAR AL WEB APP ===
def llamar_web_app(nombre_proyecto, datos_finales):
    payload = {
        'nombreProyecto': nombre_proyecto,
        'datos': datos_finales
    }
    headers = {'Content-Type': 'application/json'}
    response = requests.post(WEB_APP_URL, data=json.dumps(payload), headers=headers)
    if response.status_code == 200:
        return response.json().get('url')
    else:
        raise Exception(f"Error {response.status_code}: {response.text}")

# === SUBIR IMAGEN A DRIVE ===
@app.route('/subir-imagen', methods=['POST'])
def subir_imagen_directa():
    nombre_proyecto = request.form.get('nombre_proyecto')
    nombre_archivo = request.form.get('nombre_archivo')
    archivo = request.files.get('archivo')

    if not nombre_proyecto or not nombre_archivo or not archivo:
        return jsonify({"error": "Faltan datos"}), 400

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

# === ACTUALIZAR O CREAR PRESENTACIÓN ===
@app.route('/actualizar-presentacion', methods=['POST'])
def actualizar_presentacion():
    body = request.get_json()
    nombre_proyecto = body.get("nombre_proyecto")
    datos_fijos = body.get("datos_fijos")
    datos_variables = body.get("datos_variables")

    if not nombre_proyecto or not datos_variables:
        return jsonify({"error": "Faltan campos requeridos"}), 400

    drive, sheets = autenticar()
    carpeta_id = obtener_o_crear_carpeta(drive, nombre_proyecto)

    datos_fijos_existentes = obtener_datos_fijos(sheets, nombre_proyecto)
    if datos_fijos_existentes:
        datos_finales = {**datos_fijos_existentes, **datos_variables}
        url = llamar_web_app(nombre_proyecto, datos_finales)
        return jsonify({"mensaje": "✅ Presentación actualizada", "url": url})

    if not datos_fijos:
        return jsonify({"error": "Faltan datos fijos para proyecto nuevo"}), 400

    datos_finales = {**datos_fijos, **datos_variables}
    url = llamar_web_app(nombre_proyecto, datos_finales)

    try:
        presentacion_id = url.split('/d/')[1].split('/')[0]
    except:
        presentacion_id = ""

    agregar_proyecto_y_actualizacion(sheets, nombre_proyecto, datos_fijos, datos_variables, carpeta_id, presentacion_id)

    return jsonify({"mensaje": "✅ Proyecto nuevo creado y presentación generada", "url": url})

# === EJECUCIÓN LOCAL ===
if __name__ == '__main__':
    if not os.path.exists('temp'):
        os.makedirs('temp')
    port = int(os.environ.get('PORT', 5002))
    app.run(host='0.0.0.0', port=port, debug=True)