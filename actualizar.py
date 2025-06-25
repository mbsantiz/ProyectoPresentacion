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
CREDENTIALS_FILE = 'astute-dreamer-463903-s9-bcbff0079541.json'
WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzGf9eb4pchk9c9k-5i-DBMxNXw8-Vo-iJq3QGJpKWwUs_fDiUwCF4NAY2EUXGjeYM_/exec'

app = Flask(__name__)

# === AUTENTICACIÓN ===
def autenticar():
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE,
        scopes=['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets.readonly']
    )
    return build('drive', 'v3', credentials=creds), build('sheets', 'v4', credentials=creds)

# === BUSCAR DATOS FIJOS EN SHEETS ===
def obtener_datos_fijos(sheets, nombre_proyecto):
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='Proyectos'
    ).execute()

    for row in result.get('values', []):
        if row[0] == nombre_proyecto:
            return {
                "TITULO1": row[1],
                "TITULO2": row[2],
                "BUOWNER": row[3],
                "PM": row[4],
                "FECHAI": row[5],
                "FECHAF": row[6]
            }
    return {}

# === OBTENER ID DE CARPETA ===
def obtener_carpeta_id(drive, nombre_proyecto):
    query = f"'{CARPETA_PADRE_ID}' in parents and name = '{nombre_proyecto}' and mimeType = 'application/vnd.google-apps.folder'"
    result = drive.files().list(q=query, fields="files(id)").execute()
    items = result.get('files', [])
    return items[0]['id'] if items else None

# === SUBIR IMAGEN DIRECTAMENTE DESDE AGENTE ===
@app.route('/subir-imagen', methods=['POST'])
def subir_imagen_directa():
    nombre_proyecto = request.form.get('nombre_proyecto')
    nombre_archivo = request.form.get('nombre_archivo')
    archivo = request.files.get('archivo')

    if not nombre_proyecto or not nombre_archivo or not archivo:
        return jsonify({"error": "Faltan datos: nombre_proyecto, nombre_archivo o archivo"}), 400

    drive, _ = autenticar()
    carpeta_id = obtener_carpeta_id(drive, nombre_proyecto)

    if not carpeta_id:
        return jsonify({"error": "No se encontró la carpeta del proyecto"}), 404

    query = f"'{carpeta_id}' in parents and name = '{nombre_archivo}'"
    archivos = drive.files().list(q=query, fields="files(id)").execute().get('files', [])
    for archivo_existente in archivos:
        drive.files().delete(fileId=archivo_existente['id']).execute()

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
    response = requests.post(WEB_APP_URL, data=json.dumps(payload), headers=headers)
    if response.status_code == 200:
        return response.json().get('url')
    else:
        raise Exception(f"Error {response.status_code}: {response.text}")

# === ENDPOINT PRINCIPAL ===
@app.route('/actualizar-presentacion', methods=['POST'])
def actualizar_presentacion():
    body = request.get_json()
    nombre_proyecto = body.get("nombre_proyecto")
    datos_variables = body.get("datos_variables")

    if not nombre_proyecto or not datos_variables:
        return jsonify({"error": "Faltan campos: 'nombre_proyecto' y/o 'datos_variables'"}), 400

    drive, sheets = autenticar()
    carpeta_id = obtener_carpeta_id(drive, nombre_proyecto)

    if not carpeta_id:
        return jsonify({"error": "No se encontró la carpeta del proyecto"}), 404

    datos_fijos = obtener_datos_fijos(sheets, nombre_proyecto)
    if not datos_fijos:
        return jsonify({"error": "No se encontraron los datos fijos en Sheets"}), 404

    datos_finales = {**datos_fijos, **datos_variables}
    url = llamar_web_app(nombre_proyecto, datos_finales)

    return jsonify({"mensaje": "✅ Presentación actualizada", "url": url})

# === RUN SERVER ===
if __name__ == '__main__':
    if not os.path.exists('temp'):
        os.makedirs('temp')
    port = int(os.environ.get('PORT', 5002))
    app.run(host='0.0.0.0', port=port, debug=True)