from flask import Flask, request, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# === CONFIG ===
CARPETA_PADRE_ID = '1IqWQkzIEksw6F7bxqRA_od-1f0BHaKXJ'
SPREADSHEET_ID = '1dKEnfq7RxCt8Slk73oGhKrBmYyVayt8vP3gHJJ24wjQ'
CREDENTIALS_FILE = 'astute-dreamer-463903-s9-bcbff0079541.json'

# === FLASK APP ===
app = Flask(__name__)

# === Autenticación Google ===
def autenticar():
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE,
        scopes=[
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/spreadsheets.readonly'
        ]
    )
    return build('drive', 'v3', credentials=creds), creds

# === Verificar proyecto en Google Sheets ===
def proyecto_existe_en_sheets(creds, nombre_proyecto):
    try:
        sheets = build('sheets', 'v4', credentials=creds)
        result = sheets.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='Proyectos!A2:A'
        ).execute()
        nombres = [fila[0] for fila in result.get('values', []) if fila]
        return nombre_proyecto in nombres
    except HttpError as error:
        print("❌ Error en Sheets:", error)
        return False

# === Verificar si el proyecto existe (Drive + Sheets) ===
def verificar_existencia(nombre_proyecto):
    drive_service, creds = autenticar()

    # Buscar carpeta en Drive
    query = f"'{CARPETA_PADRE_ID}' in parents and name = '{nombre_proyecto}' and mimeType = 'application/vnd.google-apps.folder'"
    resultado = drive_service.files().list(q=query, fields="files(id)").execute()
    carpeta_existe = bool(resultado.get('files'))

    # Buscar en Google Sheets
    sheets_existe = proyecto_existe_en_sheets(creds, nombre_proyecto)

    return carpeta_existe and sheets_existe

# === Endpoint ===
@app.route('/verificar-proyecto', methods=['POST'])
def verificar_proyecto():
    data = request.get_json()
    nombre_proyecto = data.get('nombreProyecto')
    
    if not nombre_proyecto:
        return jsonify({"error": "Falta 'nombreProyecto'"}), 400

    existe = verificar_existencia(nombre_proyecto)
    return jsonify({"existe": existe})

# === Run local ===
if __name__ == '__main__':
    app.run(port=5001, debug=True)