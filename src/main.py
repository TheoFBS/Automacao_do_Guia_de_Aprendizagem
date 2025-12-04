import logging
import os
from api.automation import Automation
from config.settings import (
    CREDS,
    BOT_CREDS,
    TOKEN,
    SCOPES,
    OUTPUT_FOLDER,
    TEMPLATE_FILE
)

from flask import Flask
from flask import render_template, request, jsonify


# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("src/logs/app.log"),
        logging.StreamHandler()
    ]
)

app = Flask(__name__)
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/', methods=['GET'])
def get_word():
    return

@app.route('/processar', methods=['POST'])
def processar():
    if request.is_json:
        data = request.get_json()
        
        
        automation = Automation(CREDS, BOT_CREDS, TOKEN, SCOPES)
        output_file = os.path.join(OUTPUT_FOLDER, "documento_preenchido.docx")
        automation.process_document(TEMPLATE_FILE, output_file, data)
        
        
        return jsonify({'message': f'{data}'}), 200
    return jsonify({'error': 'deu ruim'}), 400

def main():
    output_file = os.path.join(OUTPUT_FOLDER, "documento_preenchido.docx")
    # Inicia Processo
    automation = Automation(CREDS, BOT_CREDS, TOKEN, SCOPES)
    # Processa um Documento
    automation.process_document(TEMPLATE_FILE, output_file, '1LSEerPaaokPQHR3Ov0miPCQqUGUyumobVUWIv49y4No', 'Arte!A2:J322', '5º ano', '1°')
    

if __name__ == "__main__":
    app.run(port=8000)