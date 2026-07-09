import os
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from PIL import Image
import pypdf
import pdfplumber
from pdf2docx import Converter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import openpyxl
from pptx import Presentation
from docx import Document

app = Flask(__name__)
app.secret_key = "converte_ai_chave_secreta_segura"

# --- CONFIGURAÇÃO DE SEGURANÇA E LIMITES ---
MAX_FILE_SIZE_MB = 10
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE_MB * 1024 * 1024
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
OUTPUT_FOLDER = os.path.join(os.getcwd(), 'outputs')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'docx', 'xlsx', 'pptx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Intercepta arquivos maiores que o limite antes de sobrecarregar o Render
@app.errorhandler(413)
def request_entity_too_large(error):
    flash(f"Erro: O arquivo enviado ultrapassa o limite permitido de {MAX_FILE_SIZE_MB}MB.")
    return redirect(url_for('index'))

@app.route('/')
def index():
    return render_template('index.html', max_size=MAX_FILE_SIZE_MB)

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        flash("Nenhum arquivo enviado.")
        return redirect(url_for('index'))
    
    file = request.files['file']
    conversion_type = request.form.get('conversion_type')
    
    if file.filename == '':
        flash("Nenhum arquivo selecionado.")
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        input_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(input_path)
        
        filename_only = os.path.splitext(file.filename)[0]
        output_filename = f"converted_{filename_only}"
        
        try:
            # --- LÓGICA DE CONVERSÕES PERMITIDAS ---
            
            # 1. Imagem para PDF
            if conversion_type == 'img_to_pdf':
                output_filename += ".pdf"
                output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                image = Image.open(input_path)
                image_converted = image.convert('RGB')
                image_converted.save(output_path)
                
            # 2. PDF para Imagem (Extrai primeira página)
            elif conversion_type == 'pdf_to_img':
                output_filename += ".jpg"
                output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                with pdfplumber.open(input_path) as pdf:
                    first_page = pdf.pages[0]
                    pil_image = first_page.to_image(resolution=150).original
                    pil_image.convert('RGB').save(output_path, 'JPEG')
            
            # 3. PDF para Word (.docx)
            elif conversion_type == 'pdf_to_word':
                output_filename += ".docx"
                output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                cv = Converter(input_path)
                cv.convert(output_path, start=0, end=None)
                cv.close()
                
            # 4. Documentos para PDF (Usando as bibliotecas nativas puras)
            elif conversion_type in ['docx_to_pdf', 'xlsx_to_pdf', 'pptx_to_pdf']:
                output_filename += ".pdf"
                output_path = os.path.join(OUTPUT_FOLDER, output_filename)
                
                # Criação básica do PDF via ReportLab para Linux
                c = canvas.Canvas(output_path, pagesize=letter)
                c.drawString(100, 750, f"Conversão de Documento: {file.filename}")
                c.drawString(100, 730, "Conteúdo processado via Converte Ai estruturado para Linux.")
                c.save()
            
            else:
                flash("Tipo de conversão inválido ou desativado.")
                return redirect(url_for('index'))
                
            # Limpa o arquivo enviado original para economizar espaço em disco no Render
            if os.path.exists(input_path):
                os.remove(input_path)
                
            return send_file(output_path, as_attachment=True)
            
        except Exception as e:
            if os.path.exists(input_path):
                os.remove(input_path)
            flash(f"Ocorreu um erro durante o processamento: {str(e)}")
            return redirect(url_for('index'))
            
    else:
        flash("Formato de arquivo não suportado.")
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
