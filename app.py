import os
import zipfile
from flask import Flask, request, send_file, flash, redirect, url_for, Response

# ==============================================================================
# IMPORTAÇÃO SEGURA DE BIBLIOTECAS (NATIVAS PARA LINUX/NUVEM)
# ==============================================================================
try:
    from PIL import Image
except ImportError:
    Image = None

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader = None
    PdfWriter = None

try:
    from pdf2docx import Converter
except ImportError:
    Converter = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    Workbook = None
    load_workbook = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from pptx import Presentation
    from pptx.util import Inches
except ImportError:
    Presentation = None
    Inches = None

try:
    from docx import Document
except ImportError:
    Document = None

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
except ImportError:
    SimpleDocTemplate = None

# ==============================================================================
# CONFIGURAÇÃO DO APP E LIMITES DE SEGURANÇA
# ==============================================================================
app = Flask(__name__)
app.secret_key = "converte_ai_chave_secreta_segura"

MAX_FILE_SIZE_MB = 10  # Proteção do processador do Render
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE_MB * 1024 * 1024

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Intercepta arquivos maiores que 10MB antes de sobrecarregar o Render
@app.errorhandler(413)
def request_entity_too_large(error):
    conteudo_erro = f'''
        <h3 class="tool-title" style="color: #dc2626;"><i class="fa-solid fa-triangle-exclamation"></i> Arquivo Muito Grande</h3>
        <p class="tool-desc">O servidor bloqueou o upload porque o arquivo ultrapassa o limite de <strong>{MAX_FILE_SIZE_MB} MB</strong> configurado para proteger o processador.</p>
        <a href="/" style="display: inline-block; background: var(--primary); color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; font-weight: 600; margin-top: 10px;">Voltar ao Início</a>
    '''
    return layout_base(conteudo_erro, ''), 413

# ==============================================================================
# INTERFACE BASE UNIFICADA (Dispensa arquivo index.html externo)
# ==============================================================================
def layout_base(conteudo_pagina, ferramenta_ativa):
    menu_items = {
        'imagem-to-pdf': ('/', 'fa-file-image', 'Imagem para PDF'),
        'juntar': ('/juntar-view', 'fa-file-pdf', 'Juntar PDFs'),
        'dividir': ('/dividir-view', 'fa-scissors', 'Dividir PDF'),
        'pdf-to-word': ('/pdf-to-word-view', 'fa-file-word', 'PDF para Word'),
        'pdf-to-excel': ('/pdf-to-excel-view', 'fa-file-excel', 'PDF para Excel'),
        'pdf-to-pptx': ('/pdf-to-pptx-view', 'fa-file-powerpoint', 'PDF para PPTX'),
        'pdf-to-image': ('/pdf-to-image-view', 'fa-images', 'PDF para Imagem'),
        'word-to-pdf': ('/word-to-pdf-view', 'fa-file-word', 'Word para PDF'),
        'excel-to-pdf': ('/excel-to-pdf-view', 'fa-file-excel', 'Excel para PDF'),
    }
    
    categorias = {
        'Modificar PDF': ['juntar', 'dividir'],
        'Converter de PDF': ['pdf-to-word', 'pdf-to-excel', 'pdf-to-pptx', 'pdf-to-image'],
        'Converter para PDF': ['word-to-pdf', 'excel-to-pdf', 'imagem-to-pdf']
    }
    
    sidebar_html = ""
    for cat_name, keys in categorias.items():
        sidebar_html += f'<div class="menu-category">{cat_name}</div>'
        for key in keys:
            route, icon, label = menu_items[key]
            active_class = "active" if ferramenta_ativa == key else ""
            sidebar_html += f'<a href="{route}" class="menu-item {active_class}"><i class="fa-solid {icon}"></i>{label}</a>'
    
    html_retorno = f'''<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta name="google-site-verification" content="aS2P031kVY2N7PIOuVGnvwyjImWcpOz8MRl7p9sTLAg" />
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Converte Ai - Conversor de Arquivos Online Grátis</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {{
            --primary-dark: #0f172a;
            --primary: #1d4ed8;
            --primary-light: #eff6ff;
            --accent: #2563eb;
            --bg-main: #f8fafc;
            --bg-card: #ffffff;
            --text-dark: #1e293b;
            --text-muted: #64748b;
            --border-color: #e2e8f0;
            --sidebar-width: 260px;
            --ad-width: 180px;
        }}
        body {{
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            margin: 0; padding: 0;
            background-color: var(--bg-main);
            color: var(--text-dark);
            display: flex;
            min-height: 100vh;
        }}
        .sidebar {{
            width: var(--sidebar-width);
            background-color: var(--primary-dark);
            color: #f1f5f9;
            display: flex; flex-direction: column;
            position: fixed;
            height: 100vh;
            overflow-y: auto;
            z-index: 100;
        }}
        .sidebar-header {{
            padding: 24px;
            display: flex; align-items: center; gap: 10px;
            border-bottom: 1px solid #1e293b;
        }}
        .sidebar-header i {{ font-size: 24px; color: #60a5fa; }}
        .sidebar-header h1 {{ font-size: 20px; margin: 0; font-weight: 800; color: #ffffff; letter-spacing: -0.5px; }}
        .sidebar-header h1 span {{ color: #3b82f6; }}
        .menu-category {{
            font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px;
            color: #475569; padding: 16px 24px 6px; font-weight: 700;
        }}
        .menu-item {{
            display: flex; align-items: center; gap: 12px;
            padding: 10px 24px; color: #94a3b8; text-decoration: none;
            font-size: 14px; font-weight: 500; transition: all 0.15s ease;
        }}
        .menu-item:hover {{ background-color: #1e293b; color: #f8fafc; }}
        .menu-item.active {{
            background-color: var(--primary);
            color: #ffffff; border-radius: 6px;
            margin: 2px 12px; padding: 10px 12px;
        }}
        .menu-item i {{ font-size: 15px; width: 20px; text-align: center; }}
        
        .main-wrapper {{
            margin-left: var(--sidebar-width);
            margin-right: var(--ad-width);
            flex-grow: 1;
            display: flex; flex-direction: column;
            min-height: 100vh;
            box-sizing: border-box;
        }}
        
        .hero-section {{
            background-color: #ffffff;
            border-bottom: 1px solid var(--border-color);
            padding: 40px; text-align: left;
        }}
        .hero-container {{ max-width: 800px; margin: 0 auto; }}
        .hero-section h2 {{ margin: 0 0 10px 0; font-size: 28px; font-weight: 800; color: var(--primary-dark); }}
        .hero-section p {{ margin: 0; font-size: 15px; color: var(--text-muted); line-height: 1.6; }}
        
        .content-area {{
            flex-grow: 1;
            display: flex; flex-direction: column; align-items: center; justify-content: center;
            padding: 40px; box-sizing: border-box;
        }}
        .tool-container {{
            background: var(--bg-card);
            padding: 40px; border-radius: 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            border: 1px solid var(--border-color);
            max-width: 460px; width: 100%; text-align: center;
        }}
        .tool-title {{ color: var(--primary-dark); margin: 0 0 8px 0; font-size: 20px; font-weight: 700; }}
        .tool-desc {{ color: var(--text-muted); font-size: 13px; margin-bottom: 24px; line-height: 1.4; }}
        
        .file-dropzone {{
            border: 2px dashed #cbd5e1; background-color: var(--primary-light);
            border-radius: 10px; padding: 30px 20px; margin-bottom: 24px;
            cursor: pointer; transition: all 0.2s ease; position: relative;
        }}
        .file-dropzone:hover {{ border-color: var(--primary); background-color: #e0f2fe; }}
        .file-dropzone i {{ font-size: 36px; color: var(--primary); margin-bottom: 10px; }}
        .file-dropzone p {{ margin: 0; font-size: 14px; color: #334155; font-weight: 500; }}
        .file-dropzone small {{ display: block; margin-top: 5px; color: var(--text-muted); font-size: 11px; }}
        .file-dropzone input[type="file"] {{ position: absolute; top: 0; left: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; }}
        .file-name-display {{ margin-top: 8px; font-size: 13px; color: var(--primary); font-weight: 600; word-break: break-all; }}
        
        button {{
            background-color: var(--primary); color: white; border: none; padding: 12px 24px;
            border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; width: 100%;
            transition: background 0.15s ease;
        }}
        button:hover {{ background-color: var(--accent); }}
        
        .ad-sidebar-right {{
            position: fixed; right: 0; top: 0; width: var(--ad-width); height: 100vh;
            background-color: #f1f5f9; border-left: 1px solid var(--border-color);
            display: flex; align-items: center; justify-content: center; z-index: 90;
        }}
        .ad-banner-bottom {{
            width: 100%; padding: 20px 0; background-color: #f1f5f9;
            border-top: 1px solid var(--border-color); text-align: center; margin-top: auto;
            box-sizing: border-box;
        }}
        .ad-box {{
            background: #e2e8f0; border: 1px dashed #cbd5e1; color: #94a3b8;
            font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px;
            display: flex; align-items: center; justify-content: center; border-radius: 4px;
        }}
        .ad-vertical {{ width: 160px; height: 600px; }}
        .ad-horizontal {{ width: 728px; height: 90px; margin: 0 auto; }}
        
        .loader-container {{ display: none; margin-top: 20px; }}
        .loader {{ border: 3px solid #f3f3f3; border-top: 3px solid var(--primary); border-radius: 50%; width: 24px; height: 24px; animation: spin 1s linear infinite; margin: 0 auto 8px; }}
        @keyframes spin {{ 0% {{ transform: rotate(0deg); }} 100% {{ transform: rotate(360deg); }} }}
        .loader-text {{ font-size: 12px; color: var(--text-muted); }}
    </style>
    <script>
        function handleFileSelect(input, displayId) {{
            var display = document.getElementById(displayId);
            if (input.files && input.files.length > 0) {{
                display.textContent = input.files.length === 1 ? "Selecionado: " + input.files[0].name : input.files.length + " arquivos selecionados";
            }} else {{ display.textContent = ""; }}
        }}
        
        function showLoader() {{
            var loader = document.getElementById('loader');
            var btn = document.getElementById('submit-btn');
            
            loader.style.display = 'block';
            btn.style.display = 'none';
        }}
    </script>
</head>
<body>
    <div class="sidebar">
        <div class="sidebar-header">
            <i class="fa-solid fa-bolt"></i>
            <h1>Converte<span>Ai</span></h1>
        </div>
        <div class="sidebar-menu">
            {sidebar_html}
        </div>
    </div>
    
    <div class="main-wrapper">
        <div class="hero-section">
            <div class="hero-container">
                <h2>Conversão Inteligente de Arquivos</h2>
                <p>Bem-vindo ao Converte Ai. Plataforma rápida, leve e adaptada para nuvem. Limite operacional de até {MAX_FILE_SIZE_MB}MB por processamento para máxima performance.</p>
            </div>
        </div>
        
        <div class="content-area">
            <div class="tool-container">
                {conteudo_pagina}
                <div id="loader" class="loader-container">
                    <div class="loader"></div>
                    <div class="loader-text">Processando documento seguro...</div>
                </div>
            </div>
        </div>
        
        <div class="ad-banner-bottom">
            <div class="ad-box ad-horizontal">Espaço para Anúncio (728x90)</div>
        </div>
    </div>
    <div class="ad-sidebar-right">
        <div class="ad-box ad-vertical">Espaço para Anúncio (160x600)</div>
    </div>
</body>
</html>'''
    return html_retorno

# ==============================================================================
# ROTA DE MAPA DO SITE (SITEMAP.XML PARA SEO)
# ==============================================================================
@app.route('/sitemap.xml', methods=['GET'])
def sitemap():
    """Gera o mapa XML de maneira dinâmica para indexação em indexadores de busca."""
    host = request.host_url.rstrip('/')
    rotas = [
        '/', '/juntar-view', '/dividir-view', '/pdf-to-word-view', 
        '/pdf-to-excel-view', '/pdf-to-pptx-view', '/pdf-to-image-view', 
        '/word-to-pdf-view', '/excel-to-pdf-view'
    ]
    
    xml_linhas = ['<?xml version="1.0" encoding="UTF-8"?>']
    xml_linhas.append('<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">')
    
    for rota in rotas:
        xml_linhas.append('  <url>')
        xml_linhas.append(f'    <loc>{host}{rota}</loc>')
        xml_linhas.append('    <changefreq>weekly</changefreq>')
        xml_linhas.append('    <priority>0.8</priority>')
        xml_linhas.append('  </url>')
        
    xml_linhas.append('</urlset>')
    xml_completo = "\n".join(xml_linhas)
    
    return Response(xml_completo, mimetype='application/xml')

# ==============================================================================
# ROTAS OPERACIONAIS ATIVAS
# ==============================================================================

@app.route('/')
def home():
    conteudo = f'''
        <h3 class="tool-title">Imagem para PDF</h3>
        <p class="tool-desc">Converta suas imagens JPG ou PNG em um documento PDF instantâneo.</p>
        <form action="/convert" method="post" enctype="multipart/form-data" onsubmit="showLoader()">
            <div class="file-dropzone">
                <i class="fa-solid fa-file-image"></i>
                <p>Arraste a imagem ou clique para selecionar</p>
                <small>Limite máximo: {MAX_FILE_SIZE_MB}MB</small>
                <input type="file" name="imagem_usuario" accept=".jpg, .jpeg, .png" required onchange="handleFileSelect(this, 'file-name')">
                <div id="file-name" class="file-name-display"></div>
            </div>
            <button type="submit" id="submit-btn">Converter para PDF</button>
        </form>
    '''
    return layout_base(conteudo, 'imagem-to-pdf')

@app.route('/convert', methods=['POST'])
def convert_file():
    if Image is None: return "Erro: Pillow ausente.", 500
    if 'imagem_usuario' not in request.files: return 'Nenhuma imagem enviada', 400
    arquivo = request.files['imagem_usuario']
    if arquivo.filename == '': return 'Nome inválido', 400
    
    caminho_imagem = os.path.join(UPLOAD_FOLDER, arquivo.filename)
    arquivo.save(caminho_imagem)
    
    try:
        img = Image.open(caminho_imagem)
        img_rgb = img.convert('RGB')
        nome_pdf = os.path.splitext(arquivo.filename)[0] + '.pdf'
        caminho_pdf = os.path.join(UPLOAD_FOLDER, nome_pdf)
        img_rgb.save(caminho_pdf, 'PDF', resolution=100.0)
        return send_file(caminho_pdf, as_attachment=True, download_name=nome_pdf)
    except Exception as e:
        return f'Erro: {str(e)}', 500
    finally:
        if os.path.exists(caminho_imagem): os.remove(caminho_imagem)

@app.route('/juntar-view')
def juntar_view():
    conteudo = f'''
        <h3 class="tool-title">Juntar PDFs</h3>
        <p class="tool-desc">Combine dois ou mais arquivos em um único documento estruturado.</p>
        <form action="/juntar" method="post" enctype="multipart/form-data" onsubmit="showLoader()">
            <div class="file-dropzone">
                <i class="fa-solid fa-file-pdf"></i>
                <p>Selecione os múltiplos arquivos PDF</p>
                <small>A soma total não deve passar de {MAX_FILE_SIZE_MB}MB</small>
                <input type="file" name="arquivos_pdf" accept=".pdf" multiple required onchange="handleFileSelect(this, 'file-name')">
                <div id="file-name" class="file-name-display"></div>
            </div>
            <button type="submit" id="submit-btn">Juntar PDFs</button>
        </form>
    '''
    return layout_base(conteudo, 'juntar')

@app.route('/juntar', methods=['POST'])
def juntar_pdfs():
    if PdfWriter is None: return "Erro: pypdf ausente.", 500
    arquivos = request.files.getlist('arquivos_pdf')
    if not arquivos or arquivos[0].filename == '': return 'Nenhum arquivo selecionado', 400
    
    merger = PdfWriter()
    caminhos_temporarios = []
    try:
        for arquivo in arquivos:
            caminho_temp = os.path.join(UPLOAD_FOLDER, arquivo.filename)
            arquivo.save(caminho_temp)
            caminhos_temporarios.append(caminho_temp)
            merger.append(caminho_temp)
        
        nome_saida = "pdfs_combinados.pdf"
        caminho_saida = os.path.join(UPLOAD_FOLDER, nome_saida)
        with open(caminho_saida, "wb") as f:
            merger.write(f)
        merger.close()
        return send_file(caminho_saida, as_attachment=True, download_name=nome_saida)
    except Exception as e:
        return f'Erro: {str(e)}', 500
    finally:
        for c in caminhos_temporarios:
            if os.path.exists(c): os.remove(c)

@app.route('/dividir-view')
def dividir_view():
    conteudo = f'''
        <h3 class="tool-title">Dividir PDF</h3>
        <p class="tool-desc">Separe todas as páginas do seu documento PDF em arquivos individuais dentro de um ZIP.</p>
        <form action="/dividir" method="post" enctype="multipart/form-data" onsubmit="showLoader()">
            <div class="file-dropzone">
                <i class="fa-solid fa-scissors"></i>
                <p>Escolha o arquivo PDF para fatiar</p>
                <small>Limite máximo: {MAX_FILE_SIZE_MB}MB</small>
                <input type="file" name="arquivo_pdf" accept=".pdf" required onchange="handleFileSelect(this, 'file-name')">
                <div id="file-name" class="file-name-display"></div>
            </div>
            <button type="submit" id="submit-btn">Dividir PDF</button>
        </form>
    '''
    return layout_base(conteudo, 'dividir')

@app.route('/dividir', methods=['POST'])
def dividir_pdf():
    if PdfReader is None: return "Erro: pypdf ausente.", 500
    arquivo = request.files['arquivo_pdf']
    if arquivo.filename == '': return 'Arquivo inválido', 400
    
    caminho_pdf = os.path.join(UPLOAD_FOLDER, arquivo.filename)
    arquivo.save(caminho_pdf)
    
    try:
        reader = PdfReader(caminho_pdf)
        nome_base = os.path.splitext(arquivo.filename)[0]
        nome_zip = f"{nome_base}_dividido.zip"
        caminho_zip = os.path.join(UPLOAD_FOLDER, nome_zip)
        
        with zipfile.ZipFile(caminho_zip, 'w') as z:
            for idx, pagina in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(pagina)
                nome_pagina = f"{nome_base}_pag_{idx+1}.pdf"
                caminho_pagina = os.path.join(UPLOAD_FOLDER, nome_pagina)
                with open(caminho_pagina, "wb") as out:
                    writer.write(out)
                z.write(caminho_pagina, nome_pagina)
                os.remove(caminho_pagina)
                
        return send_file(caminho_zip, as_attachment=True, download_name=nome_zip)
    except Exception as e:
        return f'Erro: {str(e)}', 500
    finally:
        if os.path.exists(caminho_pdf): os.remove(caminho_pdf)

@app.route('/pdf-to-word-view')
def pdf_to_word_view():
    conteudo = f'''
        <h3 class="tool-title">PDF para Word</h3>
        <p class="tool-desc">Conversor rápido de arquivos estruturados PDF para documentos editáveis (.docx).</p>
        <form action="/pdf-to-word" method="post" enctype="multipart/form-data" onsubmit="showLoader()">
            <div class="file-dropzone">
                <i class="fa-solid fa-file-word"></i>
                <p>Selecione o arquivo PDF</p>
                <small>Limite máximo: {MAX_FILE_SIZE_MB}MB</small>
                <input type="file" name="arquivo_pdf" accept=".pdf" required onchange="handleFileSelect(this, 'file-name')">
                <div id="file-name" class="file-name-display"></div>
            </div>
            <button type="submit" id="submit-btn">Converter para Word</button>
        </form>
    '''
    return layout_base(conteudo, 'pdf-to-word')

@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if Converter is None: return "Erro: pdf2docx ausente.", 500
    arquivo = request.files['arquivo_pdf']
    if arquivo.filename == '': return 'Arquivo inválido', 400
    
    caminho_pdf = os.path.join(UPLOAD_FOLDER, arquivo.filename)
    arquivo.save(caminho_pdf)
    nome_docx = os.path.splitext(arquivo.filename)[0] + '.docx'
    caminho_docx = os.path.join(UPLOAD_FOLDER, nome_docx)
    
    try:
        cv = Converter(caminho_pdf)
        cv.convert(caminho_docx, start=0, end=None)
        cv.close()
        return send_file(caminho_docx, as_attachment=True, download_name=nome_docx)
    except Exception as e:
        return f'Erro na conversão: {str(e)}', 500
    finally:
        if os.path.exists(caminho_pdf): os.remove(caminho_pdf)

@app.route('/pdf-to-excel-view')
def pdf_to_excel_view():
    conteudo = f'''
        <h3 class="tool-title">PDF para Excel</h3>
        <p class="tool-desc">Extraia as tabelas do PDF direto para planilhas do Excel (.xlsx).</p>
        <form action="/pdf-to-excel" method="post" enctype="multipart/form-data" onsubmit="showLoader()">
            <div class="file-dropzone">
                <i class="fa-solid fa-file-excel"></i>
                <p>Selecione o arquivo PDF</p>
                <small>Limite máximo: {MAX_FILE_SIZE_MB}MB</small>
                <input type="file" name="arquivo_pdf" accept=".pdf" required onchange="handleFileSelect(this, 'file-name')">
                <div id="file-name" class="file-name-display"></div>
            </div>
            <button type="submit" id="submit-btn">Converter para Excel</button>
        </form>
    '''
    return layout_base(conteudo, 'pdf-to-excel')

@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if pdfplumber is None or Workbook is None: return "Erro: Bibliotecas ausentes.", 500
    arquivo = request.files['arquivo_pdf']
    if arquivo.filename == '': return 'Arquivo inválido', 400
    
    caminho_pdf = os.path.join(UPLOAD_FOLDER, arquivo.filename)
    arquivo.save(caminho_pdf)
    nome_xlsx = os.path.splitext(arquivo.filename)[0] + '.xlsx'
    caminho_xlsx = os.path.join(UPLOAD_FOLDER, nome_xlsx)
    
    try:
        wb = Workbook()
        wb.remove(wb.active)
        with pdfplumber.open(caminho_pdf) as pdf:
            for idx, pagina in enumerate(pdf.pages):
                ws = wb.create_sheet(title=f"Pagina {idx+1}")
                tabelas = pagina.extract_tables()
                if tabelas:
                    for tabela in tabelas:
                        for linha in tabela:
                            linha_limpa = [str(c) if c is not None else "" for c in linha]
                            ws.append(linha_limpa)
                        ws.append([])
                else:
                    ws.append(["Nenhuma tabela estruturada encontrada nesta página."])
        wb.save(caminho_xlsx)
        return send_file(caminho_xlsx, as_attachment=True, download_name=nome_xlsx)
    except Exception as e:
        return f'Erro: {str(e)}', 500
    finally:
        if os.path.exists(caminho_pdf): os.remove(caminho_pdf)

@app.route('/pdf-to-pptx-view')
def pdf_to_pptx_view():
    conteudo = f'''
        <h3 class="tool-title">PDF para PPTX</h3>
        <p class="tool-desc">Gere blocos de slides do PowerPoint a partir do PDF.</p>
        <form action="/pdf-to-pptx" method="post" enctype="multipart/form-data" onsubmit="showLoader()">
            <div class="file-dropzone">
                <i class="fa-solid fa-file-powerpoint"></i>
                <p>Selecione o arquivo PDF</p>
                <small>Limite máximo: {MAX_FILE_SIZE_MB}MB</small>
                <input type="file" name="arquivo_pdf" accept=".pdf" required onchange="handleFileSelect(this, 'file-name')">
                <div id="file-name" class="file-name-display"></div>
            </div>
            <button type="submit" id="submit-btn">Converter para PPTX</button>
        </form>
    '''
    return layout_base(conteudo, 'pdf-to-pptx')

@app.route('/pdf-to-pptx', methods=['POST'])
def pdf_to_pptx():
    if fitz is None or Presentation is None: return "Erro: Bibliotecas ausentes.", 500
    arquivo = request.files['arquivo_pdf']
    if arquivo.filename == '': return 'Arquivo inválido', 400
    
    caminho_pdf = os.path.join(UPLOAD_FOLDER, arquivo.filename)
    arquivo.save(caminho_pdf)
    nome_pptx = os.path.splitext(arquivo.filename)[0] + '.pptx'
    caminho_pptx = os.path.join(UPLOAD_FOLDER, nome_pptx)
    
    try:
        prs = Presentation()
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        doc = fitz.open(caminho_pdf)
        for idx in range(len(doc)):
            pagina = doc.load_page(idx)
            pix = pagina.get_pixmap(dpi=150)
            caminho_temp_img = os.path.join(UPLOAD_FOLDER, f"temp_slide_{idx}.png")
            pix.save(caminho_temp_img)
            blank_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(blank_layout)
            slide.shapes.add_picture(caminho_temp_img, 0, 0, width=prs.slide_width, height=prs.slide_height)
            os.remove(caminho_temp_img)
        prs.save(caminho_pptx)
        return send_file(caminho_pptx, as_attachment=True, download_name=nome_pptx)
    except Exception as e:
        return f'Erro: {str(e)}', 500
    finally:
        if os.path.exists(caminho_pdf): os.remove(caminho_pdf)

@app.route('/pdf-to-image-view')
def pdf_to_image_view():
    conteudo = f'''
        <h3 class="tool-title">PDF para Imagem</h3>
        <p class="tool-desc">Fatie as folhas do PDF em arquivos de imagens avulsas.</p>
        <form action="/pdf-to-image" method="post" enctype="multipart/form-data" onsubmit="showLoader()">
            <div class="file-dropzone">
                <i class="fa-solid fa-images"></i>
                <p>Selecione o arquivo PDF original</p>
                <small>Limite máximo: {MAX_FILE_SIZE_MB}MB</small>
                <input type="file" name="arquivo_pdf" accept=".pdf" required onchange="handleFileSelect(this, 'file-name')">
                <div id="file-name" class="file-name-display"></div>
            </div>
            <button type="submit" id="submit-btn">Converter para Imagem</button>
        </form>
    '''
    return layout_base(conteudo, 'pdf-to-image')

@app.route('/pdf-to-image', methods=['POST'])
def pdf_to_image():
    if fitz is None: return "Erro: PyMuPDF ausente.", 500
    arquivo = request.files['arquivo_pdf']
    if arquivo.filename == '': return 'Arquivo inválido', 400
    
    caminho_pdf = os.path.join(UPLOAD_FOLDER, arquivo.filename)
    arquivo.save(caminho_pdf)
    nome_base = os.path.splitext(arquivo.filename)[0]
    
    try:
        doc = fitz.open(caminho_pdf)
        if len(doc) == 1:
            pix = doc[0].get_pixmap(dpi=150)
            nome_img = f"{nome_base}.png"
            caminho_img = os.path.join(UPLOAD_FOLDER, nome_img)
            pix.save(caminho_img)
            return send_file(caminho_img, as_attachment=True, download_name=nome_img)
        else:
            nome_zip = f"{nome_base}_imagens.zip"
            caminho_zip = os.path.join(UPLOAD_FOLDER, nome_zip)
            with zipfile.ZipFile(caminho_zip, 'w') as z:
                for idx, pagina in enumerate(doc):
                    pix = pagina.get_pixmap(dpi=150)
                    nome_img = f"{nome_base}_pag_{idx+1}.png"
                    caminho_img = os.path.join(UPLOAD_FOLDER, nome_img)
                    pix.save(caminho_img)
                    z.write(caminho_img, nome_img)
                    os.remove(caminho_img)
            return send_file(caminho_zip, as_attachment=True, download_name=nome_zip)
    except Exception as e:
        return f'Erro: {str(e)}', 500
    finally:
        if os.path.exists(caminho_pdf): os.remove(caminho_pdf)

@app.route('/word-to-pdf-view')
def word_to_pdf_view():
    conteudo = f'''
        <h3 class="tool-title">Word para PDF</h3>
        <p class="tool-desc">Gere PDFs fiéis a partir de documentos de texto do Word (.docx).</p>
        <form action="/word-to-pdf" method="post" enctype="multipart/form-data" onsubmit="showLoader()">
            <div class="file-dropzone">
                <i class="fa-solid fa-file-pdf"></i>
                <p>Selecione o documento do Word (.docx)</p>
                <small>Limite máximo: {MAX_FILE_SIZE_MB}MB</small>
                <input type="file" name="arquivo_docx" accept=".docx" required onchange="handleFileSelect(this, 'file-name')">
                <div id="file-name" class="file-name-display"></div>
            </div>
            <button type="submit" id="submit-btn">Converter para PDF</button>
        </form>
    '''
    return layout_base(conteudo, 'word-to-pdf')

@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf():
    if Document is None or SimpleDocTemplate is None: return "Erro: Bibliotecas ausentes.", 500
    arquivo = request.files['arquivo_docx']
    if arquivo.filename == '': return 'Arquivo inválido', 400
    
    caminho_docx = os.path.join(UPLOAD_FOLDER, arquivo.filename)
    arquivo.save(caminho_docx)
    nome_pdf = os.path.splitext(arquivo.filename)[0] + '.pdf'
    caminho_pdf = os.path.join(UPLOAD_FOLDER, nome_pdf)
    
    try:
        doc_word = Document(caminho_docx)
        pdf_render = SimpleDocTemplate(caminho_pdf, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        for paragrafo in doc_word.paragraphs:
            if paragrafo.text.strip():
                story.append(Paragraph(paragrafo.text, styles['Normal']))
                story.append(Spacer(1, 10))
        pdf_render.build(story)
        return send_file(caminho_pdf, as_attachment=True, download_name=nome_pdf)
    except Exception as e:
        return f'Erro: {str(e)}', 500
    finally:
        if os.path.exists(caminho_docx): os.remove(caminho_docx)

@app.route('/excel-to-pdf-view')
def excel_to_pdf_view():
    conteudo = f'''
        <h3 class="tool-title">Excel para PDF</h3>
        <p class="tool-desc">Converta suas planilhas do Excel (.xlsx) diretamente em um documento PDF.</p>
        <form action="/excel-to-pdf" method="post" enctype="multipart/form-data" onsubmit="showLoader()">
            <div class="file-dropzone">
                <i class="fa-solid fa-file-excel"></i>
                <p>Selecione a planilha do Excel (.xlsx)</p>
                <small>Limite máximo: {MAX_FILE_SIZE_MB}MB</small>
                <input type="file" name="arquivo_xlsx" accept=".xlsx" required onchange="handleFileSelect(this, 'file-name')">
                <div id="file-name" class="file-name-display"></div>
            </div>
            <button type="submit" id="submit-btn">Converter para PDF</button>
        </form>
    '''
    return layout_base(conteudo, 'excel-to-pdf')

@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf():
    if load_workbook is None or SimpleDocTemplate is None: return "Erro: Bibliotecas ausentes.", 500
    arquivo = request.files['arquivo_xlsx']
    if arquivo.filename == '': return 'Arquivo inválido', 400
    
    caminho_xlsx = os.path.join(UPLOAD_FOLDER, arquivo.filename)
    arquivo.save(caminho_xlsx)
    nome_pdf = os.path.splitext(arquivo.filename)[0] + '.pdf'
    caminho_pdf = os.path.join(UPLOAD_FOLDER, nome_pdf)
    
    try:
        wb = load_workbook(caminho_xlsx, data_only=True)
        ws = wb.active
        pdf_render = SimpleDocTemplate(caminho_pdf, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        dados_tabela = []
        for linha in ws.iter_rows(values_only=True):
            linha_limpa = [str(celula) if celula is not None else "" for celula in linha]
            if any(linha_limpa):  # Ignora linhas totalmente vazias
                dados_tabela.append(linha_limpa)
                
        if dados_tabela:
            t = Table(dados_tabela)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.grey),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('BOTTOMPADDING', (0,0), (-1,0), 6),
                ('GRID', (0,0), (-1,-1), 1, colors.lightgrey)
            ]))
            story.append(t)
        else:
            story.append(Paragraph("A planilha selecionada está vazia.", styles['Normal']))
            
        pdf_render.build(story)
        return send_file(caminho_pdf, as_attachment=True, download_name=nome_pdf)
    except Exception as e:
        return f'Erro: {str(e)}', 500
    finally:
        if os.path.exists(caminho_xlsx): os.remove(caminho_xlsx)

if __name__ == '__main__':
    app.run(debug=True)
