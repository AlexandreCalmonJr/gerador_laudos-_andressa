from flask import Flask, render_template, request, send_from_directory, flash, redirect, url_for
from docx import Document
from docx.shared import Inches
import os
import uuid
import logging
import time

# --- 1. MELHORIA: CONFIGURAÇÃO E LOGGING ---
# Configuração de logging para registrar erros em um arquivo
logging.basicConfig(
    filename='app_errors.log', 
    level=logging.ERROR, 
    format='%(asctime)s %(levelname)s %(name)s %(threadName)s : %(message)s'
)

# --- CONFIGURAÇÃO DO FLASK ---
app = Flask(__name__)
app.secret_key = 'uma-chave-secreta-muito-forte-e-dificil' # Chave mais segura

# Configuração das pastas
UPLOAD_FOLDER = 'uploads'
GENERATED_FOLDER = 'gerados'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['GENERATED_FOLDER'] = GENERATED_FOLDER
app.config['MAX_FILE_AGE_SECONDS'] = 24 * 3600 # 24 horas para limpeza

# --- LÓGICA DE GERAÇÃO DO DOCUMENTO (SEM ALTERAÇÃO) ---
def gerar_documento(dados, nome_arquivo_saida):
    try:
        doc = Document('Vistoria_Modelo.docx')

        all_paragraphs = list(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_paragraphs.extend(cell.paragraphs)

        for p in all_paragraphs:
            for key, value in dados['texto'].items():
                if key in p.text:
                    inline = p.runs
                    for i in range(len(inline)):
                        if key in inline[i].text:
                            p.runs[i].text = p.runs[i].text.replace(key, value)
        
        for p in doc.paragraphs:
            for comodo_id, imagens in dados['imagens'].items():
                marcador_imagem = f"{{{{IMAGENS_{comodo_id.upper()}}}}}"
                if marcador_imagem in p.text:
                    p.text = ''
                    for img_path in imagens:
                        try:
                            p.add_run().add_picture(img_path, width=Inches(3.0))
                        except Exception as e:
                            app.logger.error(f"Erro ao adicionar imagem {img_path}: {e}")
                            # Pula a imagem corrompida mas continua o processo
                            continue

        caminho_completo = os.path.join(app.config['GENERATED_FOLDER'], nome_arquivo_saida)
        doc.save(caminho_completo)
        return True
    except FileNotFoundError:
        app.logger.error("O template 'Vistoria_Modelo.docx' não foi encontrado.")
        flash("Erro Crítico: O arquivo modelo da vistoria não foi encontrado no servidor.")
        return False
    except Exception as e:
        app.logger.error(f"Erro desconhecido ao gerar o documento: {e}")
        flash("Ocorreu um erro inesperado ao gerar o documento.")
        return False

# --- 2. MELHORIA: FUNÇÃO DE LIMPEZA DE ARQUIVOS ANTIGOS ---
def limpar_arquivos_antigos():
    """Apaga arquivos nas pastas 'uploads' e 'gerados' mais antigos que o tempo definido."""
    now = time.time()
    max_age = app.config['MAX_FILE_AGE_SECONDS']
    
    for folder in [app.config['UPLOAD_FOLDER'], app.config['GENERATED_FOLDER']]:
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            if os.path.isfile(file_path):
                try:
                    file_age = now - os.path.getmtime(file_path)
                    if file_age > max_age:
                        os.remove(file_path)
                        app.logger.info(f"Arquivo antigo removido: {file_path}")
                except Exception as e:
                    app.logger.error(f"Erro ao tentar remover arquivo antigo {file_path}: {e}")

# --- ROTAS DA APLICAÇÃO WEB ---
@app.route('/')
def index():
    # Roda a limpeza antes de carregar a página principal
    limpar_arquivos_antigos()
    return render_template('form.html')

@app.route('/gerar', methods=['POST'])
def gerar_laudo():
    # --- 3. MELHORIA: VALIDAÇÃO DOS DADOS DE ENTRADA ---
    if 'LOCATARIO_NOME_1' not in request.form or not request.form['LOCATARIO_NOME_1'].strip():
        flash("O campo 'Nome do Locatário 1' é obrigatório para gerar o arquivo.")
        return redirect(url_for('index'))

    dados_texto = {f"{{{{{key}}}}}": value for key, value in request.form.items()}
    dados_imagens = {}
    
    comodos = ["externa", "sala", "cozinha", "banheiro_social", "quarto1", "quarto2", 
               "servico", "fundos", "gourmet", "banheiro_edicula", "laterais"]
    
    for comodo in comodos:
        arquivos = request.files.getlist(f'imagens_{comodo}')
        if arquivos and arquivos[0].filename != '':
            caminhos_salvos = []
            for arquivo in arquivos:
                nome_seguro = f"{uuid.uuid4()}_{arquivo.filename}"
                caminho = os.path.join(app.config['UPLOAD_FOLDER'], nome_seguro)
                arquivo.save(caminho)
                caminhos_salvos.append(caminho)
            dados_imagens[comodo] = caminhos_salvos

    dados_completos = {'texto': dados_texto, 'imagens': dados_imagens}

    # Nome do arquivo com ID único para evitar sobrescrever
    nome_locatario = request.form.get('LOCATARIO_NOME_1').replace(' ', '_')
    unique_id = str(uuid.uuid4())[:6]
    nome_arquivo_final = f"Laudo_Vistoria_{nome_locatario}_{unique_id}.docx"
    
    sucesso = gerar_documento(dados_completos, nome_arquivo_final)

    # Limpa apenas os uploads da sessão atual
    for comodo_imgs in dados_imagens.values():
        for path in comodo_imgs:
            if os.path.exists(path):
                os.remove(path)
            
    if sucesso:
        return redirect(url_for('resultado', filename=nome_arquivo_final))
    else:
        # A função gerar_documento já terá usado flash() para a mensagem de erro
        return redirect(url_for('index'))

@app.route('/resultado/<filename>')
def resultado(filename):
    return render_template('resultado.html', filename=filename)

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config["GENERATED_FOLDER"], filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)