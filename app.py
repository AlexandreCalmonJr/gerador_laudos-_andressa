from flask import Flask, render_template, request, send_from_directory, flash, redirect, url_for
from docx import Document
from docx.shared import Inches
import os
import uuid # Para gerar nomes de arquivo únicos

# --- CONFIGURAÇÃO DO FLASK ---
app = Flask(__name__)
app.secret_key = 'super_secret_key' # Necessário para usar 'flash messages'

# Configuração das pastas
UPLOAD_FOLDER = 'uploads'
GENERATED_FOLDER = 'gerados'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['GENERATED_FOLDER'] = GENERATED_FOLDER


# --- LÓGICA DE GERAÇÃO DO DOCUMENTO (A MESMA DE ANTES) ---
def gerar_documento(dados, nome_arquivo_saida):
    try:
        doc = Document('Vistoria_Modelo.docx')

        # Substitui texto em parágrafos e tabelas
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
        
        # Insere imagens
        for p in doc.paragraphs:
            for comodo_id, imagens in dados['imagens'].items():
                marcador_imagem = f"{{{{IMAGENS_{comodo_id.upper()}}}}}"
                if marcador_imagem in p.text:
                    p.text = ''
                    for img_path in imagens:
                        p.add_run().add_picture(img_path, width=Inches(3.0))

        caminho_completo = os.path.join(app.config['GENERATED_FOLDER'], nome_arquivo_saida)
        doc.save(caminho_completo)
        return True
    except Exception as e:
        print(f"Erro ao gerar o documento: {e}")
        return False


# --- ROTAS DA APLICAÇÃO WEB ---
@app.route('/')
def index():
    # Renderiza o formulário HTML
    return render_template('form.html')

@app.route('/gerar', methods=['POST'])
def gerar_laudo():
    # 1. Coletar dados do formulário
    dados_texto = {f"{{{{{key}}}}}": value for key, value in request.form.items()}
    dados_imagens = {}
    
    # 2. Processar e salvar as imagens enviadas
    comodos = ["externa", "sala", "cozinha", "banheiro_social", "quarto1", "quarto2", 
               "servico", "gourmet", "banheiro_edicula", "laterais"]
    
    for comodo in comodos:
        arquivos = request.files.getlist(f'imagens_{comodo}')
        if arquivos and arquivos[0].filename != '':
            caminhos_salvos = []
            for arquivo in arquivos:
                # Garante um nome de arquivo seguro e único
                nome_seguro = f"{uuid.uuid4()}_{arquivo.filename}"
                caminho = os.path.join(app.config['UPLOAD_FOLDER'], nome_seguro)
                arquivo.save(caminho)
                caminhos_salvos.append(caminho)
            dados_imagens[comodo] = caminhos_salvos

    dados_completos = {
        'texto': dados_texto,
        'imagens': dados_imagens
    }

    # 3. Gerar o nome do arquivo e chamar a função do docx
    nome_locatario = request.form.get('LOCATARIO_NOME_1', 'laudo').replace(' ', '_')
    nome_arquivo_final = f"Laudo_Vistoria_{nome_locatario}.docx"
    
    sucesso = gerar_documento(dados_completos, nome_arquivo_final)

    # 4. Limpar imagens temporárias
    for comodo in dados_imagens.values():
        for path in comodo:
            os.remove(path)
            
    if sucesso:
        # Redireciona para a página de resultado com o nome do arquivo
        return redirect(url_for('resultado', filename=nome_arquivo_final))
    else:
        flash("Ocorreu um erro ao gerar o documento.")
        return redirect(url_for('index'))

@app.route('/resultado/<filename>')
def resultado(filename):
    # Renderiza a página de sucesso
    return render_template('resultado.html', filename=filename)

@app.route('/download/<filename>')
def download_file(filename):
    # Oferece o arquivo gerado para download
    return send_from_directory(app.config["GENERATED_FOLDER"], filename, as_attachment=True)


if __name__ == '__main__':
    # Roda a aplicação em modo de desenvolvimento
    app.run(debug=True)