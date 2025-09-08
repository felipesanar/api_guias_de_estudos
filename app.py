import openpyxl
from flask import Flask, jsonify, request
from collections import defaultdict
import os
import glob
from flask_caching import Cache

app = Flask(__name__)

# Configuração de cache
cache = Cache(app, config={'CACHE_TYPE': 'simple'})

# Estrutura global para armazenar os dados processados
dados_ies = {}

def processar_arquivo_excel(nome_arquivo):
    """
    Processa o arquivo Excel usando openpyxl (sem pandas)
    """
    try:
        # Carregar o arquivo Excel
        wb = openpyxl.load_workbook(nome_arquivo, data_only=True)
        ies_abas = wb.sheetnames
        
        dados_processados = {}
        
        for ies in ies_abas:
            ws = wb[ies]
            
            # Encontrar cabeçalhos
            headers = []
            for col in range(1, ws.max_column + 1):
                cell_value = ws.cell(row=1, column=col).value
                if cell_value:
                    headers.append(str(cell_value).strip())
                else:
                    headers.append(f"coluna_{col}")
            
            # Estrutura hierárquica para os dados desta IES
            ies_estruturada = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
            
            # Processar linhas (começando da linha 2)
            for row in range(2, ws.max_row + 1):
                try:
                    # Extrair valores das células
                    valores = {}
                    for col in range(1, ws.max_column + 1):
                        if col - 1 < len(headers):
                            cell_value = ws.cell(row=row, column=col).value
                            valores[headers[col - 1]] = str(cell_value).strip() if cell_value is not None else ""
                    
                    semestre = valores.get('Semestre', '').strip()
                    materia = valores.get('Materia', '').strip()
                    tema = valores.get('Tema', '').strip()
                    subtema = valores.get('Subtema', '').strip()
                    aula = valores.get('Aula', '').strip()
                    link_aula = valores.get('Link Aula', '').strip()
                    link_pdf = valores.get('Link PDF', '').strip()
                    link_quiz = valores.get('Link Quiz', '').strip()
                    
                    # Pular linhas com dados essenciais faltantes
                    if not all([semestre, materia, tema, subtema, aula]):
                        continue
                    
                    # Criar objeto de aula
                    aula_obj = {
                        'nome': aula,
                        'link_aula': link_aula if link_aula else None,
                        'link_pdf': link_pdf if link_pdf else None,
                        'link_quiz': link_quiz if link_quiz else None
                    }
                    
                    # Adicionar à estrutura hierárquica
                    ies_estruturada[semestre][materia][tema].append({
                        'subtema': subtema,
                        'aula': aula_obj
                    })
                    
                except Exception as e:
                    print(f"Erro ao processar linha {row} na IES {ies}: {e}")
                    continue
            
            # Converter defaultdict para dict regular e organizar a estrutura
            ies_estruturada_final = {}
            for semestre, materias in ies_estruturada.items():
                semestre_dict = {}
                for materia, temas in materias.items():
                    materia_dict = {}
                    for tema, subtemas in temas.items():
                        materia_dict[tema] = subtemas
                    semestre_dict[materia] = materia_dict
                ies_estruturada_final[semestre] = semestre_dict
            
            dados_processados[ies] = ies_estruturada_final
            
        return dados_processados
        
    except Exception as e:
        print(f"Erro ao processar arquivo Excel: {e}")
        return {}

def formatar_resposta_api(dados_ies, especifica_ies=None, semestre=None):
    """
    Formata os dados para a resposta da API conforme a hierarquia solicitada
    """
    resultado = {}
    
    # Filtrar por IES específica se fornecida
    ies_para_processar = {especifica_ies: dados_ies[especifica_ies]} if especifica_ies and especifica_ies in dados_ies else dados_ies
    
    for ies_nome, ies_dados in ies_para_processar.items():
        ies_resultado = {}
        
        # Filtrar por semestre se fornecido
        if semestre and semestre in ies_dados:
            semestres_para_processar = {semestre: ies_dados[semestre]}
        else:
            semestres_para_processar = ies_dados
        
        for semestre_nome, semestre_dados in semestres_para_processar.items():
            semestre_resultado = []
            
            for materia, temas in semestre_dados.items():
                materia_dict = {
                    'materia': materia,
                    'temas': []
                }
                
                for tema, subtemas in temas.items():
                    tema_dict = {
                        'tema': tema,
                        'subtemas': []
                    }
                    
                    for subtema_info in subtemas:
                        subtema_dict = {
                            'subtema': subtema_info['subtema'],
                            'aulas': [subtema_info['aula']]
                        }
                        tema_dict['subtemas'].append(subtema_dict)
                    
                    materia_dict['temas'].append(tema_dict)
                
                semestre_resultado.append(materia_dict)
            
            ies_resultado[semestre_nome] = semestre_resultado
        
        resultado[ies_nome] = ies_resultado
    
    return resultado

# Endpoint raiz com informações da API
@app.route('/')
def home():
    # Gerar HTML com links clicáveis para todas as IES
    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>API de Conteúdos de IES</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; }
            h1 { color: #333; }
            ul { list-style-type: none; padding: 0; }
            li { margin: 10px 0; }
            a { 
                text-decoration: none; 
                color: #0366d6; 
                font-weight: bold; 
                padding: 8px 12px;
                border: 1px solid #0366d6;
                border-radius: 4px;
                display: inline-block;
            }
            a:hover { background-color: #f0f7ff; }
            .endpoint { margin-top: 30px; }
            .ies-list { margin-top: 20px; }
        </style>
    </head>
    <body>
        <h1>API de Conteúdos de IES</h1>
        <p>API para acesso aos conteúdos das IES a partir de arquivo Excel.</p>
        
        <div class="endpoint">
            <h2>Endpoints disponíveis:</h2>
            <ul>
                <li><a href="/listar-ies">/listar-ies</a> - Lista todas as IES disponíveis</li>
                <li><code>/&lt;nome_ies&gt;</code> - Todos os conteúdos de uma IES</li>
                <li><code>/&lt;nome_ies&gt;/&lt;semestre&gt;</code> - Conteúdos de uma IES por semestre</li>
            </ul>
        </div>
    """
    
    # Adicionar lista de IES com links clicáveis se os dados foram carregados
    if dados_ies:
        html += """
        <div class="ies-list">
            <h2>IES Disponíveis (clique para acessar):</h2>
            <ul>
        """
        
        for ies in sorted(dados_ies.keys()):
            html += f'<li><a href="/{ies}">{ies}</a></li>'
        
        html += """
            </ul>
        </div>
        """
    else:
        html += """
        <div class="warning">
            <h2 style="color: #d63636;">Atenção: Nenhum arquivo Excel encontrado</h2>
            <p>Por favor, faça upload de um arquivo Excel com a estrutura correta ou use o endpoint abaixo para recarregar os dados:</p>
            <form action="/recarregar-dados" method="POST">
                <button type="submit">Recarregar Dados</button>
            </form>
        </div>
        """
    
    html += """
    </body>
    </html>
    """
    
    return html

# Endpoint para listar todas as IES disponíveis
@app.route('/listar-ies')
@cache.cached(timeout=300)
def listar_ies():
    """Retorna a lista de todas as IES disponíveis na API"""
    ies_disponiveis = list(dados_ies.keys())
    return jsonify({"ies_disponiveis": ies_disponiveis})

# Rota dinâmica para acessar conteúdos por IES
@app.route('/<string:nome_ies>')
@cache.cached(timeout=300)
def get_conteudos_ies(nome_ies):
    """Retorna todos os conteúdos de uma IES específica"""
    if nome_ies not in dados_ies:
        return jsonify({"error": f"IES '{nome_ies}' não encontrada"}), 404
    
    dados_formatados = formatar_resposta_api(dados_ies, nome_ies)
    
    # Adicionar links para semestres se quisermos uma visualização HTML
    if request.args.get('format') == 'html':
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Conteúdos da IES {nome_ies}</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 40px; }}
                h1 {{ color: #333; }}
                ul {{ list-style-type: none; padding: 0; }}
                li {{ margin: 10px 0; }}
                a {{ 
                    text-decoration: none; 
                    color: #0366d6; 
                    font-weight: bold; 
                    padding: 8px 12px;
                    border: 1px solid #0366d6;
                    border-radius: 4px;
                    display: inline-block;
                }}
                a:hover {{ background-color: #f0f7ff; }}
                .back-link {{ margin-top: 20px; }}
            </style>
        </head>
        <body>
            <h1>Conteúdos da IES {nome_ies}</h1>
            <p><a href="/">← Voltar para página inicial</a></p>
            
            <h2>Semestres disponíveis:</h2>
            <ul>
        """
        
        for semestre in sorted(dados_ies[nome_ies].keys()):
            html += f'<li><a href="/{nome_ies}/{semestre}">Semestre {semestre}</a></li>'
        
        html += """
            </ul>
        </body>
        </html>
        """
        return html
    
    return jsonify(dados_formatados)

# Rota dinâmica para acessar conteúdos por IES e semestre
@app.route('/<string:nome_ies>/<string:semestre>')
@cache.cached(timeout=300)
def get_conteudos_ies_semestre(nome_ies, semestre):
    """Retorna os conteúdos de uma IES específica filtrados por semestre"""
    if nome_ies not in dados_ies:
        return jsonify({"error": f"IES '{nome_ies}' não encontrada"}), 404
    
    if semestre not in dados_ies[nome_ies]:
        return jsonify({"error": f"Semestre '{semestre}' não encontrado para a IES '{nome_ies}'"}), 404
    
    dados_formatados = formatar_resposta_api(dados_ies, nome_ies, semestre)
    
    # Adicionar visualização HTML se solicitado
    if request.args.get('format') == 'html':
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Conteúdos da IES {nome_ies} - Semestre {semestre}</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 40px; }}
                h1, h2 {{ color: #333; }}
                .back-link {{ margin-bottom: 20px; }}
                a {{ 
                    text-decoration: none; 
                    color: #0366d6; 
                }}
                a:hover {{ text-decoration: underline; }}
                .materia {{ margin-top: 20px; border-left: 4px solid #0366d6; padding-left: 15px; }}
                .tema {{ margin-left: 20px; }}
                .subtema {{ margin-left: 40px; }}
                .aula {{ margin-left: 60px; }}
            </style>
        </head>
        <body>
            <div class="back-link">
                <a href="/{nome_ies}">← Voltar para {nome_ies}</a> | 
                <a href="/">Página inicial</a>
            </div>
            
            <h1>IES {nome_ies} - Semestre {semestre}</h1>
        """
        
        for materia in dados_formatados[nome_ies][semestre]:
            html += f'<div class="materia"><h2>{materia["materia"]}</h2></div>'
            
            for tema in materia["temas"]:
                html += f'<div class="tema"><h3>{tema["tema"]}</h3></div>'
                
                for subtema in tema["subtemas"]:
                    html += f'<div class="subtema"><h4>{subtema["subtema"]}</h4></div>'
                    
                    for aula in subtema["aulas"]:
                        html += f'<div class="aula"><strong>{aula["nome"]}</strong>'
                        if aula["link_aula"]:
                            html += f' | <a href="{aula["link_aula"]}" target="_blank">Aula</a>'
                        if aula["link_pdf"]:
                            html += f' | <a href="{aula["link_pdf"]}" target="_blank">PDF</a>'
                        if aula["link_quiz"]:
                            html += f' | <a href="{aula["link_quiz"]}" target="_blank">Quiz</a>'
                        html += '</div>'
        
        html += """
        </body>
        </html>
        """
        return html
    
    return jsonify(dados_formatados)

# Endpoint para recarregar os dados sem reiniciar o servidor
@app.route('/recarregar-dados', methods=['POST'])
def recarregar_dados():
    """Recarrega os dados do arquivo Excel sem precisar reiniciar o servidor"""
    global dados_ies
    try:
        arquivos = glob.glob('*.xlsx')
        if not arquivos:
            return jsonify({"error": "Nenhum arquivo .xlsx encontrado"}), 404
        
        dados_ies = processar_arquivo_excel(arquivos[0])
        return jsonify({"status": "dados_recarregados", "ies_carregadas": list(dados_ies.keys())})
    except Exception as e:
        return jsonify({"error": f"Erro ao recarregar dados: {str(e)}"}), 500

# Handler para erros 404
@app.errorhandler(404)
def not_found(error):
    return jsonify({"error": "Endpoint não encontrado"}), 404

# Handler para erros 500
@app.errorhandler(500)
def internal_error(error):
    return jsonify({"error": "Erro interno do servidor"}), 500

# Endpoint de debug para verificar se o Flask está carregado
@app.route('/debug')
def debug():
    return jsonify({
        'status': 'flask_loaded',
        'routes': [str(rule) for rule in app.url_map.iter_rules()],
        'python_version': '3.13.4'
    })

if __name__ == '__main__':
    # Procurar por arquivos Excel no diretório atual
    arquivos = glob.glob('*.xlsx')
    
    if not arquivos:
        print("Nenhum arquivo .xlsx encontrado no diretório atual.")
        print("Por favor, coloque um arquivo Excel com a estrutura especificada na mesma pasta do script.")
        # No Render, não queremos que a aplicação falhe completamente
        # Inicializamos com dados vazios mas a API fica disponível
        dados_ies = {}
    else:
        print(f"Processando arquivo: {arquivos[0]}")
        dados_ies = processar_arquivo_excel(arquivos[0])
        
        if dados_ies:
            print(f"Dados carregados com sucesso para as IES: {list(dados_ies.keys())}")
            print("API pronta para receber requisições.")
        else:
            print("Erro ao processar o arquivo Excel. Verifique a estrutura do arquivo.")
            dados_ies = {}
    
    # No Render, use a porta fornecida pela variável de ambiente PORT
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port, debug=False)