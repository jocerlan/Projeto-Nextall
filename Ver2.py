import pandas as pd
from flask import Flask, render_template_string, request
import os
from collections import Counter
import re

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

html_template = """
<!DOCTYPE html>
<html>
<head>
    <title>Upload de Arquivo</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        form { margin: 20px 0; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .accordion { background-color: #f1f1f1; color: #444; cursor: pointer; padding: 18px; width: 100%; text-align: left; border: none; outline: none; transition: 0.4s; margin-top: 10px; }
        .active, .accordion:hover { background-color: #ddd; }
        .panel { padding: 0 18px; background-color: white; max-height: 0; overflow: hidden; transition: max-height 0.2s ease-out; }
        h2 { margin-top: 30px; }
    </style>
</head>
<body>
    <h2>Envie seu arquivo Excel</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="file" required>
        <button type="submit">Enviar</button>
    </form>
    
    {% if summary_data %}
        <h2>Resumo por Técnico</h2>
        <table>
            <tr>
                <th>Técnico</th>
                <th>Quantidade de OS</th>
                <th>Detalhes</th>
            </tr>
            {% for row in summary_data %}
            <tr>
                <td>{{ row["Técnico"] }}</td>
                <td>{{ row["Quantidade de OS"] }}</td>
                <td><button class="accordion">Ver motivos</button>
                <div class="panel">
                    <table style="width: 100%">
                        <tr>
                            <th>Motivo</th>
                            <th>Quantidade</th>
                        </tr>
                        {% for motivo in row["Motivos"] %}
                        <tr>
                            <td>{{ motivo["Motivo"] }}</td>
                            <td>{{ motivo["Quantidade"] }}</td>
                        </tr>
                        {% endfor %}
                    </table>
                </div></td>
            </tr>
            {% endfor %}
        </table>
    {% endif %}
    
    <script>
        var acc = document.getElementsByClassName("accordion");
        var i;
        
        for (i = 0; i < acc.length; i++) {
            acc[i].addEventListener("click", function() {
                this.classList.toggle("active");
                var panel = this.nextElementSibling;
                if (panel.style.maxHeight) {
                    panel.style.maxHeight = null;
                } else {
                    panel.style.maxHeight = panel.scrollHeight + "px";
                }
            });
        }
    </script>
</body>
</html>
"""

@app.route('/', methods=['GET'])
def home():
    return render_template_string(html_template)

def process_technician_names(technician_str):
    """
    Processa uma string contendo nomes de técnicos separados por vírgulas, ponto e vírgula,
    ou outros delimitadores comuns, retornando uma lista de nomes normalizados
    """
    if not isinstance(technician_str, str):
        return []
    
    # Substituir múltiplos delimitadores por vírgulas
    normalized = re.sub(r'[;/|]+', ',', technician_str)
    # Dividir por vírgula e limpar espaços
    techs = [tech.strip().lower() for tech in normalized.split(',')]
    # Remover entradas vazias
    return [tech for tech in techs if tech]

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "Nenhum arquivo enviado", 400
    file = request.files['file']
    if file.filename == '':
        return "Nenhum arquivo selecionado", 400
    
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)
    
    # Processar o arquivo
    xls = pd.ExcelFile(file_path)
    df = pd.read_excel(xls, sheet_name="Ordens de Serviço")
    df_cleaned = df.iloc[6:].reset_index(drop=True)
    df_cleaned.columns = df_cleaned.iloc[0]  # Definir a primeira linha como cabeçalho
    df_cleaned = df_cleaned[1:].reset_index(drop=True)
    
    # Remover os tipos de OS "Financeiro" e "Entrega de Carnê"
    df_cleaned = df_cleaned[~df_cleaned['Motivo'].isin(['Financeiro', 'Entrega de Carnê'])]
    
    # Criar um DataFrame para armazenar os técnicos e seus motivos
    technician_motives = []
    
    # Processar cada linha do DataFrame
    for index, row in df_cleaned.iterrows():
        motivo = row['Motivo'] if pd.notna(row['Motivo']) else "Não especificado"
        
        # Processar responsáveis
        if pd.notna(row['Responsável']):
            for tech in process_technician_names(str(row['Responsável'])):
                technician_motives.append({
                    'Técnico': tech,
                    'Motivo': motivo
                })
        
        # Processar técnicos auxiliares
        if pd.notna(row['Técnico(s) auxiliar(s)']):
            for tech in process_technician_names(str(row['Técnico(s) auxiliar(s)'])):
                technician_motives.append({
                    'Técnico': tech,
                    'Motivo': motivo
                })
    
    # Converter para DataFrame
    df_motives = pd.DataFrame(technician_motives)
    
    # Verificar se temos dados para processar
    if df_motives.empty:
        return "Nenhum técnico encontrado nos dados", 400
    
    # Calcular estatísticas por técnico
    summary_data = []
    for tech, group in df_motives.groupby('Técnico'):
        # Contar os motivos para este técnico
        motives_count = Counter(group['Motivo'])
        
        # Criar lista de motivos ordenada do mais frequente para o menos frequente
        motives_list = [{'Motivo': motivo, 'Quantidade': count} 
                        for motivo, count in motives_count.most_common()]
        
        # Adicionar a capitalização adequada para nomes de pessoas
        tech_display = ' '.join(word.capitalize() for word in tech.split())
        
        summary_data.append({
            'Técnico': tech_display,
            'Quantidade de OS': len(group),
            'Motivos': motives_list
        })
    
    # Ordenar por quantidade de OS (do maior para o menor)
    summary_data = sorted(summary_data, key=lambda x: x['Quantidade de OS'], reverse=True)
    
    return render_template_string(html_template, summary_data=summary_data)

if __name__ == '__main__':
    app.run(debug=True)