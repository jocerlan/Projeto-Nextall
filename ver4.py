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
        .info { margin: 10px 0; padding: 10px; background-color: #e7f3fe; border-left: 6px solid #2196F3; }
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
        <div class="info">Total de técnicos: {{ summary_data|length }}</div>
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
                            <th>Porcentagem</th>
                        </tr>
                        {% for motivo in row["Motivos"] %}
                        <tr>
                            <td>{{ motivo["Motivo"] }}</td>
                            <td>{{ motivo["Quantidade"] }}</td>
                            <td>{{ motivo["Porcentagem"] }}%</td>
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

def normalize_technician_name(name):
    """
    Normaliza o nome de um técnico para facilitar a comparação e evitar duplicidades.
    """
    if not isinstance(name, str):
        return ""
    
    # Remover espaços extras, converter para minúsculas
    normalized = name.strip().lower()
    return normalized

def split_compound_names(names_list):
    """
    Separa nomes compostos que possam ter sido concatenados.
    Por exemplo: "Jadiel Daniel" -> ["Jadiel", "Daniel"]
    """
    result = []
    
    # Lista de nomes conhecidos que não devem ser separados
    # (Adicione aqui nomes compostos que não devem ser divididos)
    non_separable_names = []
    
    for name in names_list:
        name = name.strip()
        if not name:
            continue
            
        # Verifica se o nome está na lista de não separáveis
        if name.lower() in [n.lower() for n in non_separable_names]:
            result.append(name)
            continue
            
        # Separa o nome por espaços
        parts = name.split()
        
        # Se o nome tem mais de uma parte, pode ser um nome composto
        if len(parts) > 1:
            result.append(parts[0])  # Primeiro nome
            result.append(' '.join(parts[1:]))  # Sobrenome ou outros nomes
        else:
            result.append(name)  # Nome único
            
    return result

def process_technician_names(technician_str):
    """
    Processa uma string contendo nomes de técnicos separados por vírgulas, ponto e vírgula,
    ou outros delimitadores comuns, retornando uma lista de nomes normalizados
    """
    if not isinstance(technician_str, str):
        return []
    
    # Substituir múltiplos delimitadores por vírgulas
    normalized = re.sub(r'[;/|]+', ',', technician_str)
    
    # Dividir por vírgula
    initial_names = [tech.strip() for tech in normalized.split(',')]
    
    # Processar nomes compostos
    split_names = split_compound_names(initial_names)
    
    # Normalizar e remover vazios
    return [normalize_technician_name(tech) for tech in split_names if tech.strip()]

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
    
    # Usar um dicionário para armazenar os dados de cada técnico
    technician_data = {}
    
    # Processar cada linha do DataFrame
    for index, row in df_cleaned.iterrows():
        motivo = row['Motivo'] if pd.notna(row['Motivo']) else "Não especificado"
        
        # Processar responsáveis
        if pd.notna(row['Responsável']):
            for tech in process_technician_names(str(row['Responsável'])):
                if tech not in technician_data:
                    technician_data[tech] = {'motivos': []}
                technician_data[tech]['motivos'].append(motivo)
        
        # Processar técnicos auxiliares
        if pd.notna(row['Técnico(s) auxiliar(s)']):
            for tech in process_technician_names(str(row['Técnico(s) auxiliar(s)'])):
                if tech not in technician_data:
                    technician_data[tech] = {'motivos': []}
                technician_data[tech]['motivos'].append(motivo)
    
    # Verificar se temos dados para processar
    if not technician_data:
        return "Nenhum técnico encontrado nos dados", 400
    
    # Converter para o formato de resumo
    summary_data = []
    for tech, data in technician_data.items():
        # Contar os motivos para este técnico
        motives_count = Counter(data['motivos'])
        total_os = sum(motives_count.values())
        
        # Criar lista de motivos ordenada do mais frequente para o menos frequente
        motives_list = []
        for motivo, count in motives_count.most_common():
            percentage = round((count / total_os) * 100, 1)
            motives_list.append({
                'Motivo': motivo,
                'Quantidade': count,
                'Porcentagem': percentage
            })
        
        # Adicionar a capitalização adequada para nomes de pessoas
        tech_display = ' '.join(word.capitalize() for word in tech.split())
        
        summary_data.append({
            'Técnico': tech_display,
            'Quantidade de OS': total_os,
            'Motivos': motives_list
        })
    
    # Ordenar por quantidade de OS (do maior para o menor)
    summary_data = sorted(summary_data, key=lambda x: x['Quantidade de OS'], reverse=True)
    
    return render_template_string(html_template, summary_data=summary_data)

if __name__ == '__main__':
    app.run(debug=True)