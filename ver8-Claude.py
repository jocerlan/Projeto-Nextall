import pandas as pd
from flask import Flask, render_template_string, request, redirect, url_for
import os
from collections import Counter
import re
import sqlite3
from datetime import datetime
import json

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def init_db():
    conn = sqlite3.connect('reports.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TIMESTAMP NOT NULL,
            filename TEXT NOT NULL,
            data JSON NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

html_template = """
<!DOCTYPE html>
<html>
<head>
    <title>Sistema de Relatórios - Nordeste Solar</title>
    <style>
        :root {
            --primary-orange: #ff8533;
            --secondary-orange: #ff9966;
            --light-orange: #fff3e6;
            --dark-orange: #cc5200;
            --text-dark: #333333;
        }
        
        body { 
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: var(--light-orange);
            color: var(--text-dark);
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }

        .header {
            background-color: var(--primary-orange);
            color: white;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 30px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        h2 {
            margin: 0;
            color: white;
            font-size: 24px;
        }

        .upload-form {
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }

        .file-input {
            display: none;
        }

        .file-label {
            background-color: var(--primary-orange);
            color: white;
            padding: 12px 20px;
            border-radius: 4px;
            cursor: pointer;
            display: inline-block;
            transition: background-color 0.3s;
        }

        .file-label:hover {
            background-color: var(--dark-orange);
        }

        button[type="submit"] {
            background-color: var(--primary-orange);
            color: white;
            border: none;
            padding: 12px 25px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
            transition: background-color 0.3s;
        }

        button[type="submit"]:hover {
            background-color: var(--dark-orange);
        }

        table {
            width: 100%;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        th {
            background-color: var(--primary-orange);
            color: white;
            padding: 15px;
            text-align: left;
        }

        td {
            padding: 12px 15px;
            border-bottom: 1px solid #eee;
        }

        tr:nth-child(even) {
            background-color: var(--light-orange);
        }

        .info {
            background-color: var(--secondary-orange);
            color: white;
            padding: 15px;
            border-radius: 8px;
            margin: 20px 0;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        .accordion {
            background-color: var(--secondary-orange);
            color: white;
            padding: 12px;
            border-radius: 4px;
            border: none;
            outline: none;
            transition: 0.3s;
            cursor: pointer;
            width: 100%;
            text-align: left;
            margin-top: 5px;
        }

        .accordion:hover {
            background-color: var(--dark-orange);
        }

        .panel {
            background-color: white;
            max-height: 0;
            overflow: hidden;
            transition: max-height 0.3s ease-out;
            border-radius: 0 0 4px 4px;
        }

        .value-highlight {
            font-weight: bold;
            color: var (--dark-orange);
        }

        .contract-btn {
            background-color: var(--secondary-orange);
            color: white;
            border: none;
            padding: 6px 12px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
        }

        .contract-btn:hover {
            background-color: var (--dark-orange);
        }

        .contracts-panel {
            margin-top: 5px;
            padding: 5px;
            background-color: var(--light-orange);
            border-radius: 4px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>Sistema de Relatórios - Comissão Técnico</h2>
        </div>

        <div class="upload-form">
            <form action="/upload" method="post" enctype="multipart/form-data">
                <input type="file" name="file" required class="file-input" id="file-input">
                <label for="file-input" class="file-label">Escolher arquivo Excel</label>
                <button type="submit">Processar arquivo</button>
            </form>
            <div style="margin-top: 20px">
                <a href="{{ url_for('reports') }}" class="file-label">Ver Relatórios Salvos</a>
            </div>
        </div>

        {% if summary_data %}
            <div class="info">
                Total de técnicos: <strong>{{ summary_data|length }}</strong>
            </div>
            <table>
                <tr>
                    <th>Técnico</th>
                    <th>Quantidade de OS</th>
                    <th>Valor Total (R$)</th>
                    <th>Detalhes</th>
                </tr>
                {% for row in summary_data %}
                <tr>
                    <td>{{ row["Técnico"] }}</td>
                    <td class="value-highlight">{{ row["Quantidade de OS"] }}</td>
                    <td class="value-highlight">R$ {{ "%.2f"|format(row["Valor Total"]) }}</td>
                    <td>
                        <button class="accordion">Ver detalhes</button>
                        <div class="panel">
                            <table style="box-shadow: none;">
                                <tr>
                                    <th>Motivo</th>
                                    <th>Quantidade</th>
                                    <th>Porcentagem</th>
                                    <th>Contratos</th>
                                </tr>
                                {% for motivo in row["Motivos"] %}
                                <tr>
                                    <td>{{ motivo["Motivo"] }}</td>
                                    <td>{{ motivo["Quantidade"] }}</td>
                                    <td>{{ motivo["Porcentagem"] }}%</td>
                                    <td>
                                        {% if motivo["Contratos"] %}
                                        <button class="contract-btn" onclick="toggleContracts(this)">
                                            Ver Contratos ({{ motivo["Contratos"]|length }})
                                        </button>
                                        <div class="contracts-panel" style="display: none;">
                                            {{ motivo["Contratos"]|join(", ") }}
                                        </div>
                                        {% else %}
                                        Sem contratos
                                        {% endif %}
                                    </td>
                                </tr>
                                {% endfor %}
                            </table>
                        </div>
                    </td>
                </tr>
                {% endfor %}
            </table>
        {% endif %}
    </div>

    <script>
        // Update file input label
        document.getElementById('file-input').addEventListener('change', function(e) {
            const fileName = e.target.files[0].name;
            const label = document.querySelector('.file-label');
            label.textContent = fileName;
        });

        // Accordion functionality
        var acc = document.getElementsByClassName("accordion");
        for (var i = 0; i < acc.length; i++) {
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

        // Toggle contracts panel
        function toggleContracts(button) {
            const panel = button.nextElementSibling;
            if (panel.style.display === "none") {
                panel.style.display = "block";
                button.textContent = "Ocultar Contratos";
            } else {
                panel.style.display = "none";
                button.textContent = "Ver Contratos (" + panel.textContent.split(",").length + ")";
            }
        }
    </script>
</body>
</html>
"""

reports_template = """
<!DOCTYPE html>
<html>
<head>
    <title>Relatórios Salvos - Nordeste Solar</title>
    <style>
        /* Copy existing styles from html_template */
        /* Add new styles: */
        .reports-list {
            background: white;
            padding: 20px;
            border-radius: 8px;
            margin-top: 20px;
        }
        .report-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px;
            border-bottom: 1px solid var(--light-orange);
        }
        .report-actions {
            display: flex;
            gap: 10px;
        }
        .delete-btn {
            background-color: #ff4444;
            color: white;
            border: none;
            padding: 6px 12px;
            border-radius: 4px;
            cursor: pointer;
        }
        .delete-btn:hover {
            background-color: #cc0000;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>Relatórios Salvos - Nordeste Solar</h2>
        </div>
        <div class="reports-list">
            {% if reports %}
                {% for report in reports %}
                    <div class="report-item">
                        <div>
                            <strong>{{ report.filename }}</strong>
                            <br>
                            <small>Data: {{ report.date }}</small>
                        </div>
                        <div class="report-actions">
                            <a href="{{ url_for('view_report', report_id=report.id) }}" class="file-label">Ver Relatório</a>
                            <form action="{{ url_for('delete_report', report_id=report.id) }}" method="post" style="display: inline;">
                                <button type="submit" class="delete-btn" onclick="return confirm('Tem certeza que deseja excluir este relatório?')">Excluir</button>
                            </form>
                        </div>
                    </div>
                {% endfor %}
            {% else %}
                <p>Nenhum relatório encontrado.</p>
            {% endif %}
        </div>
        <div style="margin-top: 20px">
            <a href="{{ url_for('home') }}" class="file-label">Voltar para Upload</a>
        </div>
    </div>
</body>
</html>
"""

@app.route('/', methods=['GET'])
def home():
    return render_template_string(html_template)

def get_first_name(full_name):
    """
    Extrai o primeiro nome de um nome completo
    """
    if not isinstance(full_name, str):
        return ""
    return full_name.split()[0].lower().strip() if full_name.split() else ""

def get_all_name_parts(full_name):
    """
    Retorna todas as partes do nome como uma lista
    """
    if not isinstance(full_name, str):
        return []
    return [part.lower().strip() for part in full_name.split()]

def normalize_technician_name(name):
    """
    Normaliza o nome de um técnico para facilitar a comparação e evitar duplicidades.
    """
    if not isinstance(name, str):
        return ""
    return name.strip().lower()

def extract_first_names(technician_str):
    """
    Extrai apenas o primeiro nome de cada técnico auxiliar
    """
    if not isinstance(technician_str, str):
        return []
    
    # Substituir múltiplos delimitadores por vírgulas
    normalized = re.sub(r'[;/|]+', ',', technician_str)
    
    # Dividir por vírgula
    names = []
    for tech in normalized.split(','):
        tech = tech.strip()
        if tech:
            # Para cada nome completo, pegar apenas o primeiro nome
            first_name = tech.split()[0] if tech.split() else ""
            if first_name:
                names.append(normalize_technician_name(first_name))
    
    # Remover duplicatas
    return list(dict.fromkeys(names))

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
    
    # Dicionário para mapear primeiros nomes aos nomes completos
    name_mapping = {}
    tech_full_names = set()

    # Primeiro passo: coletar todos os nomes completos dos responsáveis
    for index, row in df_cleaned.iterrows():
        if pd.notna(row['Responsável']):
            full_name = normalize_technician_name(str(row['Responsável']))
            first_name = get_first_name(full_name)
            name_parts = get_all_name_parts(full_name)
            
            for part in name_parts:
                name_mapping[part] = full_name
            tech_full_names.add(full_name)

    # Usar um dicionário para armazenar os dados de cada técnico
    technician_data = {}
    processed_techs = set()

    # Processar cada linha do DataFrame
    for index, row in df_cleaned.iterrows():
        motivo = row['Motivo'] if pd.notna(row['Motivo']) else "Não especificado"
        processed_techs.clear()

        # Processar responsável
        if pd.notna(row['Responsável']):
            tech = normalize_technician_name(str(row['Responsável']))
            if tech and tech not in processed_techs:
                if tech not in technician_data:
                    technician_data[tech] = {
                        "Técnico": tech,
                        "Quantidade de OS": 0,
                        "Valor Total": 0,
                        "Motivos": Counter(),
                        "Contratos": {}  # Add this line
                    }
                technician_data[tech]["Quantidade de OS"] += 1
                technician_data[tech]["Valor Total"] += 3
                technician_data[tech]["Motivos"][motivo] += 1
                processed_techs.add(tech)

                # When processing each row, add the contract ID:
                if pd.notna(row['ID Contrato']):
                    if motivo not in technician_data[tech]["Contratos"]:
                        technician_data[tech]["Contratos"][motivo] = []
                    technician_data[tech]["Contratos"][motivo].append(str(row['ID Contrato']))

        # Processar técnicos auxiliares
        if pd.notna(row['Técnico(s) auxiliar(s)']):
            aux_techs = extract_first_names(str(row['Técnico(s) auxiliar(s)']))
            for aux_tech in aux_techs:
                # Verificar se o primeiro nome corresponde a algum técnico conhecido
                matched_full_name = name_mapping.get(aux_tech)
                
                if matched_full_name and matched_full_name not in processed_techs:
                    # Usar o nome completo correspondente
                    tech = matched_full_name
                elif aux_tech not in processed_techs:
                    # Usar apenas o primeiro nome se não houver correspondência
                    tech = aux_tech

                if tech not in processed_techs:
                    if tech not in technician_data:
                        technician_data[tech] = {
                            "Técnico": tech,
                            "Quantidade de OS": 0,
                            "Valor Total": 0,
                            "Motivos": Counter(),
                            "Contratos": {}  # Add this line
                        }
                    technician_data[tech]["Quantidade de OS"] += 1
                    technician_data[tech]["Valor Total"] += 3
                    technician_data[tech]["Motivos"][motivo] += 1
                    processed_techs.add(tech)

                    # When processing each row, add the contract ID:
                    if pd.notna(row['ID Contrato']):
                        if motivo not in technician_data[tech]["Contratos"]:
                            technician_data[tech]["Contratos"][motivo] = []
                        technician_data[tech]["Contratos"][motivo].append(str(row['ID Contrato']))

    # Verificar se temos dados para processar
    if not technician_data:
        return "Nenhum técnico encontrado nos dados", 400
    
    # Converter para o formato de resumo
    summary_data = []
    for tech_info in technician_data.values():
        total_os = tech_info["Quantidade de OS"]
        motivos_list = [
            {
                "Motivo": motivo,
                "Quantidade": quantidade,
                "Porcentagem": round((quantidade / total_os) * 100, 1),
                "Contratos": tech_info["Contratos"].get(motivo, [])  # Add this line
            }
            for motivo, quantidade in tech_info["Motivos"].items()
        ]
        summary_data.append({
            "Técnico": tech_info["Técnico"].title(),
            "Quantidade de OS": total_os,
            "Valor Total": tech_info["Valor Total"],
            "Motivos": sorted(motivos_list, key=lambda x: x["Quantidade"], reverse=True)
        })
    
    summary_data.sort(key=lambda x: x["Quantidade de OS"], reverse=True)
    
    # Save report to database
    conn = sqlite3.connect('reports.db')
    c = conn.cursor()
    c.execute('''
        INSERT INTO reports (date, filename, data)
        VALUES (?, ?, ?)
    ''', (
        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        file.filename,
        json.dumps(summary_data)
    ))
    conn.commit()
    conn.close()
    
    return render_template_string(html_template, summary_data=summary_data)

# Add these new routes

@app.route('/reports')
def reports():
    conn = sqlite3.connect('reports.db')
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT * FROM reports ORDER BY date DESC')
    reports = c.fetchall()
    conn.close()
    return render_template_string(reports_template, reports=reports)

@app.route('/report/<int:report_id>')
def view_report(report_id):
    conn = sqlite3.connect('reports.db')
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT * FROM reports WHERE id = ?', (report_id,))
    report = c.fetchone()
    conn.close()
    
    if (report):
        summary_data = json.loads(report['data'])
        return render_template_string(html_template, summary_data=summary_data)
    return "Relatório não encontrado", 404

@app.route('/report/<int:report_id>/delete', methods=['POST'])
def delete_report(report_id):
    conn = sqlite3.connect('reports.db')
    c = conn.cursor()
    c.execute('DELETE FROM reports WHERE id = ?', (report_id,))
    conn.commit()
    conn.close()
    return redirect(url_for('reports'))

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=5000, debug=True)
