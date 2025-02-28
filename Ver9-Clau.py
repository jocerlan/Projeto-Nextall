import pandas as pd
from flask import Flask, render_template_string, request
import os
from collections import Counter
import re
import sqlite3
from datetime import datetime
import json

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Add this function to initialize the database
def init_db():
    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            technician_data TEXT
        )
    ''')
    conn.commit()
    conn.close()

# Initialize the database
init_db()

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

        .contracts-list {
            margin-top: 8px;
            padding: 8px;
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
                                        <button class="contract-btn" onclick="toggleContracts(this)">Ver Contratos</button>
                                        <div class="contracts-list" style="display: none;">
                                            {% for contrato in motivo["Contratos"] %}
                                                <div>{{ contrato }}</div>
                                            {% endfor %}
                                        </div>
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

        // Toggle contracts list
        function toggleContracts(button) {
            const contractsList = button.nextElementSibling;
            if (contractsList.style.display === "none") {
                contractsList.style.display = "block";
                button.textContent = "Ocultar Contratos";
            } else {
                contractsList.style.display = "none";
                button.textContent = "Ver Contratos";
            }
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

def find_matching_technician(tech_name, technician_data):
    """
    Procura por um técnico existente que corresponda ao nome fornecido
    """
    full_name, first_name = normalize_technician_name(tech_name), tech_name.split()[0].lower()
    
    if not first_name:
        return None
        
    # Primeiro procura por correspondência exata
    if (full_name in technician_data):
        return full_name
        
    # Procura por correspondência no primeiro nome ou sobrenome
    for existing_tech in technician_data.keys():
        existing_full, existing_first = normalize_technician_name(existing_tech), existing_tech.split()[0].lower()
        if first_name == existing_first or first_name in existing_full.split():
            return existing_tech
            
    # Se não encontrou, retorna o nome completo original
    return full_name

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
    
    # Lista para rastrear técnicos já contabilizados em cada OS
    processed_techs = set()
    
    # Processar cada linha do DataFrame
    for index, row in df_cleaned.iterrows():
        motivo = row['Motivo'] if pd.notna(row['Motivo']) else "Não especificado"
        contract_id = str(row['ID Contrato']) if pd.notna(row['ID Contrato']) else "N/A"
        processed_techs.clear()  # Limpar o conjunto para cada nova OS
        
        # Processar responsáveis
        if pd.notna(row['Responsável']):
            # Para responsável principal, mantemos o nome completo
            tech = normalize_technician_name(str(row['Responsável']))
            matching_tech = find_matching_technician(tech, technician_data)
            if matching_tech and matching_tech not in processed_techs:
                if matching_tech not in technician_data:
                    technician_data[matching_tech] = {
                        "Técnico": matching_tech,
                        "Quantidade de OS": 0,
                        "Valor Total": 0,
                        "Motivos": Counter(),
                        "Contratos": {}  # New structure to store contract IDs by motivo
                    }
                technician_data[matching_tech]["Quantidade de OS"] += 1
                technician_data[matching_tech]["Valor Total"] += 3  # Add 3 reais per OS
                technician_data[matching_tech]["Motivos"][motivo] += 1
                
                # Store contract IDs for each motivo
                if motivo not in technician_data[matching_tech]["Contratos"]:
                    technician_data[matching_tech]["Contratos"][motivo] = []
                technician_data[matching_tech]["Contratos"][motivo].append(contract_id)
                
                processed_techs.add(matching_tech)
        
        # Processar técnicos auxiliares (apenas primeiro nome)
        if pd.notna(row['Técnico(s) auxiliar(s)']):
            for aux_tech in extract_first_names(str(row['Técnico(s) auxiliar(s)'])):
                matching_tech = find_matching_technician(aux_tech, technician_data)
                if matching_tech and matching_tech not in processed_techs:
                    if matching_tech not in technician_data:
                        technician_data[matching_tech] = {
                            "Técnico": matching_tech,
                            "Quantidade de OS": 0,
                            "Valor Total": 0,
                            "Motivos": Counter(),
                            "Contratos": {}  # New structure to store contract IDs by motivo
                        }
                    technician_data[matching_tech]["Quantidade de OS"] += 1
                    technician_data[matching_tech]["Valor Total"] += 3  # Add 3 reais per OS
                    technician_data[matching_tech]["Motivos"][motivo] += 1
                    
                    # Store contract IDs for each motivo
                    if motivo not in technician_data[matching_tech]["Contratos"]:
                        technician_data[matching_tech]["Contratos"][motivo] = []
                    technician_data[matching_tech]["Contratos"][motivo].append(contract_id)
                    
                    processed_techs.add(matching_tech)
    
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
                "Contratos": tech_info["Contratos"].get(motivo, [])  # Add contract IDs
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
    
    # Save the processed data to the database
    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('INSERT INTO reports (date, technician_data) VALUES (?, ?)', (datetime.now().isoformat(), json.dumps(technician_data)))
    conn.commit()
    conn.close()
    
    return render_template_string(html_template, summary_data=summary_data)

@app.route('/reports', methods=['GET'])
def get_reports():
    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('SELECT id, date FROM reports ORDER BY date DESC')
    reports = c.fetchall()
    conn.close()
    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Relatórios Salvos</title>
        <style>
            body { 
                font-family: 'Segoe UI', Arial, sans-serif;
                margin: 0;
                padding: 20px;
                background-color: #fff3e6;
                color: #333333;
            }
            .container {
                max-width: 800px;
                margin: 0 auto;
                padding: 20px;
            }
            .header {
                background-color: #ff8533;
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
            table {
                width: 100%;
                background: white;
                border-radius: 8px;
                overflow: hidden;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            th, td {
                padding: 12px 15px;
                border-bottom: 1px solid #eee;
            }
            th {
                background-color: #ff8533;
                color: white;
                text-align: left;
            }
            tr:nth-child(even) {
                background-color: #fff3e6;
            }
            a {
                color: #ff8533;
                text-decoration: none;
                margin-right: 10px;
            }
            a:hover {
                text-decoration: underline;
            }
            .delete-btn {
                color: #ff3333;
                cursor: pointer;
            }
            .delete-btn:hover {
                text-decoration: underline;
            }
            .actions {
                display: flex;
                gap: 10px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h2>Relatórios Salvos</h2>
            </div>
            <table>
                <tr>
                    <th>ID</th>
                    <th>Data</th>
                    <th>Ações</th>
                </tr>
                {% for report in reports %}
                <tr id="report-{{ report[0] }}">
                    <td>{{ report[0] }}</td>
                    <td>{{ report[1] }}</td>
                    <td class="actions">
                        <a href="/report/{{ report[0] }}">Ver</a>
                        <a href="#" class="delete-btn" onclick="deleteReport({{ report[0] }})">Excluir</a>
                    </td>
                </tr>
                {% endfor %}
            </table>
        </div>

        <script>
        function deleteReport(reportId) {
            if (confirm('Tem certeza que deseja excluir este relatório?')) {
                fetch(`/report/${reportId}/delete`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    }
                }).then(response => {
                    if (response.ok) {
                        document.getElementById(`report-${reportId}`).remove();
                        console.log('Relatório excluído com sucesso');
                    } else if (response.status === 404) {
                        alert('Relatório não encontrado');
                        location.reload();
                    } else {
                        alert('Erro ao excluir o relatório');
                        console.error('Erro:', response.status);
                    }
                }).catch(error => {
                    console.error('Erro:', error);
                    alert('Erro ao excluir o relatório');
                });
            }
        }
        </script>
    </body>
    </html>
""", reports=reports)

@app.route('/report/<int:report_id>', methods=['GET'])
def get_report(report_id):
    conn = sqlite3.connect('data.db')
    c = conn.cursor()
    c.execute('SELECT technician_data FROM reports WHERE id = ?', (report_id,))
    report = c.fetchone()
    conn.close()
    if report:
        technician_data = json.loads(report[0])
        summary_data = []
        for tech_info in technician_data.values():
            total_os = tech_info["Quantidade de OS"]
            motivos_list = [
                {
                    "Motivo": motivo,
                    "Quantidade": quantidade,
                    "Porcentagem": round((quantidade / total_os) * 100, 1),
                    "Contratos": tech_info["Contratos"].get(motivo, [])  # Add contract IDs
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
        
        return render_template_string(html_template, summary_data=summary_data)
    else:
        return "Relatório não encontrado", 404

# Add this route before if __name__ == '__main__':
@app.route('/report/<int:report_id>/delete', methods=['POST'])
def delete_report(report_id):
    try:
        conn = sqlite3.connect('data.db')
        c = conn.cursor()
        
        # Verify if report exists before deletion
        c.execute('SELECT id FROM reports WHERE id = ?', (report_id,))
        if not c.fetchone():
            conn.close()
            return 'Relatório não encontrado', 404
            
        # Delete the report
        c.execute('DELETE FROM reports WHERE id = ?', (report_id,))
        conn.commit()
        conn.close()
        return '', 204
    except Exception as e:
        if conn:
            conn.rollback()
            conn.close()
        print(f"Erro ao excluir relatório: {e}")
        return 'Erro ao excluir relatório', 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
