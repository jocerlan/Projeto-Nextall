import pandas as pd
from flask import Flask, render_template_string, request, redirect, url_for, jsonify
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
    
    # Existing table
    c.execute('''
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TIMESTAMP NOT NULL,
            filename TEXT NOT NULL,
            data JSON NOT NULL
        )
    ''')
    
    # New table for removed OS
    c.execute('''
        CREATE TABLE IF NOT EXISTS removed_os (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_id INTEGER NOT NULL,
            contract_id TEXT NOT NULL,
            technician TEXT NOT NULL,
            motivo TEXT NOT NULL,
            removal_reason TEXT NOT NULL,
            removed_date TIMESTAMP NOT NULL,
            FOREIGN KEY (report_id) REFERENCES reports(id)
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

        .save-btn {
            background-color: var(--primary-orange);
            color: white;
            border: none;
            padding: 12px 25px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
            transition: background-color 0.3s;
            margin-top: 20px;
        }

        .save-btn:hover {
            background-color: var (--dark-orange);
        }
        
        .modal {
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0,0,0,0.5);
        }

        .modal-content {
            background-color: white;
            margin: 15% auto;
            padding: 20px;
            border-radius: 8px;
            width: 80%;
            max-width: 500px;
        }

        .close {
            float: right;
            cursor: pointer;
            font-size: 28px;
        }

        .form-group {
            margin-bottom: 15px;
        }

        .form-group label {
            display: block;
            margin-bottom: 5px;
        }

        .form-group textarea {
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }

        .removed-os-item {
            background-color: #ffebee;
            padding: 10px;
            margin: 5px 0;
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
                                            {% for contract in motivo["Contratos"] %}
                                            <div class="contract-item">
                                                {{ contract }}
                                                <button class="edit-btn" 
                                                        onclick="openEditModal('{{ contract }}', '{{ row['Técnico'] }}', '{{ motivo['Motivo'] }}', '{{ report_id }}')">
                                                    Remover da Comissão
                                                </button>
                                            </div>
                                            {% endfor %}
                                            
                                            <div class="removed-os-section" style="margin-top: 15px;">
                                                <h4>OS Removidas da Comissão:</h4>
                                                <div class="removed-os-list" data-technician="{{ row['Técnico'] }}" data-motivo="{{ motivo['Motivo'] }}">
                                                    <!-- Removed OS will be loaded here -->
                                                </div>
                                            </div>
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
        {% if summary_data %}
            <form action="{{ url_for('save_report') }}" method="post" style="text-align: center;">
                <input type="hidden" name="filename" value="{{ filename }}">
                <input type="hidden" name="report_data" value="{{ report_data }}">
                <button type="submit" class="save-btn">Salvar Relatório</button>
            </form>
        {% endif %}
    </div>

    <!-- Add Modal HTML -->
    <div id="editModal" class="modal">
        <div class="modal-content">
            <span class="close">&times;</span>
            <h3>Remover OS da Comissão</h3>
            <form id="removeOsForm">
                <input type="hidden" id="contractId" name="contract_id">
                <input type="hidden" id="technicianName" name="technician">
                <input type="hidden" id="osMotivo" name="motivo">
                <input type="hidden" id="reportId" name="report_id">
                
                <div class="form-group">
                    <label>Contrato: <span id="contractDisplay"></span></label>
                </div>
                <div class="form-group">
                    <label>Técnico: <span id="technicianDisplay"></span></label>
                </div>
                <div class="form-group">
                    <label>Motivo: <span id="motivoDisplay"></span></label>
                </div>
                <div class="form-group">
                    <label for="removalReason">Motivo da Remoção:</label>
                    <textarea id="removalReason" name="removal_reason" required></textarea>
                </div>
                <div class="form-group">
                    <button type="submit" class="save-btn">Confirmar Remoção</button>
                    <button type="button" class="cancel-btn">Cancelar</button>
                </div>
            </form>
        </div>
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

        function openEditModal(contractId, technician, motivo, reportId) {
            const modal = document.getElementById('editModal');
            document.getElementById('contractId').value = contractId;
            document.getElementById('technicianName').value = technician;
            document.getElementById('osMotivo').value = motivo;
            document.getElementById('reportId').value = reportId;
            
            document.getElementById('contractDisplay').textContent = contractId;
            document.getElementById('technicianDisplay').textContent = technician;
            document.getElementById('motivoDisplay').textContent = motivo;
            
            modal.style.display = 'block';
        }

        document.querySelector('.close').onclick = function() {
            document.getElementById('editModal').style.display = 'none';
        }

        document.querySelector('.cancel-btn').onclick = function() {
            document.getElementById('editModal').style.display = 'none';
        }

        document.getElementById('removeOsForm').onsubmit = function(e) {
            e.preventDefault();
            const formData = new FormData(this);
            
            fetch('/remove_os', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    location.reload();
                } else {
                    alert('Erro ao remover OS: ' + data.error);
                }
            });
        }

        // Load removed OS for each section
        function loadRemovedOS() {
            const reportId = document.querySelector('input[name="report_id"]').value;
            fetch(`/get_removed_os/${reportId}`)
                .then(response => response.json())
                .then(data => {
                    const sections = document.querySelectorAll('.removed-os-list');
                    sections.forEach(section => {
                        const technician = section.dataset.technician;
                        const motivo = section.dataset.motivo;
                        const filteredOS = data.filter(os => 
                            os.technician.toLowerCase() === technician.toLowerCase() && 
                            os.motivo === motivo
                        );
                        
                        section.innerHTML = filteredOS.map(os => `
                            <div class="removed-os-item">
                                <strong>Contrato:</strong> ${os.contract_id}<br>
                                <strong>Motivo da Remoção:</strong> ${os.removal_reason}<br>
                                <strong>Data:</strong> ${os.removed_date}
                            </div>
                        `).join('') || '<p>Nenhuma OS removida</p>';
                    });
                });
        }

        // Call when page loads
        document.addEventListener('DOMContentLoaded', loadRemovedOS);
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

        .reports-list {
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }

        .report-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 15px;
            border-bottom: 1px solid var(--light-orange);
            transition: background-color 0.3s;
        }

        .report-item:hover {
            background-color: var(--light-orange);
        }

        .report-info {
            flex: 1;
        }

        .report-title {
            font-size: 18px;
            color: var(--text-dark);
            margin-bottom: 5px;
        }

        .report-date {
            color: #666;
            font-size: 14px;
        }

        .report-actions {
            display: flex;
            gap: 10px;
            align-items: center;
        }

        .view-btn {
            background-color: var(--primary-orange);
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            text-decoration: none;
            font-size: 14px;
            transition: background-color 0.3s;
        }

        .view-btn:hover {
            background-color: var (--dark-orange);
        }

        .delete-btn {
            background-color: #ff4444;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            transition: background-color 0.3s;
        }

        .delete-btn:hover {
            background-color: #cc0000;
        }

        .back-btn {
            display: inline-block;
            background-color: var(--primary-orange);
            color: white;
            text-decoration: none;
            padding: 12px 20px;
            border-radius: 4px;
            transition: background-color 0.3s;
        }

        .back-btn:hover {
            background-color: var(--dark-orange);
        }

        .empty-message {
            text-align: center;
            color: #666;
            padding: 20px;
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
                        <div class="report-info">
                            <div class="report-title">{{ report.filename }}</div>
                            <div class="report-date">Data: {{ report.date }}</div>
                        </div>
                        <div class="report-actions">
                            <a href="{{ url_for('view_report', report_id=report.id) }}" class="view-btn">Ver Relatório</a>
                            <a href="{{ url_for('view_chart', report_id=report.id) }}" class="view-btn">Ver Gráfico</a>
                            <form action="{{ url_for('delete_report', report_id=report.id) }}" method="post" style="display: inline;">
                                <button type="submit" class="delete-btn" onclick="return confirm('Tem certeza que deseja excluir este relatório?')">Excluir</button>
                            </form>
                        </div>
                    </div>
                {% endfor %}
            {% else %}
                <div class="empty-message">
                    <p>Nenhum relatório encontrado.</p>
                </div>
            {% endif %}
        </div>
        <div style="margin-top: 20px">
            <a href="{{ url_for('home') }}" class="back-btn">Voltar para Upload</a>
        </div>
    </div>
</body>
</html>
"""

# Add this new template after the existing templates
chart_template = """
<!DOCTYPE html>
<html>
<head>
    <title>Gráfico de OS - Nordeste Solar</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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

        .chart-container {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
            height: 600px; /* Fixed height for better visualization */
        }

        .back-btn {
            display: inline-block;
            background-color: var(--primary-orange);
            color: white;
            text-decoration: none;
            padding: 12px 20px;
            border-radius: 4px;
            transition: background-color 0.3s;
        }

        .back-btn:hover {
            background-color: var(--dark-orange);
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>Gráfico de Distribuição de OS - {{ filename }}</h2>
        </div>
        <div class="chart-container">
            <canvas id="osChart"></canvas>
        </div>
        <a href="{{ url_for('reports') }}" class="back-btn">Voltar para Lista</a>
    </div>

    <script>
        const data = {{ chart_data|tojson }};
        const ctx = document.getElementById('osChart').getContext('2d');
        
        // Define custom colors for different OS types
        const customColors = [
            '#FF8533', // Primary Orange
            '#4CAF50', // Green
            '#2196F3', // Blue
            '#9C27B0', // Purple
            '#F44336', // Red
            '#FFC107', // Amber
            '#00BCD4', // Cyan
            '#795548', // Brown
            '#607D8B', // Blue Grey
            '#E91E63', // Pink
            '#673AB7', // Deep Purple
            '#3F51B5', // Indigo
            '#009688', // Teal
            '#FFEB3B', // Yellow
            '#FF5722'  // Deep Orange
        ];

        new Chart(ctx, {
            type: 'pie',
            data: {
                labels: data.labels,
                datasets: [{
                    data: data.values,
                    backgroundColor: customColors.slice(0, data.labels.length),
                    borderColor: 'white',
                    borderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'right',
                        labels: {
                            font: {
                                size: 14,
                                family: "'Segoe UI', Arial, sans-serif"
                            },
                            padding: 20,
                            usePointStyle: true,
                            pointStyle: 'circle'
                        }
                    },
                    tooltip: {
                        backgroundColor: 'rgba(255, 255, 255, 0.9)',
                        titleColor: '#333',
                        bodyColor: '#333',
                        bodyFont: {
                            size: 14,
                            family: "'Segoe UI', Arial, sans-serif"
                        },
                        padding: 12,
                        boxWidth: 10,
                        borderColor: '#ddd',
                        borderWidth: 1,
                        callbacks: {
                            label: function(context) {
                                const label = context.label || '';
                                const value = context.raw || 0;
                                return `${label}: ${value}%`;
                            }
                        }
                    }
                }
            }
        });
    </script>
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
    
    # Remove the database saving code from here
    return render_template_string(html_template, 
                                summary_data=summary_data, 
                                filename=file.filename,
                                report_data=json.dumps(summary_data))

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
    
    if report: 
        summary_data = json.loads(report['data'])
        return render_template_string(html_template, 
                                    summary_data=summary_data,
                                    report_id=report_id)
    return "Relatório não encontrado", 404

@app.route('/report/<int:report_id>/delete', methods=['POST'])
def delete_report(report_id):
    conn = sqlite3.connect('reports.db')
    c = conn.cursor()
    c.execute('DELETE FROM reports WHERE id = ?', (report_id,))
    conn.commit()
    conn.close()
    return redirect(url_for('reports'))

@app.route('/save_report', methods=['POST'])
def save_report():
    filename = request.form.get('filename')
    report_data = request.form.get('report_data')
    
    if not filename or not report_data:
        return "Dados inválidos", 400
    
    conn = sqlite3.connect('reports.db')
    c = conn.cursor()
    c.execute('''
        INSERT INTO reports (date, filename, data)
        VALUES (?, ?, ?)
    ''', (
        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        filename,
        report_data
    ))
    conn.commit()
    conn.close()
    
    return redirect(url_for('reports'))

@app.route('/report/<int:report_id>/chart')
def view_chart(report_id):
    conn = sqlite3.connect('reports.db')
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT * FROM reports WHERE id = ?', (report_id,))
    report = c.fetchone()
    conn.close()
    
    if report:
        summary_data = json.loads(report['data'])
        
        # Aggregate OS types across all technicians
        motivos_counter = Counter()
        total_os = 0
        
        for tech in summary_data:
            for motivo in tech['Motivos']:
                motivos_counter[motivo['Motivo']] += motivo['Quantidade']
                total_os += motivo['Quantidade']
        
        # Calculate percentages
        chart_data = {
            'labels': [],
            'values': []
        }
        
        for motivo, count in motivos_counter.most_common():
            percentage = round((count / total_os) * 100, 1)
            chart_data['labels'].append(motivo)
            chart_data['values'].append(percentage)
        
        return render_template_string(chart_template, 
                                    filename=report['filename'],
                                    chart_data=chart_data)
    
    return "Relatório não encontrado", 404

@app.route('/remove_os', methods=['POST'])
def remove_os():
    report_id = request.form.get('report_id')
    contract_id = request.form.get('contract_id')
    technician = request.form.get('technician')
    motivo = request.form.get('motivo')
    removal_reason = request.form.get('removal_reason')
    
    conn = sqlite3.connect('reports.db')
    c = conn.cursor()
    
    try:
        # Get current report data
        c.execute('SELECT data FROM reports WHERE id = ?', (report_id,))
        report_data = json.loads(c.fetchone()[0])
        
        # Update report data
        for tech in report_data:
            if tech['Técnico'].lower() == technician.lower():
                for m in tech['Motivos']:
                    if m['Motivo'] == motivo:
                        if contract_id in m['Contratos']:
                            m['Contratos'].remove(contract_id)
                            m['Quantidade'] -= 1
                            tech['Quantidade de OS'] -= 1
                            tech['Valor Total'] -= 3  # Adjust commission value
                            
                            # Add to removed_os table
                            c.execute('''
                                INSERT INTO removed_os 
                                (report_id, contract_id, technician, motivo, removal_reason, removed_date)
                                VALUES (?, ?, ?, ?, ?, ?)
                            ''', (
                                report_id,
                                contract_id,
                                technician,
                                motivo,
                                removal_reason,
                                datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            ))
        
        # Update report data in database
        c.execute('UPDATE reports SET data = ? WHERE id = ?',
                 (json.dumps(report_data), report_id))
        
        conn.commit()
        return jsonify({'success': True})
    
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'error': str(e)})
    finally:
        conn.close()

@app.route('/get_removed_os/<int:report_id>')
def get_removed_os(report_id):
    conn = sqlite3.connect('reports.db')
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    
    c.execute('''
        SELECT * FROM removed_os 
        WHERE report_id = ? 
        ORDER BY removed_date DESC
    ''', (report_id,))
    
    removed = [dict(row) for row in c.fetchall()]
    conn.close()
    
    return jsonify(removed)

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=5000, debug=True)
