import os
import json
import openpyxl
from collections import Counter, defaultdict
from flask import Flask, request, jsonify, render_template_string, send_file
from werkzeug.utils import secure_filename
from pipeline_auditor import procesar_presupuesto_incremental, cargar_historico

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = os.path.dirname(os.path.abspath(__file__))
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB
BD_HISTORICO_PATH = os.path.join(app.config["UPLOAD_FOLDER"], "BD_HISTORICO_BASE_VACIA.xlsx")

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>KIMSAV - Clasificador Incremental de Presupuestos</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg: #0b0f19;
      --card-bg: rgba(22, 30, 49, 0.75);
      --card-border: rgba(255, 255, 255, 0.08);
      --accent: #4f46e5;
      --accent-gradient: linear-gradient(135deg, #6366f1 0%, #3b82f6 50%, #06b6d4 100%);
      --text: #f3f4f6;
      --text-muted: #9ca3af;
      --success: #10b981;
      --warning: #f59e0b;
      --danger: #ef4444;
      --info: #38bdf8;
      --border-radius: 16px;
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      font-family: 'Outfit', sans-serif;
    }

    body {
      background-color: var(--bg);
      background-image: 
        radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.15) 0px, transparent 50%),
        radial-gradient(at 100% 100%, rgba(6, 182, 212, 0.12) 0px, transparent 50%);
      color: var(--text);
      min-height: 100vh;
      padding: 30px 20px;
    }

    .container {
      max-width: 1200px;
      margin: 0 auto;
    }

    header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 30px;
      padding-bottom: 20px;
      border-bottom: 1px solid var(--card-border);
    }

    .logo-box h1 {
      font-size: 26px;
      font-weight: 800;
      background: var(--accent-gradient);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      letter-spacing: -0.5px;
    }

    .logo-box p {
      color: var(--text-muted);
      font-size: 14px;
      margin-top: 4px;
    }

    .badge-history {
      background: rgba(99, 102, 241, 0.15);
      border: 1px solid rgba(99, 102, 241, 0.35);
      padding: 8px 16px;
      border-radius: 999px;
      font-size: 14px;
      font-weight: 600;
      color: #818cf8;
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .badge-history span {
      background: #4f46e5;
      color: white;
      padding: 2px 8px;
      border-radius: 999px;
      font-size: 12px;
    }

    .grid-layout {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 25px;
      margin-bottom: 30px;
    }

    @media (max-width: 850px) {
      .grid-layout {
        grid-template-columns: 1fr;
      }
    }

    .card {
      background: var(--card-bg);
      backdrop-filter: blur(16px);
      border: 1px solid var(--card-border);
      border-radius: var(--border-radius);
      padding: 26px;
      box-shadow: 0 10px 30px -10px rgba(0, 0, 0, 0.5);
    }

    .card h2 {
      font-size: 19px;
      font-weight: 700;
      margin-bottom: 16px;
      display: flex;
      align-items: center;
      gap: 10px;
    }

    /* Form Styles */
    .drop-zone {
      border: 2px dashed rgba(255, 255, 255, 0.18);
      border-radius: 12px;
      padding: 30px 20px;
      text-align: center;
      cursor: pointer;
      transition: all 0.25s ease;
      background: rgba(255, 255, 255, 0.02);
      margin-bottom: 18px;
    }

    .drop-zone:hover, .drop-zone.dragover {
      border-color: #6366f1;
      background: rgba(99, 102, 241, 0.08);
    }

    .drop-zone svg {
      width: 44px;
      height: 44px;
      stroke: #818cf8;
      margin-bottom: 10px;
    }

    .drop-zone p {
      font-size: 14px;
      color: var(--text-muted);
    }

    .file-info {
      margin-top: 10px;
      font-size: 13px;
      font-weight: 600;
      color: #38bdf8;
    }

    .input-row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 14px;
      margin-bottom: 14px;
    }

    .form-group {
      margin-bottom: 14px;
    }

    .form-group label {
      display: block;
      font-size: 13px;
      font-weight: 600;
      color: var(--text-muted);
      margin-bottom: 6px;
    }

    .form-group input {
      width: 100%;
      background: rgba(255, 255, 255, 0.05);
      border: 1px solid rgba(255, 255, 255, 0.12);
      border-radius: 8px;
      padding: 10px 14px;
      color: white;
      font-size: 14px;
      outline: none;
      transition: border 0.2s ease;
    }

    .form-group input:focus {
      border-color: #6366f1;
      background: rgba(255, 255, 255, 0.08);
    }

    .checkbox-group {
      display: flex;
      align-items: center;
      gap: 10px;
      margin: 16px 0;
      cursor: pointer;
    }

    .checkbox-group input {
      accent-color: #6366f1;
      width: 17px;
      height: 17px;
      cursor: pointer;
    }

    .checkbox-group label {
      font-size: 13px;
      color: var(--text-muted);
      cursor: pointer;
    }

    .btn-submit {
      width: 100%;
      padding: 13px;
      border: none;
      border-radius: 10px;
      background: var(--accent-gradient);
      color: white;
      font-size: 15px;
      font-weight: 700;
      cursor: pointer;
      box-shadow: 0 4px 15px rgba(99, 102, 241, 0.35);
      transition: all 0.2s ease;
      display: flex;
      justify-content: center;
      align-items: center;
      gap: 8px;
    }

    .btn-submit:hover:not(:disabled) {
      transform: translateY(-2px);
      box-shadow: 0 6px 20px rgba(99, 102, 241, 0.5);
    }

    .btn-submit:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }

    /* History Projects List */
    .history-list {
      max-height: 380px;
      overflow-y: auto;
      display: flex;
      flex-direction: column;
      gap: 10px;
      padding-right: 6px;
    }

    .history-item {
      background: rgba(255, 255, 255, 0.04);
      border: 1px solid rgba(255, 255, 255, 0.06);
      border-radius: 10px;
      padding: 12px 16px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      transition: background 0.2s ease;
    }

    .history-item:hover {
      background: rgba(255, 255, 255, 0.07);
    }

    .history-item-name {
      font-weight: 600;
      font-size: 15px;
      color: #e0e7ff;
    }

    .history-item-meta {
      font-size: 12px;
      color: var(--text-muted);
      margin-top: 3px;
    }

    .history-item-count {
      background: rgba(16, 185, 129, 0.15);
      color: #34d399;
      border: 1px solid rgba(16, 185, 129, 0.3);
      padding: 4px 10px;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 600;
      font-family: 'JetBrains Mono', monospace;
    }

    .empty-state {
      text-align: center;
      padding: 40px 20px;
      color: var(--text-muted);
    }

    /* Stats & Results */
    .results-section {
      display: none;
      animation: fadeIn 0.4s ease forwards;
    }

    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(10px); }
      to { opacity: 1; transform: translateY(0); }
    }

    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 15px;
      margin-bottom: 25px;
    }

    .kpi-card {
      background: rgba(255, 255, 255, 0.03);
      border: 1px solid rgba(255, 255, 255, 0.07);
      border-radius: 12px;
      padding: 16px;
      text-align: center;
    }

    .kpi-title {
      font-size: 12px;
      font-weight: 600;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .kpi-value {
      font-size: 24px;
      font-weight: 800;
      color: white;
      margin-top: 6px;
      font-family: 'JetBrains Mono', monospace;
    }

    .kpi-sub {
      font-size: 11px;
      margin-top: 4px;
      color: var(--text-muted);
    }

    .kpi-value.hist { color: #38bdf8; }
    .kpi-value.reglas { color: #818cf8; }
    .kpi-value.hoja { color: #fbbf24; }
    .kpi-value.rev { color: #f87171; }

    /* Tables & Pre */
    .table-container {
      max-height: 400px;
      overflow-y: auto;
      border: 1px solid var(--card-border);
      border-radius: 10px;
      background: rgba(0, 0, 0, 0.2);
    }

    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }

    th {
      background: rgba(255, 255, 255, 0.05);
      position: sticky;
      top: 0;
      padding: 12px 14px;
      text-align: left;
      font-weight: 700;
      color: #94a3b8;
      border-bottom: 1px solid rgba(255, 255, 255, 0.1);
    }

    td {
      padding: 10px 14px;
      border-bottom: 1px solid rgba(255, 255, 255, 0.04);
    }

    tr:hover {
      background: rgba(255, 255, 255, 0.02);
    }

    .badge-method {
      padding: 3px 8px;
      border-radius: 4px;
      font-size: 11px;
      font-weight: 600;
      text-transform: uppercase;
    }

    .badge-method.historico { background: rgba(56, 189, 248, 0.2); color: #38bdf8; }
    .badge-method.reglas { background: rgba(129, 140, 248, 0.2); color: #818cf8; }
    .badge-method.hoja { background: rgba(251, 191, 36, 0.2); color: #fbbf24; }
    .badge-method.sin_resolver { background: rgba(239, 68, 68, 0.2); color: #ef4444; }

    .download-bar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      background: rgba(99, 102, 241, 0.1);
      border: 1px solid rgba(99, 102, 241, 0.3);
      padding: 15px 20px;
      border-radius: 12px;
      margin-top: 20px;
    }

    .btn-download {
      background: #10b981;
      color: white;
      text-decoration: none;
      padding: 9px 18px;
      border-radius: 8px;
      font-weight: 700;
      font-size: 14px;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      transition: background 0.2s;
    }

    .btn-download:hover {
      background: #059669;
    }

    /* Loader */
    .spinner {
      display: none;
      width: 20px;
      height: 20px;
      border: 3px solid rgba(255, 255, 255, 0.3);
      border-radius: 50%;
      border-top-color: white;
      animation: spin 0.8s ease-in-out infinite;
    }

    @keyframes spin {
      to { transform: rotate(360deg); }
    }
  </style>
</head>
<body>

  <div class="container">
    <header>
      <div class="logo-box">
        <h1>KIMSAV Core</h1>
        <p>Sistema Incremental y Continuo de Clasificación de Presupuestos</p>
      </div>
      <div class="badge-history" id="headerHistoryBadge">
        <svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10"></path></svg>
        Histórico Acumulado: <span id="headerProjectCount">0 proyectos</span>
      </div>
    </header>

    <div class="grid-layout">
      <!-- Upload Card -->
      <div class="card">
        <h2>
          <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path></svg>
          Procesar Nuevo Presupuesto
        </h2>

        <form id="uploadForm">
          <div class="drop-zone" id="dropZone">
            <svg fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 13h6m-3-3v6m5 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg>
            <p>Arrastra tu archivo Excel aquí o haz clic para seleccionar</p>
            <input type="file" id="fileInput" name="file" accept=".xlsx,.xls" style="display:none;" required>
            <div class="file-info" id="fileInfo"></div>
          </div>

          <div class="form-group">
            <label for="proyectoInput">Nombre del Proyecto</label>
            <input type="text" id="proyectoInput" name="proyecto" placeholder="Ej: VIRIDIAN EUREKA (auto-detectado si se omite)">
          </div>

          <div class="input-row">
            <div class="form-group">
              <label for="tipoInput">Tipo de Proyecto</label>
              <input type="text" id="tipoInput" name="tipo" placeholder="Ej: Edificio Residencial">
            </div>
            <div class="form-group">
              <label for="departamentoInput">Ubicación / Depto</label>
              <input type="text" id="departamentoInput" name="departamento" placeholder="Ej: Lima">
            </div>
          </div>

          <div class="form-group">
            <label for="areaInput">Área Techada Total (m²)</label>
            <input type="number" step="0.01" id="areaInput" name="area" placeholder="Opcional (se busca en resumen)">
          </div>

          <div class="checkbox-group">
            <input type="checkbox" id="forzarInput" name="forzar">
            <label for="forzarInput">Sobrescribir si el proyecto ya existe en la base histórica</label>
          </div>

          <button type="submit" class="btn-submit" id="btnSubmit">
            <div class="spinner" id="spinner"></div>
            <span id="btnText">Clasificar e Incorporar al Histórico</span>
          </button>
        </form>
      </div>

      <!-- History Status Card -->
      <div class="card">
        <h2>
          <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
          Base Histórica Acumulativa
        </h2>
        <p style="font-size: 13px; color: var(--text-muted); margin-bottom: 16px;">
          Cada proyecto clasificado enriquece la base de datos maestra para que los siguientes proyectos se clasifiquen con mayor precisión.
        </p>
        <div class="history-list" id="historyList">
          <div class="empty-state">Cargando base histórica...</div>
        </div>
      </div>
    </div>

    <!-- Results Section -->
    <div class="card results-section" id="resultsSection">
      <h2>
        <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4"></path></svg>
        Auditoría y Resultados de Clasificación
      </h2>

      <!-- KPIs -->
      <div class="kpi-grid">
        <div class="kpi-card">
          <div class="kpi-title">Total Partidas</div>
          <div class="kpi-value" id="kpiTotal">0</div>
          <div class="kpi-sub" id="kpiTotalMonto">S/ 0.00</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-title">Por Histórico</div>
          <div class="kpi-value hist" id="kpiHist">0</div>
          <div class="kpi-sub" id="kpiHistPct">0%</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-title">Por Reglas KIMSAV</div>
          <div class="kpi-value reglas" id="kpiReglas">0</div>
          <div class="kpi-sub" id="kpiReglasPct">0%</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-title">Por Hoja / Auxiliar</div>
          <div class="kpi-value hoja" id="kpiHoja">0</div>
          <div class="kpi-sub" id="kpiHojaPct">0%</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-title">Revisión Manual</div>
          <div class="kpi-value rev" id="kpiRev">0</div>
          <div class="kpi-sub">Score Mín: <span id="kpiScoreMin">0</span></div>
        </div>
        <div class="kpi-card">
          <div class="kpi-title">Base Histórica</div>
          <div class="kpi-value" id="kpiHistBase">0 &rarr; 1</div>
          <div class="kpi-sub">Proyectos acumulados</div>
        </div>
      </div>

      <!-- Table Sample -->
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Descripción</th>
              <th>Unidad</th>
              <th>Metrado</th>
              <th>Subtotal (S/)</th>
              <th>Especialidad</th>
              <th>Subespecialidad</th>
              <th>Código</th>
              <th>Método</th>
              <th>Score</th>
              <th>Ref. Histórica</th>
            </tr>
          </thead>
          <tbody id="itemsTableBody"></tbody>
        </table>
      </div>

      <!-- Download Bar -->
      <div class="download-bar">
        <div>
          <strong style="font-size: 15px;">Archivo clasificado generado</strong>
          <p style="font-size: 12px; color: var(--text-muted); margin-top: 2px;" id="downloadSubtext">Incluye scores, métodos y proyectos de referencia.</p>
        </div>
        <div style="display: flex; gap: 10px;">
          <button type="button" class="btn-download" style="background: #4f46e5; border: none; cursor: pointer;" onclick="window.scrollTo({top: 0, behavior: 'smooth'}); fileInput.click();">
            <svg width="18" height="18" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4"></path></svg>
            Cargar Siguiente Presupuesto
          </button>
          <a href="#" class="btn-download" id="downloadBtn">
            <svg width="18" height="18" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
            Descargar Excel Clasificado
          </a>
        </div>
      </div>
    </div>
  </div>

  <script>
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');
    const proyectoInput = document.getElementById('proyectoInput');
    const uploadForm = document.getElementById('uploadForm');
    const btnSubmit = document.getElementById('btnSubmit');
    const spinner = document.getElementById('spinner');
    const btnText = document.getElementById('btnText');

    // Drag & drop handlers
    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
    dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
      if (e.dataTransfer.files.length) {
        fileInput.files = e.dataTransfer.files;
        updateFileInfo();
      }
    });

    fileInput.addEventListener('change', updateFileInfo);

    function updateFileInfo() {
      if (fileInput.files.length) {
        const file = fileInput.files[0];
        fileInfo.textContent = `Archivo seleccionado: ${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`;
        if (!proyectoInput.value) {
          const autoName = file.name.replace(/\.[^/.]+$/, "");
          proyectoInput.value = autoName;
        }
      }
    }

    // Load History
    async function loadHistory() {
      try {
        const res = await fetch('/api/historico');
        const data = await res.json();
        const listEl = document.getElementById('historyList');
        const headerCount = document.getElementById('headerProjectCount');

        headerCount.textContent = `${data.total_proyectos} ${data.total_proyectos === 1 ? 'proyecto' : 'proyectos'}`;

        if (data.proyectos.length === 0) {
          listEl.innerHTML = '<div class="empty-state">La base histórica está vacía. El primer presupuesto que proceses inaugurará el historial.</div>';
          return;
        }

        listEl.innerHTML = data.proyectos.map(p => `
          <div class="history-item">
            <div>
              <div class="history-item-name">${p.nombre}</div>
              <div class="history-item-meta">Área: ${p.area ? p.area.toLocaleString() + ' m²' : 'N/D'}</div>
            </div>
            <div class="history-item-count">${p.partidas.toLocaleString()} partidas</div>
          </div>
        `).join('');
      } catch (err) {
        console.error("Error al cargar historial:", err);
      }
    }

    // Submit Form
    uploadForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      if (!fileInput.files.length) {
        alert("Por favor selecciona un archivo Excel de presupuesto.");
        return;
      }

      btnSubmit.disabled = true;
      spinner.style.display = 'block';
      btnText.textContent = 'Procesando y Clasificando...';

      const formData = new FormData(uploadForm);
      try {
        const res = await fetch('/api/procesar', {
          method: 'POST',
          body: formData
        });

        const data = await res.json();
        if (!res.ok) {
          throw new Error(data.error || 'Error al procesar el archivo');
        }

        displayResults(data);
        await loadHistory();

        // Resetear formulario para permitir cargar el siguiente presupuesto de inmediato
        fileInput.value = '';
        fileInfo.innerHTML = '<span style="color:#10b981;">✓ Proyecto procesado y guardado. ¡Listo para cargar el siguiente!</span>';
        proyectoInput.value = '';
        document.getElementById('tipoInput').value = '';
        document.getElementById('departamentoInput').value = '';
        document.getElementById('areaInput').value = '';
        document.getElementById('forzarInput').checked = false;
      } catch (err) {
        alert("Error: " + err.message);
      } finally {
        btnSubmit.disabled = false;
        spinner.style.display = 'none';
        btnText.textContent = 'Clasificar e Incorporar al Histórico';
      }
    });

    function displayResults(data) {
      document.getElementById('resultsSection').style.display = 'block';
      document.getElementById('resultsSection').scrollIntoView({ behavior: 'smooth' });

      const total = data.total_partidas;
      document.getElementById('kpiTotal').textContent = total.toLocaleString();
      document.getElementById('kpiTotalMonto').textContent = 'S/ ' + data.suma_calculada.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      
      document.getElementById('kpiHist').textContent = data.por_historico;
      document.getElementById('kpiHistPct').textContent = total ? ((data.por_historico / total) * 100).toFixed(1) + '%' : '0%';

      document.getElementById('kpiReglas').textContent = data.por_reglas;
      document.getElementById('kpiReglasPct').textContent = total ? ((data.por_reglas / total) * 100).toFixed(1) + '%' : '0%';

      document.getElementById('kpiHoja').textContent = data.por_hoja;
      document.getElementById('kpiHojaPct').textContent = total ? ((data.por_hoja / total) * 100).toFixed(1) + '%' : '0%';

      document.getElementById('kpiRev').textContent = data.para_revision;
      document.getElementById('kpiScoreMin').textContent = data.score_min.toFixed(1);

      document.getElementById('kpiHistBase').innerHTML = `${data.proyectos_hist_antes} &rarr; <strong>${data.proyectos_hist_despues}</strong>`;

      // Download link
      const downloadBtn = document.getElementById('downloadBtn');
      downloadBtn.href = `/api/descargar?archivo=${encodeURIComponent(data.archivo_clasificado)}`;
      document.getElementById('downloadSubtext').textContent = `Archivo: ${data.archivo_clasificado}`;

      // Table Sample (Top 50 items)
      const tbody = document.getElementById('itemsTableBody');
      tbody.innerHTML = data.muestra_items.map((it, idx) => `
        <tr>
          <td style="color: #64748b;">${idx + 1}</td>
          <td style="font-weight: 500;">${it.descripcion}</td>
          <td>${it.unidad || '-'}</td>
          <td>${it.metrado !== null ? it.metrado : '-'}</td>
          <td style="font-family: 'JetBrains Mono', monospace;">S/ ${(it.subtotal_soles || 0).toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
          <td><span style="font-weight: 700; color: #a5b4fc;">${it.especialidad}</span></td>
          <td>${it.subespecialidad || '-'}</td>
          <td style="font-family: 'JetBrains Mono', monospace; font-weight: 600;">${it.codigo || '-'}</td>
          <td><span class="badge-method ${it.metodo_clasificacion || 'sin_resolver'}">${it.metodo_clasificacion || 'sin resolver'}</span></td>
          <td style="font-weight: 600;">${it.score_clasificacion !== null ? it.score_clasificacion : '-'}</td>
          <td style="font-size: 11px; color: #94a3b8;">${it.ref_proyecto ? it.ref_proyecto : '-'}</td>
        </tr>
      `).join('');
    }

    // Initial load
    loadHistory();
  </script>
</body>
</html>
"""

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/api/historico", methods=["GET"])
def get_historico():
    hist = cargar_historico(BD_HISTORICO_PATH)
    proyectos = []
    for nombre, data in hist.items():
        proyectos.append({
            "nombre": nombre,
            "partidas": len(data.get("items", [])),
            "area": data.get("area"),
        })
    return jsonify({
        "total_proyectos": len(proyectos),
        "proyectos": proyectos
    })

@app.route("/api/procesar", methods=["POST"])
def procesar():
    if "file" not in request.files:
        return jsonify({"error": "No se envió ningún archivo."}), 400
    
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Archivo no seleccionado."}), 400

    filename = secure_filename(file.filename)
    upload_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    file.save(upload_path)

    proyecto_nombre = request.form.get("proyecto") or filename.rsplit(".", 1)[0]
    tipo = request.form.get("tipo") or None
    departamento = request.form.get("departamento") or None
    area_techada = float(request.form.get("area")) if request.form.get("area") else None
    sobrescribir = request.form.get("forzar") == "on" or request.form.get("forzar") == "true"

    try:
        resultado = procesar_presupuesto_incremental(
            path_excel=upload_path,
            path_bd_historico=BD_HISTORICO_PATH,
            nombre_proyecto=proyecto_nombre,
            area_techada=area_techada,
            tipo=tipo,
            departamento=departamento,
            sobrescribir_si_existe=sobrescribir,
            guardar_en_historico=True,
            exportar_clasificado=True
        )

        por_hist = sum(1 for it in resultado.items if it.get("metodo_clasificacion") == "historico")
        por_reglas = sum(1 for it in resultado.items if it.get("metodo_clasificacion") == "reglas")
        por_hoja = sum(1 for it in resultado.items if it.get("metodo_clasificacion") == "hoja")
        sin_resolver = sum(1 for it in resultado.items if it.get("metodo_clasificacion") == "sin_resolver" or it.get("especialidad") == "SIN_CLASIFICAR")
        para_revision = sum(1 for it in resultado.items if it.get("observaciones") == "Revisar manualmente" or it.get("subespecialidad") == "Revisar manualmente")

        scores = [it.get("score_clasificacion") for it in resultado.items if isinstance(it.get("score_clasificacion"), (int, float))]
        score_prom = float(sum(scores) / len(scores)) if scores else 0.0
        score_min = float(min(scores)) if scores else 0.0

        clasificado_filename = f"{proyecto_nombre}_clasificado.xlsx"

        # Muestra de partidas para la tabla del frontend
        muestra_items = []
        for it in resultado.items[:80]:
            muestra_items.append({
                "descripcion": it.get("descripcion"),
                "unidad": it.get("unidad"),
                "metrado": it.get("metrado"),
                "subtotal_soles": round(it.get("subtotal_soles", 0), 2),
                "especialidad": it.get("especialidad"),
                "subespecialidad": it.get("subespecialidad"),
                "codigo": it.get("codigo"),
                "score_clasificacion": it.get("score_clasificacion"),
                "metodo_clasificacion": it.get("metodo_clasificacion"),
                "ref_proyecto": it.get("ref_proyecto"),
            })

        return jsonify({
            "proyecto": resultado.proyecto,
            "total_partidas": len(resultado.items),
            "suma_calculada": round(resultado.cuadre.get("suma_calculada", 0), 2),
            "por_historico": por_hist,
            "por_reglas": por_reglas,
            "por_hoja": por_hoja,
            "sin_resolver": sin_resolver,
            "para_revision": para_revision,
            "score_prom": score_prom,
            "score_min": score_min,
            "proyectos_hist_antes": resultado.proyectos_hist_antes,
            "proyectos_hist_despues": resultado.proyectos_hist_despues,
            "archivo_clasificado": clasificado_filename,
            "muestra_items": muestra_items,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/descargar", methods=["GET"])
def descargar():
    archivo = request.args.get("archivo")
    if not archivo:
        return "Archivo no especificado", 400
    
    # Prevenir Directory Traversal de manera segura
    clean_name = os.path.basename(archivo)
    target_path = os.path.join(app.config["UPLOAD_FOLDER"], clean_name)
    
    if not os.path.exists(target_path):
        # Probar con secure_filename como respaldo
        fallback_path = os.path.join(app.config["UPLOAD_FOLDER"], secure_filename(clean_name))
        if os.path.exists(fallback_path):
            target_path = fallback_path
        else:
            return "Archivo no encontrado", 404

    return send_file(target_path, as_attachment=True, download_name=clean_name)

if __name__ == "__main__":
    print("\n" + "="*60)
    print("Servidor KIMSAV Web UI activo en: http://localhost:5000")
    print("="*60 + "\n")
    app.run(host="0.0.0.0", port=5000, debug=False)
