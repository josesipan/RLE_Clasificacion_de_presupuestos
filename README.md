# KIMSAV - Sistema Incremental de Clasificación de Presupuestos

Sistema inteligente y continuo de extracción, clasificación y auditoría de presupuestos de obra civil y edificación según el catálogo maestro **KIMSAV**.

## 🚀 Características Principales

1. **Aprendizaje Incremental Continuo:**
   - Cada presupuesto procesado genera su archivo individual `<nombre>_clasificado.xlsx` y **se incorpora automáticamente** como referencia histórica en la hoja `BD_HISTORICO PPTO`.
   - La hoja `Clasificación` se mantiene intacta como catálogo maestro KIMSAV.

2. **Jerarquía Inteligente de Clasificación:**
   - **Histórico Acumulado:** Busca partidas similares en presupuestos anteriores comparando descripción, unidades y especialidad.
   - **Reglas KIMSAV (`classify_rules.py`):** Fallback robusto por coincidencia de palabras clave y pesos técnicos.
   - **Información Auxiliar:** Deducción por hoja de origen y herencia por desglose continuo.
   - **Revisión Manual:** Marcado automático cuando no hay suficiente evidencia.

3. **Control Contra Duplicados:**
   - Verificación de proyectos existentes para evitar duplicaciones accidentales en la base histórica.

4. **Interfaz Web Intuitiva (`localhost:5000`):**
   - Carga mediante *Drag & Drop*.
   - Auditoría visual en tiempo real (distribución por especialidad, método de clasificación, scores, ratios y proyectos históricos).
   - Descarga directa de archivos clasificados.

---

## 💻 Instalación y Requisitos

```bash
pip install openpyxl numpy flask
```

---

## 🖥️ Uso de la Interfaz Web (Recomendado)

Inicia el servidor local:

```bash
python app.py
```

Abre tu navegador en:
👉 **[http://localhost:5000](http://localhost:5000)**

---

## ⚙️ Uso desde la Terminal (CLI)

```bash
# Procesar un presupuesto y acumularlo en el histórico
python pipeline_auditor.py "Presupuesto_Nuevo.xlsx" --bd "BD_HISTORICO_BASE_VACIA.xlsx" --proyecto "Nombre_Proyecto"

# Reprocesar y sobrescribir un proyecto existente
python pipeline_auditor.py "Presupuesto_Nuevo.xlsx" --bd "BD_HISTORICO_BASE_VACIA.xlsx" --proyecto "Nombre_Proyecto" --forzar
```

---

## 🐍 Uso en Python / Google Colab

```python
from pipeline_auditor import procesar_presupuesto_incremental

# Procesar presupuesto
resultado = procesar_presupuesto_incremental(
    path_excel="Presupuesto.xlsx",
    path_bd_historico="BD_HISTORICO_BASE_VACIA.xlsx",
    nombre_proyecto="PROYECTO_A"
)
```
