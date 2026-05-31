# 🏛️ CEPLAN Checker

Sistema de revisión inteligente de Planes Estratégicos Institucionales (PEI) del Estado peruano, impulsado por IA (Claude de Anthropic).

---

## ¿Qué hace?

Analiza documentos PEI (.docx, .pdf, .txt) y genera un informe de observaciones con:

- **Semáforo** por componente: ✅ CORRECTO · ⚠️ REVISAR · ❌ ERROR
- **Puntaje global** del PEI (0–100)
- **Nivel de riesgo** (BAJO / MEDIO / ALTO)
- **Observaciones y recomendaciones** por cada componente (Misión, OEI, AEI, Indicadores, etc.)
- **Reporte PDF descargable** con todas las observaciones

---

## Instalación rápida

### 1. Clona o descarga este proyecto

```bash
# Crea una carpeta para el proyecto
mkdir ceplan_checker
cd ceplan_checker
```

### 2. Instala las dependencias

```bash
pip install -r requirements.txt
```

O manualmente:

```bash
pip install streamlit anthropic python-docx PyPDF2 reportlab
```

### 3. Obtén tu API Key de Anthropic

1. Ve a [console.anthropic.com](https://console.anthropic.com)
2. Crea una cuenta o inicia sesión
3. Ve a "API Keys" → "Create Key"
4. Copia tu clave (empieza con `sk-ant-...`)

> 💡 El primer mes tiene créditos gratis para pruebas.

### 4. Ejecuta la aplicación

```bash
streamlit run app.py
```

La app se abrirá automáticamente en tu navegador en `http://localhost:8501`

---

## Uso

1. Ingresa tu **API Key de Anthropic** en el panel lateral
2. Escribe el nombre de la **entidad** (opcional)
3. **Sube tu documento PEI** (.docx, .pdf) o pega el texto directamente
4. Haz clic en **"Analizar PEI con IA"**
5. Revisa el semáforo de resultados
6. **Descarga el informe PDF** con todas las observaciones

---

## Componentes que analiza

| Componente | Descripción |
|------------|-------------|
| **Misión** | Redacción, alineación normativa y claridad |
| **Situación futura deseada** | Coherencia con la visión institucional |
| **OEI** | Formulación de Objetivos Estratégicos Institucionales |
| **AEI** | Acciones Estratégicas Institucionales |
| **Indicadores** | Existencia, medibilidad y pertinencia |
| **Ruta estratégica** | Coherencia entre objetivos, acciones e indicadores |

---

## Stack tecnológico

| Componente | Tecnología |
|------------|-----------|
| Frontend / UI | Streamlit |
| IA revisora | Anthropic Claude (claude-sonnet-4) |
| Lectura .docx | python-docx |
| Lectura .pdf | PyPDF2 |
| Generación PDF | ReportLab |

---

## Estructura del proyecto

```
ceplan_checker/
├── app.py              # Aplicación principal
├── requirements.txt    # Dependencias
└── README.md           # Este archivo
```

---

## Basado en

- Guía CEPLAN para el Planeamiento Institucional (Directiva vigente)
- Metodología Design Thinking (Proyecto Final — Generación del Modelo de Negocio 2026)

---

*Desarrollado por Jazmin Hurtado · Proyecto Final PEI · 2026*
