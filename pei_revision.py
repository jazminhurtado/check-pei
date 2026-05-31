"""
CEPLAN Checker — Revisora de PEI con IA
Desarrollado con: Python + Streamlit + Anthropic Claude API
"""

import streamlit as st
import anthropic
import io
import json
from datetime import datetime

# ─── Intentar importar librerías opcionales ────────────────────────────────
try:
    import docx
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import PyPDF2
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.colors import HexColor
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# ─── Configuración de página ────────────────────────────────────────────────
st.set_page_config(
    page_title="CEPLAN Checker",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CSS personalizado ──────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Sora', sans-serif;
    }
    
    .main { background-color: #F8F9FB; }
    
    .header-banner {
        background: linear-gradient(135deg, #1a3a5c 0%, #0d2137 60%, #102040 100%);
        color: white;
        padding: 2.5rem 2rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        border-left: 6px solid #E8B84B;
        box-shadow: 0 8px 32px rgba(26,58,92,0.18);
    }
    .header-banner h1 {
        font-size: 2rem;
        font-weight: 700;
        margin: 0 0 0.3rem 0;
        letter-spacing: -0.5px;
    }
    .header-banner p {
        font-size: 1rem;
        opacity: 0.75;
        margin: 0;
    }
    .header-badge {
        display: inline-block;
        background: #E8B84B;
        color: #1a3a5c;
        font-weight: 700;
        font-size: 0.7rem;
        padding: 3px 10px;
        border-radius: 20px;
        margin-bottom: 0.8rem;
        letter-spacing: 1px;
        text-transform: uppercase;
    }
    
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 1.4rem 1.2rem;
        text-align: center;
        border: 1px solid #E5E9F0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .metric-card .value {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1a3a5c;
        font-family: 'IBM Plex Mono', monospace;
    }
    .metric-card .label {
        font-size: 0.78rem;
        color: #6B7898;
        text-transform: uppercase;
        letter-spacing: 0.8px;
    }
    
    .semaforo-verde   { background: #E8F8F0; border: 2px solid #2ECC71; border-radius: 10px; padding: 1rem 1.2rem; margin: 0.5rem 0; }
    .semaforo-amarillo{ background: #FFFBEA; border: 2px solid #F1C40F; border-radius: 10px; padding: 1rem 1.2rem; margin: 0.5rem 0; }
    .semaforo-rojo    { background: #FEF0EE; border: 2px solid #E74C3C; border-radius: 10px; padding: 1rem 1.2rem; margin: 0.5rem 0; }
    
    .seccion-titulo {
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        color: #1a3a5c;
        background: #EEF2F8;
        border-left: 4px solid #1a3a5c;
        padding: 6px 14px;
        border-radius: 0 6px 6px 0;
        margin: 1.5rem 0 0.8rem 0;
        display: inline-block;
    }
    
    .info-box {
        background: #EEF6FF;
        border: 1px solid #B8D4F5;
        border-radius: 10px;
        padding: 1rem 1.3rem;
        font-size: 0.9rem;
        color: #1a3a5c;
    }
    
    div[data-testid="stFileUploader"] {
        border: 2px dashed #A8C0DC;
        border-radius: 12px;
        background: #F5F9FF;
    }
    
    .stButton > button {
        background: linear-gradient(90deg, #1a3a5c, #2556A0);
        color: white;
        border: none;
        padding: 0.6rem 2rem;
        font-weight: 600;
        font-family: 'Sora', sans-serif;
        border-radius: 8px;
        font-size: 0.95rem;
        transition: all 0.2s;
        box-shadow: 0 4px 12px rgba(26,58,92,0.2);
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(26,58,92,0.3);
    }
    
    .footer-note {
        text-align: center;
        color: #9BA5BC;
        font-size: 0.78rem;
        margin-top: 3rem;
        padding-top: 1.5rem;
        border-top: 1px solid #E5E9F0;
    }
</style>
""", unsafe_allow_html=True)


# ─── Funciones de extracción de texto ───────────────────────────────────────

def extraer_texto_docx(archivo_bytes: bytes) -> str:
    if not DOCX_AVAILABLE:
        return ""
    doc = docx.Document(io.BytesIO(archivo_bytes))
    parrafos = []
    for para in doc.paragraphs:
        texto = para.text.strip()
        if texto:
            parrafos.append(texto)
    return "\n".join(parrafos)


def extraer_texto_pdf(archivo_bytes: bytes) -> str:
    if not PDF_AVAILABLE:
        return ""
    reader = PyPDF2.PdfReader(io.BytesIO(archivo_bytes))
    texto_total = []
    for pagina in reader.pages:
        texto = pagina.extract_text()
        if texto:
            texto_total.append(texto)
    return "\n".join(texto_total)


def extraer_texto_txt(archivo_bytes: bytes) -> str:
    return archivo_bytes.decode("utf-8", errors="ignore")


# ─── Función principal de análisis IA ───────────────────────────────────────

def analizar_pei_con_ia(texto_pei: str, entidad: str) -> dict:
    """Llama a la API de Anthropic para analizar el PEI."""

    cliente = anthropic.Anthropic()

    prompt_sistema = """Eres un experto revisor de Planes Estratégicos Institucionales (PEI) del Estado peruano, con profundo conocimiento de la normativa CEPLAN.

Tu tarea es revisar el documento PEI que se te proporciona y generar un informe estructurado en JSON con el siguiente formato EXACTO (responde SOLO con el JSON, sin texto adicional ni backticks):

{
  "puntaje_global": <número 0-100>,
  "nivel_riesgo": "<BAJO|MEDIO|ALTO>",
  "resumen_ejecutivo": "<2-3 oraciones sobre el estado general del PEI>",
  "componentes": {
    "mision": {
      "estado": "<CORRECTO|REVISAR|ERROR>",
      "puntaje": <0-100>,
      "observaciones": ["<obs1>", "<obs2>"],
      "recomendaciones": ["<rec1>", "<rec2>"]
    },
    "situacion_futura_deseada": {
      "estado": "<CORRECTO|REVISAR|ERROR>",
      "puntaje": <0-100>,
      "observaciones": ["<obs1>"],
      "recomendaciones": ["<rec1>"]
    },
    "objetivos_estrategicos_institucionales": {
      "estado": "<CORRECTO|REVISAR|ERROR>",
      "puntaje": <0-100>,
      "observaciones": ["<obs1>"],
      "recomendaciones": ["<rec1>"]
    },
    "acciones_estrategicas_institucionales": {
      "estado": "<CORRECTO|REVISAR|ERROR>",
      "puntaje": <0-100>,
      "observaciones": ["<obs1>"],
      "recomendaciones": ["<rec1>"]
    },
    "indicadores": {
      "estado": "<CORRECTO|REVISAR|ERROR>",
      "puntaje": <0-100>,
      "observaciones": ["<obs1>"],
      "recomendaciones": ["<rec1>"]
    },
    "ruta_estrategica": {
      "estado": "<CORRECTO|REVISAR|ERROR>",
      "puntaje": <0-100>,
      "observaciones": ["<obs1>"],
      "recomendaciones": ["<rec1>"]
    }
  },
  "errores_criticos": ["<error crítico 1>", "<error crítico 2>"],
  "fortalezas": ["<fortaleza 1>", "<fortaleza 2>"],
  "listo_para_ceplan": <true|false>
}

Evalúa con rigor según la Guía CEPLAN para el Planeamiento Institucional vigente. Si el documento está incompleto, señálalo claramente. Si falta algún componente, márcalo como ERROR."""

    prompt_usuario = f"""Entidad: {entidad if entidad else 'No especificada'}

DOCUMENTO PEI A REVISAR:
{texto_pei[:8000]}"""

    respuesta = cliente.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2500,
        system=prompt_sistema,
        messages=[{"role": "user", "content": prompt_usuario}]
    )

    texto_respuesta = respuesta.content[0].text.strip()

    # Limpiar posibles backticks
    texto_respuesta = texto_respuesta.replace("```json", "").replace("```", "").strip()

    resultado = json.loads(texto_respuesta)
    return resultado


# ─── Generación de reporte PDF ───────────────────────────────────────────────

def generar_reporte_pdf(analisis: dict, entidad: str) -> bytes:
    if not REPORTLAB_AVAILABLE:
        return None

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=50, leftMargin=50, topMargin=60, bottomMargin=50)

    estilos = getSampleStyleSheet()
    navy = HexColor("#1a3a5c")
    dorado = HexColor("#E8B84B")

    estilo_titulo = ParagraphStyle("titulo", parent=estilos["Title"], textColor=navy, fontSize=20, spaceAfter=6, fontName="Helvetica-Bold")
    estilo_subtitulo = ParagraphStyle("sub", parent=estilos["Normal"], textColor=HexColor("#555"), fontSize=10, spaceAfter=16)
    estilo_seccion = ParagraphStyle("sec", parent=estilos["Heading2"], textColor=navy, fontSize=12, spaceBefore=14, spaceAfter=4, fontName="Helvetica-Bold")
    estilo_normal = ParagraphStyle("norm", parent=estilos["Normal"], fontSize=9.5, leading=14, spaceAfter=4)
    estilo_obs = ParagraphStyle("obs", parent=estilos["Normal"], fontSize=9, leading=13, leftIndent=12, textColor=HexColor("#444"))

    elementos = []
    elementos.append(Paragraph("🏛️ CEPLAN Checker", estilo_titulo))
    elementos.append(Paragraph(f"Informe de Revisión del PEI — {entidad or 'Entidad no especificada'}", estilo_subtitulo))
    elementos.append(Paragraph(f"Generado el {datetime.now().strftime('%d/%m/%Y a las %H:%M')}", estilo_subtitulo))
    elementos.append(Spacer(1, 10))

    # Resumen ejecutivo
    elementos.append(Paragraph("RESUMEN EJECUTIVO", estilo_seccion))
    elementos.append(Paragraph(analisis.get("resumen_ejecutivo", "—"), estilo_normal))
    elementos.append(Spacer(1, 6))

    # Tabla de puntaje
    puntaje = analisis.get("puntaje_global", 0)
    riesgo = analisis.get("nivel_riesgo", "—")
    listo = "SÍ ✓" if analisis.get("listo_para_ceplan") else "NO ✗"
    datos_tabla = [
        ["PUNTAJE GLOBAL", "NIVEL DE RIESGO", "LISTO PARA CEPLAN"],
        [f"{puntaje}/100", riesgo, listo],
    ]
    tabla = Table(datos_tabla, colWidths=[150, 150, 150])
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("FONTSIZE", (0, 1), (-1, 1), 13),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("BACKGROUND", (0, 1), (-1, 1), HexColor("#EEF2F8")),
        ("ROWBACKGROUNDS", (0, 1), (-1, 1), [HexColor("#EEF2F8")]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    elementos.append(tabla)
    elementos.append(Spacer(1, 14))

    # Componentes
    elementos.append(Paragraph("ANÁLISIS POR COMPONENTE", estilo_seccion))
    nombres = {
        "mision": "Misión Institucional",
        "situacion_futura_deseada": "Situación Futura Deseada",
        "objetivos_estrategicos_institucionales": "Objetivos Estratégicos (OEI)",
        "acciones_estrategicas_institucionales": "Acciones Estratégicas (AEI)",
        "indicadores": "Indicadores",
        "ruta_estrategica": "Ruta Estratégica"
    }
    iconos_estado = {"CORRECTO": "✅", "REVISAR": "⚠️", "ERROR": "❌"}

    for clave, nombre in nombres.items():
        comp = analisis.get("componentes", {}).get(clave, {})
        if comp:
            estado = comp.get("estado", "—")
            icono = iconos_estado.get(estado, "•")
            puntaje_comp = comp.get("puntaje", 0)
            elementos.append(Paragraph(f"{icono} {nombre} — {estado} ({puntaje_comp}/100)", ParagraphStyle("comp", parent=estilos["Normal"], fontSize=10, fontName="Helvetica-Bold", textColor=navy, spaceBefore=8, spaceAfter=3)))

            for obs in comp.get("observaciones", []):
                elementos.append(Paragraph(f"• {obs}", estilo_obs))
            for rec in comp.get("recomendaciones", []):
                elementos.append(Paragraph(f"→ {rec}", ParagraphStyle("rec", parent=estilos["Normal"], fontSize=9, leading=13, leftIndent=12, textColor=HexColor("#1a6a3a"))))

    # Errores críticos y fortalezas
    errores = analisis.get("errores_criticos", [])
    if errores:
        elementos.append(Spacer(1, 8))
        elementos.append(Paragraph("ERRORES CRÍTICOS", estilo_seccion))
        for e in errores:
            elementos.append(Paragraph(f"❌ {e}", ParagraphStyle("err", parent=estilos["Normal"], fontSize=9.5, leading=14, textColor=HexColor("#C0392B"))))

    fortalezas = analisis.get("fortalezas", [])
    if fortalezas:
        elementos.append(Spacer(1, 8))
        elementos.append(Paragraph("FORTALEZAS DEL PEI", estilo_seccion))
        for f in fortalezas:
            elementos.append(Paragraph(f"✅ {f}", ParagraphStyle("fort", parent=estilos["Normal"], fontSize=9.5, leading=14, textColor=HexColor("#1a6a3a"))))

    elementos.append(Spacer(1, 20))
    elementos.append(Paragraph("Generado por CEPLAN Checker — Sistema IA de Revisión de PEI | Basado en normativa CEPLAN vigente", ParagraphStyle("foot", parent=estilos["Normal"], fontSize=8, textColor=HexColor("#999"), alignment=1)))

    doc.build(elementos)
    return buffer.getvalue()


# ─── Interfaz principal ──────────────────────────────────────────────────────

def main():

    # Header
    st.markdown("""
    <div class="header-banner">
        <div class="header-badge">Beta v1.0</div>
        <h1>🏛️ CEPLAN Checker</h1>
        <p>Sistema de revisión inteligente de Planes Estratégicos Institucionales (PEI) del Estado peruano</p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Configuración")

        api_key = st.text_input(
            "API Key de Anthropic",
            type="password",
            placeholder="sk-ant-...",
            help="Obtén tu clave en console.anthropic.com"
        )

        entidad = st.text_input("🏢 Nombre de la Entidad", placeholder="Ej: Municipalidad Distrital de San Bartolo")

        st.markdown("---")
        st.markdown("### 📋 ¿Qué analiza el sistema?")
        componentes_lista = [
            "✅ Misión institucional",
            "✅ Situación futura deseada",
            "✅ Objetivos Estratégicos (OEI)",
            "✅ Acciones Estratégicas (AEI)",
            "✅ Indicadores",
            "✅ Ruta estratégica",
        ]
        for c in componentes_lista:
            st.markdown(f"<small>{c}</small>", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("<small style='color:#999'>Basado en la Guía CEPLAN para el Planeamiento Institucional vigente</small>", unsafe_allow_html=True)

        #st.markdown("---")
        #st.markdown("### 📦 Dependencias requeridas")
        #st.code("pip install streamlit anthropic python-docx PyPDF2 reportlab", language="bash")

    # Área principal
    col1, col2 = st.columns([3, 2])

    with col1:
        st.markdown('<div class="seccion-titulo">📁 Cargar Documento PEI</div>', unsafe_allow_html=True)

        formatos_soportados = []
        if DOCX_AVAILABLE:
            formatos_soportados.append("docx")
        if PDF_AVAILABLE:
            formatos_soportados.append("pdf")
        formatos_soportados.append("txt")

        archivo = st.file_uploader(
            "Sube tu archivo PEI",
            type=formatos_soportados,
            help="Formatos: " + ", ".join([f".{f}" for f in formatos_soportados])
        )

        if not DOCX_AVAILABLE or not PDF_AVAILABLE:
            st.warning(f"⚠️ Algunas librerías no están instaladas. Instala con: `pip install python-docx PyPDF2`")

        # Texto manual como alternativa
        with st.expander("✍️ O pega el texto del PEI directamente"):
            texto_manual = st.text_area(
                "Texto del PEI",
                height=200,
                placeholder="Pega aquí el contenido de tu Plan Estratégico Institucional..."
            )

    with col2:
        st.markdown('<div class="seccion-titulo">🚀 Iniciar Revisión</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class="info-box">
            <strong>¿Cómo funciona?</strong><br><br>
            1️⃣ Sube tu documento PEI<br>
            2️⃣ La IA analiza cada componente<br>
            3️⃣ Recibes un semáforo de errores<br>
            4️⃣ Descarga el informe con observaciones
        </div>
        """, unsafe_allow_html=True)

        btn_analizar = st.button("🔍 Analizar PEI con IA", use_container_width=True)

    # ─── Procesamiento ───────────────────────────────────────────────────────
    if btn_analizar:
        if not api_key:
            st.error("⚠️ Ingresa tu API Key de Anthropic en el panel lateral para continuar.")
            return

        # Extraer texto
        texto_pei = ""

        if archivo:
            bytes_archivo = archivo.read()
            extension = archivo.name.split(".")[-1].lower()

            with st.spinner("📄 Extrayendo texto del documento..."):
                if extension == "docx" and DOCX_AVAILABLE:
                    texto_pei = extraer_texto_docx(bytes_archivo)
                elif extension == "pdf" and PDF_AVAILABLE:
                    texto_pei = extraer_texto_pdf(bytes_archivo)
                elif extension == "txt":
                    texto_pei = extraer_texto_txt(bytes_archivo)

        elif texto_manual and texto_manual.strip():
            texto_pei = texto_manual.strip()

        if not texto_pei or len(texto_pei) < 50:
            st.error("❌ No se pudo extraer texto suficiente. Verifica el archivo o pega el texto manualmente.")
            return

        st.success(f"✅ Texto extraído: {len(texto_pei):,} caracteres")

        # Configurar cliente con API key del usuario
        import os
        os.environ["ANTHROPIC_API_KEY"] = api_key

        # Analizar con IA
        with st.spinner("🤖 Analizando tu PEI con IA... esto puede tomar 15-30 segundos"):
            try:
                analisis = analizar_pei_con_ia(texto_pei, entidad)
            except json.JSONDecodeError as e:
                st.error(f"❌ Error al procesar respuesta de la IA: {e}")
                return
            except Exception as e:
                st.error(f"❌ Error al conectar con la API: {str(e)}")
                return

        # ─── Mostrar resultados ───────────────────────────────────────────────

        st.markdown("---")
        st.markdown("## 📊 Resultados del Análisis")

        # Métricas principales
        puntaje = analisis.get("puntaje_global", 0)
        riesgo = analisis.get("nivel_riesgo", "—")
        listo = analisis.get("listo_para_ceplan", False)

        col_a, col_b, col_c, col_d = st.columns(4)
        with col_a:
            st.markdown(f'<div class="metric-card"><div class="value">{puntaje}</div><div class="label">Puntaje Global /100</div></div>', unsafe_allow_html=True)
        with col_b:
            color_riesgo = {"BAJO": "#2ECC71", "MEDIO": "#F1C40F", "ALTO": "#E74C3C"}.get(riesgo, "#888")
            st.markdown(f'<div class="metric-card"><div class="value" style="color:{color_riesgo}">{riesgo}</div><div class="label">Nivel de Riesgo</div></div>', unsafe_allow_html=True)
        with col_c:
            icono_listo = "✅" if listo else "❌"
            st.markdown(f'<div class="metric-card"><div class="value">{icono_listo}</div><div class="label">Listo para CEPLAN</div></div>', unsafe_allow_html=True)
        with col_d:
            n_errores = len(analisis.get("errores_criticos", []))
            st.markdown(f'<div class="metric-card"><div class="value" style="color:#E74C3C">{n_errores}</div><div class="label">Errores Críticos</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Resumen ejecutivo
        st.markdown('<div class="seccion-titulo">📝 Resumen Ejecutivo</div>', unsafe_allow_html=True)
        st.info(analisis.get("resumen_ejecutivo", "—"))

        # Semáforo por componente
        st.markdown('<div class="seccion-titulo">🚦 Semáforo por Componente</div>', unsafe_allow_html=True)

        nombres_comp = {
            "mision": "📌 Misión Institucional",
            "situacion_futura_deseada": "🎯 Situación Futura Deseada",
            "objetivos_estrategicos_institucionales": "📈 Objetivos Estratégicos (OEI)",
            "acciones_estrategicas_institucionales": "⚙️ Acciones Estratégicas (AEI)",
            "indicadores": "📊 Indicadores",
            "ruta_estrategica": "🗺️ Ruta Estratégica"
        }

        estilos_estado = {
            "CORRECTO": ("semaforo-verde", "✅ CORRECTO"),
            "REVISAR": ("semaforo-amarillo", "⚠️ REVISAR"),
            "ERROR": ("semaforo-rojo", "❌ ERROR"),
        }

        componentes_data = analisis.get("componentes", {})

        for clave, nombre in nombres_comp.items():
            comp = componentes_data.get(clave, {})
            if not comp:
                continue

            estado = comp.get("estado", "REVISAR")
            puntaje_c = comp.get("puntaje", 0)
            clase_css, etiqueta = estilos_estado.get(estado, ("semaforo-amarillo", "⚠️ REVISAR"))

            with st.expander(f"{nombre} — {etiqueta} ({puntaje_c}/100)"):
                st.markdown(f'<div class="{clase_css}">', unsafe_allow_html=True)

                obs = comp.get("observaciones", [])
                if obs:
                    st.markdown("**🔍 Observaciones:**")
                    for o in obs:
                        st.markdown(f"• {o}")

                recs = comp.get("recomendaciones", [])
                if recs:
                    st.markdown("**💡 Recomendaciones:**")
                    for r in recs:
                        st.markdown(f"→ {r}")

                st.markdown('</div>', unsafe_allow_html=True)

        # Errores críticos
        errores_criticos = analisis.get("errores_criticos", [])
        if errores_criticos:
            st.markdown('<div class="seccion-titulo">🚨 Errores Críticos</div>', unsafe_allow_html=True)
            for e in errores_criticos:
                st.error(f"❌ {e}")

        # Fortalezas
        fortalezas = analisis.get("fortalezas", [])
        if fortalezas:
            st.markdown('<div class="seccion-titulo">💪 Fortalezas del PEI</div>', unsafe_allow_html=True)
            for f in fortalezas:
                st.success(f"✅ {f}")

        # Descargar reporte PDF
        st.markdown("---")
        st.markdown('<div class="seccion-titulo">📥 Descargar Informe</div>', unsafe_allow_html=True)

        if REPORTLAB_AVAILABLE:
            pdf_bytes = generar_reporte_pdf(analisis, entidad)
            if pdf_bytes:
                nombre_archivo = f"CEPLAN_Checker_{(entidad or 'entidad').replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                st.download_button(
                    label="📄 Descargar Reporte PDF",
                    data=pdf_bytes,
                    file_name=nombre_archivo,
                    mime="application/pdf"
                )
        else:
            # Fallback: descargar JSON
            json_str = json.dumps(analisis, ensure_ascii=False, indent=2)
            st.download_button(
                label="📋 Descargar Análisis (JSON)",
                data=json_str.encode("utf-8"),
                file_name=f"ceplan_analisis_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                mime="application/json"
            )
            st.info("Instala `reportlab` para generar reportes PDF: `pip install reportlab`")

    # Footer
    st.markdown("""
    <div class="footer-note">
        🏛️ CEPLAN Checker · Proyecto Design Thinking · Generación del Modelo de Negocio 2026<br>
        Desarrollado con Python · Streamlit · Anthropic Claude API
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
