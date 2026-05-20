import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
from datetime import datetime
import os
import shutil
import zipfile
from io import BytesIO
from pathlib import Path
from dataclasses import dataclass
from typing import Optional


# ─────────────────────────────────────────────
# Configuración de página
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Generador de Facturas",
    page_icon="🧾",
    layout="centered",
)

st.markdown("""
<style>
    /* Fuente y fondo general */
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&family=DM+Mono&display=swap');

    html, body, [class*="css"] {
        font-family: 'DM Sans', sans-serif;
    }

    /* Encabezado */
    .app-header {
        background: linear-gradient(135deg, #1a1f36 0%, #2d3561 100%);
        color: white;
        border-radius: 16px;
        padding: 2rem 2.5rem;
        margin-bottom: 2rem;
    }
    .app-header h1 { margin: 0; font-size: 1.8rem; font-weight: 600; }
    .app-header p  { margin: 0.4rem 0 0; opacity: 0.7; font-size: 0.95rem; }

    /* Tarjetas de sección */
    .section-card {
        background: #f8f9fc;
        border: 1px solid #e4e7f0;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1.2rem;
    }
    .section-title {
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #6b7280;
        margin-bottom: 0.8rem;
    }

    /* Métricas de resultado */
    .metric-row { display: flex; gap: 1rem; margin: 1rem 0; }
    .metric-box {
        flex: 1;
        background: white;
        border: 1px solid #e4e7f0;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
    }
    .metric-box .value { font-size: 2rem; font-weight: 700; color: #1a1f36; }
    .metric-box .label { font-size: 0.8rem; color: #6b7280; margin-top: 0.2rem; }

    /* Botón primario */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #1a1f36, #2d3561);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        font-weight: 600;
        width: 100%;
        transition: opacity 0.2s;
    }
    div.stButton > button[kind="primary"]:hover { opacity: 0.88; }

    /* Barra de progreso */
    .stProgress > div > div > div { background: #2d3561; border-radius: 4px; }

    /* Alertas */
    .stAlert { border-radius: 10px; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Constantes y configuración de campos
# ─────────────────────────────────────────────

OUTPUT_DIR = Path("output_facturas")

# Mapa: campo_plantilla → (x, y)
FIELD_POSITIONS: dict[str, tuple[int, int]] = {
    "Factura electrónica de venta N°": (220, 118),
    "Nombre":                           (150, 133),
    "Cédula":                           (150, 155),
    "Municipio":                        (150, 185),
    "Localidad":                        (150, 205),
    "Tipo de lectura":                  (150, 256),
    "Referencia de pago/NUI":           (490, 135),
    "Días facturados":                  (485, 339),
    "Deuda anterior":                   (485, 359),
    "Otros conceptos":                  (485, 382),
    "1+2 Valor total a cancelar":       (485, 404),
    "Fecha de emisión":                 (485, 425),
    "Pago oportuno":                    (485, 446),
    "Suspension a partir de":           (485, 468),
    "Cargos facturados Mes":            (485, 512),
    "Costo mensual prestación de servicios": (485, 533),
    "Valor refacturacion":              (485, 555),
    "Valor por mora":                   (485, 580),
    "Interés por mora":                 (485, 600),
    "Valor subsidio":                   (485, 620),
    "Total servicio":                   (485, 645),
}

# Columnas requeridas en el Excel
REQUIRED_COLUMNS = {
    "Factura", "Nombre", "Cedula", "Municipio", "Localidad",
    "Tipo de Lectura", "Referencia de pago/NUI", "Dias facturados",
    "Deuda anterior", "1+2 Valor total a cancelar", "Fecha de emision",
    "Pago oportuno", "Suspension a partir de", "Cargos facturados Mes",
    "Costo mensual preatacion de servicios", "Valor por mora",
    "Valor subsidio", "Total Servicio",
}


# ─────────────────────────────────────────────
# Helpers de formato
# ─────────────────────────────────────────────

def fmt_fecha(valor) -> str:
    """Formatea un valor como fecha DD/MM/YYYY."""
    if pd.notna(valor):
        if isinstance(valor, datetime):
            return valor.strftime("%d/%m/%Y")
        return str(valor)
    return ""


def fmt_moneda(valor) -> str:
    """Formatea un número como moneda colombiana: $1.234.567"""
    try:
        return f"${int(valor):,}".replace(",", ".")
    except (ValueError, TypeError):
        return "$0"


# ─────────────────────────────────────────────
# Lógica principal de generación
# ─────────────────────────────────────────────

def build_field_values(row: pd.Series) -> dict[str, str]:
    """Construye el diccionario campo → texto a insertar."""
    return {
        "Factura electrónica de venta N°":      str(row["Factura"]),
        "Nombre":                                str(row["Nombre"]),
        "Cédula":                                str(row["Cedula"]),
        "Municipio":                             str(row["Municipio"]),
        "Localidad":                             str(row["Localidad"]),
        "Tipo de lectura":                       str(row["Tipo de Lectura"]),
        "Referencia de pago/NUI":                str(row["Referencia de pago/NUI"]).strip(),
        "Días facturados":                       str(row["Dias facturados"]),
        "Deuda anterior":                        fmt_moneda(row["Deuda anterior"]),
        "Otros conceptos":                       "$0",
        "1+2 Valor total a cancelar":            fmt_moneda(row["1+2 Valor total a cancelar"]),
        "Fecha de emisión":                      fmt_fecha(row["Fecha de emision"]),
        "Pago oportuno":                         fmt_fecha(row["Pago oportuno"]),
        "Suspension a partir de":                fmt_fecha(row["Suspension a partir de"]),
        "Cargos facturados Mes":                 str(row["Cargos facturados Mes"]).strip(),
        "Costo mensual prestación de servicios": fmt_moneda(row["Costo mensual preatacion de servicios"]),
        "Valor refacturacion":                   "$0",
        "Valor por mora":                        fmt_moneda(row["Valor por mora"]),
        "Interés por mora":                      "$0",
        "Valor subsidio":                        fmt_moneda(row["Valor subsidio"]),
        "Total servicio":                        fmt_moneda(row["Total Servicio"]),
    }


def build_output_filename(row: pd.Series) -> str:
    """Genera el nombre del PDF a partir de los datos de la fila."""
    referencia = str(row.get("Referencia de pago/NUI", "SINREF")).strip()
    cargos = str(row.get("Cargos facturados Mes", "NA")).strip().upper().replace(" ", "_")
    return f"Factura_{referencia}_{cargos}.pdf"


def fill_pdf(
    row: pd.Series,
    template_path: Path,
    font_path: Path,
    output_dir: Path,
) -> Path:
    """Rellena la plantilla PDF con los datos de una fila y guarda el resultado."""
    doc = fitz.open(str(template_path))
    page = doc[0]
    fontname = "custom_font"
    page.insert_font(fontname=fontname, fontfile=str(font_path))

    field_values = build_field_values(row)

    for campo, (x, y) in FIELD_POSITIONS.items():
        texto = field_values.get(campo, "")
        page.insert_text((x, y), texto, fontsize=8.5, fontname=fontname, color=(0, 0, 0))

    output_path = output_dir / build_output_filename(row)
    doc.save(str(output_path))
    doc.close()
    return output_path


def validate_columns(df: pd.DataFrame) -> list[str]:
    """Retorna las columnas requeridas que faltan en el DataFrame."""
    return sorted(REQUIRED_COLUMNS - set(df.columns))


def create_zip(file_paths: list[Path]) -> bytes:
    """Empaqueta una lista de archivos en un ZIP en memoria."""
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in file_paths:
            zf.write(path, arcname=path.name)
    buffer.seek(0)
    return buffer.read()


def cleanup_output_dir() -> None:
    """Elimina el directorio de salida temporal."""
    shutil.rmtree(OUTPUT_DIR, ignore_errors=True)


# ─────────────────────────────────────────────
# Interfaz de usuario
# ─────────────────────────────────────────────

st.markdown("""
<div class="app-header">
    <h1>🧾 Generador de Facturas PDF</h1>
    <p>Carga tu Excel, plantilla PDF y fuente para generar todas las facturas en segundos.</p>
</div>
""", unsafe_allow_html=True)

# — Uploads —
st.markdown('<div class="section-card"><div class="section-title">📂 Archivos de entrada</div>', unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)
with col1:
    excel_file   = st.file_uploader("Excel con datos", type=["xlsx"], label_visibility="visible")
with col2:
    pdf_template = st.file_uploader("Plantilla PDF",   type=["pdf"],  label_visibility="visible")
with col3:
    font_file    = st.file_uploader("Fuente (.ttf)",   type=["ttf"],  label_visibility="visible")

st.markdown("</div>", unsafe_allow_html=True)

# Estado de sesión
if "generated_files" not in st.session_state:
    st.session_state.generated_files = []
if "zip_data" not in st.session_state:
    st.session_state.zip_data = None

# — Botón de generación —
if st.button("🚀 Generar Facturas", type="primary"):
    if not all([excel_file, pdf_template, font_file]):
        st.error("❗ Por favor sube los tres archivos antes de continuar.")
    else:
        try:
            df = pd.read_excel(excel_file)
        except Exception as e:
            st.error(f"No se pudo leer el archivo Excel: {e}")
            st.stop()

        missing = validate_columns(df)
        if missing:
            st.error(
                "El Excel no contiene las siguientes columnas requeridas:\n\n"
                + "\n".join(f"  • `{c}`" for c in missing)
            )
            st.stop()

        # Preparar directorio y archivos temporales
        OUTPUT_DIR.mkdir(exist_ok=True)
        template_path = OUTPUT_DIR / "template.pdf"
        font_path     = OUTPUT_DIR / "font.ttf"
        template_path.write_bytes(pdf_template.read())
        font_path.write_bytes(font_file.read())

        total = len(df)
        progress_bar = st.progress(0, text="Iniciando…")
        generated: list[Path] = []
        errors: list[str] = []

        for i, (_, row) in enumerate(df.iterrows(), start=1):
            try:
                path = fill_pdf(row, template_path, font_path, OUTPUT_DIR)
                generated.append(path)
            except Exception as e:
                ref = row.get("Referencia de pago/NUI", f"fila {i}")
                errors.append(f"Fila {i} (ref. {ref}): {e}")

            progress_bar.progress(i / total, text=f"Generando factura {i} de {total}…")

        progress_bar.empty()

        st.session_state.generated_files = generated
        st.session_state.zip_data = create_zip(generated) if generated else None

        # Resumen
        st.markdown(f"""
        <div class="metric-row">
            <div class="metric-box">
                <div class="value">{len(generated)}</div>
                <div class="label">Facturas generadas</div>
            </div>
            <div class="metric-box">
                <div class="value">{len(errors)}</div>
                <div class="label">Errores</div>
            </div>
            <div class="metric-box">
                <div class="value">{total}</div>
                <div class="label">Total en Excel</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if generated:
            st.success(f"✅ {len(generated)} factura(s) generadas correctamente.")
        if errors:
            with st.expander(f"⚠️ {len(errors)} error(es) durante la generación"):
                for err in errors:
                    st.warning(err)

# — Descarga ZIP —
if st.session_state.zip_data:
    st.markdown('<div class="section-card"><div class="section-title">📦 Descargar resultados</div>', unsafe_allow_html=True)

    if st.download_button(
        label="📁 Descargar todas las facturas (.zip)",
        data=st.session_state.zip_data,
        file_name="Facturas_generadas.zip",
        mime="application/zip",
        type="primary",
    ):
        cleanup_output_dir()
        st.session_state.generated_files = []
        st.session_state.zip_data = None
        st.success("🧹 Archivos temporales eliminados tras la descarga.")

    st.markdown("</div>", unsafe_allow_html=True)
