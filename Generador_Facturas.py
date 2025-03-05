import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
from datetime import datetime
import os

def formatear_fecha(fecha):
    if pd.notna(fecha):
        return fecha.strftime("%d/%m/%Y") if isinstance(fecha, datetime) else str(fecha)
    return ""

def formatear_moneda(valor):
    try:
        valor = int(valor)
        return f"${valor:,.0f}".replace(",", ".")
    except:
        return "$0"

def llenar_pdf(datos, pdf_template, font_path):
    doc = fitz.open(pdf_template)
    page = doc[0]
    fontname = "custom_font"
    page.insert_font(fontname=fontname, fontfile=font_path)
    
    campos = {
        "Factura electrónica de venta N°": (220, 118),
        "Nombre": (150, 133),
        "Cédula": (150, 155),
        "Municipio": (150, 185),
        "Localidad": (150, 205),
        "Tipo de lectura": (150, 256),
        "Referencia de pago/NUI": (490, 135),
        "Días facturados": (485, 339),
        "Deuda anterior": (485, 359),
        "Otros conceptos": (485, 382),
        "1+2 Valor total a cancelar": (485, 404),
        "Fecha de emisión": (485, 425),
        "Pago oportuno": (485, 446),
        "Suspension a partir de": (485, 468),
        "Cargos facturados mes": (485, 512),
        "Costo mensual prestación de servicios": (485, 533),
        "Valor refacturacion": (485,555),
        "Valor por mora": (485, 580),
        "Interés por mora": (485, 600),
        "Valor subsidio": (485, 620),
        "Total servicio": (485, 645)
    }
    
    valores = {
        "Factura electrónica de venta N°": str(datos["Factura"]),
        "Nombre": datos["Nombre"],
        "Cédula": str(datos["Cedula"]),
        "Municipio": datos["Municipio"],
        "Localidad": datos["Localidad"],
        "Tipo de lectura": datos["Tipo de Lectura"],
        "Referencia de pago/NUI": str(datos["Referencia de pago/NUI"]),
        "Días facturados": str(datos["Dias facturados"]),
        "Deuda anterior": formatear_moneda(datos["Deuda anterior"]),
        "1+2 Valor total a cancelar": formatear_moneda(datos["1+2 Valor total a cancelar"]),
        "Fecha de emisión": formatear_fecha(datos["Fecha de emision"]),
        "Pago oportuno": formatear_fecha(datos["Pago oportuno"]),
        "Suspension a partir de": formatear_fecha(datos["Suspension a partir de"]),
        "Cargos facturados mes": formatear_moneda(datos["Cargos facturados Mes"]),
        "Costo mensual prestación de servicios": formatear_moneda(datos["Costo mensual preatacion de servicios"]),
        "Valor por mora": formatear_moneda(datos["Valor por mora"]),
        "Valor subsidio": formatear_moneda(datos["Valor subsidio"]),
        "Total servicio": formatear_moneda(datos["Total Servicio"]),
    }
    
    for campo, (x, y) in campos.items():
        page.insert_text((x, y), valores[campo], fontsize=8.5, fontname=fontname, color=(0, 0, 0))
    
    output_path = f"Factura_{datos['Referencia de pago/NUI']}.pdf"
    doc.save(output_path)
    doc.close()
    return output_path

st.title("Generador de Facturas en PDF")

excel_file = st.file_uploader("Subir archivo Excel", type=["xlsx"])
pdf_template = st.file_uploader("Subir plantilla PDF", type=["pdf"])
font_file = st.file_uploader("Subir fuente personalizada", type=["ttf"])

generated_pdfs = []

if st.button("Generar Facturas"):
    if excel_file and pdf_template and font_file:
        df = pd.read_excel(excel_file)
        pdf_template_path = "template.pdf"
        font_path = "font.ttf"
        with open(pdf_template_path, "wb") as f:
            f.write(pdf_template.read())
        with open(font_path, "wb") as f:
            f.write(font_file.read())
        for _, row in df.iterrows():
            pdf_path = llenar_pdf(row, pdf_template_path, font_path)
            generated_pdfs.append(pdf_path)
        st.success("Facturas generadas correctamente.")
    else:
        st.error("Por favor, suba todos los archivos antes de generar las facturas.")

if generated_pdfs:
    for pdf in generated_pdfs:
        with open(pdf, "rb") as f:
            st.download_button(label=f"Descargar {pdf}", data=f, file_name=pdf)