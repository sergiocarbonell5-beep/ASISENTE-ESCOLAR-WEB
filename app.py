# -*- coding: utf-8 -*-
"""
SISTEMA INTEGRAL EDUCATIVO — C.E.R. SIRAVITA (PDF FIDELIDAD 100% SOBRE PLANTILLA IMAGEN)
=======================================================================================
Funcionalidades: PDF montado sobre imagen oficial, Excel, Días Hábiles (Lunes a Viernes),
Días No Lectivos, Alertas de WhatsApp y Gestión Multidocente.
"""

import os
import re
import io
import base64
import datetime
import random
import urllib.parse
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Excel
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# PDF con ReportLab
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors

from supabase import create_client, Client

# ================================================================
# CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ================================================================
st.set_page_config(
    page_title="Asistente Educativo - C.E.R. Siravita",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

APP_TITLE = "Asistente Educativo Sergio Carbonell"
NOMBRE_ESCUELA = "C.E.R. Siravita"
SEDE_DEFECTO = "Chicago Alto"

PUNTOS_BASE = 10
PUNTOS_EXTRA_PUNTUALIDAD = 5
CUPOS_PUNTUALIDAD = 5

MENSAJES_ANIMO = [
    "¡Sigue así, campeón!",
    "¡Cada día sumas más!",
    "¡Eres un ejemplo para tus compañeros!",
    "¡La constancia es la clave del éxito!",
    "¡Un día más, un paso más cerca de tu meta!",
    "¡Tu esfuerzo no pasa desapercibido!",
]

MATERIAS_LISTA = [
    "Inglés", "Matemáticas", "Español", "Sociales", 
    "Naturales", "Artística", "Ética", "Religión", "Informática"
]

MESES_ESPANOL = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

# ================================================================
# CONEXIÓN CON SUPABASE
# ================================================================
@st.cache_resource
def init_supabase() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_supabase()

if "profesor" not in st.session_state:
    st.session_state.profesor = None

# ================================================================
# AUTENTICACIÓN
# ================================================================
def registrar_profesor(nombre, email, password):
    email_clean = email.strip().lower()
    res = supabase.table("profesores").select("id").eq("email", email_clean).execute()
    if res.data:
        return False, "El correo electrónico ya está registrado."
    
    data = {"nombre": nombre.strip(), "email": email_clean, "password": password.strip()}
    supabase.table("profesores").insert(data).execute()
    return True, "¡Profesor registrado exitosamente! Ya puedes iniciar sesión."

def login_profesor(email, password):
    email_clean = email.strip().lower()
    res = supabase.table("profesores").select("*").eq("email", email_clean).eq("password", password.strip()).execute()
    if res.data:
        return res.data[0]
    return None

# ================================================================
# CONSULTAS A LA NUBE
# ================================================================
def obtener_grados(profesor_id):
    res = supabase.table("grados").select("*").eq("profesor_id", profesor_id).order("nombre").execute()
    return res.data

def agregar_grado(nombre, profesor_id):
    supabase.table("grados").insert({"nombre": nombre, "profesor_id": profesor_id}).execute()

def editar_grado(grado_id, nuevo_nombre):
    supabase.table("grados").update({"nombre": nuevo_nombre}).eq("id", grado_id).execute()

def eliminar_grado(grado_id):
    supabase.table("grados").delete().eq("id", grado_id).execute()

def obtener_estudiantes(grado_id, profesor_id):
    res = supabase.table("estudiantes").select("*").eq("grado_id", grado_id).eq("profesor_id", profesor_id).eq("activo", 1).order("nombre").execute()
    return res.data

def agregar_estudiante(nombre, telefono_acudiente, grado_id, profesor_id):
    supabase.table("estudiantes").insert({
        "nombre": nombre, 
        "telefono_acudiente": telefono_acudiente,
        "grado_id": grado_id, 
        "profesor_id": profesor_id,
        "puntos": 0, "racha": 0, "activo": 1
    }).execute()

def editar_estudiante(estudiante_id, nuevo_nombre, nuevo_telefono):
    supabase.table("estudiantes").update({
        "nombre": nuevo_nombre,
        "telefono_acudiente": nuevo_telefono
    }).eq("id", estudiante_id).execute()

def eliminar_estudiante(estudiante_id):
    supabase.table("estudiantes").update({"activo": 0}).eq("id", estudiante_id).execute()

def ya_registrado_hoy(estudiante_id, fecha):
    res = supabase.table("registros").select("id").eq("estudiante_id", estudiante_id).eq("fecha", fecha).execute()
    return len(res.data) > 0

def es_dia_no_lectivo(fecha_str, grado_id, profesor_id):
    try:
        res = supabase.table("dias_no_lectivos").select("*").eq("fecha", fecha_str).eq("grado_id", grado_id).eq("profesor_id", profesor_id).execute()
        return res.data[0] if res.data else None
    except Exception:
        return None

def registrar_dia_no_lectivo(fecha_str, motivo, grado_id, profesor_id):
    try:
        supabase.table("dias_no_lectivos").insert({
            "fecha": fecha_str,
            "motivo": motivo,
            "grado_id": grado_id,
            "profesor_id": profesor_id
        }).execute()
        return True
    except Exception:
        return False

# REGISTRO DE ASISTENCIA
def registrar_asistencia(estudiante, grado_id, profesor_id):
    hoy = datetime.date.today().isoformat()
    hora_str = datetime.datetime.now().strftime("%H:%M:%S")
    
    if ya_registrado_hoy(estudiante["id"], hoy):
        return None

    res_hoy = supabase.table("registros").select("id", count="exact").eq("fecha", hoy).eq("grado_id", grado_id).execute()
    orden = (res_hoy.count or 0) + 1
    
    puntos_extra = PUNTOS_EXTRA_PUNTUALIDAD if orden <= CUPOS_PUNTUALIDAD else 0
    puntos_totales = PUNTOS_BASE + puntos_extra
    
    ayer = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
    nueva_racha = estudiante["racha"] + 1 if estudiante["ultima_fecha"] == ayer else 1
    nuevos_puntos = estudiante["puntos"] + puntos_totales
    
    supabase.table("estudiantes").update({
        "puntos": nuevos_puntos, 
        "racha": nueva_racha, 
        "ultima_fecha": hoy,
        "total_asistencias": estudiante.get("total_asistencias", 0) + 1
    }).eq("id", estudiante["id"]).execute()

    supabase.table("registros").insert({
        "estudiante_id": estudiante["id"], 
        "fecha": hoy, 
        "hora": hora_str,
        "puntos_obtenidos": puntos_totales, 
        "orden_llegada": orden,
        "grado_id": grado_id, 
        "profesor_id": profesor_id
    }).execute()

    return {
        "nombre": estudiante["nombre"], 
        "puntos_ganados": puntos_totales,
        "puntos_extra": puntos_extra, 
        "puntos_totales": nuevos_puntos, 
        "racha": nueva_racha
    }

def crear_link_whatsapp(telefono, nombre_estudiante, nombre_profesor):
    if not telefono:
        return None
    num_limpio = re.sub(r'\D', '', str(telefono))
    if len(num_limpio) == 10:
        num_limpio = "57" + num_limpio
        
    mensaje = f"Estimado acudiente, le saluda el/la profe {nombre_profesor} del {NOMBRE_ESCUELA}. Le informamos que el/la estudiante *{nombre_estudiante}* no ha registrado su asistencia a clases en el día de hoy. Por favor confirmar novedades. Muchas gracias."
    msg_encoded = urllib.parse.quote(mensaje)
    return f"https://wa.me/{num_limpio}?text={msg_encoded}"

# ================================================================
# GENERACIÓN DE PDF EXACTO SOBRE PLANTILLA DE IMAGEN
# ================================================================
def generar_pdf_asistencia_oficial(grado_nombre, profesor_nombre, registros_mes, estudiantes_lista, mes_nombre, sede_nombre=SEDE_DEFECTO):
    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter # 612 x 792 pt

    # 1. Dibujar Imagen Oficial de Fondo
    ruta_imagen = "plantilla_asistencia.png"
    if os.path.exists(ruta_imagen):
        p.drawImage(ruta_imagen, 0, 0, width=width, height=height)
    else:
        p.setFont("Helvetica-Bold", 10)
        p.drawString(40, height - 40, "CENTRO EDUCATIVO RURAL SIRAVITA - CONTROL ASISTENCIA")

    # 2. Estampar Datos Generales (Sede y Mes)
    p.setFont("Helvetica-Bold", 9)
    p.setFillColor(colors.HexColor("#000000"))
    p.drawString(115, 415, str(sede_nombre).upper())
    p.drawString(335, 415, str(mes_nombre).upper())

    # 3. Coordenadas Calibradas para las Filas de la Tabla
    x_alumnos = 12
    x_grado = 86
    x_dias_inicio = 132.5
    w_dia = 13.5
    x_total = 560

    y_fila_inicio = 344  # Alineado dentro de las casillas de la tabla
    h_fila = 18.8        # Distancia entre renglones

    for idx, est in enumerate(estudiantes_lista[:13]):
        y_pos = y_fila_inicio - (idx * h_fila)
        
        # Nombre del Alumno
        p.setFont("Helvetica-Bold", 6.5)
        p.setFillColor(colors.HexColor("#1A237E"))
        p.drawString(x_alumnos, y_pos, str(est["nombre"])[:20])

        # Grado
        p.setFont("Helvetica", 6.5)
        p.setFillColor(colors.HexColor("#333333"))
        p.drawString(x_grado, y_pos, str(grado_nombre)[:8])

        # Asistencias por día (✓)
        tot_asist = 0
        for d in range(1, 32):
            asistio = any(r["estudiante_id"] == est["id"] and int(r["fecha"].split("-")[2]) == d for r in registros_mes)
            if asistio:
                tot_asist += 1
                x_check = x_dias_inicio + ((d - 1) * w_dia)
                p.setFont("Helvetica-Bold", 8)
                p.setFillColor(colors.HexColor("#2E7D32"))
                p.drawString(x_check + 1, y_pos, "✓")

        # Total
        p.setFont("Helvetica-Bold", 8)
        p.setFillColor(colors.HexColor("#000000"))
        p.drawString(x_total, y_pos, str(tot_asist))

    # 4. Firma del Docente
    p.setFont("Helvetica-Bold", 9)
    p.setFillColor(colors.HexColor("#000000"))
    p.drawString(120, 28, str(profesor_nombre).upper())

    p.showPage()
    p.save()
    buffer.seek(0)
    return buffer

# GENERACIÓN EXCEL
def generar_excel_asistencia_oficial(grado_nombre, profesor_nombre, registros_mes, estudiantes_lista, mes_nombre, sede_nombre=SEDE_DEFECTO):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Control Asistencia"
    ws.views.sheetView[0].showGridLines = True

    VERDE_OSCURO = "004D25"
    VERDE_LEMA = "2E7D32"
    VERDE_HEADER_TABLA = "1B432C"
    AMARILLO_CREMA = "FFF9C4"
    
    fill_header_tabla = PatternFill(start_color=VERDE_HEADER_TABLA, end_color=VERDE_HEADER_TABLA, fill_type="solid")
    fill_row_top = PatternFill(start_color=AMARILLO_CREMA, end_color=AMARILLO_CREMA, fill_type="solid")

    font_encabezado_bold = Font(name="Calibri", size=10, bold=True, color="000000")
    font_lema = Font(name="Calibri", size=10, bold=True, italic=True, color=VERDE_LEMA)
    font_subtitulos = Font(name="Calibri", size=11, bold=True, color="000000")
    font_th = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=9)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='B0BEC5'),
        right=Side(style='thin', color='B0BEC5'),
        top=Side(style='thin', color='B0BEC5'),
        bottom=Side(style='thin', color='B0BEC5')
    )

    ws.merge_cells("A1:AH1")
    ws["A1"] = "REPÚBLICA DE COLOMBIA"
    ws["A1"].font = font_encabezado_bold
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:AH2")
    ws["A2"] = "SECRETARÍA DE EDUCACIÓN DEPARTAMENTAL NORTE DE SANTANDER"
    ws["A2"].font = font_encabezado_bold
    ws["A2"].alignment = align_center

    ws.merge_cells("A3:AH3")
    ws["A3"] = "CENTRO EDUCATIVO RURAL SIRAVITA"
    ws["A3"].font = font_encabezado_bold
    ws["A3"].alignment = align_center

    ws.merge_cells("A4:AH4")
    ws["A4"] = "MUNICIPIO DE ARBOLEDAS"
    ws["A4"].font = font_encabezado_bold
    ws["A4"].alignment = align_center

    ws.merge_cells("A5:AH5")
    ws["A5"] = "DANE 254051000139 | DECRETO DE CREACIÓN 00252 DEL 12 DE ABRIL DE 2005"
    ws["A5"].font = Font(name="Calibri", size=8)
    ws["A5"].alignment = align_center

    ws.merge_cells("A6:AH6")
    ws["A6"] = "Lema: Con Escuela nueva y metodología activa para una educación proactiva."
    ws["A6"].font = font_lema
    ws["A6"].alignment = align_center

    ws.merge_cells("A8:AH8")
    ws["A8"] = "REGISTRO CONTROL ASISTENCIA"
    ws["A8"].font = Font(name="Calibri", size=14, bold=True, color=VERDE_OSCURO)
    ws["A8"].alignment = align_center

    ws.merge_cells("A10:AH10")
    ws["A10"] = f"SEDE: {sede_nombre.upper()}     |     MES DE: {mes_nombre.upper()}     |     AÑO: 2026"
    ws["A10"].font = font_subtitulos
    ws["A10"].alignment = align_left

    headers = ["ALUMNOS", "GRADO"] + [str(i) for i in range(1, 32)] + ["TOTAL"]
    
    for col_idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=12, column=col_idx, value=text)
        cell.font = font_th
        cell.fill = fill_header_tabla
        cell.alignment = align_center
        cell.border = thin_border

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=13, column=col_idx)
        cell.fill = fill_row_top
        cell.border = thin_border

    row_start = 14
    for est in estudiantes_lista:
        ws.cell(row=row_start, column=1, value=est["nombre"]).font = font_data
        ws.cell(row=row_start, column=1).alignment = align_left
        ws.cell(row=row_start, column=1).border = thin_border

        ws.cell(row=row_start, column=2, value=grado_nombre).font = font_data
        ws.cell(row=row_start, column=2).alignment = align_center
        ws.cell(row=row_start, column=2).border = thin_border

        asistencias_total = 0
        for dia in range(1, 32):
            col_idx = dia + 2
            asistio = any(r["estudiante_id"] == est["id"] and int(r["fecha"].split("-")[2]) == dia for r in registros_mes)
            val = "✓" if asistio else ""
            if asistio:
                asistencias_total += 1

            cell_d = ws.cell(row=row_start, column=col_idx, value=val)
            cell_d.font = Font(name="Calibri", size=10, bold=True, color="2E7D32")
            cell_d.alignment = align_center
            cell_d.border = thin_border

        cell_t = ws.cell(row=row_start, column=34, value=asistencias_total)
        cell_t.font = Font(name="Calibri", size=9, bold=True)
        cell_t.alignment = align_center
        cell_t.border = thin_border

        row_start += 1

    while row_start < 28:
        for col_idx in range(1, 35):
            c = ws.cell(row=row_start, column=col_idx)
            c.border = thin_border
        row_start += 1

    ws.cell(row=30, column=1, value=f"DOCENTE: {profesor_nombre.upper()}").font = font_subtitulos
    ws.cell(row=32, column=1, value="Documento Oficial - Uso Académico").font = Font(name="Calibri", size=8, italic=True)

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 12
    for col in range(3, 34):
        col_letter = get_column_letter(col)
        ws.column_dimensions[col_letter].width = 3.5
    ws.column_dimensions['AH'].width = 8

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ================================================================
# VENTANA CELEBRACIÓN CON SALUDO DE MAÑANA Y AUDIO INSTANTÁNEO
# ================================================================
@st.dialog("🎉 ¡Asistencia Registrada!", width="large")
def ventana_celebracion(res):
    nombre = res["nombre"]
    puntos_ganados = res["puntos_ganados"]
    puntos_totales = res["puntos_totales"]
    racha = res["racha"]
    
    saludo_hablado = "Buenos días"
    frase_animo = random.choice(MENSAJES_ANIMO)
    texto_voz = f"Buenos días, bienvenido {nombre}. Has ganado {puntos_ganados} puntos."

    st.markdown(f"""
        <div style="background-color: #22C55E; color: white; padding: 25px; border-radius: 20px; text-align: center;">
            <h2 style="color: #FFE066; margin:0;">🌟 ¡{saludo_hablado}, {nombre}! 🌟</h2>
            <h1 style="color: #FFE066; font-size: 40px; margin: 10px 0;">{nombre}</h1>
            <p style="font-size: 20px; margin:5px;">✨ +{puntos_ganados} puntos ganados</p>
            <p style="font-size: 22px; font-weight: bold; margin:5px;">🔥 Racha: {racha} día(s) | Puntos: {puntos_totales}</p>
            <h3 style="color: #FFE066; margin-top:10px;">{frase_animo}</h3>
        </div>
    """, unsafe_allow_html=True)

    components.html(f"""
        <script>
            (function() {{
                if ('speechSynthesis' in window) {{
                    window.speechSynthesis.cancel();
                    var msg = new SpeechSynthesisUtterance("{texto_voz}");
                    msg.lang = 'es-ES';
                    msg.rate = 1.1;
                    window.speechSynthesis.speak(msg);
                }}
            }})();
        </script>
    """, height=0)

    st.write("")
    if st.button("☑️ ¡Entendido!", use_container_width=True, type="primary"):
        st.rerun()

# ================================================================
# VISTA: LOGIN / REGISTRO
# ================================================================
if st.session_state.profesor is None:
    st.title("🏫 " + NOMBRE_ESCUELA)
    st.subheader("Plataforma Multidocente de Gestión Educativa")
    
    tab_login, tab_registro = st.tabs(["🔑 Iniciar Sesión", "📝 Registrar nuevo Profesor"])
    
    with tab_login:
        with st.form("form_login"):
            email = st.text_input("Correo electrónico:")
            password = st.text_input("Contraseña:", type="password")
            btn_login = st.form_submit_button("Ingresar al Panel")
            if btn_login:
                prof = login_profesor(email, password)
                if prof:
                    st.session_state.profesor = prof
                    st.success(f"¡Bienvenido(a), Profe {prof['nombre']}!")
                    st.rerun()
                else:
                    st.error("Correo o contraseña incorrectos.")

    with tab_registro:
        with st.form("form_reg"):
            nom_reg = st.text_input("Nombre Completo:")
            email_reg = st.text_input("Correo electrónico:")
            pass_reg = st.text_input("Crear Contraseña:", type="password")
            btn_reg = st.form_submit_button("Crear Cuenta de Docente")
            if btn_reg:
                if nom_reg and email_reg and pass_reg:
                    ok, msg = registrar_profesor(nom_reg, email_reg, pass_reg)
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)
                else:
                    st.warning("Completa todos los campos.")

# ================================================================
# VISTA: PANEL PRINCIPAL
# ================================================================
else:
    prof = st.session_state.profesor
    
    with st.sidebar:
        st.title(f"👨‍🏫 Profe: {prof['nombre']}")
        st.caption(f"📧 {prof['email']}")
        
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            st.session_state.profesor = None
            st.rerun()
            
        st.markdown("---")
        
        grados = obtener_grados(prof["id"])
        if not grados:
            st.warning("Aún no tienes cursos creados.")
            grado_sel_id = None
            grado_sel_nombre = ""
        else:
            nombres_grados = [g["nombre"] for g in grados]
            grado_sel_nombre = st.selectbox("Selecciona un Grado:", nombres_grados)
            grado_sel_id = next(g["id"] for g in grados if g["nombre"] == grado_sel_nombre)

        st.markdown("---")
        
        # GESTIÓN DE CURSOS Y ESTUDIANTES
        with st.expander("⚙️ Gestionar Cursos / Grados"):
            nuevo_grado_txt = st.text_input("Nombre de nuevo curso:", placeholder="Ej. Grado Cuarto")
            if st.button("➕ Crear Curso", use_container_width=True):
                if nuevo_grado_txt.strip():
                    agregar_grado(nuevo_grado_txt.strip(), prof["id"])
                    st.success("Curso creado.")
                    st.rerun()
            
            if grado_sel_id:
                st.markdown("---")
                edit_g_nom = st.text_input("Editar nombre del curso actual:", value=grado_sel_nombre)
                if st.button("✏️ Guardar Nombre Curso", use_container_width=True):
                    editar_grado(grado_sel_id, edit_g_nom.strip())
                    st.success("Nombre actualizado.")
                    st.rerun()
                    
                if st.button("🗑️ Eliminar Curso Actual", type="primary", use_container_width=True):
                    eliminar_grado(grado_sel_id)
                    st.success("Curso eliminado.")
                    st.rerun()

        st.markdown("---")
        with st.expander("👤 Agregar / Editar Estudiante"):
            if grado_sel_id:
                nom_est = st.text_input("Nombre completo de alumno:")
                tel_est = st.text_input("Teléfono WhatsApp Acudiente:", placeholder="Ej: 3001234567")
                if st.button("➕ Guardar Alumno", use_container_width=True):
                    if nom_est.strip():
                        agregar_estudiante(nom_est.strip(), tel_est.strip(), grado_sel_id, prof["id"])
                        st.success("Alumno guardado.")
                        st.rerun()
                
                st.markdown("---")
                estudiantes_lista = obtener_estudiantes(grado_sel_id, prof["id"])
                if estudiantes_lista:
                    dict_est = {e["nombre"]: e for e in estudiantes_lista}
                    est_edit_nom = st.selectbox("Selecciona alumno a editar:", list(dict_est.keys()))
                    est_obj = dict_est[est_edit_nom]
                    
                    nuevo_nom_val = st.text_input("Nuevo nombre:", value=est_obj["nombre"])
                    nuevo_tel_val = st.text_input("Nuevo teléfono:", value=est_obj.get("telefono_acudiente") or "")
                    
                    col_e1, col_e2 = st.columns(2)
                    with col_e1:
                        if st.button("✏️ Actualizar", use_container_width=True):
                            editar_estudiante(est_obj["id"], nuevo_nom_val.strip(), nuevo_tel_val.strip())
                            st.success("Datos actualizados.")
                            st.rerun()
                    with col_e2:
                        if st.button("🗑️ Eliminar", use_container_width=True):
                            eliminar_estudiante(est_obj["id"])
                            st.success("Alumno eliminado.")
                            st.rerun()

    # PESTAÑAS PRINCIPALES
    st.title(f"📋 Asistente Educativo — {grado_sel_nombre if grado_sel_nombre else 'Crea un curso'}")

    if grado_sel_id:
        t_docs, t_asistencia, t_notas, t_convivencia, t_lideres = st.tabs([
            "📄 Documentos Institucionales",
            "📋 Registro de Asistencia", 
            "📝 Calificaciones",
            "⚖️ Convivencia",
            "🏆 Tabla de Líderes"
        ])

        # 1. DOCUMENTOS INSTITUCIONALES
        with t_docs:
            st.subheader("📁 Repositorio de Documentos por Materia")
            tipo_doc_sel = st.radio("Categoría:", ["Planes de Área", "Planes de Clase", "Guías Educativas"], horizontal=True)
            materia_doc = st.selectbox("Materia:", MATERIAS_LISTA)
            
            archivo_subido = st.file_uploader(f"Subir documento a {tipo_doc_sel} ({materia_doc}):", type=["pdf", "docx", "pptx", "xlsx", "txt"])
            if archivo_subido is not None:
                if st.button("💾 Guardar Documento en la Nube"):
                    bytes_data = archivo_subido.getvalue()
                    b64_str = base64.b64encode(bytes_data).decode('utf-8')
                    
                    supabase.table("documentos").insert({
                        "nombre": archivo_subido.name,
                        "materia": materia_doc,
                        "tipo_doc": tipo_doc_sel,
                        "contenido_b64": b64_str,
                        "grado_id": grado_sel_id,
                        "profesor_id": prof["id"]
                    }).execute()
                    st.success("¡Documento guardado permanentemente en la nube!")
                    st.rerun()

            st.markdown("---")
            st.write(f"**Documentos de {materia_doc} en {tipo_doc_sel}:**")
            res_docs = supabase.table("documentos").select("*").eq("materia", materia_doc).eq("tipo_doc", tipo_doc_sel).eq("grado_id", grado_sel_id).eq("profesor_id", prof["id"]).execute()
            
            if not res_docs.data:
                st.info(f"Aún no se han subido documentos para {materia_doc}.")
            else:
                for doc in res_docs.data:
                    c1, c2 = st.columns([3, 1])
                    with c1:
                        st.write(f"📄 **{doc['nombre']}** ({doc['created_at'][:10]})")
                    with c2:
                        bytes_dec = base64.b64decode(doc['contenido_b64'])
                        st.download_button("⬇️ Descargar", data=bytes_dec, file_name=doc['nombre'], use_container_width=True)

        # 2. REGISTRO DE ASISTENCIA CON DÍAS HÁBILES Y FORMATO PDF/EXCEL
        with t_asistencia:
            estudiantes = obtener_estudiantes(grado_sel_id, prof["id"])
            hoy_dt = datetime.date.today()
            hoy_str = hoy_dt.isoformat()
            dia_semana = hoy_dt.weekday()
            es_fin_semana = dia_semana >= 5
            
            registro_no_lectivo = es_dia_no_lectivo(hoy_str, grado_sel_id, prof["id"])
            
            total_est = len(estudiantes) if estudiantes else 0
            asistieron_count = 0
            estudiantes_ausentes = []
            
            if estudiantes:
                for est in estudiantes:
                    if ya_registrado_hoy(est["id"], hoy_str):
                        asistieron_count += 1
                    else:
                        estudiantes_ausentes.append(est)
            
            faltaron_count = total_est - asistieron_count

            col_titulo, col_resumen = st.columns([1.6, 1])
            
            with col_titulo:
                st.subheader("Marcar Asistencia Diaria")
                st.caption(f"📅 Fecha actual: {hoy_str} ({'Fin de Semana - Sin Clases' if es_fin_semana else 'Día Hábil'})")
            
            with col_resumen:
                if total_est > 0:
                    c_red, c_green = st.columns(2)
                    with c_red:
                        st.markdown(f"""
                            <div style="background-color: #FFEBEE; border: 2px solid #EF5350; border-radius: 12px; padding: 8px 12px; text-align: center;">
                                <span style="color: #C62828; font-weight: bold; font-size: 13px;">✖ Faltaron</span><br>
                                <span style="color: #C62828; font-size: 26px; font-weight: bold;">{faltaron_count}</span><br>
                                <span style="color: #B71C1C; font-size: 11px;">{'estudiante' if faltaron_count == 1 else 'estudiantes'}</span>
                            </div>
                        """, unsafe_allow_html=True)
                    with c_green:
                        st.markdown(f"""
                            <div style="background-color: #E8F5E9; border: 2px solid #66BB6A; border-radius: 12px; padding: 8px 12px; text-align: center;">
                                <span style="color: #2E7D32; font-weight: bold; font-size: 13px;">✅ Asistieron</span><br>
                                <span style="color: #2E7D32; font-size: 26px; font-weight: bold;">{asistieron_count}</span><br>
                                <span style="color: #1B5E20; font-size: 11px;">{'estudiante' if asistieron_count == 1 else 'estudiantes'}</span>
                            </div>
                        """, unsafe_allow_html=True)

            st.write("")

            # BOTONES DE EXPORTACIÓN
            if estudiantes:
                mes_actual_nombre = MESES_ESPANOL[hoy_dt.month - 1]
                res_reg = supabase.table("registros").select("*").eq("grado_id", grado_sel_id).execute()
                
                col_ex1, col_ex2 = st.columns(2)
                with col_ex1:
                    pdf_bytes = generar_pdf_asistencia_oficial(
                        grado_nombre=grado_sel_nombre,
                        profesor_nombre=prof['nombre'],
                        registros_mes=res_reg.data or [],
                        estudiantes_lista=estudiantes,
                        mes_nombre=mes_actual_nombre,
                        sede_nombre=SEDE_DEFECTO
                    )
                    st.download_button(
                        label="📄 Exportar Planilla Oficial en PDF (Formato Imagen)",
                        data=pdf_bytes,
                        file_name=f"Planilla_Oficial_Asistencia_{grado_sel_nombre}_{mes_actual_nombre}_2026.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        type="primary"
                    )
                with col_ex2:
                    excel_bytes = generar_excel_asistencia_oficial(
                        grado_nombre=grado_sel_nombre,
                        profesor_nombre=prof['nombre'],
                        registros_mes=res_reg.data or [],
                        estudiantes_lista=estudiantes,
                        mes_nombre=mes_actual_nombre,
                        sede_nombre=SEDE_DEFECTO
                    )
                    st.download_button(
                        label="📊 Exportar Planilla Oficial en Excel (.xlsx)",
                        data=excel_bytes,
                        file_name=f"Control_Asistencia_{grado_sel_nombre}_{mes_actual_nombre}_2026.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            st.markdown("---")

            # REGISTRO DE DÍAS NO LECTIVOS / SIN CLASE
            with st.expander("🚫 Registrar Día Sin Clase (Festivo, Paro, Jornada Pedagógica)", expanded=False):
                if registro_no_lectivo:
                    st.warning(f"⚠️ El día de hoy ({hoy_str}) fue registrado como DÍA SIN CLASE: **{registro_no_lectivo['motivo']}**")
                else:
                    motivo_txt = st.text_input("Motivo de la suspensión de clases:", placeholder="Ej. Jornada Pedagógica de Docentes / Clima")
                    if st.button("📌 Declarar Día Sin Clase"):
                        if motivo_txt.strip():
                            if registrar_dia_no_lectivo(hoy_str, motivo_txt.strip(), grado_sel_id, prof["id"]):
                                st.success("Día registrado como no lectivo. No afectará las asistencias.")
                                st.rerun()

            st.markdown("---")

            # NOTIFICACIONES A AUSENTES POR WHATSAPP
            if estudiantes_ausentes and not es_fin_semana and not registro_no_lectivo:
                with st.expander("📲 Notificar ausencias a Acudientes por WhatsApp", expanded=False):
                    st.caption("Haz clic en el botón al lado de cada estudiante faltante para abrir su chat con la notificación lista:")
                    for aus in estudiantes_ausentes:
                        tel = aus.get("telefono_acudiente")
                        c_a1, c_a2 = st.columns([2, 1])
                        with c_a1:
                            st.write(f"🔴 **{aus['nombre']}** - Tel: `{tel if tel else 'Sin registrar'}`")
                        with c_a2:
                            link_wa = crear_link_whatsapp(tel, aus['nombre'], prof['nombre'])
                            if link_wa:
                                st.link_button("📲 Enviar WhatsApp", link_wa, use_container_width=True)
                            else:
                                st.button("⚠️ Falta Teléfono", disabled=True, key=f"no_tel_{aus['id']}", use_container_width=True)

            # CONTROL DE ASISTENCIA DIARIA (LUNES A VIERNES)
            if es_fin_semana:
                st.info("🌴 Hoy es fin de semana (Sábado/Domingo). El registro de asistencia no está activo.")
            elif registro_no_lectivo:
                st.info(f"🚫 Día declarado sin actividades escolares por: **{registro_no_lectivo['motivo']}**.")
            elif not estudiantes:
                st.info("No hay alumnos registrados en este curso. Agrégalos desde la barra lateral.")
            else:
                cols = st.columns(3)
                for i, est in enumerate(estudiantes):
                    registrado = ya_registrado_hoy(est["id"], hoy_str)
                    col = cols[i % 3]
                    with col:
                        with st.container(border=True):
                            if registrado:
                                st.success(f"✅ {est['nombre']}")
                                st.caption(f"Puntos: {est['puntos']} | Racha: 🔥 {est['racha']} días")
                                st.button("Registrado", key=f"btn_{est['id']}", disabled=True)
                            else:
                                st.markdown(f"**👤 {est['nombre']}**")
                                st.caption(f"Puntos: {est['puntos']} | Racha: 🔥 {est['racha']} días")
                                if st.button("✋ Marcar Asistencia", key=f"btn_{est['id']}"):
                                    res = registrar_asistencia(est, grado_sel_id, prof["id"])
                                    if res:
                                        st.balloons()
                                        ventana_celebracion(res)

        # 3. CALIFICACIONES
        with t_notas:
            st.subheader("📝 Registro de Calificaciones")
            estudiantes = obtener_estudiantes(grado_sel_id, prof["id"])
            
            if estudiantes:
                col_n1, col_n2, col_n3, col_n4 = st.columns(4)
                with col_n1:
                    dict_e = {e["nombre"]: e["id"] for e in estudiantes}
                    est_nota_nom = st.selectbox("Estudiante:", list(dict_e.keys()))
                with col_n2:
                    mat_nota = st.selectbox("Materia de nota:", MATERIAS_LISTA)
                with col_n3:
                    val_nota = st.number_input("Nota (1.0 a 5.0):", min_value=1.0, max_value=5.0, value=5.0, step=0.1)
                with col_n4:
                    periodo_nota = st.selectbox("Periodo:", ["Periodo 1", "Periodo 2", "Periodo 3", "Periodo 4"])
                
                if st.button("💾 Guardar Calificación", use_container_width=True):
                    supabase.table("calificaciones").insert({
                        "estudiante_id": dict_e[est_nota_nom],
                        "materia": mat_nota,
                        "nota": val_nota,
                        "periodo": periodo_nota,
                        "profesor_id": prof["id"]
                    }).execute()
                    st.success("Nota guardada exitosamente.")
                    st.rerun()

                st.markdown("---")
                st.write("**Historial de Notas Registradas:**")
                res_notas = supabase.table("calificaciones").select("estudiante_id, materia, nota, periodo").eq("profesor_id", prof["id"]).execute()
                if res_notas.data:
                    id_to_name = {e["id"]: e["nombre"] for e in estudiantes}
                    df_notas = pd.DataFrame([
                        {"Estudiante": id_to_name.get(n["estudiante_id"], "N/A"), "Materia": n["materia"], "Nota": n["nota"], "Periodo": n["periodo"]}
                        for n in res_notas.data if n["estudiante_id"] in id_to_name
                    ])
                    st.dataframe(df_notas, use_container_width=True)

        # 4. CONVIVENCIA
        with t_convivencia:
            st.subheader("⚖️ Observador del Estudiante / Convivencia")
            estudiantes = obtener_estudiantes(grado_sel_id, prof["id"])
            
            if estudiantes:
                dict_e = {e["nombre"]: e["id"] for e in estudiantes}
                c_c1, c_c2 = st.columns(2)
                with c_c1:
                    est_conv_nom = st.selectbox("Estudiante a registrar:", list(dict_e.keys()))
                with c_c2:
                    tipo_conv = st.selectbox("Tipo de anotación:", ["Positivo / Reconocimiento", "Llamado de atención", "Falta grave"])
                
                desc_conv = st.text_area("Descripción de la situación:")
                if st.button("📝 Guardar Anotación", use_container_width=True):
                    if desc_conv.strip():
                        supabase.table("convivencia").insert({
                            "estudiante_id": dict_e[est_conv_nom],
                            "fecha": datetime.date.today().isoformat(),
                            "tipo": tipo_conv,
                            "descripcion": desc_conv.strip(),
                            "profesor_id": prof["id"]
                        }).execute()
                        st.success("Anotación guardada en el observador.")
                        st.rerun()

                st.markdown("---")
                res_conv = supabase.table("convivencia").select("*").eq("profesor_id", prof["id"]).execute()
                if res_conv.data:
                    id_to_name = {e["id"]: e["nombre"] for e in estudiantes}
                    df_c = pd.DataFrame([
                        {"Fecha": c["fecha"], "Estudiante": id_to_name.get(c["estudiante_id"], "N/A"), "Tipo": c["tipo"], "Anotación": c["descripcion"]}
                        for c in res_conv.data if c["estudiante_id"] in id_to_name
                    ])
                    st.dataframe(df_c, use_container_width=True)

        # 5. TABLA DE LÍDERES
        with t_lideres:
            st.subheader("🏆 Clasificación General por Puntos y Rachas")
            estudiantes = obtener_estudiantes(grado_sel_id, prof["id"])
            if estudiantes:
                df = pd.DataFrame([
                    {"Estudiante": e["nombre"], "Puntos Acumulados": e["puntos"], "Racha Actual": f"🔥 {e['racha']} días", "Asistencias Totales": e.get("total_asistencias", 0)}
                    for e in estudiantes
                ]).sort_values(by="Puntos Acumulados", ascending=False)
                
                st.dataframe(df, use_container_width=True)
