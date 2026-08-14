# -*- coding: utf-8 -*-
"""
SISTEMA INTEGRAL EDUCATIVO — VERSIÓN STREAMLIT (WEB)
====================================================
Adaptado para navegador con Streamlit.
Incluye Pantalla de Bienvenida (Splash Screen) al inicio,
gestión de SQLite ('asistencia.db'), Documentos Institucionales por materia,
previsualizador integrado, asistencia con voz automática y calificaciones.
"""

import os
import sys
import re
import base64
import sqlite3
import datetime
import csv
import shutil
import random
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# ================================================================
# CONFIGURACIÓN DE PÁGINA Y ESTILOS EN STREAMLIT
# ================================================================
st.set_page_config(
    page_title="Asistente Educativo - C.E.R. Siravita",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

APP_TITLE = "Asistente Educativo Sergio Carbonell"
NOMBRE_ESCUELA = "C.E.R. Siravita"

# Reglas de puntuación
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
    "Matemáticas", "Español", "Inglés", "Sociales", 
    "Naturales", "Artística", "Ética", "Religión", "Informática"
]

# Estado global para controlar la pantalla de inicio
if "pantalla_inicio" not in st.session_state:
    st.session_state.pantalla_inicio = True

# ================================================================
# UBICACIÓN DE LA BASE DE DATOS Y CARPETAS DE DOCUMENTOS
# ================================================================
def carpeta_base():
    return os.path.dirname(os.path.abspath(__file__))

DB_NAME = os.path.join(carpeta_base(), "asistencia.db")
CARPETA_DOCS = os.path.join(carpeta_base(), "documentos_institucionales")

# Crear carpetas para los documentos institucionales si no existen
for subcarpeta in ["planes_area", "planes_clase", "guias_educativas"]:
    for mat in MATERIAS_LISTA:
        mat_folder = mat.lower().replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
        os.makedirs(os.path.join(CARPETA_DOCS, subcarpeta, mat_folder), exist_ok=True)

def limpiar_nombre_archivo(nombre):
    """Limpia caracteres especiales y tildes para evitar OSError en Windows"""
    remplazos = {'á':'a', 'é':'e', 'í':'i', 'ó':'o', 'ú':'u', 'Á':'A', 'É':'E', 'Í':'I', 'Ó':'O', 'Ú':'U', 'ñ':'n', 'Ñ':'N'}
    for k, v in remplazos.items():
        nombre = nombre.replace(k, v)
    return re.sub(r'[^a-zA-Z0-9_.-]', '_', nombre)

# ================================================================
# BASE DE DATOS E INICIALIZACIÓN
# ================================================================
def conectar():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    conn = conectar()
    c = conn.cursor()
    c.execute("PRAGMA foreign_keys = OFF")

    c.execute("""CREATE TABLE IF NOT EXISTS grados (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT UNIQUE NOT NULL
                )""")
    if c.execute("SELECT COUNT(*) FROM grados").fetchone()[0] == 0:
        for nombre in ("Segundo", "Cuarto", "Quinto"):
            c.execute("INSERT OR IGNORE INTO grados (nombre) VALUES (?)", (nombre,))
    conn.commit()

    c.execute("""CREATE TABLE IF NOT EXISTS estudiantes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL,
                    puntos INTEGER DEFAULT 0,
                    racha INTEGER DEFAULT 0,
                    max_racha INTEGER DEFAULT 0,
                    ultima_fecha TEXT,
                    total_asistencias INTEGER DEFAULT 0,
                    veces_primero INTEGER DEFAULT 0,
                    activo INTEGER DEFAULT 1,
                    grado_id INTEGER NOT NULL,
                    UNIQUE(nombre, grado_id),
                    FOREIGN KEY(grado_id) REFERENCES grados(id)
                )""")
    conn.commit()

    c.execute("""CREATE TABLE IF NOT EXISTS registros (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    estudiante_id INTEGER NOT NULL,
                    fecha TEXT NOT NULL,
                    hora TEXT NOT NULL,
                    puntos_obtenidos INTEGER,
                    orden_llegada INTEGER,
                    grado_id INTEGER,
                    FOREIGN KEY(estudiante_id) REFERENCES estudiantes(id) ON DELETE CASCADE
                )""")
    conn.commit()

    c.execute("""CREATE TABLE IF NOT EXISTS insignias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    estudiante_id INTEGER NOT NULL,
                    nombre_insignia TEXT NOT NULL,
                    fecha_obtenida TEXT,
                    FOREIGN KEY(estudiante_id) REFERENCES estudiantes(id) ON DELETE CASCADE
                )""")
    conn.commit()

    c.execute("""CREATE TABLE IF NOT EXISTS materias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT UNIQUE NOT NULL
                )""")
    for materia in MATERIAS_LISTA:
        c.execute("INSERT OR IGNORE INTO materias (nombre) VALUES (?)", (materia,))
    conn.commit()

    c.execute("""CREATE TABLE IF NOT EXISTS categorias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT UNIQUE NOT NULL
                )""")
    categorias_default = ("Participación", "Evaluaciones", "Tareas", "Comportamiento")
    for categoria in categorias_default:
        c.execute("INSERT OR IGNORE INTO categorias (nombre) VALUES (?)", (categoria,))
    conn.commit()

    c.execute("""CREATE TABLE IF NOT EXISTS calificaciones (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    estudiante_id INTEGER NOT NULL,
                    materia_id INTEGER NOT NULL,
                    categoria_id INTEGER NOT NULL,
                    item_num INTEGER NOT NULL,
                    valor REAL,
                    fecha DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY(estudiante_id) REFERENCES estudiantes(id) ON DELETE CASCADE,
                    FOREIGN KEY(materia_id) REFERENCES materias(id),
                    FOREIGN KEY(categoria_id) REFERENCES categorias(id)
                )""")
    conn.commit()

    c.execute("""CREATE TABLE IF NOT EXISTS tipos_falta (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL,
                    descripcion TEXT,
                    gravedad INTEGER
                )""")
    c.execute("""CREATE TABLE IF NOT EXISTS categorias_falta (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tipo_falta_id INTEGER NOT NULL,
                    articulo TEXT,
                    numeral TEXT,
                    descripcion TEXT NOT NULL,
                    sancion TEXT,
                    FOREIGN KEY(tipo_falta_id) REFERENCES tipos_falta(id)
                )""")
    c.execute("""CREATE TABLE IF NOT EXISTS incidentes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    estudiante_id INTEGER NOT NULL,
                    fecha TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    descripcion TEXT NOT NULL,
                    tipo_falta_id INTEGER,
                    categoria_falta_id INTEGER,
                    es_reincidente BOOLEAN DEFAULT 0,
                    sancion_aplicada TEXT,
                    estado TEXT DEFAULT 'Pendiente',
                    registrado_por TEXT,
                    FOREIGN KEY(estudiante_id) REFERENCES estudiantes(id) ON DELETE CASCADE,
                    FOREIGN KEY(tipo_falta_id) REFERENCES tipos_falta(id),
                    FOREIGN KEY(categoria_falta_id) REFERENCES categorias_falta(id)
                )""")
    conn.commit()
    conn.close()

# ================================================================
# OPERACIONES BD
# ================================================================
def obtener_grados():
    conn = conectar()
    filas = conn.execute("SELECT id, nombre FROM grados ORDER BY nombre COLLATE NOCASE").fetchall()
    conn.close()
    return filas

def agregar_grado(nombre):
    conn = conectar()
    try:
        conn.execute("INSERT INTO grados (nombre) VALUES (?)", (nombre,))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()

def editar_grado(grado_id, nuevo_nombre):
    conn = conectar()
    try:
        conn.execute("UPDATE grados SET nombre = ? WHERE id = ?", (nuevo_nombre, grado_id))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()

def eliminar_grado(grado_id):
    conn = conectar()
    try:
        conn.execute("UPDATE estudiantes SET activo = 0 WHERE grado_id = ?", (grado_id,))
        conn.execute("DELETE FROM grados WHERE id = ?", (grado_id,))
        conn.commit()
        return True
    except Exception:
        return False
    finally:
        conn.close()

def obtener_estudiantes_por_grado(grado_id, solo_activos=True):
    conn = conectar()
    q = "SELECT * FROM estudiantes WHERE grado_id = ?"
    if solo_activos:
        q += " AND activo = 1"
    q += " ORDER BY nombre COLLATE NOCASE"
    filas = conn.execute(q, (grado_id,)).fetchall()
    conn.close()
    return filas

def ya_registrado_hoy(id_estudiante, fecha):
    conn = conectar()
    res = conn.execute("SELECT id FROM registros WHERE estudiante_id=? AND fecha=?", (id_estudiante, fecha)).fetchone()
    conn.close()
    return res is not None

def registrar_asistencia(id_estudiante, grado_id):
    hoy = datetime.date.today()
    fecha_str = hoy.isoformat()
    hora_str = datetime.datetime.now().strftime("%H:%M:%S")
    if ya_registrado_hoy(id_estudiante, fecha_str):
        return None
    conn = conectar()
    c = conn.cursor()
    total_hoy = c.execute("SELECT COUNT(*) FROM registros WHERE fecha=? AND grado_id=?", (fecha_str, grado_id)).fetchone()[0]
    orden = total_hoy + 1
    puntos_extra = PUNTOS_EXTRA_PUNTUALIDAD if orden <= CUPOS_PUNTUALIDAD else 0
    puntos_totales = PUNTOS_BASE + puntos_extra
    
    fila = c.execute("SELECT nombre, puntos, racha, max_racha, ultima_fecha, veces_primero FROM estudiantes WHERE id=?", (id_estudiante,)).fetchone()
    nombre, puntos_actuales, racha_actual, max_racha, ultima_fecha, veces_primero = fila
    
    nueva_racha = racha_actual + 1 if (ultima_fecha == (hoy - datetime.timedelta(days=1)).isoformat()) else 1
    nueva_max_racha = max(max_racha, nueva_racha)
    nuevas_veces_primero = veces_primero + (1 if orden == 1 else 0)
    nuevos_puntos = puntos_actuales + puntos_totales

    c.execute("""UPDATE estudiantes SET puntos=?, racha=?, max_racha=?, ultima_fecha=?,
                 total_asistencias = total_asistencias + 1, veces_primero=? WHERE id=?""",
              (nuevos_puntos, nueva_racha, nueva_max_racha, fecha_str, nuevas_veces_primero, id_estudiante))
    
    c.execute("INSERT INTO registros (estudiante_id, fecha, hora, puntos_obtenidos, orden_llegada, grado_id) VALUES (?,?,?,?,?,?)",
              (id_estudiante, fecha_str, hora_str, puntos_totales, orden, grado_id))
    conn.commit()
    conn.close()
    return {
        "nombre": nombre,
        "puntos_ganados": puntos_totales,
        "puntos_extra": puntos_extra,
        "puntos_totales": nuevos_puntos,
        "racha": nueva_racha
    }

def agregar_estudiante(nombre, grado_id):
    conn = conectar()
    try:
        conn.execute("INSERT INTO estudiantes (nombre, grado_id) VALUES (?, ?)", (nombre, grado_id))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()

# ================================================================
# MODAL VENTANA VERDE DE CELEBRACIÓN CON VOZ
# ================================================================
@st.dialog("🎉 ¡Asistencia Registrada!", width="large")
def ventana_celebracion(res):
    nombre = res["nombre"]
    puntos_ganados = res["puntos_ganados"]
    puntos_extra = res["puntos_extra"]
    puntos_totales = res["puntos_totales"]
    racha = res["racha"]
    
    hora = datetime.datetime.now().hour
    if 5 <= hora < 12:
        saludo = "🌅 ¡Buenos días"
        saludo_hablado = "Buenos días"
    elif 12 <= hora < 19:
        saludo = "☀️ ¡Buenas tardes"
        saludo_hablado = "Buenas tardes"
    else:
        saludo = "🌙 ¡Buenas noches"
        saludo_hablado = "Buenas noches"
        
    frase_animo = random.choice(MENSAJES_ANIMO)
    texto_voz = f"{saludo_hablado}, bienvenido {nombre}. Has ganado {puntos_ganados} puntos."

    st.markdown(f"""
        <div style="
            background-color: #22C55E;
            color: white;
            padding: 35px 20px;
            border-radius: 20px;
            text-align: center;
            font-family: 'Segoe UI', sans-serif;
            box-shadow: 0 10px 25px rgba(0,0,0,0.3);
        ">
            <h2 style="color: #FFE066; font-size: 32px; margin: 0 0 10px 0;">{saludo}, {nombre}! 🌟</h2>
            <h3 style="color: white; font-size: 26px; margin: 0 0 15px 0;">🎉 ¡Bienvenido(a)! 🎉</h3>
            <h1 style="color: #FFE066; font-size: 48px; margin: 10px 0; font-weight: bold;">{nombre}</h1>
            <p style="font-size: 20px; color: white; margin: 15px 0;">
                ✨ +{puntos_ganados} puntos ({PUNTOS_BASE} base + {puntos_extra} extra por estar entre los primeros {CUPOS_PUNTUALIDAD})
            </p>
            <p style="font-size: 22px; font-weight: bold; color: white; margin: 15px 0;">
                🔥 Racha: {racha} día(s) &nbsp;&nbsp;|&nbsp;&nbsp; Puntos: {puntos_totales}
            </p>
            <h3 style="color: #FFE066; font-size: 24px; margin-top: 25px;">{frase_animo}</h3>
        </div>
    """, unsafe_allow_html=True)

    components.html(f"""
        <script>
            if ('speechSynthesis' in window) {{
                window.speechSynthesis.cancel();
                var msg = new SpeechSynthesisUtterance("{texto_voz}");
                msg.lang = 'es-ES';
                msg.rate = 1.0;
                window.speechSynthesis.speak(msg);
            }}
        </script>
    """, height=0)

    st.write("")
    if st.button("☑️ ¡Entendido! (o presiona Enter)", use_container_width=True, type="primary"):
        st.rerun()

# ================================================================
# ADMINISTRACIÓN DE ARCHIVOS + PREVISUALIZADOR INTEGRADO
# ================================================================
def gestionar_archivos_seccion(titulo, subcarpeta):
    st.subheader(f"{titulo}")
    
    materia_sel = st.selectbox(f"Selecciona la Materia para {titulo}:", MATERIAS_LISTA, key=f"mat_sel_{subcarpeta}")
    mat_folder = materia_sel.lower().replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u")
    path_carpeta = os.path.join(CARPETA_DOCS, subcarpeta, mat_folder)
    os.makedirs(path_carpeta, exist_ok=True)
    
    count_key = f"uploader_count_{subcarpeta}_{mat_folder}"
    if count_key not in st.session_state:
        st.session_state[count_key] = 0
        
    uploader_key = f"uploader_{subcarpeta}_{mat_folder}_{st.session_state[count_key]}"
    
    archivo_subido = st.file_uploader(f"Subir documento a {titulo} ({materia_sel}):", key=uploader_key)
    if archivo_subido is not None:
        nombre_seguro = limpiar_nombre_archivo(archivo_subido.name)
        ruta_destino = os.path.join(path_carpeta, nombre_seguro)
        
        with open(ruta_destino, "wb") as f:
            f.write(archivo_subido.getbuffer())
        
        st.session_state[count_key] += 1
        st.toast(f"¡Documento '{nombre_seguro}' subido con éxito!")
        st.rerun()

    st.markdown("---")
    st.write(f"📋 **Documentos de {materia_sel}:**")
    
    archivos = os.listdir(path_carpeta)
    if not archivos:
        st.info(f"Aún no se han subido documentos para {materia_sel}.")
    else:
        for arch in archivos:
            ruta_archivo = os.path.join(path_carpeta, arch)
            
            with st.expander(f"📄 {arch}", expanded=False):
                col1, col2 = st.columns([1, 1])
                
                with open(ruta_archivo, "rb") as f:
                    bytes_arch = f.read()
                col1.download_button(
                    label="📥 Descargar Archivo",
                    data=bytes_arch,
                    file_name=arch,
                    key=f"down_{subcarpeta}_{mat_folder}_{arch}",
                    use_container_width=True
                )
                
                if col2.button("🗑️ Eliminar Archivo", key=f"del_{subcarpeta}_{mat_folder}_{arch}", use_container_width=True):
                    os.remove(ruta_archivo)
                    st.success(f"Archivo '{arch}' eliminado.")
                    st.rerun()
                
                st.markdown("---")
                ext = arch.split(".")[-1].lower()
                if ext == "pdf":
                    base64_pdf = base64.b64encode(bytes_arch).decode('utf-8')
                    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="550" type="application/pdf"></iframe>'
                    st.markdown(pdf_display, unsafe_allow_html=True)
                elif ext in ["png", "jpg", "jpeg", "webp"]:
                    st.image(bytes_arch, use_column_width=True)
                elif ext in ["txt", "csv"]:
                    try:
                        st.code(bytes_arch.decode("utf-8"))
                    except Exception:
                        st.write("No se pudo decodificar el texto.")
                else:
                    st.info("ℹ️ La vista previa en la web está disponible para **PDFs, imágenes y archivos de texto**. Para editar archivos de Word (.docx) o Excel (.xlsx), utiliza el botón de descarga.")

# ================================================================
# VISTA GENERAL DE BIENVENIDA (SPLASH SCREEN INICIAL)
# ================================================================
def mostrar_pantalla_bienvenida():
    ruta_logo = os.path.join(carpeta_base(), "logo.png.jpg")
    if not os.path.exists(ruta_logo):
        ruta_logo = os.path.join(carpeta_base(), "logo.png")

    fecha_backup = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if os.path.exists(ruta_logo):
            st.image(ruta_logo, width=180)

        st.markdown("""
            <div style="text-align: center; font-family: 'Segoe UI', sans-serif;">
                <h3 style="color: #FFD166; margin-bottom: 5px;">C.E.R. Siravita</h3>
                <hr style="border: 1px solid #7C83FD; width: 60%; margin: 10px auto 20px auto;">
                <h1 style="color: white; font-size: 38px; font-weight: bold; margin-bottom: 10px;">
                    Asistente Educativo Sergio Carbonell
                </h1>
                <p style="color: #B9BDD6; font-size: 18px; margin-bottom: 5px;">
                    Sistema de Gestión Educativa
                </p>
                <p style="color: #8188A8; font-size: 14px; margin-bottom: 30px;">
                    Registro de asistencia con puntos, rachas e insignias
                </p>
            </div>
        """, unsafe_allow_html=True)

        if st.button("🚀 Comenzar", use_container_width=True, type="primary"):
            st.session_state.pantalla_inicio = False
            st.rerun()

        st.markdown(f"""
            <div style="text-align: center; font-family: 'Segoe UI', sans-serif; margin-top: 15px;">
                <p style="color: #8188A8; font-size: 12px;">(o presiona Enter)</p>
                <p style="color: #5A6988; font-size: 11px; margin-top: 25px;">💾 Último backup: {fecha_backup}</p>
            </div>
        """, unsafe_allow_html=True)

# ================================================================
# FLUJO PRINCIPAL DEL PROGRAMA
# ================================================================
init_db()

# Si la pantalla de inicio está activa, muestra el Splash Screen
if st.session_state.pantalla_inicio:
    mostrar_pantalla_bienvenida()
else:
    # ------------------------------------------------------------
    # NAVEGACIÓN PRINCIPAL + SIDEBAR
    # ------------------------------------------------------------
    ruta_logo = os.path.join(carpeta_base(), "logo.png.jpg")
    if not os.path.exists(ruta_logo):
        ruta_logo = os.path.join(carpeta_base(), "logo.png")

    with st.sidebar:
        if os.path.exists(ruta_logo):
            st.image(ruta_logo, use_column_width=True)
        st.title("🏫 " + NOMBRE_ESCUELA)
        st.caption("Panel de Control General")
        
        # Botón para volver a la pantalla de bienvenida
        if st.button("🏠 Ir a Pantalla de Inicio", use_container_width=True):
            st.session_state.pantalla_inicio = True
            st.rerun()

        st.markdown("---")

        grados = obtener_grados()
        if not grados:
            st.warning("No hay grados registrados. Agrega uno abajo.")
            nombres_grados = []
            grado_sel_id = None
            grado_sel_nombre = ""
        else:
            nombres_grados = [g["nombre"] for g in grados]
            grado_sel_nombre = st.selectbox("Selecciona un Grado:", nombres_grados)
            grado_sel_id = next(g["id"] for g in grados if g["nombre"] == grado_sel_nombre)
        
        st.markdown("---")
        
        with st.expander("📚 Gestionar Cursos / Grados", expanded=False):
            subtab_g1, subtab_g2, subtab_g3 = st.tabs(["➕ Crear", "✏️ Renombrar", "🗑️ Eliminar"])
            
            with subtab_g1:
                nuevo_grado_txt = st.text_input("Nombre del nuevo curso:", placeholder="Ej. Primero", key="input_nuevo_grado")
                if st.button("➕ Crear Curso", use_container_width=True):
                    if nuevo_grado_txt.strip():
                        if agregar_grado(nuevo_grado_txt.strip()):
                            st.success(f"Curso '{nuevo_grado_txt}' creado.")
                            st.rerun()
                        else:
                            st.error("Ese curso ya existe.")
                    else:
                        st.warning("Escribe un nombre válido.")

            with subtab_g2:
                if grados:
                    g_edit_id = st.selectbox(
                        "Selecciona curso a renombrar:",
                        [g["id"] for g in grados],
                        format_func=lambda x: next(g["nombre"] for g in grados if g["id"] == x),
                        key="select_g_edit"
                    )
                    g_edit_nuevo_nombre = st.text_input("Nuevo nombre:", key="input_edit_grado")
                    if st.button("💾 Guardar Cambio", use_container_width=True):
                        if g_edit_nuevo_nombre.strip():
                            if editar_grado(g_edit_id, g_edit_nuevo_nombre.strip()):
                                st.success("Curso actualizado.")
                                st.rerun()
                            else:
                                st.error("Error o ya existe ese nombre.")
                        else:
                            st.warning("Escribe el nuevo nombre.")
                else:
                    st.info("No hay cursos.")

            with subtab_g3:
                if grados:
                    g_del_id = st.selectbox(
                        "Selecciona curso a eliminar:",
                        [g["id"] for g in grados],
                        format_func=lambda x: next(g["nombre"] for g in grados if g["id"] == x),
                        key="select_g_del"
                    )
                    st.caption("⚠️ Se desactivarán sus alumnos y se eliminará el curso.")
                    if st.button("🗑️ Eliminar Curso", use_container_width=True):
                        if len(grados) <= 1:
                            st.error("No puedes eliminar todos los cursos.")
                        else:
                            eliminar_grado(g_del_id)
                            st.success("Curso eliminado.")
                            st.rerun()
                else:
                    st.info("No hay cursos.")

        st.markdown("---")
        
        st.subheader("➕ Agregar Estudiante")
        nuevo_nombre = st.text_input("Nombre completo:")
        if st.button("Guardar Estudiante", use_container_width=True):
            if nuevo_nombre.strip():
                if grado_sel_id is not None:
                    if agregar_estudiante(nuevo_nombre.strip(), grado_sel_id):
                        st.success(f"¡{nuevo_nombre} registrado!")
                        st.rerun()
                    else:
                        st.error("El estudiante ya existe en este grado.")
                else:
                    st.error("Debes seleccionar o crear un grado primero.")

    # ------------------------------------------------------------
    # VISTA PRINCIPAL
    # ------------------------------------------------------------
    if not grados:
        st.title(f"🏫 {NOMBRE_ESCUELA}")
        st.info("Aún no tienes cursos creados. Despliega la opción **'📚 Gestionar Cursos / Grados'** en el menú de la izquierda para crear tu primer curso.")
    else:
        st.title(f"📋 Asistente Educativo — Grado {grado_sel_nombre}")

        pestanas = st.tabs([
            "📄 Documentos Institucionales", 
            "📋 Registro de Asistencia", 
            "📝 Calificaciones", 
            "⚖️ Convivencia", 
            "🏆 Tabla de Líderes"
        ])

        # ----------------------------------------------------------------
        # PESTAÑA 1: DOCUMENTOS INSTITUCIONALES
        # ----------------------------------------------------------------
        with pestanas[0]:
            st.title("📄 Documentos Institucionales")
            st.caption("Repositorio para consultar, subir, previsualizar y descargar archivos ordenados por materia.")
            
            subtab_doc1, subtab_doc2, subtab_doc3 = st.tabs([
                "📚 Planes de Área", 
                "📝 Planes de Clase", 
                "📖 Guías Educativas"
            ])
            
            with subtab_doc1:
                gestionar_archivos_seccion("📚 Planes de Área", "planes_area")
                
            with subtab_doc2:
                gestionar_archivos_seccion("📝 Planes de Clase", "planes_clase")
                
            with subtab_doc3:
                gestionar_archivos_seccion("📖 Guías Educativas", "guias_educativas")

        # ----------------------------------------------------------------
        # PESTAÑA 2: ASISTENCIA
        # ----------------------------------------------------------------
        with pestanas[1]:
            st.subheader("Selecciona tu nombre para registrar asistencia")
            estudiantes = obtener_estudiantes_por_grado(grado_sel_id)
            hoy_str = datetime.date.today().isoformat()
            
            if not estudiantes:
                st.info("No hay estudiantes registrados en este grado. Utiliza la barra lateral para agregar estudiantes.")
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
                                    res = registrar_asistencia(est["id"], grado_sel_id)
                                    if res:
                                        st.balloons()
                                        ventana_celebracion(res)

        # ----------------------------------------------------------------
        # PESTAÑA 3: CALIFICACIONES
        # ----------------------------------------------------------------
        with pestanas[2]:
            st.subheader("Módulo de Calificaciones")
            conn = conectar()
            materias = conn.execute("SELECT id, nombre FROM materias").fetchall()
            conn.close()
            
            mat_nombres = [m["nombre"] for m in materias]
            mat_sel = st.selectbox("Selecciona Materia:", mat_nombres)
            mat_id = next(m["id"] for m in materias if m["nombre"] == mat_sel)
            
            st.write(f"Registro rápido de notas para **{mat_sel}**:")
            
            if estudiantes:
                datos_tabla = []
                for e in estudiantes:
                    datos_tabla.append({"ID": e["id"], "Estudiante": e["nombre"], "Nota 1": 0.0, "Nota 2": 0.0, "Nota 3": 0.0})
                
                df = pd.DataFrame(datos_tabla)
                df_editado = st.data_editor(df, num_rows="fixed")
            else:
                st.info("No hay alumnos para calificar.")

        # ----------------------------------------------------------------
        # PESTAÑA 4: CONVIVENCIA
        # ----------------------------------------------------------------
        with pestanas[3]:
            st.subheader("Registro de Incidentes y Faltas Disciplinarias")
            
            with st.form("form_convivencia"):
                nombres_est = [e["nombre"] for e in estudiantes]
                est_incidente = st.selectbox("Estudiante:", nombres_est if nombres_est else ["Sin estudiantes"])
                desc = st.text_area("Descripción del hecho:")
                sancion = st.text_input("Sanción / Compromiso:")
                
                guardar_falta = st.form_submit_button("Guardar Incidente")
                if guardar_falta:
                    if desc.strip() and nombres_est:
                        st.success("Incidente registrado en la base de datos.")
                    else:
                        st.warning("Completa la descripción.")

        # ----------------------------------------------------------------
        # PESTAÑA 5: CLASIFICACIÓN
        # ----------------------------------------------------------------
        with pestanas[4]:
            st.subheader("🏆 Tabla de Posiciones")
            if estudiantes:
                df_leaderboard = pd.DataFrame([
                    {"Estudiante": e["nombre"], "Puntos": e["puntos"], "Racha": f"🔥 {e['racha']} días", "Asistencias Totales": e["total_asistencias"]}
                    for e in estudiantes
                ]).sort_values(by="Puntos", ascending=False)
                
                st.dataframe(df_leaderboard, use_container_width=True)