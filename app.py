# -*- coding: utf-8 -*-
"""
SISTEMA INTEGRAL EDUCATIVO — CONEXIÓN NUBE SUPABASE (MULTI-DOCENTE)
===================================================================
Base de datos persistente en Supabase (nada se borra al reiniciar).
Gestión de Login / Registro de Profesores con aislamiento de datos.
"""

import os
import re
import base64
import datetime
import random
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
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

# ================================================================
# CONEXIÓN CON SUPABASE
# ================================================================
@st.cache_resource
def init_supabase() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase = init_supabase()

# Variables de estado
if "profesor" not in st.session_state:
    st.session_state.profesor = None

# ================================================================
# FUNCIONES DE AUTENTICACIÓN
# ================================================================
def registrar_profesor(nombre, email, password):
    email_clean = email.strip().lower()
    res = supabase.table("profesores").select("id").eq("email", email_clean).execute()
    if res.data:
        return False, "El correo electrónico ya está registrado."
    
    data = {
        "nombre": nombre.strip(),
        "email": email_clean,
        "password": password.strip()
    }
    supabase.table("profesores").insert(data).execute()
    return True, "¡Profesor registrado exitosamente! Ya puedes iniciar sesión."

def login_profesor(email, password):
    email_clean = email.strip().lower()
    res = supabase.table("profesores").select("*").eq("email", email_clean).eq("password", password.strip()).execute()
    if res.data:
        return res.data[0]
    return None

# ================================================================
# CONSULTAS NUBE DE DATOS (FILTRADAS POR PROFESOR)
# ================================================================
def obtener_grados(profesor_id):
    res = supabase.table("grados").select("*").eq("profesor_id", profesor_id).order("nombre").execute()
    return res.data

def agregar_grado(nombre, profesor_id):
    supabase.table("grados").insert({"nombre": nombre, "profesor_id": profesor_id}).execute()

def obtener_estudiantes_por_grado(grado_id, profesor_id):
    res = supabase.table("estudiantes").select("*").eq("grado_id", grado_id).eq("profesor_id", profesor_id).eq("activo", 1).order("nombre").execute()
    return res.data

def agregar_estudiante(nombre, grado_id, profesor_id):
    data = {
        "nombre": nombre,
        "grado_id": grado_id,
        "profesor_id": profesor_id,
        "puntos": 0,
        "racha": 0,
        "activo": 1
    }
    supabase.table("estudiantes").insert(data).execute()

def ya_registrado_hoy(estudiante_id, fecha):
    res = supabase.table("registros").select("id").eq("estudiante_id", estudiante_id).eq("fecha", fecha).execute()
    return len(res.data) > 0

def registrar_asistencia(estudiante, grado_id, profesor_id):
    hoy = datetime.date.today().isoformat()
    hora_str = datetime.datetime.now().strftime("%H:%M:%S")
    
    if ya_registrado_hoy(estudiante["id"], hoy):
        return None

    # Contar cuántos han llegado hoy en este grado
    res_hoy = supabase.table("registros").select("id").eq("fecha", hoy).eq("grado_id", grado_id).execute()
    orden = len(res_hoy.data) + 1
    
    puntos_extra = PUNTOS_EXTRA_PUNTUALIDAD if orden <= CUPOS_PUNTUALIDAD else 0
    puntos_totales = PUNTOS_BASE + puntos_extra
    
    # Calcular racha
    ayer = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
    nueva_racha = estudiante["racha"] + 1 if estudiante["ultima_fecha"] == ayer else 1
    nuevos_puntos = estudiante["puntos"] + puntos_totales
    
    # Actualizar Estudiante
    supabase.table("estudiantes").update({
        "puntos": nuevos_puntos,
        "racha": nueva_racha,
        "ultima_fecha": hoy,
        "total_asistencias": estudiante.get("total_asistencias", 0) + 1
    }).eq("id", estudiante["id"]).execute()

    # Insertar Registro
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

# ================================================================
# MODAL VENTANA CELEBRACIÓN CON VOZ
# ================================================================
@st.dialog("🎉 ¡Asistencia Registrada!", width="large")
def ventana_celebracion(res):
    nombre = res["nombre"]
    puntos_ganados = res["puntos_ganados"]
    puntos_extra = res["puntos_extra"]
    puntos_totales = res["puntos_totales"]
    racha = res["racha"]
    
    hora = datetime.datetime.now().hour
    saludo_hablado = "Buenos días" if 5 <= hora < 12 else ("Buenas tardes" if 12 <= hora < 19 else "Buenas noches")
    frase_animo = random.choice(MENSAJES_ANIMO)
    texto_voz = f"{saludo_hablado}, bienvenido {nombre}. Has ganado {puntos_ganados} puntos."

    st.markdown(f"""
        <div style="background-color: #22C55E; color: white; padding: 30px; border-radius: 20px; text-align: center;">
            <h2 style="color: #FFE066;">🌟 ¡{saludo_hablado}, {nombre}! 🌟</h2>
            <h1 style="color: #FFE066; font-size: 45px;">{nombre}</h1>
            <p style="font-size: 20px;">✨ +{puntos_ganados} puntos ganados</p>
            <p style="font-size: 22px; font-weight: bold;">🔥 Racha: {racha} día(s) | Puntos: {puntos_totales}</p>
            <h3 style="color: #FFE066;">{frase_animo}</h3>
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
    if st.button("☑️ ¡Entendido!", use_container_width=True, type="primary"):
        st.rerun()

# ================================================================
# INTERFAZ DE LOGIN / REGISTRO
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
# PANEL PRINCIPAL DEL PROFESOR
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
            grado_sel_nombre = st.selectbox("Selecciona tu Grado:", nombres_grados)
            grado_sel_id = next(g["id"] for g in grados if g["nombre"] == grado_sel_nombre)
            
        st.markdown("---")
        with st.expander("➕ Crear Nuevo Curso"):
            nuevo_grado_txt = st.text_input("Nombre del Curso:", placeholder="Ej. Grado Cuarto")
            if st.button("Guardar Curso", use_container_width=True):
                if nuevo_grado_txt.strip():
                    agregar_grado(nuevo_grado_txt.strip(), prof["id"])
                    st.success("¡Curso creado!")
                    st.rerun()

        st.markdown("---")
        st.subheader("➕ Agregar Estudiante")
        nom_est = st.text_input("Nombre de alumno:")
        if st.button("Guardar Alumno", use_container_width=True):
            if nom_est.strip() and grado_sel_id:
                agregar_estudiante(nom_est.strip(), grado_sel_id, prof["id"])
                st.success("Alumno guardado.")
                st.rerun()

    # VISTA DE PESTAÑAS
    st.title(f"📋 Panel Educativo — {grado_sel_nombre if grado_sel_nombre else 'Selecciona o crea un curso'}")

    if grado_sel_id:
        p1, p2 = st.tabs(["📋 Registro de Asistencia", "🏆 Tabla de Posiciones"])
        
        with p1:
            st.subheader("Selecciona un estudiante para marcar asistencia hoy")
            estudiantes = obtener_estudiantes_por_grado(grado_sel_id, prof["id"])
            hoy_str = datetime.date.today().isoformat()
            
            if not estudiantes:
                st.info("No hay alumnos en este curso. Agrégalos en la barra lateral.")
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

        with p2:
            st.subheader("🏆 Clasificación de Puntos y Rachas")
            estudiantes = obtener_estudiantes_por_grado(grado_sel_id, prof["id"])
            if estudiantes:
                df = pd.DataFrame([
                    {"Estudiante": e["nombre"], "Puntos": e["puntos"], "Racha": f"🔥 {e['racha']} días"}
                    for e in estudiantes
                ]).sort_values(by="Puntos", ascending=False)
                st.dataframe(df, use_container_width=True)
