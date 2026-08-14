# -*- coding: utf-8 -*-
"""
SISTEMA INTEGRAL EDUCATIVO — C.E.R. SIRAVITA (SUPABASE NUBE COMPLETO Y OPTIMIZADO)
=====================================================================
Todas las funcionalidades: Documentos, Asistencia Rápida, Calificaciones, 
Convivencia, Tabla de Líderes y Gestión Completa de Cursos/Alumnos.
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
    "Inglés", "Matemáticas", "Español", "Sociales", 
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

if "profesor" not in st.session_state:
    st.session_state.profesor = None

# ================================================================
# AUTENTICACIÓN DE PROFESORES
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
# CONSULTAS BASE DE DATOS NUBE
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

def agregar_estudiante(nombre, grado_id, profesor_id):
    supabase.table("estudiantes").insert({
        "nombre": nombre, "grado_id": grado_id, "profesor_id": profesor_id,
        "puntos": 0, "racha": 0, "activo": 1
    }).execute()

def editar_estudiante(estudiante_id, nuevo_nombre):
    supabase.table("estudiantes").update({"nombre": nuevo_nombre}).eq("id", estudiante_id).execute()

def eliminar_estudiante(estudiante_id):
    supabase.table("estudiantes").update({"activo": 0}).eq("id", estudiante_id).execute()

def ya_registrado_hoy(estudiante_id, fecha):
    res = supabase.table("registros").select("id").eq("estudiante_id", estudiante_id).eq("fecha", fecha).execute()
    return len(res.data) > 0

# REGISTRO DE ASISTENCIA RÁPIDO Y OPTIMIZADO
def registrar_asistencia(estudiante, grado_id, profesor_id):
    hoy = datetime.date.today().isoformat()
    hora_str = datetime.datetime.now().strftime("%H:%M:%S")
    
    if ya_registrado_hoy(estudiante["id"], hoy):
        return None

    # Consulta directa rápida del orden
    res_hoy = supabase.table("registros").select("id", count="exact").eq("fecha", hoy).eq("grado_id", grado_id).execute()
    orden = (res_hoy.count or 0) + 1
    
    puntos_extra = PUNTOS_EXTRA_PUNTUALIDAD if orden <= CUPOS_PUNTUALIDAD else 0
    puntos_totales = PUNTOS_BASE + puntos_extra
    
    ayer = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
    nueva_racha = estudiante["racha"] + 1 if estudiante["ultima_fecha"] == ayer else 1
    nuevos_puntos = estudiante["puntos"] + puntos_totales
    
    # Actualizaciones inmediatas
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

    # Voz sintetizada inmediata
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
                if st.button("➕ Guardar Alumno", use_container_width=True):
                    if nom_est.strip():
                        agregar_estudiante(nom_est.strip(), grado_sel_id, prof["id"])
                        st.success("Alumno guardado.")
                        st.rerun()
                
                st.markdown("---")
                estudiantes_lista = obtener_estudiantes(grado_sel_id, prof["id"])
                if estudiantes_lista:
                    dict_est = {e["nombre"]: e for e in estudiantes_lista}
                    est_edit_nom = st.selectbox("Selecciona alumno a editar:", list(dict_est.keys()))
                    est_obj = dict_est[est_edit_nom]
                    
                    nuevo_nom_val = st.text_input("Nuevo nombre:", value=est_obj["nombre"])
                    col_e1, col_e2 = st.columns(2)
                    with col_e1:
                        if st.button("✏️ Actualizar", use_container_width=True):
                            editar_estudiante(est_obj["id"], nuevo_nom_val.strip())
                            st.success("Nombre actualizado.")
                            st.rerun()
                    with col_e2:
                        if st.button("🗑️ Eliminar", use_container_width=True):
                            eliminar_estudiante(est_obj["id"])
                            st.success("Alumno eliminado.")
                            st.rerun()

    # PESTAÑAS PRINCIPALES DEL SISTEMA
    st.title(f"📋 Asistente Educativo — {grado_sel_nombre if grado_sel_nombre else 'Crea un curso'}")

    if grado_sel_id:
        t_docs, t_asistencia, t_notas, t_convivencia, t_lideres = st.tabs([
            "📄 Documentos Institucionales",
            "📋 Registro de Asistencia", 
            "📝 Calificaciones",
            "⚖️ Convivencia",
            "🏆 Tabla de Líderes"
        ])

        # -------------------------------------------------------------
        # 1. DOCUMENTOS INSTITUCIONALES
        # -------------------------------------------------------------
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

        # -------------------------------------------------------------
        # 2. REGISTRO DE ASISTENCIA
        # -------------------------------------------------------------
        with t_asistencia:
            st.subheader("Marcar Asistencia Diaria")
            estudiantes = obtener_estudiantes(grado_sel_id, prof["id"])
            hoy_str = datetime.date.today().isoformat()
            
            if not estudiantes:
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

        # -------------------------------------------------------------
        # 3. CALIFICACIONES
        # -------------------------------------------------------------
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

        # -------------------------------------------------------------
        # 4. CONVIVENCIA
        # -------------------------------------------------------------
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

        # -------------------------------------------------------------
        # 5. TABLA DE LÍDERES
        # -------------------------------------------------------------
        with t_lideres:
            st.subheader("🏆 Clasificación General por Puntos y Rachas")
            estudiantes = obtener_estudiantes(grado_sel_id, prof["id"])
            if estudiantes:
                df = pd.DataFrame([
                    {"Estudiante": e["nombre"], "Puntos Acumulados": e["puntos"], "Racha Actual": f"🔥 {e['racha']} días", "Asistencias Totales": e.get("total_asistencias", 0)}
                    for e in estudiantes
                ]).sort_values(by="Puntos Acumulados", ascending=False)
                
                st.dataframe(df, use_container_width=True)
