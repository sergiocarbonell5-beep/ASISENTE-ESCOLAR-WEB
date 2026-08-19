# ================================================================
        # 1. REPOSITORIO DE DOCUMENTOS INSTITUCIONALES (SOLUCIÓN DEFINITIVA)
        # ================================================================
        with t_docs:
            st.subheader("📁 Repositorio de Documentos Institucionales")
            st.caption("Gestiona el Plan de Área general o los Ejes Temáticos por periodo escolar.")

            tipo_doc_sel = st.radio("Categoría de Documento:", ["Planes de Área", "Ejes Temáticos"], horizontal=True)

            try:
                res_docs_db = supabase.table("documentos").select("*").eq("profesor_id", prof["id"]).execute()
                docs_existentes = res_docs_db.data or []
            except Exception as err_db:
                docs_existentes = []

            # Mapeo robusto multi-filtro
            dict_planes_map = {}
            dict_ejes_map = {}

            for d_item in docs_existentes:
                mat_k = normalizar_texto(d_item.get("materia"))
                t_doc = str(d_item.get("tipo_doc", ""))
                per_val = str(d_item.get("periodo") or "")

                # Extraer periodo si está guardado dentro del tipo_doc (ej: "Ejes Temáticos P3")
                if not per_val or per_val == "None":
                    match_p = re.search(r'P([1-4])', t_doc)
                    if match_p:
                        per_val = match_p.group(1)

                if "Plan" in t_doc or t_doc == "Planes de Área":
                    dict_planes_map[mat_k] = d_item
                elif "Eje" in t_doc or "Ejes" in t_doc:
                    dict_ejes_map[(mat_k, per_val)] = d_item

            st.markdown("---")
            v_key = st.session_state.upload_ver

            # MODO 1: PLANES DE ÁREA
            if tipo_doc_sel == "Planes de Área":
                st.markdown(f"### 📑 Planes de Área Generales — {grado_sel_nombre}")
                
                col_m_head, col_doc_h = st.columns([2.5, 5])
                with col_m_head:
                    st.markdown("**Asignatura / Área**")
                with col_doc_h:
                    st.markdown("**Documento Plan de Área (Anual / Consolidado)**")

                st.markdown("---")

                for m_idx, mat_nombre in enumerate(MATERIAS_LISTA):
                    key_mat_norm = normalizar_texto(mat_nombre)
                    c_m, c_d = st.columns([2.5, 5])
                    
                    with c_m:
                        st.write(f"**{m_idx + 1}. {mat_nombre}**")
                    
                    with c_d:
                        doc_guardado = dict_planes_map.get(key_mat_norm)
                        
                        if doc_guardado:
                            col_info, col_btn = st.columns([3, 2])
                            with col_info:
                                st.success(f"📄 **{doc_guardado['nombre']}**")
                            with col_btn:
                                bytes_dec = base64.b64decode(doc_guardado['contenido_b64'])
                                st.download_button(
                                    label="⬇️ Descargar Plan",
                                    data=bytes_dec,
                                    file_name=doc_guardado['nombre'],
                                    key=f"down_pa_{m_idx}",
                                    use_container_width=True
                                )
                        else:
                            f_up = st.file_uploader(
                                label=f"Subir Plan {mat_nombre}",
                                type=["pdf", "docx", "pptx", "xlsx", "txt"],
                                key=f"up_pa_{m_idx}_v{v_key}",
                                label_visibility="collapsed"
                            )
                            if f_up is not None:
                                bytes_data = f_up.getvalue()
                                b64_str = base64.b64encode(bytes_data).decode('utf-8')
                                
                                payload = {
                                    "nombre": f_up.name,
                                    "materia": mat_nombre,
                                    "tipo_doc": "Planes de Área",
                                    "contenido_b64": b64_str,
                                    "grado_id": grado_sel_id,
                                    "profesor_id": prof["id"]
                                }
                                
                                try:
                                    supabase.table("documentos").insert(payload).execute()
                                    st.session_state.upload_ver += 1
                                    st.toast(f"✅ ¡Plan guardado para {mat_nombre}!")
                                    st.rerun()
                                except Exception as ex_db:
                                    st.error(f"Error al subir: {ex_db}")

                    st.divider()

            # MODO 2: EJES TEMÁTICOS (CORREGIDO PARA MOSTRAR BOTÓN DE DESCARGA)
            else:
                st.markdown(f"### 📑 Matriz de Ejes Temáticos — {grado_sel_nombre}")

                col_m_head, col_p1_h, col_p2_h, col_p3_h, col_p4_h = st.columns([2.5, 1.8, 1.8, 1.8, 1.8])
                with col_m_head:
                    st.markdown("**Asignatura / Área**")
                with col_p1_h:
                    st.markdown("**1️⃣ Primer Periodo**")
                with col_p2_h:
                    st.markdown("**2️⃣ Segundo Periodo**")
                with col_p3_h:
                    st.markdown("**3️⃣ Tercer Periodo**")
                with col_p4_h:
                    st.markdown("**4️⃣ Cuarto Periodo**")

                st.markdown("---")

                for m_idx, mat_nombre in enumerate(MATERIAS_LISTA):
                    key_mat_norm = normalizar_texto(mat_nombre)
                    cols_p = st.columns([2.5, 1.8, 1.8, 1.8, 1.8])
                    
                    with cols_p[0]:
                        st.write(f"**{m_idx + 1}. {mat_nombre}**")

                    for p_num, p_col in zip(["1", "2", "3", "4"], cols_p[1:]):
                        with p_col:
                            doc_guardado = dict_ejes_map.get((key_mat_norm, p_num))
                            
                            if doc_guardado:
                                st.success(f"📄 {doc_guardado['nombre'][:12]}...")
                                bytes_dec = base64.b64decode(doc_guardado['contenido_b64'])
                                st.download_button(
                                    label="⬇️ Descargar",
                                    data=bytes_dec,
                                    file_name=doc_guardado['nombre'],
                                    key=f"down_et_{m_idx}_{p_num}",
                                    use_container_width=True
                                )
                            else:
                                f_up = st.file_uploader(
                                    label=f"Subir P{p_num}",
                                    type=["pdf", "docx", "pptx", "xlsx", "txt"],
                                    key=f"up_et_{m_idx}_{p_num}_v{v_key}",
                                    label_visibility="collapsed"
                                )
                                if f_up is not None:
                                    bytes_data = f_up.getvalue()
                                    b64_str = base64.b64encode(bytes_data).decode('utf-8')
                                    
                                    # Formato de doble respaldo
                                    payload = {
                                        "nombre": f_up.name,
                                        "materia": mat_nombre,
                                        "tipo_doc": f"Ejes Temáticos P{p_num}",
                                        "periodo": p_num,
                                        "contenido_b64": b64_str,
                                        "grado_id": grado_sel_id,
                                        "profesor_id": prof["id"]
                                    }
                                    
                                    try:
                                        supabase.table("documentos").insert(payload).execute()
                                        st.session_state.upload_ver += 1
                                        st.toast(f"✅ ¡{f_up.name} guardado en Periodo {p_num}!")
                                        st.rerun()
                                    except Exception:
                                        # Si la columna 'periodo' no existe en Supabase
                                        try:
                                            payload.pop("periodo", None)
                                            supabase.table("documentos").insert(payload).execute()
                                            st.session_state.upload_ver += 1
                                            st.toast(f"✅ ¡{f_up.name} guardado en Periodo {p_num}!")
                                            st.rerun()
                                        except Exception as ex_err:
                                            st.error(f"Error en Supabase: {ex_err}")

                    st.divider()
