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
    p.drawString(115, 415, str(sede_nombre).upper()) # Ajustado sobre la línea SEDE
    p.drawString(335, 415, str(mes_nombre).upper())  # Ajustado sobre la línea MES DE

    # 3. Coordenadas Calibradas para las Filas de la Tabla
    x_alumnos = 12
    x_grado = 86
    x_dias_inicio = 132.5
    w_dia = 13.5
    x_total = 560

    y_fila_inicio = 344  # Bajado para no pisar el encabezado verde
    h_fila = 18.8        # Distancia exacta entre renglones

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
