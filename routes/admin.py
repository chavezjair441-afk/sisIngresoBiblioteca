from flask import Blueprint, render_template, request, jsonify, send_file
import pandas as pd
import io
from datetime import datetime
from db import get_db_connection

admin_bp = Blueprint('admin', __name__, url_prefix='/admin')

@admin_bp.route('/')
def admin_dashboard():
    conn = get_db_connection()
    cursor = conn.cursor()

    # 1. Totales
    cursor.execute("SELECT COUNT(*) FROM RegistroIngresos WHERE CAST(FechaHora AS DATE) = CAST(GETDATE() AS DATE)")
    total_hoy = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM RegistroIngresos WHERE VisitanteID IS NOT NULL AND CAST(FechaHora AS DATE) = CAST(GETDATE() AS DATE)")
    total_visitantes = cursor.fetchone()[0]

    # 2. Por Piso
    cursor.execute("SELECT Piso, COUNT(*) FROM RegistroIngresos WHERE CAST(FechaHora AS DATE) = CAST(GETDATE() AS DATE) GROUP BY Piso")
    pisos_dict = {row[0]: row[1] for row in cursor.fetchall()}

    # 3. Gráfico Horas
    cursor.execute("""
        SELECT DATEPART(HOUR, FechaHora) as Hora, COUNT(*) 
        FROM RegistroIngresos WHERE CAST(FechaHora AS DATE) = CAST(GETDATE() AS DATE) 
        GROUP BY DATEPART(HOUR, FechaHora) ORDER BY Hora
    """)
    datos_horas = cursor.fetchall()
    chart_horas_labels = [f"{row[0]}:00" for row in datos_horas]
    chart_horas_values = [row[1] for row in datos_horas]

    # 4. Top Orígenes (Unificado)
    cursor.execute("""
        SELECT TOP 5 Origen, COUNT(*) as Cantidad FROM (
            SELECT A.Escuela as Origen FROM RegistroIngresos R JOIN Alumnos A ON R.AlumnoID = A.AlumnoID 
            WHERE CAST(R.FechaHora AS DATE) = CAST(GETDATE() AS DATE)
            UNION ALL
            SELECT V.Institucion as Origen FROM RegistroIngresos R JOIN Visitantes V ON R.VisitanteID = V.VisitanteID 
            WHERE CAST(R.FechaHora AS DATE) = CAST(GETDATE() AS DATE)
        ) as T GROUP BY Origen ORDER BY Cantidad DESC
    """)
    datos_escuelas = cursor.fetchall()
    chart_escuelas_labels = [row[0] for row in datos_escuelas]
    chart_escuelas_values = [row[1] for row in datos_escuelas]

    # 5. Tabla Últimos
    cursor.execute("""
        SELECT TOP 10 COALESCE(A.NombreCompleto, V.NombreCompleto), R.Piso, FORMAT(R.FechaHora, 'HH:mm:ss'), 
        COALESCE(A.Escuela, V.Institucion), CASE WHEN R.VisitanteID IS NOT NULL THEN 'Visitante' ELSE 'Alumno' END
        FROM RegistroIngresos R
        LEFT JOIN Alumnos A ON R.AlumnoID = A.AlumnoID
        LEFT JOIN Visitantes V ON R.VisitanteID = V.VisitanteID
        ORDER BY R.FechaHora DESC
    """)
    ultimos = cursor.fetchall()
    conn.close()

    return render_template('admin_dashboard.html', total_hoy=total_hoy, total_visitantes=total_visitantes, 
                           pisos=pisos_dict, ultimos=ultimos, labels_horas=chart_horas_labels, 
                           data_horas=chart_horas_values, labels_escuelas=chart_escuelas_labels, 
                           data_escuelas=chart_escuelas_values)

# --- RUTAS DE EXCEL Y REPORTES ---

@admin_bp.route('/subir_excel', methods=['POST'])
def subir_excel():
    if 'archivo_excel' not in request.files: return jsonify({'status': 'error', 'msg': 'Falta archivo'})
    file = request.files['archivo_excel']
    if file.filename == '': return jsonify({'status': 'error', 'msg': 'Nombre vacío'})

    try:
        df = pd.read_excel(file, dtype={'DNI': str})
        conn = get_db_connection()
        cursor = conn.cursor()
        contador = 0
        
        for _, row in df.iterrows():
            dni = str(row.get('DNI', '')).strip()
            nombre = row.get('APELLIDOS Y NOMBRE', '')
            codigo = str(row.get('CODIGO DE MATRICULA', '')).strip()
            escuela = row.get('ESCUELA', '')
            semestre = row.get('SEMESTRE', '')
            
            if not dni or len(dni) < 5: continue

            # Lógica Upsert (Insertar o Actualizar)
            cursor.execute("SELECT AlumnoID FROM Alumnos WHERE DNI = ?", (dni,))
            if cursor.fetchone():
                cursor.execute("UPDATE Alumnos SET NombreCompleto=?, CodigoMatricula=?, Escuela=?, Semestre=?, Estado=1 WHERE DNI=?", 
                               (nombre, codigo, escuela, semestre, dni))
            else:
                cursor.execute("INSERT INTO Alumnos (NombreCompleto, DNI, CodigoMatricula, Escuela, Semestre) VALUES (?,?,?,?,?)", 
                               (nombre, dni, codigo, escuela, semestre))
            contador += 1
            
        conn.commit()
        conn.close()
        return jsonify({'status': 'success', 'msg': f'Procesados {contador} alumnos.'})
    except Exception as e:
        return jsonify({'status': 'error', 'msg': str(e)})

@admin_bp.route('/reporte_hoy')
def descargar_reporte():
    conn = get_db_connection()
    sql = """
    SELECT R.RegistroID, A.NombreCompleto, A.DNI, A.Escuela, R.Piso, FORMAT(R.FechaHora, 'HH:mm:ss') as Hora
    FROM RegistroIngresos R JOIN Alumnos A ON R.AlumnoID = A.AlumnoID
    WHERE CAST(R.FechaHora AS DATE) = CAST(GETDATE() AS DATE) ORDER BY R.FechaHora DESC
    """
    df = pd.read_sql(sql, conn)
    conn.close()
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return send_file(output, download_name=f"Reporte_{datetime.now().date()}.xlsx", as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# --- AGREGAR AL FINAL DE routes/admin.py ---

@admin_bp.route('/reporte_rango')
def reporte_rango():
    # Recibimos las fechas desde la URL (ej: ?inicio=2026-02-01&fin=2026-02-05)
    fecha_inicio = request.args.get('inicio')
    fecha_fin = request.args.get('fin')

    if not fecha_inicio or not fecha_fin:
        return "Error: Debes seleccionar ambas fechas", 400

    conn = get_db_connection()
    
    # Consulta SQL filtrando por rango de fechas
    sql = """
    SELECT 
        R.RegistroID as ID,
        COALESCE(A.NombreCompleto, V.NombreCompleto) as Persona,
        COALESCE(A.DNI, V.DNI) as DNI,
        COALESCE(A.Escuela, V.Institucion) as Origen,
        CASE WHEN R.VisitanteID IS NOT NULL THEN 'VISITANTE' ELSE 'ALUMNO' END as Tipo,
        R.Piso as Piso_Acceso,
        R.Turno,
        FORMAT(R.FechaHora, 'HH:mm:ss') as Hora,
        FORMAT(R.FechaHora, 'dd/MM/yyyy') as Fecha
    FROM RegistroIngresos R
    LEFT JOIN Alumnos A ON R.AlumnoID = A.AlumnoID
    LEFT JOIN Visitantes V ON R.VisitanteID = V.VisitanteID
    WHERE CAST(R.FechaHora AS DATE) >= ? 
    AND CAST(R.FechaHora AS DATE) <= ?
    ORDER BY R.FechaHora DESC
    """
    
    # Ejecutamos la consulta enviando las fechas
    # IMPORTANTE: Asegúrate de tener instalado openpyxl (pip install openpyxl)
    df = pd.read_sql(sql, conn, params=(fecha_inicio, fecha_fin))
    conn.close()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Reporte_Rango')
    output.seek(0)

    return send_file(output, 
                    download_name=f"Reporte_{fecha_inicio}_al_{fecha_fin}.xlsx", 
                    as_attachment=True,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')