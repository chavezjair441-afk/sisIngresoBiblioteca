import os
import io
from flask import Flask, render_template, request, jsonify, send_file
import pyodbc
import pandas as pd
from dotenv import load_dotenv
from datetime import datetime

# 1. CARGAR CONFIGURACIÓN SEGURA
load_dotenv()

app = Flask(__name__)

# 2. FUNCIÓN DE CONEXIÓN A SQL SERVER
def get_db_connection():
    try:
        conn_str = (
            f"DRIVER={os.getenv('DB_DRIVER')};"
            f"SERVER={os.getenv('DB_SERVER')};"
            f"DATABASE={os.getenv('DB_DATABASE')};"
            f"Trusted_Connection={os.getenv('DB_TRUSTED_CONNECTION')};"
        )
        return pyodbc.connect(conn_str)
    except Exception as e:
        print(f"--- ERROR DE CONEXIÓN BD --- : {e}")
        return None

# ==========================================
# RUTAS DE LOS PISOS (INGRESO)
# ==========================================

@app.route('/')
def index():
    return """
    <div style="text-align: center; font-family: sans-serif; padding-top: 50px;">
        <h1>Servidor Biblioteca UNDAC Activo</h1>
        <p>Accesos:</p>
        <a href="/piso1">Piso 1</a> | <a href="/piso2">Piso 2</a> | <a href="/piso3">Piso 3</a>
        <br><br>
        <a href="/admin"><b>PANEL ADMINISTRADOR</b></a>
    </div>
    """

@app.route('/piso1')
def piso1(): return render_template('ingreso.html', piso=1)

@app.route('/piso2')
def piso2(): return render_template('ingreso.html', piso=2)

@app.route('/piso3')
def piso3(): return render_template('ingreso.html', piso=3)

# API: PROCESAR EL ESCANEO (INGRESO)
@app.route('/procesar_ingreso', methods=['POST'])
def procesar_ingreso():
    data = request.json
    codigo = data.get('codigo')
    piso = data.get('piso')

    if not codigo: return jsonify({'status': 'error', 'msg': 'Código vacío'})

    conn = get_db_connection()
    if not conn: return jsonify({'status': 'error', 'msg': 'Error BD'})

    try:
        cursor = conn.cursor()
        sql = """
        DECLARE @out_msg nvarchar(100);
        DECLARE @out_nombre nvarchar(250);
        DECLARE @out_escuela nvarchar(100);
        DECLARE @out_semestre varchar(10);
        EXEC sp_RegistrarIngreso ?, ?, @out_msg OUTPUT, @out_nombre OUTPUT, @out_escuela OUTPUT, @out_semestre OUTPUT;
        SELECT @out_msg, @out_nombre, @out_escuela, @out_semestre;
        """
        cursor.execute(sql, (codigo, piso))
        row = cursor.fetchone()
        conn.commit()
        conn.close()

        if row and row[0] == 'ACCESO CONCEDIDO':
            return jsonify({
                'status': 'success', 'msg': row[0], 'alumno': row[1], 'escuela': row[2], 'semestre': row[3]
            })
        else:
            return jsonify({'status': 'error', 'msg': row[0] if row else 'Error'})

    except Exception as e:
        return jsonify({'status': 'error', 'msg': str(e)})


# ==========================================
# ZONA ADMINISTRADOR (CON GRÁFICOS)
# ==========================================

@app.route('/admin')
def admin_dashboard():
    conn = get_db_connection()
    cursor = conn.cursor()

    # 1. Total Hoy (Alumnos + Visitantes)
    cursor.execute("SELECT COUNT(*) FROM RegistroIngresos WHERE CAST(FechaHora AS DATE) = CAST(GETDATE() AS DATE)")
    total_hoy = cursor.fetchone()[0]

    # 1.1 NUEVO: Total Visitantes Hoy
    cursor.execute("SELECT COUNT(*) FROM RegistroIngresos WHERE VisitanteID IS NOT NULL AND CAST(FechaHora AS DATE) = CAST(GETDATE() AS DATE)")
    total_visitantes = cursor.fetchone()[0]

    # 2. Por Piso
    cursor.execute("""
        SELECT Piso, COUNT(*) 
        FROM RegistroIngresos 
        WHERE CAST(FechaHora AS DATE) = CAST(GETDATE() AS DATE) 
        GROUP BY Piso
    """)
    pisos_dict = {row[0]: row[1] for row in cursor.fetchall()}

    # 3. Gráfico Horas (Este cuenta a todos por igual, no cambia)
    cursor.execute("""
        SELECT DATEPART(HOUR, FechaHora) as Hora, COUNT(*) as Cantidad
        FROM RegistroIngresos
        WHERE CAST(FechaHora AS DATE) = CAST(GETDATE() AS DATE)
        GROUP BY DATEPART(HOUR, FechaHora) ORDER BY Hora
    """)
    datos_horas = cursor.fetchall()
    chart_horas_labels = [f"{row[0]}:00" for row in datos_horas]
    chart_horas_values = [row[1] for row in datos_horas]

    # 4. NUEVO: Gráfico Top Orígenes (Mezcla Escuelas e Instituciones)
    # Usamos UNION ALL para juntar a los alumnos y a los visitantes en una sola lista
    cursor.execute("""
        SELECT TOP 5 Origen, COUNT(*) as Cantidad
        FROM (
            SELECT A.Escuela as Origen 
            FROM RegistroIngresos R 
            JOIN Alumnos A ON R.AlumnoID = A.AlumnoID 
            WHERE CAST(R.FechaHora AS DATE) = CAST(GETDATE() AS DATE)
            
            UNION ALL
            
            SELECT V.Institucion as Origen 
            FROM RegistroIngresos R 
            JOIN Visitantes V ON R.VisitanteID = V.VisitanteID 
            WHERE CAST(R.FechaHora AS DATE) = CAST(GETDATE() AS DATE)
        ) as T
        GROUP BY Origen 
        ORDER BY Cantidad DESC
    """)
    datos_escuelas = cursor.fetchall()
    chart_escuelas_labels = [row[0] for row in datos_escuelas]
    chart_escuelas_values = [row[1] for row in datos_escuelas]

    # 5. NUEVO: Tabla Últimos Ingresos (Mezclada)
    # Usamos COALESCE para decir: "Si no hay nombre de alumno, usa el de visitante"
    cursor.execute("""
        SELECT TOP 10 
            COALESCE(A.NombreCompleto, V.NombreCompleto) as Nombre,
            R.Piso, 
            FORMAT(R.FechaHora, 'HH:mm:ss'), 
            COALESCE(A.Escuela, V.Institucion) as Origen,
            CASE WHEN R.VisitanteID IS NOT NULL THEN 'Visitante' ELSE 'Alumno' END as Tipo
        FROM RegistroIngresos R
        LEFT JOIN Alumnos A ON R.AlumnoID = A.AlumnoID
        LEFT JOIN Visitantes V ON R.VisitanteID = V.VisitanteID
        ORDER BY R.FechaHora DESC
    """)
    ultimos = cursor.fetchall()
    
    conn.close()

    return render_template('admin_dashboard.html', 
                           total_hoy=total_hoy, 
                           total_visitantes=total_visitantes, 
                           pisos=pisos_dict, 
                           ultimos=ultimos,
                           labels_horas=chart_horas_labels, data_horas=chart_horas_values,
                           labels_escuelas=chart_escuelas_labels, data_escuelas=chart_escuelas_values)

# RUTA NUEVA: PÁGINA DE VISITANTES
@app.route('/admin/visitantes')
def admin_visitantes_page():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Consultamos ID, Nombre, DNI, Institucion, Correo
    cursor.execute("""
        SELECT VisitanteID, NombreCompleto, DNI, Institucion, Correo 
        FROM Visitantes 
        ORDER BY NombreCompleto ASC
    """)
    lista_visitantes = cursor.fetchall()
    conn.close()

    # Enviamos la lista a la plantilla
    return render_template('admin_visitantes.html', visitantes=lista_visitantes)

# SUBIDA DE EXCEL
@app.route('/admin/subir_excel', methods=['POST'])
def subir_excel():
    if 'archivo_excel' not in request.files:
        return jsonify({'status': 'error', 'msg': 'No se envió ningún archivo'})
    
    file = request.files['archivo_excel']
    if file.filename == '':
        return jsonify({'status': 'error', 'msg': 'Nombre de archivo vacío'})

    try:
        # Leemos el Excel
        # dtype={'DNI': str} asegura que el DNI no pierda ceros iniciales
        df = pd.read_excel(file, dtype={'DNI': str})
        
        # --- MAPEO DE COLUMNAS (TRADUCTOR) ---
        # Aquí le decimos: "Lo que en Excel se llama X, en mi código es Y"
        # Basado en tu imagen:
        # Excel: 'APELLIDOS Y NOMBRE' -> BD: NombreCompleto
        # Excel: 'CODIGO DE MATRICULA' -> BD: CodigoMatricula
        
        conn = get_db_connection()
        cursor = conn.cursor()
        contador = 0
        
        for _, row in df.iterrows():
            # Usamos .get('NOMBRE_EXACTO_EN_EXCEL', '')
            dni = str(row.get('DNI', '')).strip()
            # En la imagen dice "APELLIDOS Y NOMBRE"
            nombre = row.get('APELLIDOS Y NOMBRE', '') 
            # En la imagen dice "CODIGO DE MATRICULA"
            codigo = str(row.get('CODIGO DE MATRICULA', '')).strip()
            escuela = row.get('ESCUELA', '')
            semestre = row.get('SEMESTRE', '')

            # Validar que al menos tenga DNI
            if not dni or len(dni) < 5:
                continue

            # Verificar si existe para actualizar o insertar
            cursor.execute("SELECT AlumnoID FROM Alumnos WHERE DNI = ?", (dni,))
            data = cursor.fetchone()

            if data:
                # Si existe, ACTUALIZAMOS sus datos (por si cambió de semestre o escuela)
                cursor.execute("""
                    UPDATE Alumnos 
                    SET NombreCompleto = ?, CodigoMatricula = ?, Escuela = ?, Semestre = ?, Estado = 1
                    WHERE DNI = ?
                """, (nombre, codigo, escuela, semestre, dni))
            else:
                # Si no existe, INSERTAMOS
                cursor.execute("""
                    INSERT INTO Alumnos (NombreCompleto, DNI, CodigoMatricula, Escuela, Semestre)
                    VALUES (?, ?, ?, ?, ?)
                """, (nombre, dni, codigo, escuela, semestre))
            
            contador += 1
            
        conn.commit()
        conn.close()
        return jsonify({'status': 'success', 'msg': f'Se procesaron {contador} alumnos correctamente.'})

    except Exception as e:
        print(f"Error procesando Excel: {e}")
        return jsonify({'status': 'error', 'msg': f'Error en el archivo: {str(e)}'})
    


# ... (resto de tu código) ...

# --- NUEVA RUTA PARA GENERAR REPORTE EXCEL ---
@app.route('/admin/reporte_hoy')
def descargar_reporte():
    conn = get_db_connection()
    
    # Consulta: Trae todos los ingresos DE HOY con detalles
    sql = """
    SELECT 
        R.RegistroID as ID,
        A.NombreCompleto as Alumno,
        A.DNI,
        A.Escuela,
        R.Piso as Piso_Acceso,
        FORMAT(R.FechaHora, 'HH:mm:ss') as Hora_Ingreso,
        FORMAT(R.FechaHora, 'dd/MM/yyyy') as Fecha
    FROM RegistroIngresos R
    JOIN Alumnos A ON R.AlumnoID = A.AlumnoID
    WHERE CAST(R.FechaHora AS DATE) = CAST(GETDATE() AS DATE)
    ORDER BY R.FechaHora DESC
    """
    
    # Usamos Pandas para convertir SQL a Excel en 1 línea
    df = pd.read_sql(sql, conn)
    conn.close()

    # Crear el archivo en memoria (RAM) para no llenar tu disco duro
    output = io.BytesIO()
    # Requiere instalar: pip install xlsxwriter (Si falla, usa openpyxl que ya tienes)
    writer = pd.ExcelWriter(output, engine='openpyxl') 
    df.to_excel(writer, index=False, sheet_name='Ingresos_Hoy')
    writer.close()
    output.seek(0)

    # Enviar al navegador
    from flask import send_file
    return send_file(output, 
                     download_name=f"Reporte_Ingresos_{datetime.now().strftime('%Y-%m-%d')}.xlsx", 
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    @app.route('/admin/reporte_rango')
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
            A.NombreCompleto as Alumno,
            A.DNI,
            A.Escuela,
            R.Piso as Piso_Acceso,
            FORMAT(R.FechaHora, 'HH:mm:ss') as Hora,
            FORMAT(R.FechaHora, 'dd/MM/yyyy') as Fecha
        FROM RegistroIngresos R
        JOIN Alumnos A ON R.AlumnoID = A.AlumnoID
        WHERE CAST(R.FechaHora AS DATE) >= ? 
        AND CAST(R.FechaHora AS DATE) <= ?
        ORDER BY R.FechaHora DESC
        """
        
        # Ejecutamos la consulta enviando las fechas
        df = pd.read_sql(sql, conn, params=(fecha_inicio, fecha_fin))
        conn.close()

        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='openpyxl') 
        df.to_excel(writer, index=False, sheet_name='Reporte_Rango')
        writer.close()
        output.seek(0)

        from flask import send_file
        return send_file(output, 
                        download_name=f"Reporte_{fecha_inicio}_al_{fecha_fin}.xlsx", 
                        as_attachment=True,
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    

    # ==========================================
# ZONA DE GESTIÓN DE VISITANTES / EXTERNOS
# ==========================================

# 1. Registrar UN solo visitante (Formulario Manual)
@app.route('/admin/agregar_visitante', methods=['POST'])
def agregar_visitante():
    data = request.json
    dni = data.get('dni')
    nombre = data.get('nombre')
    correo = data.get('correo')
    institucion = data.get('institucion')

    # Regla: Si institución viene vacío, poner "Sin Institución"
    if not institucion or institucion.strip() == "":
        institucion = "Sin Institución"

    if not dni or not nombre:
        return jsonify({'status': 'error', 'msg': 'DNI y Nombre son obligatorios'})

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        # Verificar si ya existe en Visitantes
        cursor.execute("SELECT 1 FROM Visitantes WHERE DNI = ?", (dni,))
        if cursor.fetchone():
             return jsonify({'status': 'error', 'msg': 'Este DNI ya está registrado como visitante.'})

        # Insertar
        cursor.execute("""
            INSERT INTO Visitantes (NombreCompleto, DNI, Correo, Institucion)
            VALUES (?, ?, ?, ?)
        """, (nombre, dni, correo, institucion))
        conn.commit()
        return jsonify({'status': 'success', 'msg': 'Visitante registrado correctamente.'})
    except Exception as e:
        return jsonify({'status': 'error', 'msg': str(e)})
    finally:
        conn.close()

# RUTA PARA ACTUALIZAR VISITANTE
@app.route('/admin/editar_visitante', methods=['POST'])
def editar_visitante():
    data = request.json
    id_vis = data.get('id')
    dni = data.get('dni')
    nombre = data.get('nombre')
    correo = data.get('correo')
    institucion = data.get('institucion')

    if not id_vis or not dni or not nombre:
        return jsonify({'status': 'error', 'msg': 'Faltan datos obligatorios'})

    # Normalizar institución
    if not institucion or institucion.strip() == "":
        institucion = "Sin Institución"

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE Visitantes
            SET NombreCompleto = ?, DNI = ?, Correo = ?, Institucion = ?
            WHERE VisitanteID = ?
        """, (nombre, dni, correo, institucion, id_vis))
        
        conn.commit()
        return jsonify({'status': 'success', 'msg': 'Datos actualizados correctamente.'})
    except Exception as e:
        return jsonify({'status': 'error', 'msg': str(e)})
    finally:
        conn.close()

# 2. Cargar Excel de Visitantes
@app.route('/admin/subir_excel_visitantes', methods=['POST'])
def subir_excel_visitantes():
    if 'archivo_excel_vis' not in request.files: return jsonify({'status': 'error', 'msg': 'Falta archivo'})
    file = request.files['archivo_excel_vis']
    if file.filename == '': return jsonify({'status': 'error', 'msg': 'Nombre vacío'})

    try:
        # Leemos Excel. Columnas esperadas: NOMBRE, DNI, CORREO, INSTITUCION
        df = pd.read_excel(file, dtype={'DNI': str})
        df = df.fillna('') # Rellena vacíos con texto vacío
        
        conn = get_db_connection()
        cursor = conn.cursor()
        insertados = 0
        
        for _, row in df.iterrows():
            dni = row.get('DNI', '')
            nombre = row.get('NOMBRE', '') # Asegúrate que el Excel tenga este encabezado
            correo = row.get('CORREO', '')
            inst = row.get('INSTITUCION', '')

            # Lógica automática: Si inst está vacío, es "Sin Institución"
            if not inst or str(inst).strip() == "":
                inst = "Sin Institución"

            if not dni or len(str(dni).strip()) == 0: continue
            
            # Insertar si no existe
            cursor.execute("IF NOT EXISTS (SELECT 1 FROM Visitantes WHERE DNI = ?) INSERT INTO Visitantes (NombreCompleto, DNI, Correo, Institucion) VALUES (?, ?, ?, ?)", 
                           (nombre, dni, correo, inst))
            if cursor.rowcount > 0: insertados += 1
            
        conn.commit()
        conn.close()
        return jsonify({'status': 'success', 'msg': f'Se registraron {insertados} visitantes nuevos.'})
    except Exception as e:
        return jsonify({'status': 'error', 'msg': str(e)})

# ==========================================
# RUTA PARA DESCARGAR PLANTILLAS VACÍAS (SOLO ENCABEZADOS)
# ==========================================
@app.route('/admin/plantilla/<tipo>')
def descargar_plantilla(tipo):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    
    if tipo == 'alumnos':
        # Definimos SOLO los nombres de las columnas
        if tipo == 'alumnos':
        # AHORA USAMOS LOS NOMBRES OFICIALES DE LA IMAGEN
            columnas = [
                'APELLIDOS Y NOMBRE', 
                'DNI', 
                'CODIGO DE MATRICULA', 
                'CORREO INSTITUCIONAL', # Lo agregamos para que sea igual, aunque no lo usemos
                'ESCUELA', 
                'FACULTAD', 
                'SEMESTRE'
            ]
            nombre_archivo = "Plantilla_Alumnos_Oficial.xlsx"
        
    elif tipo == 'visitantes':
        # Definimos SOLO los nombres de las columnas
        columnas = ['NOMBRE', 'DNI', 'CORREO', 'INSTITUCION']
        nombre_archivo = "Plantilla_Visitantes.xlsx"
    
    else:
        return "Tipo de plantilla no válido"

    # Creamos el DataFrame VACÍO, solo con los encabezados
    df = pd.DataFrame(columns=columnas)
    
    # Guardamos el Excel
    df.to_excel(writer, index=False)
    writer.close()
    output.seek(0)
    
    return send_file(output, 
                     download_name=nombre_archivo, 
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True, host='172.16.2.169', port=5000)