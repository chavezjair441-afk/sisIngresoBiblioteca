from flask import Blueprint, render_template, request, jsonify
from db import get_db_connection

# Definimos el Blueprint
ingreso_bp = Blueprint('ingreso', __name__)

@ingreso_bp.route('/')
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

@ingreso_bp.route('/piso1')
def piso1(): return render_template('ingreso.html', piso=1)

@ingreso_bp.route('/piso2')
def piso2(): return render_template('ingreso.html', piso=2)

@ingreso_bp.route('/piso3')
def piso3(): return render_template('ingreso.html', piso=3)

# API: PROCESAR EL ESCANEO
@ingreso_bp.route('/procesar_ingreso', methods=['POST'])
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
        DECLARE @out_semestre varchar(20);
        
        -- Ejecutamos el procedimiento y capturamos los datos de salida
        EXEC sp_RegistrarIngreso ?, ?, @out_msg OUTPUT, @out_nombre OUTPUT, @out_escuela OUTPUT, @out_semestre OUTPUT;
        
        SELECT @out_msg, @out_nombre, @out_escuela, @out_semestre;
        """
        cursor.execute(sql, (codigo, piso))
        row = cursor.fetchone()
        conn.commit()
        conn.close()

        if row:
            mensaje = row[0]
            # Extraemos los datos siempre (aunque venga mensaje de error de turno)
            nombre = row[1]
            escuela = row[2]
            semestre = row[3]

            # CASO 1: ÉXITO (Entra por primera vez en el turno)
            if 'CONCEDIDO' in mensaje: 
                return jsonify({
                    'status': 'success', 
                    'msg': mensaje, 
                    'alumno': nombre, 
                    'escuela': escuela, 
                    'semestre': semestre
                })
            
            # CASO 2: YA REGISTRADO (Advertencia)
            # AQUÍ ESTÁ EL TRUCO: Enviamos 'warning' pero TAMBIÉN los datos del alumno
            elif 'YA REGISTRADO' in mensaje:
                return jsonify({
                    'status': 'warning', 
                    'msg': mensaje, 
                    'alumno': nombre, 
                    'escuela': escuela, 
                    'semestre': semestre
                })

            # CASO 3: NO ENCONTRADO O ERROR REAL
            else:
                return jsonify({'status': 'error', 'msg': mensaje})
        
        return jsonify({'status': 'error', 'msg': 'Error desconocido en BD'})

    except Exception as e:
        return jsonify({'status': 'error', 'msg': str(e)})