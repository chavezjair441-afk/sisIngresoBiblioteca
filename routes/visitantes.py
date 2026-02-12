from flask import Blueprint, render_template, request, jsonify, send_file
import pandas as pd
import io
from db import get_db_connection

visitantes_bp = Blueprint('visitantes', __name__, url_prefix='/admin')

@visitantes_bp.route('/visitantes')
def admin_visitantes_page():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT VisitanteID, NombreCompleto, DNI, Institucion, Correo FROM Visitantes ORDER BY NombreCompleto ASC")
    lista = cursor.fetchall()
    conn.close()
    return render_template('admin_visitantes.html', visitantes=lista)

@visitantes_bp.route('/agregar_visitante', methods=['POST'])
def agregar_visitante():
    data = request.json
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        inst = data.get('institucion') or "Sin Institución"
        
        cursor.execute("SELECT 1 FROM Visitantes WHERE DNI = ?", (data.get('dni'),))
        if cursor.fetchone(): return jsonify({'status': 'error', 'msg': 'DNI ya existe'})

        cursor.execute("INSERT INTO Visitantes (NombreCompleto, DNI, Correo, Institucion) VALUES (?,?,?,?)",
                       (data.get('nombre'), data.get('dni'), data.get('correo'), inst))
        conn.commit()
        return jsonify({'status': 'success', 'msg': 'Guardado'})
    except Exception as e:
        return jsonify({'status': 'error', 'msg': str(e)})
    finally:
        conn.close()

# Aquí puedes añadir las rutas 'editar_visitante' y 'subir_excel_visitantes' siguiendo la misma lógica