import pandas as pd
import pyodbc
import os
from dotenv import load_dotenv

load_dotenv()

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
        print(f"Error conectando a SQL: {e}")
        return None

# --- FUNCIONES HELPER PARA OBTENER IDs ---
def obtener_id_facultad(cursor, nombre_facultad):
    if not nombre_facultad: return None
    cursor.execute("SELECT FacultadID FROM Facultades WHERE NombreFacultad = ?", (nombre_facultad,))
    row = cursor.fetchone()
    if row: return row[0]
    # Si no existe, crear
    cursor.execute("INSERT INTO Facultades (NombreFacultad) VALUES (?)", (nombre_facultad,))
    cursor.execute("SELECT SCOPE_IDENTITY()")
    return cursor.fetchone()[0]

def obtener_id_escuela(cursor, nombre_escuela, facultad_id):
    if not nombre_escuela: return None
    cursor.execute("SELECT EscuelaID FROM Escuelas WHERE NombreEscuela = ?", (nombre_escuela,))
    row = cursor.fetchone()
    if row: return row[0]
    # Si no existe, crear
    cursor.execute("INSERT INTO Escuelas (NombreEscuela, FacultadID) VALUES (?, ?)", (nombre_escuela, facultad_id))
    cursor.execute("SELECT SCOPE_IDENTITY()")
    return cursor.fetchone()[0]

def obtener_id_semestre(cursor, nombre_semestre):
    if not nombre_semestre: return None
    # Normalizar semestre (evitar duplicados por espacios)
    nombre_semestre = str(nombre_semestre).strip().upper()
    cursor.execute("SELECT SemestreID FROM Semestres WHERE NombreSemestre = ?", (nombre_semestre,))
    row = cursor.fetchone()
    if row: return row[0]
    # Si no existe, crear
    cursor.execute("INSERT INTO Semestres (NombreSemestre) VALUES (?)", (nombre_semestre,))
    cursor.execute("SELECT SCOPE_IDENTITY()")
    return cursor.fetchone()[0]

def cargar_excel():
    archivo = 'lista_alumnos.xlsx'
    if not os.path.exists(archivo):
        print(f"❌ No encuentro '{archivo}'")
        return

    print("📂 Leyendo Excel...")
    try:
        df = pd.read_excel(archivo, dtype={'DNI': str, 'CODIGO DE MATRICULA': str, 'SEMESTRE': str})
        df = df.fillna('')
    except Exception as e:
        print(f"❌ Error Excel: {e}")
        return

    conn = get_db_connection()
    if not conn: return
    cursor = conn.cursor()
    
    total = 0
    print(f"🚀 Procesando {len(df)} alumnos...")

    for _, row in df.iterrows():
        try:
            nombre = row.get('APELLIDOS Y NOMBRE', '')
            dni = str(row.get('DNI', '')).strip()
            codigo = str(row.get('CODIGO DE MATRICULA', '')).strip()
            escuela_txt = row.get('ESCUELA', '')
            facultad_txt = row.get('FACULTAD', '')
            semestre_txt = row.get('SEMESTRE', '')
            
            # --- LA MAGIA: OBTENER IDs EN LUGAR DE TEXTO ---
            fac_id = obtener_id_facultad(cursor, facultad_txt)
            esc_id = obtener_id_escuela(cursor, escuela_txt, fac_id)
            sem_id = obtener_id_semestre(cursor, semestre_txt)

            if not dni: continue

            # Insertar o Actualizar usando IDs
            # Verificamos si existe el DNI
            cursor.execute("SELECT AlumnoID FROM Alumnos WHERE DNI = ?", (dni,))
            existe = cursor.fetchone()

            if existe:
                # Actualizar referencias
                cursor.execute("""
                    UPDATE Alumnos 
                    SET NombreCompleto=?, EscuelaID=?, SemestreID=?
                    WHERE DNI=?
                """, (nombre, esc_id, sem_id, dni))
            else:
                # Insertar nuevo con IDs
                cursor.execute("""
                    INSERT INTO Alumnos (NombreCompleto, DNI, CodigoMatricula, EscuelaID, SemestreID, Estado)
                    VALUES (?, ?, ?, ?, ?, 1)
                """, (nombre, dni, codigo, esc_id, sem_id))
            
            total += 1
            if total % 50 == 0: print(f"   ✅ {total} procesados...")

        except Exception as e:
            print(f"⚠️ Error en DNI {dni}: {e}")

    conn.commit()
    conn.close()
    print("🏁 Carga completada.")

if __name__ == '__main__':
    cargar_excel()