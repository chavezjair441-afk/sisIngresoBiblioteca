import pandas as pd
import pyodbc
import os
from dotenv import load_dotenv

# 1. Cargar configuración
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

def cargar_excel():
    archivo = 'lista_alumnos.xlsx'
    
    # Verificamos si existe el archivo
    if not os.path.exists(archivo):
        print(f"❌ No encuentro el archivo '{archivo}'. Asegúrate de ponerlo en la misma carpeta.")
        return

    print("📂 Leyendo archivo Excel... esto puede tardar unos segundos...")
    
    # LEEMOS EL EXCEL
    # dtype={'DNI': str} es VITAL para que no borre los ceros a la izquierda (ej: 060...)
    try:
        df = pd.read_excel(archivo, dtype={'DNI': str, 'CODIGO DE MATRICULA': str, 'SEMESTRE': str})
    except Exception as e:
        print(f"❌ Error leyendo el Excel: {e}")
        return

    # Limpieza de datos básica (rellenar vacíos con texto vacío para no dar error en SQL)
    df = df.fillna('')

    conn = get_db_connection()
    if not conn:
        return

    cursor = conn.cursor()
    total_insertados = 0
    total_errores = 0
    
    print(f"🚀 Iniciando carga de {len(df)} alumnos...")

    # RECORREMOS CADA FILA DEL EXCEL
    for index, row in df.iterrows():
        try:
            # Mapeamos las columnas del Excel (Izquierda) a variables
            nombre = row['APELLIDOS Y NOMBRE']
            dni = row['DNI']
            codigo = row['CODIGO DE MATRICULA']
            correo_inst = row['CORREO INSTITUCIONAL']
            correo_pers = row['CORREO PERSONAL']
            escuela = row['ESCUELA']
            facultad = row['FACULTAD']
            semestre = row['SEMESTRE']

            # Validar que tenga DNI (si la fila está vacía, la saltamos)
            if not dni or len(str(dni).strip()) == 0:
                continue

            # Query de Inserción
            # Usamos una validación simple en SQL para no insertar si ya existe el DNI
            sql = """
            IF NOT EXISTS (SELECT 1 FROM Alumnos WHERE DNI = ?)
            BEGIN
                INSERT INTO Alumnos (NombreCompleto, DNI, CodigoMatricula, CorreoInstitucional, CorreoPersonal, Escuela, Facultad, Semestre)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            END
            """
            
            cursor.execute(sql, (dni, nombre, dni, codigo, correo_inst, correo_pers, escuela, facultad, semestre))
            
            # Si se insertó una fila, aumentamos contador (cursor.rowcount nos dice cuántas filas afectó)
            if cursor.rowcount > 0:
                total_insertados += 1
                # Imprimir progreso cada 50 alumnos
                if total_insertados % 50 == 0:
                    print(f"   ✅ Van {total_insertados} alumnos...")
            
        except Exception as e:
            print(f"⚠️ Error en fila {index + 2} (DNI: {dni}): {e}")
            total_errores += 1

    conn.commit()
    conn.close()

    print("\n" + "="*40)
    print(f"🏁 PROCESO TERMINADO")
    print(f"✅ Insertados correctamente: {total_insertados}")
    print(f"⚠️ Duplicados o Errores:    {total_errores}")
    print("="*40)

if __name__ == '__main__':
    cargar_excel()