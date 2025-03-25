import pyodbc
import pandas as pd
from datetime import datetime

def query_to_df(cursor) -> pd.DataFrame:
    columns = [column[0] for column in cursor.description]
    data = cursor.fetchall()
    return pd.DataFrame(data, columns=columns)

# Cadena de conexión (usa tu DSN o conexión directa si configuraste MSOLAP como ODBC)
conn_str_base = (
    "Driver={MSOLAP};"
    "Password=Temp123!;"
    "Persist Security Info=True;"
    "User ID=SALUD\\DGIS15;"
    "Data Source=pwidgis03.salud.gob.mx;"
    "Update Isolation Level=2;"
)

try:
    # Conexión inicial sin catálogo
    conn = pyodbc.connect(conn_str_base, timeout=600)
    cursor = conn.cursor()
    
    cursor.execute("SELECT [catalog_name] FROM $system.DBSCHEMA_CATALOGS")
    catalogos_df = query_to_df(cursor)
    print("Catálogos disponibles:")
    print(catalogos_df.to_string())
    conn.close()

    # Obtener cubos de cada catálogo
    cubos_totales = []
    for catalogo in catalogos_df.iloc[:, 0]:
        try:
            print(f"Consultando cubos en catálogo: {catalogo}")
            conn_str = conn_str_base + f"Initial Catalog={catalogo};"
            conn = pyodbc.connect(conn_str, timeout=600)
            cursor = conn.cursor()

            cursor.execute("SELECT CUBE_NAME FROM $system.MDSCHEMA_CUBES")
            df_cubos = query_to_df(cursor)
            df_cubos['DATABASE'] = catalogo
            cubos_totales.append(df_cubos)
            print(df_cubos.to_string())
            conn.close()
        except Exception as ce:
            print(f"Error al consultar cubos del catálogo {catalogo}: {ce}")
            if conn:
                conn.close()

    if cubos_totales:
        final_df = pd.concat(cubos_totales, ignore_index=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        final_df.to_csv(f'cubos_olap_{timestamp}.csv', index=False, encoding='utf-8')
        print("Todos los cubos han sido guardados en CSV.")

except Exception as e:
    print(f"Error general: {e}")
