import pandas as pd
import win32com.client


def query_olap(connection_string, query):
    conn = win32com.client.Dispatch("ADODB.Connection")
    rs = win32com.client.Dispatch("ADODB.Recordset")

    conn.Open(connection_string)
    rs.Open(query, conn)

    fields = [rs.Fields.Item(i).Name for i in range(rs.Fields.Count)]
    data = []

    while not rs.EOF:
        row = [rs.Fields.Item(i).Value for i in range(rs.Fields.Count)]
        data.append(row)
        rs.MoveNext()

    rs.Close()
    conn.Close()

    return pd.DataFrame(data, columns=fields)


# Cadena de conexión
connection_string = (
    "Provider=MSOLAP.8;"
    "Data Source=pwidgis03.salud.gob.mx;"
    "User ID=SALUD\\DGIS15;"
    "Password=Temp123!;"
    "Persist Security Info=True;"
    "Initial Catalog=SIS_2024;"  # Pruébalo con una base específica si sabes una
    "Connect Timeout=60;"
)

query = "SELECT [CATALOG_NAME] FROM $system.DBSCHEMA_CATALOGS"

try:
    df_catalogos = query_olap(connection_string, query)
    print("Catálogos disponibles:")
    print(df_catalogos.to_string(index=False))

    cubos_totales = []

    for catalogo in df_catalogos.iloc[:, 0]:
        print(f"\nConsultando cubos en catálogo: {catalogo}")
        try:
            conn_cat = connection_string + f"Initial Catalog={catalogo};"
            df_cubos = query_olap(conn_cat, "SELECT CUBE_NAME FROM $system.MDSCHEMA_CUBES")

            if df_cubos.empty:
                print(f"Sin cubos en {catalogo}")
            else:
                df_cubos['DATABASE'] = catalogo
                cubos_totales.append(df_cubos)
                print(df_cubos.to_string(index=False))

        except Exception as inner_error:
            print(f"⚠ Error al consultar cubos en {catalogo}: {inner_error}")
            continue  # seguir con el siguiente catálogo

    if cubos_totales:
        cubos_df = pd.concat(cubos_totales, ignore_index=True)
        
        # Guardar como CSV
        cubos_df.to_csv("cubos_olap_win32com.csv", index=False, encoding="utf-8")
        print("✔ Datos guardados en cubos_olap_win32com.csv")
        
        # Guardar como Excel
        cubos_df.to_excel("cubos_olap_win32com.xlsx", index=False, engine="openpyxl")
        print("✔ Datos también guardados en cubos_olap_win32com.xlsx")

except Exception as e:
    print(f"❌ Error general: {e}")
