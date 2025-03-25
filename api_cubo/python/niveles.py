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
    "Initial Catalog=SIS_2024;"
    "Connect Timeout=60;"
)

query = "SELECT * FROM $SYSTEM.MDSCHEMA_LEVELS"

try:
    df_levels = query_olap(connection_string, query)

    print("\nNiveles disponibles en el cubo SIS_2024:")
    print(df_levels.to_string(index=False))

    # Guardar resultados
    df_levels.to_csv("niveles_olap.csv", index=False, encoding="utf-8")
    df_levels.to_excel("niveles_olap.xlsx", index=False, engine="openpyxl")
    print("\n✔ Resultados guardados en niveles_olap.csv y niveles_olap.xlsx")

except Exception as e:
    print(f"\n❌ Error al consultar niveles: {e}")
