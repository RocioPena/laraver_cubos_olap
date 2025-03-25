import pandas as pd
import win32com.client

def query_olap(connection_string, query):
    """ Ejecuta una consulta OLAP y devuelve un DataFrame """
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

# Cadena de conexión al cubo SIS_2024
connection_string = (
    "Provider=MSOLAP.8;"
    "Data Source=pwidgis03.salud.gob.mx;"
    "User ID=SALUD\\DGIS15;"
    "Password=Temp123!;"
    "Persist Security Info=True;"
    "Initial Catalog=SIS_2024;"
    "Connect Timeout=60;"
)

# Consulta de jerarquías
query = "SELECT * FROM $SYSTEM.MDSCHEMA_HIERARCHIES"

try:
    df_hierarchies = query_olap(connection_string, query)
    print("✅ Jerarquías del cubo SIS_2024:")
    print(df_hierarchies.to_string(index=False))

    # Guardar en Excel y CSV
    df_hierarchies.to_csv("hierarchies_olap.csv", index=False, encoding="utf-8")
    df_hierarchies.to_excel("hierarchies_olap.xlsx", index=False, engine="openpyxl")
    print("✔ Resultados guardados en hierarchies_olap.csv y hierarchies_olap.xlsx")

except Exception as e:
    print(f"❌ Error durante la consulta: {e}")
