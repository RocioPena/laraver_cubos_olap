import pandas as pd
import win32com.client


def query_olap(connection_string, query):
    """ Ejecuta una consulta a OLAP y devuelve un DataFrame """
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


# 🔸 Conexión al cubo específico "Recursos"
cubo = 'SIS_2024'
connection_string = (
    "Provider=MSOLAP.8;"
    "Data Source=pwidgis03.salud.gob.mx;"
    "User ID=SALUD\\DGIS15;"
    "Password=Temp123!;"
    "Persist Security Info=True;"
    "Update Isolation Level=2;"
    f"Initial Catalog={cubo};"
)

# 🔸 Consulta de encabezados del cubo
query = "SELECT * FROM $SYSTEM.MDSCHEMA_MEASURES"

try:
    print(f"\nConsultando los encabezados del cubo '{cubo}'...")
    df_encabezados = query_olap(connection_string, query)
    
    print(f"Encabezados del cubo '{cubo}':")
    print(df_encabezados.to_string(index=False))

    # 🔸 Guardar como Excel
    df_encabezados.to_excel(f"consulta_cubos/encabezados_{cubo}.xlsx", index=False, engine="openpyxl")
    print(f"\n✔ Archivo guardado como encabezados_{cubo}.xlsx")

except Exception as e:
    print(f"❌ Error al consultar el cubo: {e}")
