import pandas as pd
import win32com.client

def query_sql(connection_string, query):
    conn = win32com.client.Dispatch("ADODB.Connection")
    rs = win32com.client.Dispatch("ADODB.Recordset")

    conn.Open(connection_string)
    rs.Open(query, conn)

    columns = [rs.Fields.Item(i).Name for i in range(rs.Fields.Count)]
    data = []

    while not rs.EOF:
        row = [rs.Fields.Item(i).Value for i in range(rs.Fields.Count)]
        data.append(row)
        rs.MoveNext()

    rs.Close()
    conn.Close()

    return pd.DataFrame(data, columns=columns)

# Configura tu conexión

cubo = 'SIS_2024'
conn_str = (
    "Provider=MSOLAP.8;"
    "Data Source=pwidgis03.salud.gob.mx;"
    "User ID=SALUD\\DGIS15;"
    "Password=Temp123!;"
    "Persist Security Info=True;"
    "Update Isolation Level=2;"
    f"Initial Catalog={cubo};"
)
# conn_str = (
#     "Provider=SQLOLEDB;"
#     "Data Source=TU_SERVIDOR;"  # ← Cambia esto por el nombre real del servidor
#     "Initial Catalog=SIS_2024;"
#     "Integrated Security=SSPI;"
# )

# busca lo que hay en unidades
# query = "SELECT * FROM [SIS_2024].[$DIM UNIDAD]"

# busca lo que hay en unidades pero en hidalgo
# query = """
# SELECT * 
# FROM $SYSTEM.MDSCHEMA_MEMBERS 
# WHERE [DIMENSION_UNIQUE_NAME] = '[DIM UNIDAD]' AND [MEMBER_NAME] = 'HIDALGO'
# """


# busca lo que hay en VARIABLES
query = "SELECT * FROM [SIS_2024].[$DIM TIEMPO]"


try:
    df = query_sql(conn_str, query)

    # Mostrar las primeras filas
    print(df.head())

    # Guardar como Excel
    df.to_excel("consulta_cubos/dim_tiempo.xlsx", index=False, engine="openpyxl")
    print("\n✔ Datos guardados en dim_unidad.xlsx")

except Exception as e:
    print(f"❌ Error durante la consulta: {e}")
