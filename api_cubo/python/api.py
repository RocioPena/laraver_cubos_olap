from fastapi import FastAPI
from fastapi.responses import JSONResponse
import pandas as pd
import win32com.client
import pythoncom

app = FastAPI()

# Configuración global del cubo y conexión
CUBO = 'SIS_2024'
CONNECTION_STRING = (
    "Provider=MSOLAP.8;"
    "Data Source=pwidgis03.salud.gob.mx;"
    "User ID=SALUD\\DGIS15;"
    "Password=Temp123!;"
    "Persist Security Info=True;"
    "Update Isolation Level=2;"
    f"Initial Catalog={CUBO};"
)


def query_olap(connection_string: str, query: str) -> pd.DataFrame:
    # Iniciar COM en este hilo
    pythoncom.CoInitialize()
    
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

    # Finalizar COM
    pythoncom.CoUninitialize()

    return pd.DataFrame(data, columns=fields)
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


@app.get("/medidas")
def get_measures():
    """Consulta los encabezados del cubo (MEASURES) y los devuelve en JSON"""
    try:
        query = "SELECT * FROM $SYSTEM.MDSCHEMA_MEASURES"
        df = query_olap(CONNECTION_STRING, query)
        data = df.to_dict(orient="records")
        return JSONResponse(content=data)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


@app.get("/exportar_excel")
def export_measures_excel():
    """Consulta las medidas y guarda archivo Excel en el servidor"""
    try:
        query = "SELECT * FROM $SYSTEM.MDSCHEMA_MEASURES"
        df = query_olap(CONNECTION_STRING, query)
        path = f"consulta_cubos/encabezados_{CUBO}.xlsx"
        df.to_excel(path, index=False, engine="openpyxl")
        return {"status": "ok", "archivo": path}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
