# ---------- main.py (API FastAPI con encabezado flexible) ----------
from fastapi import FastAPI
from fastapi.responses import JSONResponse
from pydantic import BaseModel
import pandas as pd
import win32com.client
import pythoncom
import math

app = FastAPI()

CUBO = "SIS_2024"
CONNECTION_STRING = (
    "Provider=MSOLAP.8;"
    "Data Source=pwidgis03.salud.gob.mx;"
    "User ID=SALUD\\DGIS15;"
    "Password=Temp123!;"
    "Persist Security Info=True;"
    f"Initial Catalog={CUBO};"
    "Connect Timeout=60;"
)

variables_df = pd.read_excel("consulta_cubos/dim_variables.xlsx")
clave_to_variable = dict(
    zip(
        variables_df['[$DIM VARIABLES].[CLAVE]'],
        variables_df['[$DIM VARIABLES].[Variable]']
    )
)

class ConsultaDinamica(BaseModel):
    variables: list[str] | None = None
    variables_clave: list[str] | None = None
    unidades: list[str]
    fechas: list[str] | None = None
    filtros_where: list[str] | None = None
    entidad_filtro: str | None = None
    modo_encabezado: str | None = "nombre"  # nuevo: "nombre" o "codigo+nombre"

def query_olap(connection_string: str, query: str) -> pd.DataFrame:
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
    pythoncom.CoUninitialize()
    return pd.DataFrame(data, columns=fields)

def sanitize_result(data):
    if isinstance(data, float) and (math.isnan(data) or data == float("inf") or data == float("-inf")):
        return None
    elif isinstance(data, list):
        return [sanitize_result(x) for x in data]
    elif isinstance(data, dict):
        return {k: sanitize_result(v) for k, v in data.items()}
    return data

def ejecutar_consulta(payload: ConsultaDinamica) -> list[dict]:
    if not payload.variables and payload.variables_clave:
        payload.variables = [
            f"[DIM VARIABLES].[Variable].[{clave_to_variable[clave]}]"
            for clave in payload.variables_clave if clave in clave_to_variable
        ]

    if not payload.variables:
        raise ValueError("Debes proporcionar 'variables' o 'variables_clave' válidas.")

    if payload.entidad_filtro:
        filtro_mdx = (
            f"FILTER(\n"
            f"  {{ [DIM UNIDAD].[Entidad Municipio Localidad].[Entidad Municipio Localidad].MEMBERS }},\n"
            f"  [DIM UNIDAD].[Entidad Municipio Localidad].CurrentMember.Properties(\"Entidad\") = \"{payload.entidad_filtro}\"\n"
            f")"
        )
        mdx = (
            "SELECT "
            f"{{ {', '.join(payload.variables)} }} ON COLUMNS,\n"
            f"{{ {filtro_mdx} }} ON ROWS\n"
            f"FROM [{CUBO}]"
        )
    else:
        mdx = (
            "SELECT "
            f"{{ {', '.join(payload.variables)} }} ON COLUMNS,\n"
            f"{{ {', '.join(payload.unidades)} }} ON ROWS\n"
            f"FROM [{CUBO}]"
        )

    if payload.fechas or payload.filtros_where:
        filtros = []
        if payload.fechas:
            filtros.extend(payload.fechas)
        if payload.filtros_where:
            filtros.extend(payload.filtros_where)
        mdx += f"\nWHERE ( {', '.join(filtros)} )"

    df = query_olap(CONNECTION_STRING, mdx)

    renamed_columns = {}
    for col in df.columns:
        if "Variable].&[VBC" in col:
            clave = col.split("&[")[-1].replace("]", "")
            nombre = clave_to_variable.get(clave, clave)
            if payload.modo_encabezado == "codigo+nombre":
                renamed_columns[col] = f"{clave} - {nombre}"
            else:
                renamed_columns[col] = nombre
    df.rename(columns=renamed_columns, inplace=True)

    return sanitize_result(df.to_dict(orient="records"))

@app.post("/consulta_avanzada")
def consulta_dinamica(payload: ConsultaDinamica):
    try:
        result = ejecutar_consulta(payload)
        return JSONResponse(content=result)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
