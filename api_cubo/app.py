from filtrado_post import ejecutar_consulta, ConsultaDinamica
import pandas as pd
from datetime import datetime
import os

# Crear carpeta si no existe
carpeta = "resultado_consulta"
os.makedirs(carpeta, exist_ok=True)

# 👉 Configurar el modo de encabezado: "nombre" o "codigo+nombre"
modo_encabezado = "nombre"  # cambia aquí según prefieras

# Crear el payload
payload = ConsultaDinamica(
    variables_clave=["VBC01", "VBC02", "VBC03", "VBC51", "VBC52", "VBC53", "VBC54", "VBC55"],
    unidades=["[DIM UNIDAD].[CLUES].[HGIMB002403]"],
    modo_encabezado=modo_encabezado
)

# Ejecutar consulta
resultado = ejecutar_consulta(payload)

# Convertir a DataFrame
df = pd.DataFrame(resultado)

# Renombrar columna CLUES (opcional)
for col in df.columns:
    if "CLUES" in col and "MEMBER_CAPTION" in col:
        df.rename(columns={col: "CLUES"}, inplace=True)

# Guardar con timestamp
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
excel_file = os.path.join(carpeta, f"resultado_consulta_{timestamp}.xlsx")
csv_file = os.path.join(carpeta, f"resultado_consulta_{timestamp}.csv")

df.to_excel(excel_file, index=False)
df.to_csv(csv_file, index=False)

print(f"✅ Consulta exportada exitosamente:")
print(f"  ➤ Excel: {excel_file}")
print(f"  ➤ CSV:   {csv_file}")