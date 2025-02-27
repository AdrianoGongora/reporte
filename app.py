import pymssql
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from dotenv import load_dotenv
import os

# Cargar variables de entorno
load_dotenv()

# Parámetros de conexión
SERVER = os.getenv("DB_SERVER")
DATABASE = os.getenv("DB_DATABASE")
USER = os.getenv("DB_USER")
PASSWORD = os.getenv("DB_PASSWORD")

# Conectar a SQL Server
conn = pymssql.connect(server=SERVER, user=USER, password=PASSWORD, database=DATABASE)
cursor = conn.cursor()

# Consulta SQL
query = """
SELECT    
  CAST(DC.fecha_emision AS DATE) AS Fecha_Emision,    
  TD.descripcion AS Tipo_Documento,    
  COUNT(DC.iddocumentocontable) AS TOTAL,    
  SUM(CASE WHEN DC.idestadosunat = 8 THEN 1 ELSE 0 END) AS NODISPONIBLE,    
  SUM(CASE WHEN DC.idestadosunat = 4 THEN 1 ELSE 0 END) AS ENPROCESO_4,    
  SUM(CASE WHEN DC.idestadosunat = 6 THEN 1 ELSE 0 END) AS IGNORADA_6,    
  SUM(CASE WHEN DC.idestadosunat = 5 THEN 1 ELSE 0 END) AS RECHAZADA_5,    
  SUM(CASE WHEN DC.idestadosunat IN (2,3) THEN 1 ELSE 0 END) AS ACEPTADO 
FROM dbo.documento_contable DC        
INNER JOIN dbo.tipo_documento TD ON DC.idtipo_documento = TD.idtipo_documento
WHERE idempresa NOT IN (SELECT identidad FROM entidad_parametro WHERE idparametro = 3022)  
AND DC.idtipo_documento IN (1003, 1004, 1006, 1005, 1011, 1012, 3005, 3008)  
AND DC.idestado = 2  
AND DC.fecha_emision >= '2025-02-19'
GROUP BY CAST(DC.fecha_emision AS DATE), TD.descripcion
ORDER BY CAST(DC.fecha_emision AS DATE);
"""

# Ejecutar consulta y cargar datos en un DataFrame
df = pd.read_sql(query, conn)

# Cerrar conexión
conn.close()

# Guardar datos en Excel
file_name = "Reporte_SQL.xlsx"
df.to_excel(file_name, index=False, sheet_name="Datos")

# Cargar el archivo de Excel
wb = load_workbook(file_name)
ws = wb["Datos"]

# Lista de tipos de documentos únicos
tipos_documento = df["Tipo_Documento"].unique()

# Crear gráficos de barras apiladas por tipo de documento
image_paths = []
for tipo in tipos_documento:
    df_tipo = df[df["Tipo_Documento"] == tipo]  # Filtrar datos por tipo de documento

    estados = ["NODISPONIBLE", "ENPROCESO_4", "IGNORADA_6", "RECHAZADA_5", "ACEPTADO"]

    # Crear gráfico de barras apiladas
    plt.figure(figsize=(10, 5))
    bottom_values = [0] * len(df_tipo)  # Lista para acumular valores en las barras apiladas

    for estado in estados:
        plt.bar(df_tipo["Fecha_Emision"], df_tipo[estado], label=estado, bottom=bottom_values)
        bottom_values = [x + y for x, y in zip(bottom_values, df_tipo[estado])]  # Acumular valores

    plt.xlabel("Fecha de Emisión")
    plt.ylabel("Cantidad")
    plt.title(f"Estados de {tipo}")
    plt.xticks(rotation=45)
    plt.legend()
    plt.grid(axis="y")

    # Guardar gráfico como imagen
    image_path = f"grafico_{tipo}.png"
    plt.savefig(image_path, bbox_inches="tight")
    image_paths.append(image_path)
    plt.close()  # Cerrar la figura para evitar sobreposición

# Insertar imágenes en el archivo Excel
columna_imagen = "J"
fila_imagen = 2
for image_path in image_paths:
    img = Image(image_path)
    ws.add_image(img, f"{columna_imagen}{fila_imagen}")
    fila_imagen += 20  # Espacio para el siguiente gráfico

# Guardar cambios en el Excel
wb.save(file_name)

print(f"Reporte generado exitosamente: {file_name}")
