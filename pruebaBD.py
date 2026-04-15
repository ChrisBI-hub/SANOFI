# Corre esto por separado para ver los nombres reales
import pyodbc
import pandas as pd

server   = '150.1.1.152'
database = 'SIR'
username = 'ConsultaBD'
password = '5D$bc#kM&5W2T8J40?s%'

conn = pyodbc.connect(
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={server};DATABASE={database};"
    f"UID={username};PWD={password};TrustServerCertificate=yes;"
)

with open("SANOFI_V5.sql", encoding="latin-1") as f:
    query = f.read()

df = pd.read_sql(query, conn)
conn.close()

# Imprimir columnas numeradas para identificarlas fácil
for i, col in enumerate(df.columns):
    print(f"  {i:02d}. '{col}'")