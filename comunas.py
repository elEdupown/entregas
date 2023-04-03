import pandas as pd
import openpyxl

# Hacer dataframe de todos los archivos y obtener una lista con todas las comunas de la columna 'COMUNA'
df = pd.concat([pd.read_excel('postSDmatriz.xlsx', sheet_name='transito'), 
pd.read_excel('ott.xlsx', sheet_name='transito'),
pd.read_excel('ecommerceSDmatriz.xlsx', sheet_name='transito'),
pd.read_excel('ecommerceEDmatriz.xlsx', sheet_name='transito'),
pd.read_excel('postEDmatriz.xlsx', sheet_name='transito'),
pd.read_excel('teleSDmatriz.xlsx', sheet_name='transito'),
pd.read_excel('teleEDmatriz.xlsx', sheet_name='transito'),
pd.read_excel('business.xlsx', sheet_name='transito'),
pd.read_excel('crm.xlsx', sheet_name='transito'),])

regionesYcomunas = df[['REGION', 'COMUNA']]
# Generar dataframe con todas las regiones y comunas, sin repetir duplas
df = pd.DataFrame(regionesYcomunas, columns=['REGION', 'COMUNA'])
df = df.drop_duplicates()
df.to_excel('regionesYcomunas.xlsx', sheet_name='regionesYcomunas', index=False)

regiones = df['REGION'].unique().tolist()
# Crear dataframe con todas las regiones
df_regiones = pd.DataFrame(regiones, columns=['REGION'])

df_regiones.to_excel('regiones.xlsx', sheet_name='regiones', index=False)

# Crear diferentes dataframes con las comunas por REGION, dejando solo la columna 'COMUNA'
for region in regiones:
    df_reg = df[df['REGION'] == region]
    df_reg = df_reg[['COMUNA']]
    df_reg = df_reg.drop_duplicates()
    df_reg = df_reg.sort_values(by=['COMUNA'])
    df_reg.to_excel('comunas/{}.xlsx'.format(region), sheet_name='comunas', index=False)