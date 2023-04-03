import tkinter as tk
import pandas as pd
from datetime import datetime, timedelta
from collections import deque

def obtener_tipo_pedido():
    """Obtiene el tipo de pedido seleccionado por el usuario."""
    tipos_validos = ["Postventa", "Televenta", "Ecommerce", "Business", "CRM", "OTT"]
    while True:
        tipo_pedido = input(f"Seleccione el tipo de pedido ({', '.join(tipos_validos)}): ")
        if tipo_pedido in tipos_validos:
            return tipo_pedido
        else:
            print("Tipo de pedido no válido. Intente nuevamente.")

def obtener_region_y_comuna(regiones, comunas, excel_file):
    """Obtiene la región y la comuna seleccionadas por el usuario."""
    while True:
        region = input(f"Seleccione la región ({', '.join(regiones)}): ")
        if region in regiones:
            break
        else:
            print("Región no válida. Intente nuevamente.")
    
    while True:
        df = pd.read_excel(excel_file, sheet_name='transito')
        comunas = df[df['REGION'] == region]['COMUNA'].unique().tolist()
        comuna = input(f"Seleccione la comuna ({', '.join(comunas)}): ")
        if comuna in comunas:
            return region, comuna
        else:
            print("Comuna no válida. Intente nuevamente.")

def obtener_tipo_entrega():
    """Obtiene el tipo de entrega seleccionado por el usuario."""
    while True:
        tipo_entrega = input("Seleccione el tipo de entrega (Express, Standard): ")
        if tipo_entrega in ["Express", "Standard"]:
            return tipo_entrega
        else:
            print("Tipo de entrega no válido. Intente nuevamente.")

def calcular_tiempo(region, comuna, excel_file, etiqueta):
    # Obtener la fecha actual
    fecha_actual = datetime.now()

    # Cargar el archivo de Excel y obtener la hoja correspondiente
    df = pd.read_excel(excel_file, sheet_name="transito")
    # Obtener filas de la hoja que tengan la etiqueta ingresada y la región ingresada
    sheet = df[(df['ETIQUETA'] == etiqueta) & (df['REGION'] == region)]

    # obtener el tiempo de espera de la hoja
    tiempo_espera_str = sheet[sheet['COMUNA'] == comuna]['TIEMPO_ESPERA'].values[0]
    # obtener los días hábiles de la hoja
    dias_habiles = sheet[sheet['COMUNA'] == comuna]['DIAS_HABILES'].values[0]
    # parsear los días hábiles a una lista enlazada de enteros
    dias_habiles = deque(int(dia) for dia in dias_habiles.split(","))

    # Convertir el tiempo de espera a dias
    if tiempo_espera_str.endswith("d"):
        tiempo_espera = int(tiempo_espera_str[:-2])
    elif tiempo_espera_str.endswith("h"):
        tiempo_espera = int(tiempo_espera_str[:-2]) / 24
    else:
        raise ValueError("Formato de tiempo de espera inválido")

    # Obtener el número de día de fecha_actual
    dia_actual = fecha_actual.weekday()

    # Obtener la hora actual de la fecha_actual
    hora_actual = fecha_actual.hour

    # Verificar si la hora actual es mayor a las 12:00 p.m. (mediodía)
    if hora_actual >= 12:
        # Obtener el siguiente día hábil si el día actual no está en días hábiles o si es un domingo
        while True:
            fecha_actual += timedelta(days=1)
            dia_actual = fecha_actual.weekday()
            if dia_actual in dias_habiles and dia_actual != 6:
                break

    # Verificar si el tiempo de espera es 1 día y si el siguiente día hábil está en días hábiles
    if tiempo_espera == 1 and (dia_actual+1) % 7 in dias_habiles:
        fecha_entrega = fecha_actual + timedelta(days=1)
    else:
        # Establecer la fecha_entrega como la fecha_actual más el tiempo de espera en días
        for i in range(tiempo_espera):
            fecha_entrega = fecha_actual + timedelta(days=i+1)
            if fecha_entrega.weekday() == 6:
                tiempo_espera += 1
                fecha_entrega += timedelta(days=1)

            elif fecha_entrega.weekday() == 5 and hora_actual >= 12:
                tiempo_espera += 2
                fecha_entrega += timedelta(days=2)

    # Obtener la fecha de entrega en formato de string
    fecha_entrega_str = fecha_entrega.strftime("%d/%m/%Y")

    # Devolver la fecha de entrega estimada
    return fecha_entrega_str





def main():

    excel_files = {
        ("Postventa", "Express"): "postEDmatriz.xlsx",
        ("Postventa", "Standard"): "postSDmatriz.xlsx",
        ("Televenta", "Express"): "teleEDmatriz.xlsx",
        ("Televenta", "Standard"): "teleSDmatriz.xlsx",
        ("Ecommerce", "Express"): "ecommerceEDmatriz.xlsx",
        ("Ecommerce", "Standard"): "ecommerceSDmatriz.xlsx",
        ("Business", "Express"): "business.xlsx",
        ("CRM", "Express"): "crm.xlsx",
        ("OTT", "Express"): "ott.xlsx"
    }



    # Pedir al usuario que seleccione el tipo de pedido
    tipo_pedido = obtener_tipo_pedido()
    
    
    
    # Pedir al usuario que seleccione el tipo de entrega (solo para algunos tipos de pedido)
    tipo_entrega = None
    if tipo_pedido in ["Televenta", "Postventa", "Ecommerce"]:
        tipo_entrega = obtener_tipo_entrega()
    elif tipo_pedido in ["Business", "CRM", "OTT"]:
        tipo_entrega = "Express"

    excel_file = excel_files.get((tipo_pedido, tipo_entrega))

    print(excel_file)

    etiqueta = input("Ingrese la etiqueta (EquipoSIM, SIM): ")

    # Ir al excel y obtener en una lista las regiones de la columna REGION (sin repetir) que tengan la etiqueta ingresada
    regiones_disponibles = pd.read_excel(excel_file, sheet_name="transito")[pd.read_excel(excel_file, sheet_name="transito")['ETIQUETA'] == etiqueta]['REGION'].unique().tolist()
    # En excel, en la columna COMUNA, obtener las comunas de la región seleccionada (sin repetir) que tengan la etiqueta ingresada
    comunas_disponibles = pd.read_excel(excel_file, sheet_name="transito")[pd.read_excel(excel_file, sheet_name="transito")['ETIQUETA'] == etiqueta]['COMUNA'].unique().tolist()

    # Pedir al usuario que seleccione la región y la comuna
    region, comuna = obtener_region_y_comuna(regiones_disponibles, comunas_disponibles, excel_file)

    # Esperar a que el usuario presione una tecla para continuar
    input("Presione una tecla para continuar...")

    # Calcular el tiempo de entrega
    tiempo_estimado = calcular_tiempo(region, comuna, excel_file, etiqueta)
    
    # Mostrar el resultado al usuario
    print(f"El tiempo de entrega estimado para un pedido de tipo {tipo_pedido} en {comuna}, {region} es de {tiempo_estimado}")

if __name__ == "__main__":
    main()