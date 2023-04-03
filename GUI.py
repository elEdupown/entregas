import tkinter as tk
import pandas as pd
from datetime import datetime, timedelta
from collections import deque

import tkinter as tk

def obtener_tipo_pedido():
    tipos_validos = ["Postventa", "Televenta", "Ecommerce", "Business", "CRM", "OTT"]
    
    def validar_tipo_pedido():
        tipo_pedido = tipo_pedido_var.get()
        if tipo_pedido in tipos_validos:
            ventana.destroy()
        else:
            mensaje.config(text="Tipo de pedido no válido. Intente nuevamente.")
    
    ventana = tk.Tk()
    ventana.title("Seleccione el tipo de pedido")
    
    tipo_pedido_var = tk.StringVar()
    opciones = tipos_validos.copy()
    opciones.insert(0, "")
    tipo_pedido_var.set("")
    
    tk.Label(ventana, text="Tipo de pedido:").grid(row=0, column=0)
    tipo_pedido_menu = tk.OptionMenu(ventana, tipo_pedido_var, *opciones)
    tipo_pedido_menu.grid(row=0, column=1)
    
    mensaje = tk.Label(ventana, text="")
    mensaje.grid(row=1, column=0, columnspan=2)
    
    boton_aceptar = tk.Button(ventana, text="Aceptar", command=validar_tipo_pedido)
    boton_aceptar.grid(row=2, column=1)
    
    ventana.mainloop()
    
    return tipo_pedido_var.get()

def obtener_region_y_comuna(regiones, etiqueta, excel_file):
    """Obtiene la región y la comuna seleccionadas por el usuario."""

    # Crear la ventana gráfica
    root = tk.Tk()
    root.title("Selección de región y comuna")

    # Variables para almacenar los valores seleccionados
    var_region = tk.StringVar()
    var_comuna = tk.StringVar()

    # Crear el widget de selección de región
    label_region = tk.Label(root, text="Seleccione la región:")
    label_region.pack()

    def actualizar_comunas(*args):
        """Actualiza el menú de selección de comunas según la región elegida."""
        region_seleccionada = var_region.get()
        df = pd.read_excel(excel_file, sheet_name='transito')
        # Buscar las comunas en columna COMUNA que correspondan a la region de columna REGION y a la etiqueta de columna ETIQUETA
        comunas_region = df[(df['REGION'] == region_seleccionada) & (df['ETIQUETA'] == etiqueta)]['COMUNA'].unique()
        menu_comuna['menu'].delete(0, 'end')
        for comuna in comunas_region:
            menu_comuna['menu'].add_command(label=comuna, command=tk._setit(var_comuna, comuna))

    menu_region = tk.OptionMenu(root, var_region, *regiones, command=actualizar_comunas)
    menu_region.pack()

    # Crear el widget de selección de comuna
    label_comuna = tk.Label(root, text="Seleccione la comuna:")
    label_comuna.pack()

    menu_comuna = tk.OptionMenu(root, var_comuna, "")
    menu_comuna.pack()

    # Función para obtener los valores seleccionados
    def obtener_seleccion():
        region = var_region.get()
        comuna = var_comuna.get()
        root.destroy()
        return region, comuna

    # Crear el botón de confirmación
    boton_confirmar = tk.Button(root, text="Confirmar", command=obtener_seleccion)
    boton_confirmar.pack()

    # Mostrar la ventana gráfica
    root.mainloop()

    # Filtrar las comunas según la región seleccionada
    region_seleccionada = var_region.get()

    # Retornar los valores seleccionados
    return region_seleccionada, var_comuna.get()

def obtener_tipo_entrega():
    """Obtiene el tipo de entrega seleccionado por el usuario."""
    tipos_validos = ["Express", "Standard"]

    root = tk.Tk()
    # root.title("Seleccionar tipo de entrega")
    # root.geometry("300x100")
    # root.resizable(False, False)
    respuesta = tk.StringVar()
    root.destroy()
    def seleccionar_tipo_entrega(tipo_entrega):
        if tipo_entrega in tipos_validos:
            respuesta.set(tipo_entrega)
            root.destroy()
    
    root = tk.Tk()
    root.title("Seleccionar tipo de entrega")
    root.geometry("300x150")
    root.resizable(False, False)

    label = tk.Label(root, text="Seleccione el tipo de entrega:", font=("Arial", 12))
    label.pack(pady=5)

    express_button = tk.Button(root, text="Express", width=10, command=lambda: seleccionar_tipo_entrega("Express"))
    express_button.pack(side=tk.LEFT, padx=10)

    standard_button = tk.Button(root, text="Standard", width=10, command=lambda: seleccionar_tipo_entrega("Standard"))
    standard_button.pack(side=tk.RIGHT, padx=10)

    root.mainloop()
    
    return respuesta.get()

def calcular_tiempo(region, comuna, excel_file, etiqueta):
    # Obtener la fecha actual
    fecha_actual = datetime.today()

    # Cargar el archivo de Excel y obtener la hoja correspondiente
    df = pd.read_excel(excel_file, sheet_name="transito")
    # Obtener filas de la hoja que tengan la etiqueta ingresada, la región ingresada y la comuna ingresada
    df = df[(df['ETIQUETA'] == etiqueta) & (df['REGION'] == region) & (df['COMUNA'] == comuna)]
    # obtener el tiempo de espera de la hoja
    tiempo_espera_str = df[df['COMUNA'] == comuna]['TIEMPO_ESPERA'].values[0]
    # Obtener los días hábiles de la hoja
    dias_habiles = df[df['COMUNA'] == comuna]['DIAS_HABILES'].values[0]
    # Parsear los días hábiles a una lista enlazada de enteros y restar 1 a cada elemento
    dias_habiles = deque(int(dia)-1 for dia in dias_habiles.split(","))

    # Convertir el tiempo de espera a dias
    if tiempo_espera_str.endswith("d"):
        tiempo_espera = int(tiempo_espera_str[:-2])
    elif tiempo_espera_str.endswith("h"):
        tiempo_espera = int(tiempo_espera_str[:-2]) / 24
    else:
        raise ValueError("Formato de tiempo de espera inválido")
    
    print(f"Tiempo de espera: {tiempo_espera} días")
    # Obtener el número de día hora de fecha_actual
    dia_actual, hora_actual = fecha_actual.weekday(), fecha_actual.hour
    fecha_entrega = fecha_actual
    
    # Si el tiempo de espera es menor a 1 día y el día actual está en días hábiles
    if tiempo_espera < 1 and dia_actual in dias_habiles and hora_actual <= 12:
        fecha_entrega = fecha_actual + timedelta(hours=tiempo_espera*24)
        fecha_entrega_str = "hoy "+ fecha_entrega.strftime("%d/%m/%Y") + " a las " + fecha_entrega.strftime("%H") + " apróx"
        return fecha_entrega_str
    
    # if dia_actual in dias_habiles:
    #     if hora_actual > 12:
    #         for i in range(tiempo_espera):

    #     else:


        
    # Verificar si el tiempo de espera es 1 día y si el siguiente día hábil está en días hábiles
    if tiempo_espera == 1 and dia_actual in dias_habiles:
        if hora_actual <= 12:
            fecha_entrega = fecha_actual
        else:
            fecha_entrega +=timedelta(days=1)
    if tiempo_espera > 1 and dia_actual not in dias_habiles:
        fecha_entrega += timedelta(days=1)
        for i in range(tiempo_espera):
            fecha_entrega += timedelta(days=1)
            if fecha_entrega.weekday() == 6:  # domingo
                tiempo_espera += 1
                fecha_entrega += timedelta(days=1)
            elif fecha_entrega.weekday() == 5 and hora_actual >= 12:  # sábado después del mediodía
                tiempo_espera += 2
                fecha_entrega += timedelta(days=2)
            elif fecha_entrega.weekday() not in dias_habiles and fecha_entrega.weekday() != 6:  # días no hábiles
                fecha_entrega += timedelta(days=1)

    # Obtener la fecha de entrega en formato de string
    fecha_entrega_str = fecha_entrega.strftime("%d/%m/%Y")

    # Devolver la fecha de entrega estimada
    return fecha_entrega_str

def obtener_etiqueta():
    """Obtiene la etiqueta seleccionada por el usuario."""
    # Crear la ventana gráfica
    root = tk.Tk()
    root.title("Selección de etiqueta")
    root.geometry("300x150")
    # Variables para almacenar los valores seleccionados
    var_etiqueta = tk.StringVar()

    # Crear el widget de selección de etiqueta
    label_etiqueta = tk.Label(root, text="Seleccione la etiqueta:")
    label_etiqueta.pack()

    opciones_etiqueta = ["EquipoSIM", "SIM"]
    menu_etiqueta = tk.OptionMenu(root, var_etiqueta, *opciones_etiqueta)
    menu_etiqueta.pack()

    # Crear el botón de confirmación
    boton_confirmar = tk.Button(root, text="Confirmar", command=root.quit)
    boton_confirmar.pack()

    root.mainloop()
    root.destroy()
    return var_etiqueta.get()

def mostrar_tiempo_estimado(tipo_pedido, comuna, region, tiempo_estimado):
    """Muestra el tiempo estimado de entrega en una etiqueta."""
    
    # Crear la ventana gráfica
    root = tk.Tk()
    root.title("Tiempo estimado de entrega")

    # Crear la etiqueta para mostrar el mensaje
    mensaje_label = tk.Label(root, text=f"El tiempo de entrega estimado para el pedido de tipo {tipo_pedido} en {comuna}, {region} es para {tiempo_estimado}", 
                             font=("Arial", 12),
                             height=5)
    mensaje_label.pack()

    # Mostrar la ventana gráfica
    root.mainloop()



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

    etiqueta = obtener_etiqueta()

    # Ir al excel y obtener en una lista las regiones de la columna REGION (sin repetir) que tengan la etiqueta ingresada
    regiones_disponibles = pd.read_excel(excel_file, sheet_name="transito")[pd.read_excel(excel_file, sheet_name="transito")['ETIQUETA'] == etiqueta]['REGION'].unique().tolist()
    # Pedir al usuario que seleccione la región y la comuna
    region, comuna = obtener_region_y_comuna(regiones_disponibles, etiqueta, excel_file)
    # Calcular el tiempo de entrega
    tiempo_estimado = calcular_tiempo(region, comuna, excel_file, etiqueta)
    
    # Mostrar el resultado al usuario
    mostrar_tiempo_estimado(tipo_pedido, comuna, region, tiempo_estimado)

    # print(f"El tiempo de entrega estimado para el pedido de tipo {tipo_pedido} en {comuna}, {region} es para {tiempo_estimado}")

if __name__ == "__main__":
    main()