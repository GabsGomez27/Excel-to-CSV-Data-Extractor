from pathlib import Path
from tkinter import filedialog
from tkinter import Tk, Button, PhotoImage, Label, Frame
import sys

import pandas as pd
import os
import numpy as np

### CREACIÓN DEL CODIGO POR TKINTER DESIGNER

#creación de rutas para los assets
if getattr(sys, "frozen", False):
    OUTPUT_PATH = Path(sys._MEIPASS)
else:
    OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / "assets" / "frame0"
def relative_to_assets(path: str) -> str:
    return str(ASSETS_PATH / Path(path))


# =======================================================
# Creación de funciones y variables de extracción de datos
# =======================================================


ruta = ""
#Crea el dataframe maestro
master = pd.DataFrame(columns=[
    'Fecha',
    '#',
    'Cliente',
    'C.I',
    'Direccion',
    'Telefono',
    'Correo',
    'Poblacion',
    'Fecha de venta',
    'Mes',
    '# Factura',
    'Asesor',
    'Area',
    'Kit adquirido',
    'Plan',
    'Plan adquirido',
    'Monto del plan',
    'Monto Pagado',
    'Monto Exonerado',
    'Metodo de pago',
    'Referencia',
    'Promoción',
    'Fecha de Instalación',
    'Validación Instalación',
    'Información Contraloria',
    'Observaciones Adicionales',
    'Coordenadas',
])

def extraer_datos(datos_cliente, master_df):
  ###extrae los datos importantes del cliente
  ### Datos cliente = DataFrame (df equivalente a la ficha del cliente)
  ### master_df = DataFrame (df equivalente al maestro)

  valores = {
      'Fecha': pd.to_datetime(datos_cliente.iloc[2,14]).strftime('%d/%m/%Y'),
      '#': np.nan,
      'Cliente': datos_cliente.iloc[6,9],
      'C.I': datos_cliente.iloc[7,9],
      'Direccion': datos_cliente.iloc[29,4],
      'Telefono': datos_cliente.iloc[12,4],
      'Correo': datos_cliente.iloc[13,4],
      'Poblacion': np.nan,
      'Fecha de venta': pd.to_datetime(datos_cliente.iloc[2,14]).strftime('%d/%m/%Y'),
      'Mes': pd.to_datetime(datos_cliente.iloc[2,14]).strftime('%m'),
      '# Factura': np.nan,
      'Asesor': datos_cliente.iloc[19,3],
      'Area': np.nan,
      'Kit adquirido': datos_cliente.iloc[25,8],
      'Plan': np.nan,
      'Plan adquirido': datos_cliente.iloc[36,3],
      'Monto del plan': datos_cliente.iloc[15,17],
      'Monto Pagado': np.nan,
      'Monto Exonerado': np.nan,
      'Metodo de pago': np.nan,
      'Referencia': np.nan,
      'Promoción': 'SI',
      'Fecha de Instalación': np.nan,
      'Validación Instalación': np.nan,
      'Información Contraloria': np.nan,
      'Observaciones Adicionales': 'COMODATO',
      'Coordenadas': datos_cliente.iloc[31,4],
  }

  df=pd.DataFrame(valores, index=[0])
  master_df = pd.concat([master_df,df],axis=0).sort_values(by='Fecha', ascending=True).reset_index(drop=True)
  return master_df

def procesar_datos(ruta_carpeta, master_df):
    path_folder = ruta_carpeta
    folder = [f for f in os.listdir(path_folder) if f.endswith('.xlsx')]

    for file_name in folder:
        filepath = os.path.join(path_folder,file_name)
        datos_cliente = pd.read_excel(filepath, header=None)
        master_df = extraer_datos(datos_cliente, master_df)
    return master_df

def exportar_csv(master_df, ruta_exportacion):
   master_df.to_csv(os.path.join(ruta_exportacion, 'master_data.csv'), index=False, sep=';')

# ================================
# Creación de funciones de botones
# ================================

def button_1_command():
    # This opens the dialog and returns the path as a string
    folder_selected = filedialog.askdirectory()
    global ruta
    ruta = folder_selected
    global master

    if folder_selected:

        master = procesar_datos(folder_selected, master)

        print(f"Datos extraidos de: {folder_selected}")
        status_label.config(text=f"Datos extraidos de:\n{folder_selected}", fg="#66BB6A", font=("Segoe UI SemiBold", 16 * -1))
    else:
        print("User cancelled the selection")
        status_label.config(text="Ninguna carpeta \nseleccionada.", fg="#AAAAAA", font=("Segoe UI SemiBold", 20 * -1))

def button_2_command():
    folder_selected = filedialog.askdirectory()
    global ruta
    ruta = folder_selected
    global master
    
    exportar_csv(master, folder_selected)

    if folder_selected:
        hint_label.config(text=f"Datos exportados a:\n{folder_selected}", fg="#66BB6A")
        # Aquí se puede agregar la lógica para exportar los datos procesados
    else:
        print("User cancelled the selection for export")
        hint_label.config(text="Ninguna carpeta \nseleccionada para exportación.", fg="#AAAAAA", font=("Segoe UI SemiBold", 12 * -1))



# ================================
# Creación de la ventana principal
# ================================

window = Tk()

window.geometry("640x480")
window.configure(bg = "#1e1e1e")
window.title('WAVE DATA HELPER')

# Load main image and keep references
image_image_1 = PhotoImage(file=relative_to_assets("image_1.png"))

# Logo image (was canvas image)
logo_label = Label(window, image=image_image_1, bg="#1e1e1e")
logo_label.place(x=20, y=20, anchor='nw')

window.iconphoto(False, image_image_1)

# Title (was canvas text)
title_label = Label(window, text="WAVE DATA HELPER", fg="#AAAAAA", bg="#1e1e1e", font=("Segoe UI SemiBold", 24 * -1), anchor='nw')
title_label.place(x=100.0, y=25.0)

# Status text
status_label = Label(window, text="Ninguna carpeta \nseleccionada.", fg="#AAAAAA", bg="#1e1e1e", font=("Segoe UI SemiBold", 20 * -1), justify='left', anchor='nw')
status_label.place(x=20.0, y=154.0)

# Hint text
hint_label = Label(window, text="Al hacer clic, podrá elegir la ruta \nde exportación.", fg="#AAAAAA", bg="#1e1e1e", font=("Segoe UI SemiBold", 12 * -1), justify='left', anchor='nw')
hint_label.place(x=390.0, y=326.0)

button_image_1 = PhotoImage(file=relative_to_assets("button_1.png"))
button_1 = Button(window, image=button_image_1, borderwidth=0, highlightthickness=0, command=button_1_command, relief="flat")
button_1.place(x=24.0, y=240.0, width=233.0, height=60.0)

button_image_2 = PhotoImage(file=relative_to_assets("button_2.png"))
button_2 = Button(window, image=button_image_2, borderwidth=0, highlightthickness=0, command=button_2_command, relief="flat")
button_2.place(x=390.0, y=240.0, width=225.0, height=60.0)

# Vertical divider (was rectangle on canvas)
divider = Frame(window, bg="#D9D9D9", width=2, height=149)
divider.place(x=318.0, y=194.0)

window.resizable(False, False)
window.mainloop()

### FIN DEL CÓDIGO CREADO POR TKINTER DESIGNER

### EXTRACCIÓN DE DATOS Y PROCESAMIENTO

