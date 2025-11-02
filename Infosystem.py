from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk
import socket
import uuid
import platform
import psutil
import wmi
import csv
import os
import sys

tabla = None

# Función para obtener la ruta correcta de los recursos cuando el programa está empaquetado
def obtener_ruta_recurso(rel_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, rel_path)
    return os.path.join(os.path.abspath("."), rel_path)

def obtener_datos_equipo():
    c = wmi.WMI()
    nombre_equipo = socket.gethostname()
    try:
        ip = socket.gethostbyname(socket.gethostname())
    except:
        ip = "No detectada"
    mac1 = ':'.join(['{:02x}'.format((uuid.getnode() >> ele) & 0xff) for ele in range(40, -8, -8)])
    mac2 = "No detectada"
    procesador = platform.processor()
    memoria = str(round(psutil.virtual_memory().total / (1024**3), 2)) + " GB"
    disco = str(round(psutil.disk_usage('/').total / (1024**3), 2)) + " GB"
    try:
        bios = c.Win32_BIOS()[0]
        modelo = c.Win32_ComputerSystem()[0].Model
        marca = c.Win32_ComputerSystem()[0].Manufacturer
        serie = bios.SerialNumber
    except:
        modelo = marca = serie = "No detectado"
    dvd = "No"
    try:
        unidades = c.Win32_CDROMDrive()
        if unidades:
            dvd = "Sí"
    except:
        pass
    return [serie, modelo, marca, ip, mac1, mac2, nombre_equipo, procesador, disco, memoria, dvd]

def guardar_csv_multiple(datos, archivo='equipos_autodetectados.csv'):
    file_exists = os.path.isfile(archivo)
    with open(archivo, 'a', newline='') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["SERIE", "MODELO", "MARCA", "IP", "MAC 1", "MAC 2", "NOMBRE EQUIPO", "PROCESADOR", "DISCO DURO", "MEMORIA RAM", "DVD-ROM"])
        writer.writerow(datos)

def recolectar_y_actualizar():
    global tabla
    info = obtener_datos_equipo()
    guardar_csv_multiple(info)
    tabla.insert("", "end", values=info)
    print("Datos del equipo guardados.")

def crear_ventana():
    global tabla

    ventana = tk.Tk()
    ventana.title("InfoSystem")
    ventana.configure(bg="#1e1e1e")
    ventana.geometry("1300x680")

    contenedor = tk.Frame(ventana, bg="#1e1e1e")
    contenedor.pack(fill="both", expand=True)

    # Frame izquierdo con imagen
    frame_izquierdo = tk.Frame(contenedor, width=250, bg="#1e1e1e")
    frame_izquierdo.pack(side="left", fill="y", padx=10, pady=10)

    try:
        ruta_logo = obtener_ruta_recurso("logo.png")
        imagen_original = Image.open(ruta_logo)
        imagen_redimensionada = imagen_original.resize((200, 200), Image.LANCZOS)
        imagen_tk = ImageTk.PhotoImage(imagen_redimensionada)
        label_imagen = tk.Label(frame_izquierdo, image=imagen_tk, bg="#1e1e1e")
        label_imagen.image = imagen_tk
        label_imagen.pack(pady=20)
    except Exception as e:
        print(f"No se pudo cargar la imagen: {e}")

    # Frame derecho con tabla
    frame_derecho = tk.Frame(contenedor, bg="#1e1e1e")
    frame_derecho.pack(side="left", fill="both", expand=True, padx=10, pady=10)

    frame_canvas = tk.Frame(frame_derecho, bg="#1e1e1e")
    frame_canvas.pack(fill="both", expand=True)

    canvas = tk.Canvas(frame_canvas, bg="#1e1e1e", highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)

    scrollbar_y = ttk.Scrollbar(frame_canvas, orient="vertical", command=canvas.yview)
    scrollbar_y.pack(side="right", fill="y")

    scrollbar_x = ttk.Scrollbar(ventana, orient="horizontal", command=canvas.xview)
    scrollbar_x.pack(side="bottom", fill="x")

    canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    frame_interno = tk.Frame(canvas, bg="#1e1e1e")
    canvas.create_window((0, 0), window=frame_interno, anchor="nw")

    def actualizar_scroll(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    frame_interno.bind("<Configure>", actualizar_scroll)

    columnas = ["SERIE", "MODELO", "MARCA", "IP", "MAC 1", "MAC 2", "NOMBRE EQUIPO", "PROCESADOR", "DISCO DURO", "MEMORIA RAM", "DVD-ROM"]
    tabla = ttk.Treeview(frame_interno, columns=columnas, show="headings", height=20)
    tabla.pack(expand=True, fill="both", padx=10, pady=10)

    estilo = ttk.Style()
    estilo.theme_use("clam")
    estilo.configure("Treeview",
                     background="#2e2e2e",
                     foreground="white",
                     rowheight=25,
                     fieldbackground="#2e2e2e")
    estilo.map("Treeview",
               background=[("selected", "#33FF99")],
               foreground=[("selected", "#000000")])

    for col in columnas:
        tabla.heading(col, text=col)
        tabla.column(col, width=150, anchor="center", stretch=True)

    boton_recolectar = tk.Button(ventana, text="Recolectar Datos de un Equipo", command=recolectar_y_actualizar,
                                 font=("Arial", 12), bg="#4CAF50", fg="white", relief="solid", bd=2)
    boton_recolectar.pack(pady=10)

    etiqueta_desarrollador = tk.Label(ventana, text="Developer: Bishox",
                                      font=("Orbitron", 14, "bold"), fg="#00FF00",
                                      bg="#1e1e1e", relief="ridge", bd=2,
                                      padx=12, pady=6)
    etiqueta_desarrollador.pack(side="bottom", pady=10)

    ventana.mainloop()

crear_ventana()
