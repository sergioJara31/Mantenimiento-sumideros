import tkinter as tk
from tkinter import ttk, filedialog, Image
from PIL import Image, ImageTk
import threading
import time

from kobo_api import proceso,obtener_uid, crear_sesion, obtener_formulario, procesar_datos, descargar_fotos, crear_excel, insertar_imagenes 

def centrar_ventana(ventana, ancho, alto):
    # 1. Obtener las dimensiones de la pantalla
    pantalla_ancho = ventana.winfo_screenwidth()
    pantalla_alto = ventana.winfo_screenheight()
    
    # 2. Calcular las coordenadas X e Y para centrar
    x = int((pantalla_ancho - ancho) / 2)
    y = int((pantalla_alto - alto) / 2)
    
    # 3. Aplicar la geometría a la ventana
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
# -------------------------
# Ventana
# -------------------------

ventana = tk.Tk()

ventana.title("Descarga de Imagenes y Creación de Excel")

centrar_ventana(ventana, 600, 400) 

imagen = Image.open("logo.jpeg").convert("RGBA")
ventana.iconbitmap("logo.ico")
# Ajustarla al tamaño de la ventana
imagen = imagen.resize((600, 400), Image.Resampling.LANCZOS)
imagen.putalpha(60)

# Convertirla al formato de Tkinter
fondo = ImageTk.PhotoImage(imagen)
# Colocarla como fondo
canvas = tk.Canvas(
    ventana,
    width=600,
    height=400,
    highlightthickness=0,
    bd=0
)
canvas.pack(fill="both", expand=True)

canvas.create_image(0, 0, image=fondo, anchor="nw")

canvas.create_text(
    300,
    40,
    text="Mantenimiento de sumideros",
    fill="darkblue",
    font=("Arial", 16, "bold")
)

canvas.create_text(
    150,
    80,
    text="Actualizar carpeta existente",
    fill="blue",
    font=("Arial", 14)
)

canvas.create_text(
    280,
    120,
    text="Selecciona una carpeta ya creada que desee actualizar (solo la carpeta del semestre)",
    fill="navy",
    font=("Arial", 10, )
)

def direccionArchivoExcel():
    carpeta_destino = filedialog.askdirectory(title="Seleccionar Carpeta para Guardar el Excel")
    if carpeta_destino:
        return carpeta_destino
    else:
        return ""
    
def crearExcel(crearArchivo):
    carpeta = direccionArchivoExcel()
    if(carpeta):
        iniciar_proceso(carpeta, crearArchivo)
    else:
        print("eror al conseguir direccion")


btn_actualizar = tk.Button(
    ventana,
    text="Actualizar",
    width=15,
    command=lambda: crearExcel(False)
)

canvas.create_window(
    120,          # x
    180,          # y
    window=btn_actualizar 
)


def iniciar_proceso(rutaCarpeta, crearArchivo, ventana=ventana ):
    # Crear ventana de progreso
    ventanaBarra = tk.Toplevel(ventana)
    ventanaBarra.protocol("WM_DELETE_WINDOW", lambda: None)
    ventanaBarra.title("Procesando")
    centrar_ventana(ventanaBarra, 350, 100)
    ventanaBarra.resizable(False, False)

    lbl_estado = tk.Label(
        ventanaBarra,
        text="Generando el archivo de Excel...\n"
             "Este proceso puede tardar unos minutos, por favor espere."
    )   
    lbl_estado.pack(pady=10)

    barra = ttk.Progressbar(
        ventanaBarra,
        mode="indeterminate",
        length=300
    )
    barra.pack(pady=10)
    barra.start(10)

    def tarea():
        try:
           resultadoProceso=proceso(rutaCarpeta, crearArchivo)
        finally:
            # Volver al hilo principal para cerrar la ventanaBarra
            if(resultadoProceso):
                ventanaBarra.after(0, lambda:finalizar(True))
            else:
                ventanaBarra.after(0, lambda:finalizar(False))

    def finalizar(resultado):
        barra.stop()
        barra.pack_forget()
        if resultado:
            lbl_estado.config(text="Proceso finalizado correctamente.", fg="green")
        else:
             lbl_estado.config(text="Ha ocurrido un error por favor vuelva a intertarlo", fg="red")
    
    ttk.Button(
            ventanaBarra, 
            text="Cerrar",
            command=ventanaBarra.destroy
        ).pack(pady=10)

    threading.Thread(target=tarea, daemon=True).start()



canvas.create_text(
    120,
    220,
    text="Crear nueva carpeta",
    fill="blue",
    font=("Arial", 14)
)

canvas.create_text(
    280,
    250,
    text="Selecciona una carpeta donde se creara el documento de excel y la carpeta de fotos",
    fill="navy",
    font=("Arial", 10, )
)


btn_crear = tk.Button(
    ventana,
    text="Crear Excel",
    width=15,
    command=lambda: crearExcel(True)
    #command=direccionArchivoExcel
    
)
canvas.create_window(
    120,          # x
    300,          # y
    window= btn_crear
)

ventana.resizable(False, False)
ventana.mainloop()

