import tkinter as tk
from tkinter import filedialog
import os
import csv
import time
import pyautogui


#! Codigo para sacar el .TXT y CSV
def extraer_eventos(log_path, output_txt_path, output_csv_path=None):
    try:
        with open(log_path, 'r', encoding='latin-1') as log_file:
            lines = log_file.readlines()
        eventos = []
        evento_actual = None

        for line in lines:
            if line.lstrip().startswith("25-"):  # Nueva entrada de log
                if evento_actual:
                    eventos.append(evento_actual)  # Guardar evento anterior
                fecha = line[0:9].strip()
                hora_completa = line[10:23].strip()
                proceso = line[22:38].strip()
                codigo_aci = line[38:56].strip()
                mensaje = line[65:70].strip()  
                complemento = line[70:].strip()  # Capturar el mensaje completo desde la posición 70
                
                evento_actual = [fecha, hora_completa, proceso, codigo_aci, mensaje, complemento]  
            elif evento_actual:  
                evento_actual[-1] += " " + line.strip()  # Agregar texto al mensaje

        if evento_actual:
            eventos.append(evento_actual)  

        if eventos:
            with open(output_txt_path, 'w', encoding='utf-8') as output_file:
                for evento in eventos:
                    output_file.write("|".join(evento) + "\n")  # Separar por |
            print(f"Archivo TXT creado con éxito en: {output_txt_path}")

            if output_csv_path:
                with open(output_csv_path, 'w', newline='', encoding='utf-8') as output_file:
                    csv_writer = csv.writer(output_file, delimiter='|') 
                    csv_writer.writerow(["Fecha", "Hora Completa", "Proceso", "Codigo ACI", "Mensaje", "Complemento"])
                    csv_writer.writerows(eventos)
                print(f"Archivo CSV creado con éxito en: {output_csv_path}")

        else:
            print("No se encontraron líneas válidas. No se creó el archivo de salida.")

    except Exception as e:
        print(f"Error al procesar el archivo: {e}")

#! Boton para seleccionar archivo
def seleccionar_archivo():
    global archivo_seleccionado
    archivo_seleccionado = filedialog.askopenfilename(filetypes=[("Todos los archivos", "*.*")])
    if archivo_seleccionado:
        print(f"Archivo seleccionado .TXT: {archivo_seleccionado}")

#! Boton para guardar el archivo en carpeta deseada bajo el nombre deseado
def guardar_archivo():
    if archivo_seleccionado:
        archivo_guardado = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Archivo de texto", "*.txt")])
        if archivo_guardado:
            output_csv_path = None
            if var_csv.get():
                output_csv_path = archivo_guardado.replace(".txt", "_CSV.csv")
            extraer_eventos(archivo_seleccionado, archivo_guardado, output_csv_path)

#! Converciones a CSV EXCEL 
def convertir_a_csv():
    if var_csv.get():
        print("Convertir a CSV")


#! Abrir el archivo PowerBI
def abrir_power_bi():
    # Ruta al archivo de Power BI
    ruta_power_bi = r"C:\\Users\\Rmprieto\\OneDrive - Redeban Multicolor\\Proyecto RBM\\Dios.pbix"
    os.startfile(ruta_power_bi)
    print("Abriendo Power BI...")

    # Esperar a que Power BI se abra (ajusta el tiempo según sea necesario)
    time.sleep(50)  # Aumentar el tiempo de espera

    # Simular la pulsación de teclas para abrir el apartado de "Transformar datos"
    pyautogui.hotkey('alt')
    time.sleep(2)  # Aumentar el tiempo de espera
    pyautogui.press('h')
    time.sleep(2) 
    pyautogui.press('t')
    time.sleep(2)  
    pyautogui.press('d')
    time.sleep(2)
    pyautogui.press('down')
    time.sleep(2)
    pyautogui.press('enter')
    print("Navegando a 'Transformar datos' en Power BI")


#! GUI      
# Crear ventana
root = tk.Tk()
root.title("Proyecto de grado")
root.geometry("400x400")

# Botón para abrir el explorador de archivos
btn_seleccionar = tk.Button(root, text="Seleccionar Archivo", command=seleccionar_archivo)
btn_seleccionar.pack(pady=20)

# Botón para guardar el archivo seleccionado
btn_guardar = tk.Button(root, text="Guardar Archivo", command=guardar_archivo)
btn_guardar.pack(pady=20)

# Checkbutton para convertir a CSV
var_csv = tk.BooleanVar()
chk_convertir_csv = tk.Checkbutton(root, text="Convertir a CSV", variable=var_csv, command=convertir_a_csv)
chk_convertir_csv.pack(pady=20)

# Botón para abrir Power BI y actualizar tablas
btn_abrir_power_bi = tk.Button(root, text="Abrir Power BI y Actualizar Tablas", command=abrir_power_bi)
btn_abrir_power_bi.pack(pady=50)

root.mainloop()