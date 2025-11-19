import serial
import time
import tkinter as tk
from tkinter import messagebox, simpledialog
import pandas as pd
import os

# Configurar puerto y velocidad
puerto = 'COM5'   # 👈 cambia si tu Arduino usa otro puerto
baud_rate = 9600

# Crear ventana principal
root = tk.Tk()
root.title("Nariz Electrónica - Proyecto Cacao")
root.geometry("500x400")
root.config(bg="#1a98a0")

# Conectar con Arduino
try:
    arduino = serial.Serial(puerto, baud_rate, timeout=1)
    time.sleep(2)
    conectado = True
except Exception as e:
    conectado = False
    messagebox.showerror("Error", f"No se pudo conectar al Arduino:\n{e}")

# Etiquetas
titulo = tk.Label(root, text="🌱 Nariz Electrónica de Cacao", font=("Arial", 16, "bold"), bg="#1a98a0", fg="white")
titulo.pack(pady=10)

estado = tk.Label(root, text="Presiona 'Iniciar Lectura' para comenzar", font=("Arial", 12), bg="#1a98a0", fg="white")
estado.pack(pady=5)

resultado = tk.Label(root, text="", font=("Consolas", 12), bg="#1a98a0", fg="white", justify="left")
resultado.pack(pady=10)

# Función principal
def iniciar_lectura():
    if not conectado:
        messagebox.showwarning("Sin conexión", "Arduino no está conectado.")
        return

    estado.config(text="⏳ Realizando medición, por favor espera...")
    arduino.write(b'1')  # Enviar '1' al Arduino para iniciar la lectura

    datos = []
    tiempo_inicio = time.time()
    while True:
        linea = arduino.readline().decode('utf-8').strip()
        if linea:
            print(linea)
            datos.append(linea)
        if "Normalizados" in linea:
            break
        if time.time() - tiempo_inicio > 10:
            break

    if datos:
        resumen = "\n".join(datos[-3:])
        resultado.config(text=resumen)
        estado.config(text="✅ Medición completada.")
        procesar_y_guardar(datos)
    else:
        resultado.config(text="Sin respuesta del Arduino.")
        estado.config(text="⚠️ No se recibieron datos.")

def procesar_y_guardar(datos):
    """Extrae los valores de las líneas y guarda en Excel."""
    try:
        # Buscar las líneas relevantes
        linea_mq = [l for l in datos if "MQ1:" in l and "Normalizados" not in l][0]
        linea_norm = [l for l in datos if "Normalizados" in l][0]

        # Extraer valores de sensores
        mq_vals = [float(x.split(":")[1].strip()) for x in linea_mq.split("|")]
        mq_norm = [float(x.split(":")[1].strip()) for x in linea_norm.replace("Normalizados ->", "").split("|")]

        # Pedir aroma identificado
        aroma = simpledialog.askstring("Aroma detectado", "Ingrese el aroma identificado (Ej: Floral, Fuerte, Suave):")

        # Crear DataFrame con columnas ordenadas
        df = pd.DataFrame([{
            "MQ1N": mq_norm[0],
            "MQ2N": mq_norm[1],
            "MQ3N": mq_norm[2],
            "MQ1": mq_vals[0],
            "MQ2": mq_vals[1],
            "MQ3": mq_vals[2],
            "Aroma": aroma
        }])

        archivo = "lecturas_cacao.xlsx"

        # Agregar o crear Excel
        if os.path.exists(archivo):
            df_existente = pd.read_excel(archivo)
            df_final = pd.concat([df_existente, df], ignore_index=True)
        else:
            df_final = df

        df_final.to_excel(archivo, index=False, engine='openpyxl')
        print(f"💾 Datos guardados correctamente en {archivo}")
        messagebox.showinfo("Guardado", "Datos guardados exitosamente en Excel.")

    except Exception as e:
        messagebox.showerror("Error", f"Error al procesar o guardar datos:\n{e}")

# Botón
boton = tk.Button(root, text="Iniciar Lectura", font=("Arial", 12, "bold"), bg="white", fg="#1a98a0",
                  activebackground="#148085", command=iniciar_lectura)
boton.pack(pady=20)

# Ejecutar interfaz
root.mainloop()

# Cerrar conexión
if conectado:
    arduino.close()
