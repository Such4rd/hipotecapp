from tkinter import ttk
from tkinter import Tk
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd


# Variable global para almacenar los resultados de cada cuota
tabla_cuotas = []

def verificarValores():
    for key, entry in entries.items():
        valor = entry.get()
        if not valor:
            messagebox.showerror("Error", f"El campo '{key}' no puede estar vacío.")
            return False
        try:
            if key in ["fin", "tin_f1", "tin_f", "tin_v", "euribor"]:  # Son valores flotantes
                float(valor)
            elif key in ["cuotas", "t_fijo"]:  # Son valores enteros
                int(valor)
        except ValueError:
            messagebox.showerror("Error", f"El campo '{key}' debe contener un número válido.")
            return False
    return True

def calcularInteres():
    
    global tabla_cuotas  # Para asegurarnos de que la tabla se puede acceder fuera de la función
    tabla_cuotas = []  # Reiniciar la tabla al iniciar un nuevo cálculo

    # Verificar si todos los campos están completados
    if not verificarValores():
        return
    # Obtener los valores ingresados en los Entry
    valores = {
        "fin": float(entries["fin"].get()), #cantidad a financiar
        "tin_f1": float(entries["tin_f1"].get()),  # Corregido
        "tin_f": float(entries["tin_f"].get()),
        "tin_v": float(entries["tin_v"].get()),
        "euribor": float(entries["euribor"].get()),
        "cuotas": int(entries["cuotas"].get()) * 12,
        "t_fijo": int(entries["t_fijo"].get()) * 12
    }

    # Obtener amortizaciones
    amortizaciones_lista = []
    
    for monto_entry, cuotas_entry in amortizaciones:
        try:
            monto = float(monto_entry.get()) if monto_entry.get() else 0
            cuotas = int(cuotas_entry.get()) * 12 if cuotas_entry.get() else 0

            if monto <= 0 or cuotas <= 0:
                messagebox.showerror("Error", "Los valores de amortización y cuotas deben ser positivos y mayores que cero")
                return

            # Si todo es correcto, añadir la amortización a la lista
            amortizaciones_lista.append((monto, cuotas))

        except ValueError:
            messagebox.showerror("Error de formato", "Por favor, ingresa valores numéricos válidos en las amortizaciones.")
            return
    
    n = valores["cuotas"]  # cuotas pendientes
    total_pagado = 0
    total_intereses = 0
    es_fijo = True
    cuota_v = 0
    cuota_f = 0

    for i in range(1, valores["cuotas"]+1):
        if i <= 12:
            tin_e = valores["tin_f1"] / 12 / 100  # Corregido
        elif i <= valores["t_fijo"]:
            tin_e = valores["tin_f"] / 12 / 100
        else:
            es_fijo = False
            tin_e = (valores["tin_v"] + valores["euribor"]) / 12 / 100

        cuota = valores["fin"] * tin_e * (1 + tin_e) ** n / ((1 + tin_e) ** n - 1)
        if es_fijo:
            cuota_f = cuota
        else :
            cuota_v = cuota

        intereses = valores["fin"] * tin_e
        cap_amortizado = cuota - intereses
        valores["fin"] -= cap_amortizado  # vamos restando el capital amortizado al total financiado
        n -= 1  # restamos las cuotas también
        total_intereses += intereses
        total_pagado += cuota

        # Aplicar amortizaciones múltiples
        for monto, cuotas_amortizacion in amortizaciones_lista:
            if i == cuotas_amortizacion:
                valores["fin"] -= monto
                print(f"Amortización de {monto} realizada en la cuota {i}")
                if valores["fin"] <= 0:
                    actualizarResultados(total_intereses, total_pagado, cuota_v, cuota_f)
                    messagebox.showinfo("Amortización", "La deuda ha sido completamente pagada.")
                    break  # Si la deuda se paga completamente, salir del bucle
        
        # Guardar la información de la cuota en la tabla
        tabla_cuotas.append({
            "Cuota N°": i,
            "Cuota": round(cuota, 2),
            "Intereses": round(intereses, 2),
            "Amortizado": round(cap_amortizado, 2),
            "Por pagar": round(valores['fin'], 2)
        })
            # Imprimir la información de la cuota
        print(f"Cuota {i} : {round(cuota, 2)} - Intereses: {round(intereses, 2)} - Amortizado: {round(cap_amortizado, 2)} - Por pagar: {round(valores['fin'], 2)}")

        # Establecer el resultado en los Entry correspondientes
    actualizarResultados(total_intereses, total_pagado, cuota_v, cuota_f)
    btn_excel.pack(pady=5)
        
def actualizarResultados(total_intereses, total_pagado, cuota_v, cuota_f):
    # Actualizar campos de resultados
    configurarResultado(interese_entry, total_intereses)
    configurarResultado(pagado_entry, total_pagado)
    configurarResultado(cuota_v_entry, cuota_v)
    configurarResultado(cuota_f_entry, cuota_f)

def configurarResultado(entry, valor):
    # Permite insertar el valor y luego vuelve a readonly
    entry.configure(state="normal")
    limpiarEntry(entry)
    entry.insert(0, str(round(valor, 2)))
    entry.configure(state="readonly")   

def limpiarEntry(entry):
    entry.delete(0, 'end')


def generar_excel():
    
    if not tabla_cuotas:
        messagebox.showerror("Error", "No hay datos para exportar. Realiza un cálculo primero.")
        return

    # Preguntar al usuario dónde guardar el archivo
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    if not filepath:
        return  # Si el usuario cancela la operación

    # Crear un DataFrame a partir de la tabla de cuotas
    df = pd.DataFrame(tabla_cuotas)

    # Guardar el DataFrame en un archivo Excel
    df.to_excel(filepath, index=False)

    messagebox.showinfo("Éxito", f"El archivo Excel ha sido guardado en {filepath}")

root = Tk()
root.title("Hipotecapp (Calculadora de intereses)")

# Estilos
style = ttk.Style()
style.theme_use("clam")  # Cambia el tema del estilo
style.configure("TButton", padding=10, width=15)

# Establecer un estilo para etiquetas en negrita
style.configure("Bold.TLabel", font=('Arial', 10, 'bold'), background="pink")

# Configurar el color de fondo a rosa para los marcos
style.configure("Pink.TFrame", background="pink")

# Configurar el estilo del borde de los marcos
style.configure("Rounded.TFrame", background="pink", borderwidth=2, relief="sunken", bordercolor="grey")

# Stilo para el boton
style.configure("Custom.TButton", padding=10, width=15, background="lightblue", foreground="black", font=("Helvetica", 12, "bold"))


frm = ttk.Frame(root, padding=10, style="Pink.TFrame")
frm.pack(expand=True, fill='both')  # Hace que el marco se expanda en todas las direcciones

# Crear un frame para los campos a rellenar
formulario_frame = ttk.Frame(frm, padding=10, style="Rounded.TFrame")
formulario_frame.pack(side='left', fill='both', expand=True)

# Crear los Entry para los valores
entries = {
    "fin": "Cantidad a financiar",
    "tin_f1": "Interés primer año",
    "t_fijo": "Tiempo fijo (años)",
    "tin_f": "Interés período FIJO",
    "tin_v": "Interés período VARIABLE",
    "euribor": "EURIBOR",
    "cuotas": "Años"
}

for key, value in entries.items():
    entry_frame = ttk.Frame(formulario_frame, style="Rounded.TFrame")  # Crear un frame para el label y el entry
    entry_frame.pack(fill='x', padx=5, pady=5)  # Empaquetar el frame

    ttk.Label(entry_frame, text=value, style="Bold.TLabel").pack(side='left',padx=5,pady=6)  # Aplicar el estilo de fuente en negrita
    entries[key] = ttk.Entry(entry_frame, width=25)  # Establecer un ancho predeterminado
    entries[key].pack(side='right', fill='x', padx=5, pady=5)  # Empaquetar el entry


# En el frame de formulario
amortizaciones_frame = ttk.Frame(formulario_frame, style="Rounded.TFrame")
amortizaciones_frame.pack(fill='x', padx=5, pady=5)

# Agregar un botón para añadir nuevas amortizaciones
def agregar_amortizacion():
    frame_amortizacion = ttk.Frame(amortizaciones_frame, style="Rounded.TFrame")
    frame_amortizacion.pack(fill='x', padx=5, pady=5)
    ttk.Label(frame_amortizacion, text="Monto Amortización", style="Bold.TLabel").pack(side='left', padx=5, pady=6)
    monto_entry = ttk.Entry(frame_amortizacion, width=15)
    monto_entry.pack(side='left', padx=5, pady=5)
    
    ttk.Label(frame_amortizacion, text="Cuotas (desde)", style="Bold.TLabel").pack(side='left', padx=5, pady=6)
    cuotas_entry = ttk.Entry(frame_amortizacion, width=15)
    cuotas_entry.pack(side='left', padx=5, pady=5)
    
    amortizaciones.append((monto_entry, cuotas_entry))


# Lista para almacenar amortizaciones
amortizaciones = []

# Crear un frame para los campos de resultado
resultado_frame = ttk.Frame(frm, padding=10, style="Rounded.TFrame")
resultado_frame.pack(side='right', fill='both', expand=True)

# Crear los Entry para mostrar los resultados
intereses_frame = ttk.Frame(resultado_frame, style="Rounded.TFrame") #frame total intereses
intereses_frame.pack(fill='x', padx=5, pady=5)
ttk.Label(intereses_frame, text="Total Intereses", style="Bold.TLabel").pack(anchor='w',padx=5,pady=6)  # Aplicar el estilo de fuente en negrita
interese_entry = ttk.Entry(intereses_frame, state="readonly", width=30)  # Establecer un ancho predeterminado
interese_entry.pack(side='right', fill='x', expand=True, padx=5,pady=6)

pagado_frame = ttk.Frame(resultado_frame, style="Rounded.TFrame") #frame total pagado
pagado_frame.pack(fill='x', padx=5, pady=5)
ttk.Label(pagado_frame, text="Total Pagado", style="Bold.TLabel").pack(anchor='w',padx=5,pady=6)  # Aplicar el estilo de fuente en negrita
pagado_entry = ttk.Entry(pagado_frame, state="readonly", width=30)  # Establecer un ancho predeterminado
pagado_entry.pack(side='right', fill='x', expand=True, padx=5,pady=6)

cuota_v_frame = ttk.Frame(resultado_frame, style="Rounded.TFrame") #frame cuotas variable
cuota_v_frame.pack(fill='x', padx=5, pady=5)
ttk.Label(cuota_v_frame, text="Cuota variable", style="Bold.TLabel").pack(anchor='w',padx=5,pady=6)  # Aplicar el estilo de fuente en negrita
cuota_v_entry = ttk.Entry(cuota_v_frame, state="readonly", width=30)  # Establecer un ancho predeterminado
cuota_v_entry.pack(side='right', fill='x', expand=True, padx=5,pady=6)

cuota_f_frame = ttk.Frame(resultado_frame, style="Rounded.TFrame") #frame cuotas fijas
cuota_f_frame.pack(fill='x', padx=5, pady=5)
ttk.Label(cuota_f_frame, text="Cuota fija", style="Bold.TLabel").pack(anchor='w',padx=5,pady=6)  # Aplicar el estilo de fuente en negrita
cuota_f_entry = ttk.Entry(cuota_f_frame, state="readonly", width=30)  # Establecer un ancho predeterminado
cuota_f_entry.pack(side='right', fill='x', expand=True, padx=5,pady=6)

# Botón para calcular
btn_calcular = ttk.Button(resultado_frame, text="Calcular", command=calcularInteres, style="Custom.TButton")
btn_calcular.pack(padx=5, pady=5)

# Botón para añadir amortización
agregar_button = ttk.Button(resultado_frame, text="Amortizar", command=agregar_amortizacion, style="Custom.TButton")
agregar_button.pack(padx=5,pady=5)

# Botón para generar el Excel en la interfaz
btn_excel = ttk.Button(resultado_frame, text="Generar Excel", command=generar_excel, style="Custom.TButton")
btn_excel.pack(padx=5,pady=5)
btn_excel.pack_forget()  # Ocultar el botón inicialmente

root.mainloop()
