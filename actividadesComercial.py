import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime, timedelta

# Configuración inicial
file_path = r"C:\Users\Admin Local\Downloads\Informe de Actividades.csv"
save_path = "C:/Users/Admin Local/OneDrive - PUNTO EMPLEO S.A/Automatizacion/Proyecto reporte de actividades COMERCIAL/Comercial_reporte_de_actividades"

# Verificar que la ruta de guardado exista
if not os.path.exists(save_path):
    os.makedirs(save_path)

# Lista de días festivos
dias_festivos = [
    '2024-08-07', 
    '2024-08-19',
    '2024-10-14',
    '2024-11-04',
    '2024-11-11',
    '2024-12-25',
    '2025-01-01',
    '2025-01-06',
    '2025-02-17',
    '2025-03-01',
    '2025-03-08',
    '2025-03-24',
    '2025-04-17',
    '2025-04-18',
    '2025-04-26',
    '2025-05-01',
    '2025-06-02',
    '2025-06-23',
    '2025-06-30',
    '2025-08-07',
    '2025-08-18',
    '2025-09-20',
    '2025-10-13',
    '2025-11-03',
    '2025-11-17',
    '2025-12-08',
    '2025-12-25',
    '2026-01-01',
]

# Cargar archivo CSV
try:
    data = pd.read_csv(file_path)
    print(f"Datos cargados con éxito, columnas: {data.columns}")
except FileNotFoundError:
    print(f"El archivo en la ruta {file_path} no se encontró.")
    exit(1)

# Convertir la columna de fechas a formato datetime
data['Calendarios Start Date and Time'] = pd.to_datetime(data['Calendarios Start Date and Time'])

# Obtener la fecha de ejecución actual
fecha_actual = datetime.now()

# Calcular el lunes pasado (hace 7 días) a las 7 a.m.
def obtener_lunes_pasado(fecha_actual):
    # Encontrar el lunes más reciente
    dia_semana = fecha_actual.weekday()
    # Retroceder hasta el lunes pasado (hace 7 días)
    lunes_reciente = fecha_actual - timedelta(days=dia_semana + 7)  # Retroceder al lunes de la semana anterior (+7 días)
    lunes_reciente = lunes_reciente.replace(hour=7, minute=0, second=0, microsecond=0)
    return lunes_reciente

# Calcular el lunes pasado a las 7 a.m.
lunes_pasado_7am = obtener_lunes_pasado(fecha_actual)

# Calcular días festivos dentro del rango (desde hace 7 días hasta hoy)
dias_festivos_en_semana = [pd.to_datetime(day) for day in dias_festivos if lunes_pasado_7am <= pd.to_datetime(day) <= fecha_actual]
festivos = len(dias_festivos_en_semana)

# Función para calcular el objetivo semanal ajustado
def calcular_objetivo_semanal(persona, festivos): 
    if persona in ["Randy Vizcaino", "Kattya Cordero"]:
        objetivo = 12
    else:
        objetivo = 10
    objetivo_ajustado = objetivo - (2 * festivos)
    return max(objetivo_ajustado, 0)  # El objetivo no puede ser negativo

# Definir los objetivos por persona  
personas = ["Andrea Sanchez", "Kattya Cordero", "Natalia Giraldo", "Randy Vizcaino"]
objetivos = {persona: calcular_objetivo_semanal(persona, festivos) for persona in personas}

# Filtrar los datos desde el lunes pasado a las 7 a.m.
try:
    filtered_data = data[
        (data['Calendarios Start Date and Time'] >= lunes_pasado_7am) &
        (data['Calendarios Estado actividad'] == 'Realizada') &
        (data['Calendarios Asignado a'].isin(personas))
    ]
    print(f"Filtrado realizado con éxito, número de registros: {len(filtered_data)}")
except KeyError as e:
    print(f"Error en la filtración: {e}")
    exit(1)

# Verificar si hay datos después del filtrado
if filtered_data.empty:
    print("No hay datos para graficar en el período seleccionado.")
    exit(1)

# Crear tabla resumida de actividades por persona
try:
    tabla_resumen = filtered_data.groupby(['Calendarios Asignado a', 'Calendarios Tipo de actividad']).size().unstack(fill_value=0)
    tabla_resumen['Total'] = tabla_resumen[['Reunión', 'Video Conferencia']].sum(axis=1)  # Sumar Reunión y Video Conferencia
    tabla_resumen['Total general'] = tabla_resumen.sum(axis=1)
    tabla_resumen['Objetivo Semanal'] = tabla_resumen.index.map(lambda x: objetivos.get(x, 10))
    tabla_resumen['Festivos'] = festivos  # Agregar columna de festivos
    print("Tabla resumen creada con éxito")
except Exception as e:
    print(f"Error al crear la tabla resumen: {e}")
    exit(1)

# Verificar las columnas disponibles para graficar
print(f"Columnas disponibles para graficar: {tabla_resumen.columns}")

# Formatear las fechas para el título
fecha_inicio = lunes_pasado_7am.strftime('%d %b')
fecha_fin = fecha_actual.strftime('%d %b')


# Crear la gráfica solo si las columnas necesarias existen
grafica_path = os.path.join(save_path, "Grafica_actividades_Semana.png")
try:
    # Verificar si hay columnas numéricas disponibles para graficar
    columnas_graficar = ['Reunión', 'Video Conferencia', 'Objetivo Semanal', 'Total']
    columnas_existentes = [col for col in columnas_graficar if col in tabla_resumen.columns]

    if columnas_existentes:
        fig, ax = plt.subplots(figsize=(10, 6))  # Ajusta el tamaño de la figura para evitar recortes
        colores = ['blue', 'gray', 'orange', 'green']  # Colores para las barras
        tabla_resumen[columnas_existentes].plot(kind='bar', stacked=False, ax=ax, color=colores)

        # Añadir los valores encima de cada barra
        for p in ax.patches:
            ax.annotate(str(int(p.get_height())), 
                        (p.get_x() + p.get_width() / 2., p.get_height()), 
                        ha='center', va='center', xytext=(0, 10), 
                        textcoords='offset points', color='black')

        # Anotar días festivos en la parte superior izquierda de la gráfica
        for dia_festivo in dias_festivos_en_semana:
            ax.annotate(f'Festivo {dia_festivo.strftime("%Y-%m-%d")}',
                        xy=(0, 1), xycoords='axes fraction',  # Coordenadas para la esquina superior izquierda
                        xytext=(10, -10),
                        textcoords='offset points',
                        ha='left', va='top',
                        arrowprops=dict(arrowstyle='->', lw=1.5),
                        color='red',
                        fontsize=8)

        # Ajustar el título de la gráfica para incluir las fechas
        ax.set_title(f'Actividades Semanales del {fecha_inicio} al {fecha_fin} con {festivos} Día(s) Festivo(s)')
        plt.xticks(rotation=45, ha='right')  # Rotar etiquetas para evitar que se corten

        # Ajustar la leyenda (fuera de la gráfica, en la parte superior derecha)
        ax.legend(loc='upper right', fontsize='small', title="Actividades", title_fontsize='small',
                  bbox_to_anchor=(1.15, 1))  # Leyenda fuera de la gráfica

        plt.tight_layout()  # Ajusta automáticamente los márgenes de la gráfica
        plt.savefig(grafica_path)
        plt.close()
        print(f"Gráfica guardada en: {grafica_path}")
    else:
        print("No se encontraron columnas adecuadas para graficar.")
except Exception as e:
    print(f"Error al generar la gráfica: {e}")
    exit(1)

# Guardar la tabla
tabla_path = os.path.join(save_path, "tabla_actividades_semana.xlsx")
try:
    with pd.ExcelWriter(tabla_path, engine='xlsxwriter') as writer:
        tabla_resumen.to_excel(writer, sheet_name='Resumen Actividades')
    print(f"Tabla guardada en: {tabla_path}")
except Exception as e:
    print(f"Error al guardar la tabla: {e}")
    exit(1)
