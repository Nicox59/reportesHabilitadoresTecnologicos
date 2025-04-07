import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Border, Side

def listar_archivos_csv(carpeta):
    """Obtiene la lista de archivos CSV en la carpeta especificada."""
    return [f for f in os.listdir(carpeta) if f.endswith(".csv")]

def leer_csv(ruta_archivo):
    """Lee un archivo CSV y devuelve un DataFrame."""
    try:
        return pd.read_csv(ruta_archivo, encoding="utf-8", on_bad_lines="skip")
    except Exception as e:
        print(f"Error al leer {ruta_archivo}: {e}")
        return None

def extraer_carrera(nombre_archivo):
    """Extrae la carrera desde el nombre del archivo."""
    partes = nombre_archivo.split("_")
    return partes[4][:4] if len(partes) > 4 else "N/A"

def extraer_sede(ruta_archivo):
    """Extrae la sede desde la carpeta donde está el archivo."""
    return os.path.basename(os.path.dirname(ruta_archivo))

def procesar_datos(df, nombre_archivo, ruta_archivo, columnas_a_eliminar):
    """Elimina columnas, agrega 'Carrera' y 'Sede', renombra columnas y agrega 'Estado'."""
    try:
        carrera = extraer_carrera(nombre_archivo)
        sede = extraer_sede(ruta_archivo)  # Obtener sede desde la ruta
        
        df.insert(4, "Carrera", carrera)  # Inserta 'Carrera' en la posición 4
        df.insert(5, "Sede", sede)  # Inserta 'Sede' justo después de 'Carrera'
        df = df.drop(df.columns[columnas_a_eliminar], axis=1)

        # Renombrar las columnas de índices 6 a 10 (después del ajuste de índices por 'Carrera' y 'Sede')
        nuevos_nombres = {
            df.columns[6]: "Cuestionario de Habilitación - Internet de las cosas",
            df.columns[7]: "Cuestionario de Habilitación - Robótica",
            df.columns[8]: "Cuestionario de Habilitación - Fabricación 3D",
            df.columns[9]: "Cuestionario de Habilitación - Inteligencia Artificial",
            df.columns[10]: "Cuestionario de Habilitación - Realidad Virtual y Aumentada",
        }
        df = df.rename(columns=nuevos_nombres)

        # Agregar la columna 'Estado' según las notas de los cuestionarios
        def evaluar_estado(filas):
            for valor in filas:
                if pd.notna(valor) and valor > 4.0:
                    return "Habilitado"
            return "No habilitado"

        df["Estado"] = df.iloc[:, 6:11].apply(evaluar_estado, axis=1)  # Ajustado por 'Carrera' y 'Sede'
        return df
    except Exception as e:
        print(f"Error al procesar los datos: {e}")
        return df

def guardar_como_excel(df, ruta_guardado):
    """Guarda el DataFrame como archivo Excel."""
    try:
        with pd.ExcelWriter(ruta_guardado, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Datos")
        print(f" Archivo guardado: {ruta_guardado}")
    except Exception as e:
        print(f" Error al guardar el archivo Excel: {e}")

def procesar_csvs_y_guardar_excel(carpeta):
    archivos_csv = listar_archivos_csv(carpeta)
    archivos_generados = []

    if not archivos_csv:
        print("⚠️ No se encontraron archivos CSV en la carpeta.")
        return []
    
    for archivo in archivos_csv:
        ruta_archivo = os.path.join(carpeta, archivo)
        print(f"\n Procesando archivo: {archivo}")
        
        datos = leer_csv(ruta_archivo)
        if datos is not None:
            columnas_a_eliminar = [6, 7, 8, 9, 10]  # Ajustado por la nueva posición de columnas
            datos = procesar_datos(datos, archivo, ruta_archivo, columnas_a_eliminar)
            
            ruta_guardado = os.path.join(carpeta, archivo.replace(".csv", "_modificado.xlsx"))
            guardar_como_excel(datos, ruta_guardado)
            archivos_generados.append(ruta_guardado)

    return archivos_generados

def unir_excels_y_guardar(archivos_excel, carpeta):
    """Une todos los archivos Excel generados en un solo archivo con una hoja."""
    lista_df = []

    for archivo in archivos_excel:
        try:
            df = pd.read_excel(archivo)
            lista_df.append(df)
        except Exception as e:
            print(f"Error al leer {archivo}: {e}")

    if lista_df:
        df_final = pd.concat(lista_df, ignore_index=True)  # Une archivos en filas
        
        # Crear la tabla resumen por carrera y sede
        resumen = df_final.groupby('Carrera')['Estado'].value_counts().unstack(fill_value=0)
        
        # Crear una nueva tabla con la estructura solicitada
        tabla_resumen = pd.DataFrame({
            'Carrera': resumen.index.get_level_values(0),
            'Habilitados': resumen['Habilitado'],
            'No habilitados': resumen['No habilitado']
        })

        # Añadir la fila total
        total = tabla_resumen[['Habilitados', 'No habilitados']].sum()
        total['Carrera'] = 'Total'
        tabla_resumen = pd.concat([tabla_resumen, total.to_frame().T], ignore_index=True)

        # Guardar el archivo consolidado
        ruta_consolidado = os.path.join(carpeta, "consolidado.xlsx")
        
        with pd.ExcelWriter(ruta_consolidado, engine="openpyxl") as writer:
            # Guardar los datos principales en la hoja "Datos"
            df_final.to_excel(writer, index=False, sheet_name="Datos")

            # Guardar la tabla resumen comenzando en la columna N y fila 8 (N8 en Excel)
            tabla_resumen.to_excel(writer, index=False, header=True, startrow=7, startcol=13, sheet_name="Datos")
            
            # Obtener el libro de trabajo y la hoja activa
            workbook = writer.book
            sheet = workbook["Datos"]
            
            # Convertir el rango de los datos en una tabla (solo para los datos principales)
            tabla_range = f"A1:{chr(64+len(df_final.columns))}{len(df_final)+1}"
            tabla_principal = Table(displayName="TablaPrincipal", ref=tabla_range)
            tabla_principal.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=False, showColumnStripes=False
            )
            sheet.add_table(tabla_principal)
            
            # Obtener el rango de la tabla resumen que empieza en N8
            min_row = 8  # Fila 8
            min_col = 14  # Columna N
            max_row = min_row + len(tabla_resumen)  # Número total de filas de la tabla de resumen
            max_col = 16  # Columna P

            # Aplicar bordes a todo el rango de la tabla de resumen (incluyendo "Total")
            for row in sheet.iter_rows(min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col):
                for cell in row:
                    cell.border = Border(
                        left=Side(border_style="thin", color="000000"),
                        right=Side(border_style="thin", color="000000"),
                        top=Side(border_style="thin", color="000000"),
                        bottom=Side(border_style="thin", color="000000")
                    )

        print(f"Archivo consolidado guardado en: {ruta_consolidado}")

# Ejecutar el proceso
carpetas = [
    r"C:\Users\LabiPC2\Documents\Leer archivos\archivos\BES",
    r"C:\Users\LabiPC2\Documents\Leer archivos\archivos\TPC",
    r"C:\Users\LabiPC2\Documents\Leer archivos\archivos\PAP"
]

archivos_generados_totales = []
for carpeta in carpetas:
    archivos_generados_totales.extend(procesar_csvs_y_guardar_excel(carpeta))

if archivos_generados_totales:
    unir_excels_y_guardar(archivos_generados_totales, r"C:\Users\LabiPC2\Documents\Leer archivos\archivos")

# Mostrar mensaje de éxito
tk.Tk().withdraw()
messagebox.showinfo("Proceso Completado", "Operación lista")
