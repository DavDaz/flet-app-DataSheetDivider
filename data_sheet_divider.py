# Combined_script.py
import os
import pandas as pd
from dotenv import load_dotenv
import flet as ft

# Cargar variables de entorno desde el archivo .env
load_dotenv()

def exportar_ventanas_xlsx(file_path):
    # Cargar el archivo xlsx
    xls = pd.ExcelFile(file_path)
    
    # Obtener el nombre del archivo sin la extensión
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    
    # Crear una carpeta para almacenar los resultados
    folder_name = f"resultados-{file_name}"
    os.makedirs(folder_name, exist_ok=True)
    
    # Iterar sobre todas las hojas en el archivo xlsx
    for sheet_name in xls.sheet_names:
        # Leer cada hoja en un DataFrame
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Generar el nombre de archivo exportado
        export_file_name = f"{file_name}-{sheet_name}.xlsx"
        export_file_path = os.path.join(folder_name, export_file_name)
        
        # Exportar la ventana a un archivo xlsx en la carpeta de resultados
        df.to_excel(export_file_path, index=False)

    print("Se han exportado las ventanas por xlsx separados.")

def main(page):
    def btn_click(e):
        if not input_excel.value or not input_column.value or not input_carpet.value:
            input_excel.error_text = "Por favor ingresa todos los valores"
            page.update()
        else:
            archivo_excel = input_excel.value
            nombre_sede_column = input_column.value
            carpeta_resultado = input_carpet.value

            # Verificar si el archivo Excel existe
            if not os.path.isfile(archivo_excel):
                page.clean()
                page.add(ft.Text("El archivo Excel especificado no se encuentra. Por favor, verifique la ruta y vuelva a intentarlo."))
            else:
                # Leer el archivo .xlsx
                df = pd.read_excel(archivo_excel)

                # Verificar si el nombre de la columna existe
                if nombre_sede_column not in df.columns:
                    page.clean()
                    page.add(ft.Text("El nombre de la columna especificado no existe en el archivo Excel. Por favor, verifique el nombre de la columna y vuelva a intentarlo."))
                else:
                    # Obtener una lista de todos los corregimientos únicos
                    corregimientos_unicos = df[nombre_sede_column].unique()

                    # Crear la carpeta 'resultado' si no existe
                    if not os.path.exists(carpeta_resultado):
                        os.makedirs(carpeta_resultado)

                    # Crear un nuevo archivo Excel en la carpeta 'resultado' y guardar cada bloque de datos filtrado por corregimiento en una hoja diferente
                    ruta_resultado = os.path.join(carpeta_resultado, archivo_excel)
                    with pd.ExcelWriter(ruta_resultado) as writer:
                        for corregimiento in corregimientos_unicos:
                            df_corregimiento = df[df[nombre_sede_column] == corregimiento]
                            df_corregimiento.to_excel(writer, sheet_name=corregimiento, index=False)

                    print("Se han exportado los datos filtrados por corregimiento en hojas diferentes en la carpeta 'resultado'.")

                    # Ruta del archivo xlsx
                    file_path = os.path.join(carpeta_resultado, archivo_excel)

                    # Llamar a la función para exportar las ventanas
                    exportar_ventanas_xlsx(file_path)

                    # Mostrar mensaje de completado
                    page.clean()
                    page.add(ft.Text("¡Proceso completado! Las ventanas se han exportado por separado en archivos Excel en la carpeta 'resultado'."))

    input_excel = ft.TextField(label="Nombre del Archivo Excel")
    input_column = ft.TextField(label="Nombre de la Columna")
    input_carpet = ft.TextField(label="Nombre de la carpeta para guardar el resultado")

    page.add(input_excel, input_column, input_carpet, ft.ElevatedButton("Ejecutar Proceso!", on_click=btn_click))

ft.app(target=main)
