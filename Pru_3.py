import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import pip
pip.main(["install", "openpyxl"])

# Definir ruta del archivo Excel existente 
excel_file_path = ''

# Frontend
st.title('Adjuntar archivos Excel')
# Agregar una imagen desde una ruta de archivo local
st.image('GitHub\Imagen\logo.jpg', caption='Cargar Informacion', clamp=True)

# Botón para enviar la información cargada
submitted = st.button('Enviar')

# Permitir al usuario cargar múltiples archivos Excel adjuntos
uploaded_files = st.file_uploader("Cargar archivos Excel adjuntos", accept_multiple_files=True, type='xlsx')

if uploaded_files:
    try:
        # Lista para almacenar los DataFrames de los archivos cargados
        dfs = []

        # Iterar sobre cada archivo cargado
        for uploaded_file in uploaded_files:
            # Leer el archivo Excel adjunto en un DataFrame
            df_new_data = pd.read_excel(uploaded_file)

            # Agregar DataFrame a la lista
            dfs.append(df_new_data)

        # Concatenar todos los DataFrames en uno solo
        all_data = pd.concat(dfs, ignore_index=True)

        # Ruta del archivo Excel existente
        excel_file_path = r'C:\Users\usuario\Desktop\Python_Streamlist\GitHub\Data.xlsx'

        # Nombre de la hoja de cálculo
        sheet_name = 'Hoja1'

        # Leer datos existentes
        existing_data = pd.read_excel(excel_file_path, sheet_name=sheet_name)

        # Agregar los nuevos datos a los datos existentes
        all_data = pd.concat([existing_data, all_data], ignore_index=True)

        # Escribir datos desde fila 1
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            all_data.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)

        # Mostrar mensaje de éxito si se hizo clic en el botón "Enviar"
        if submitted:
            st.success('Datos adjuntos agregados exitosamente!')

    except FileNotFoundError:
        st.error("No se encontró el archivo Excel.")
  
    except Exception as e:
        st.error("Ocurrió un error: " + str(e))

else:
    st.info("Por favor, carga archivos Excel para agregar datos adjuntos.")
