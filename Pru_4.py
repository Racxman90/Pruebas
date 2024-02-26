import streamlit as st
import pandas as pd
import base64
import io

# Frontend
st.title('Combinar Archivos Excel')
st.sidebar.image(r'Imagen\logo.jpg', use_column_width=True)

# Subheader para la sección de carga de archivos
st.sidebar.subheader("Cargar Archivos Excel")
st.sidebar.write("Utilice el botón a continuación para cargar archivos Excel:")

# Botón para cargar archivos
uploaded_files = st.sidebar.file_uploader("Seleccionar archivos Excel", 
                                          accept_multiple_files=True, 
                                          type='xlsx')

if uploaded_files:
    # Contenedor para mostrar los nombres de los archivos cargados
    st.sidebar.subheader("Archivos Cargados:")
    for file in uploaded_files:
        st.sidebar.write(file.name)

    try:
        # Procesar los archivos cargados
        dfs = []
        for file in uploaded_files:
            df = pd.read_excel(file)
            dfs.append(df)

        # Concatenar los DataFrames
        data = pd.concat(dfs, ignore_index=True)

        # Leer el archivo Excel existente, si existe
        existing_data = pd.DataFrame()
        excel_file_path = 'Data.xlsx'  # Ruta relativa al archivo Excel
        if st.sidebar.checkbox("¿Tiene un archivo existente para combinar?"):
            existing_data = pd.read_excel(excel_file_path)

        # Concatenar los nuevos datos al final de los datos existentes
        all_data = pd.concat([existing_data, data], ignore_index=True)

        # Eliminar espacios en blanco de los nombres de las columnas
        all_data.columns = all_data.columns.str.strip()

        # Eliminar espacios en blanco de los datos en las columnas
        all_data = all_data.applymap(lambda x: x.strip() if isinstance(x, str) else x)

        # Reorganizar las columnas en el orden deseado: Nombre, Mail, Edad
        all_data = all_data[['Nombre', 'Email', 'Edad']]

        # Crear un objeto ExcelWriter
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            all_data.to_excel(writer, index=False, sheet_name='Hoja1')

        processed_data = output.getvalue()

        # Generar el enlace de descarga
        b64 = base64.b64encode(processed_data).decode()
        btn = f'<a href="data:application/octet-stream;base64,{b64}" download="Nuevo_Archivo.xlsx">Descargar Archivo Combinado</a>'
        st.sidebar.markdown(btn, unsafe_allow_html=True)

        # Mostrar una vista previa de los datos
        st.subheader("Vista Previa de los Datos:")
        st.write(all_data)

    except FileNotFoundError:
        st.error("No se encontró el archivo Excel existente.")

    except Exception as e:
        st.error(f"Ocurrió un error: {str(e)}")
else:
    st.info('Por favor, cargue archivos Excel para combinar.')
