import streamlit as st
import pandas as pd

# Frontend 
st.title('Datos Excel Combinados')
st.image('Imagen\logo.jpg', clamp=True) 

uploaded_files = st.file_uploader("Cargar archivos Excel", accept_multiple_files=True, type='xlsx')

if uploaded_files:
  dfs = []
  for file in uploaded_files: 
    df = pd.read_excel(file)
    dfs.append(df)
  
  data = pd.concat(dfs, ignore_index=True)

  # Crear un nuevo archivo Excel 
  new_file = 'nuevo_archivo.xlsx'
  sheet_name = 'Hoja1'

  with pd.ExcelWriter(new_file) as writer: 
    data.to_excel(writer, sheet_name=sheet_name, index=False)

  # Opción de descarga
  with open(new_file, 'rb') as f:
    btn = st.download_button(label='Descargar datos combinados', data=f, file_name=new_file)

else:
  st.info('Por favor sube archivos Excel')
