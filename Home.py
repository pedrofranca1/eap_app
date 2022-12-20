import streamlit as st
import eap_roadmap
import eap_compare
import pandas as pd
import base64
import io
import components.login as authenticator
import openpyxl


st.set_page_config(page_title="EAP Mero 2", page_icon=":house:", layout="wide") 
with open("css\main.css" ) as f:
   st.markdown(f"<style>{f.read()}</style>",unsafe_allow_html = True)

  
name, authentication_status, username = authenticator.run_authenticator()

if "visibility" not in st.session_state:
    st.session_state.visibility = "visible"
    st.session_state.disabled = False

if st.session_state["authenticated"] == None:
    st.warning("Please enter your username and password")

if st.session_state["authenticated"] == True: 
   st.sidebar.image("./Figures/technip_logo.png", use_column_width=True)
   st.sidebar.title(f"Welcome {name}")   
   
  
   uploaded_file = st.file_uploader('Choose a Mero 2 EAP file', type='xlsx')
   uploaded_file2 = st.file_uploader('Choose a Mero 2 EAP file from last month', type='xlsx')
      
 
   
   if uploaded_file is not None:
       #st.header("Sum Results")
       # Input Excel
       st.session_state.resultados = eap_roadmap.roadmap(uploaded_file) 
       
       # Cria Colunas
       df = pd.DataFrame(st.session_state.resultados.items(), columns=['LEVEL', 'SUM']) 
       df[['SUM', 'STATUS']] = df['SUM'].str.split(", ", 1, expand=True)
       df[['LEVEL', 'DESCRIPTION']] = df['LEVEL'].str.split(' -- ', 1, expand=True)
       
       option = st.sidebar.selectbox(
              "Filter table",
              (" ", "CORRECT SUM", "INCORRECT SUM")
          )
             
       
       optionMonth = st.sidebar.selectbox(
              "Choose EAP Month",
              (" ", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov","Dec"),
              label_visibility=st.session_state.visibility,
              disabled=st.session_state.disabled
          )
       
       
       optionYear = st.sidebar.selectbox(
              "Choose EAP year",
              (" ", "2020", "2021", "2022", "2023", "2024"),
              label_visibility=st.session_state.visibility,
              disabled=st.session_state.disabled
          )
       
      
       
       finalDate = optionMonth + optionYear           
           
             
           
       if option == " ":       
           st.dataframe(df.iloc[:,[0,3,1,2]].style.applymap(eap_roadmap.color_background, subset=['STATUS']))
       elif option == "INCORRECT SUM":
           df.query("STATUS == 'INCORRECT SUM'", inplace = True)
           st.dataframe(df.iloc[:,[0,3,1,2]].style.applymap(eap_roadmap.color_background, subset=['STATUS']))  
       else: 
           df.query("STATUS == 'CORRECT SUM'", inplace = True)
           st.dataframe(df.iloc[:,[0,3,1,2]].style.applymap(eap_roadmap.color_background, subset=['STATUS'])) 
           
   
       if uploaded_file2 is not None:
           print("Upload realizado arquivo 2")
           
           #Compare 2 files:
           check = st.sidebar.checkbox("Date selected", key="disabled")
           
           if check:                                      
              #st.sidebar.markdown("Click [here](uploaded_file) to download the file.")
              #st.sidebar.download_button(label="Click here to download the file.", data=compareFiles, file_name='myresults.xlsx', mime='text/xlsx') 
              
              st.markdown("Click the button below to download the manipulated file")
              wb = openpyxl.load_workbook(uploaded_file)
           
              if st.button("Download"): 
                  ws = wb['EAP Mero 2']  
                  wb.remove(wb['EAP Mero 2'])
                  wb.create_sheet('EAP Mero 2', 0)
                  new_ws = wb['EAP Mero 2']  
                  ws = eap_compare.write_to_excel(uploaded_file, uploaded_file2, finalDate)
                  for row in ws.rows:
                     for cell in row:
                         new_ws[cell.coordinate] = cell.value  
                
                  wb.save(filename='manipulated.xlsx')
                  print("Binário criado")
                  with open('manipulated.xlsx', 'rb') as f:
                      excel_file = f.read()  
                  buf = io.BytesIO(excel_file)
                  excel_file_b64 = base64.b64encode(excel_file).decode('utf-8')
                  st.markdown(f'<a href="data:application/octet-stream;base64,{excel_file_b64}" download="manipulated.xlsx">Download manipulated file</a>', unsafe_allow_html=True)                                                     
       
   else:
       st.warning('you need to upload a excel file.')
       
       optionMonth = st.sidebar.selectbox(
              "Choose EAP Month",
              ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov","Dec"),
              disabled = True,
              label_visibility = "hidden"
          )
       
       optionYear = st.sidebar.selectbox(
              "Choose EAP year",
              ("20", "21", "22", "23", "24"),
              disabled = True,
              label_visibility = "hidden"
          )
       
       add_selectbox = st.sidebar.selectbox(
           "Filter table",
           (" ","CORRECT SUM", "INCORRECT SUM"),
           disabled = True,
           label_visibility = "hidden"
       )
       
   authenticator.button_logout()

       
   

 
 

    



