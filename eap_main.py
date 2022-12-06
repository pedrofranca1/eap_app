import streamlit as st
import eap_roadmap
import warnings
import pandas as pd
import numpy as np

############################################################################### 

st.set_page_config(layout="wide")
st.title("EAP Mero 2")

     
uploaded_file = st.file_uploader('Choose a file', type='xlsx')



if uploaded_file is not None:
    st.header("Sum results")
    # Input Excel
    st.session_state.resultados = eap_roadmap.roadmap(uploaded_file) 
    
    # Cria Colunas
    df = pd.DataFrame(st.session_state.resultados.items(), columns=['LEVEL', 'SUM']) 
    df[['SUM', 'STATUS']] = df['SUM'].str.split(", ", 1, expand=True)
    df[['LEVEL', 'DESCRIPTION']] = df['LEVEL'].str.split(' -- ', 1, expand=True)   
# =============================================================================
#     hide_dataframe_row_index = """
#             <style>
#             .row_heading.level0 {display:none}
#             .blank {display:none}
#             </style>
#             """
#     st.markdown(hide_dataframe_row_index, unsafe_allow_html=True)
# =============================================================================

    option = st.sidebar.selectbox(
        "Filter table",
        (" ", "CORRECT SUM", "INCORRECT SUM")
    )
    if option == " ":       
        st.dataframe(df.iloc[:,[0,3,1,2]].style.applymap(eap_roadmap.color_background, subset=['STATUS']))
    elif option == "INCORRECT SUM":
        df.query("STATUS == 'INCORRECT SUM'", inplace = True)
        st.dataframe(df.iloc[:,[0,3,1,2]].style.applymap(eap_roadmap.color_background, subset=['STATUS']))  
    else: 
        df.query("STATUS == 'CORRECT SUM'", inplace = True)
        st.dataframe(df.iloc[:,[0,3,1,2]].style.applymap(eap_roadmap.color_background, subset=['STATUS']))  
else:
    st.warning('you need to upload a excel file.')
    add_selectbox = st.sidebar.selectbox(
        "Filter table",
        (" ","CORRECT SUM", "INCORRECT SUM"),
        disabled = True,
        label_visibility = "hidden"
    )

    
    
with open("CSS\main.css" ) as f:
   st.markdown(f"<style>{f.read()}</style>",unsafe_allow_html = True)
# =============================================================================
# st.header("Results")
# for v in st.session_state.results:
#      st.write(v)
# =============================================================================
  






# =============================================================================
# with st.sidebar:
#     add_radio = st.radio(
#         "Choose a shipping method",
#         ("Standard (5-15 days)", "Express (2-5 days)")
#     )
# 
# =============================================================================

###############################################################################    



