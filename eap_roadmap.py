import pandas as pd
import warnings
import numpy as np
import math


def roadmap(file):
    # Ignora avisos de compilação
    warnings.simplefilter("ignore")

    #Define caminho aonde esta o arquivo    

    # Lê a tabela  
    dataframe = pd.read_excel(file, 0)
    
    # Encontra o index da primeira linha da tabela
    valor_index = dataframe.index[dataframe['MERO 2 - ESTRUTURA ANALÍTICA DE PROJETO'] == "DESCRIÇÃO"][0]

    # Apaga as linhas de fora da tabela
    dataframe.drop(dataframe.index[0:valor_index],axis=0,inplace=True)

    # Renomeia as colunas
    dataframe.rename(columns=dataframe.iloc[0], inplace = True)
   
    # Reseta o index da tabela
    dataframe.reset_index(drop=True,inplace=True)

    # Apaga a primeira linha da tabela (já que ela está duplicada no cabeçalho)
    dataframe.drop([0, 1],axis=0,inplace=True)
    dataframe.reset_index(drop=True,inplace=True)
     
    # Preencher coluna 'ITEM'
    dataframe['ITEM'] = dataframe['ITEM'].replace(np.nan, 0)

    index = 0
    
    for row in dataframe.itertuples():  
        if row[7] != 0:
           temp = row[7]
           number = temp[1:]
           if temp.find(" ") != -1:
              number = number.split(" ", 1)[0]
           if temp.find(".") != -1: 
              numberList = number.split(".")
        elif row[7] == 0 and math.isnan(row[13]) == False:
            number = temp[1:]
            if temp.find(".") != -1: 
               numberList = number.split(".")
        elif row[7] == 0 and math.isnan(row[14]) == False:
            number = temp[1:]
            if temp.find(".") != -1: 
               numberList = number.split(".")
        elif row[7] == 0 and math.isnan(row[15]) == False:
            number = temp[1:]
            if temp.find(".") != -1: 
               numberList = number.split(".")

        #Transformação Nivel 5:
        if math.isnan(row[13]) == False:      
            if number.count(".") == 1:
                finalTemp = temp[0] + number + "." + "1"
                dataframe.iloc[index, 6] = finalTemp
                temp = finalTemp
            elif number.count(".") == 2:
                newNumber = int(numberList[2]) + 1
                finalTemp = temp[0] + numberList[0] + "." + numberList[1] + "." + str(newNumber)
                dataframe.iloc[index, 6] = finalTemp
                temp = finalTemp  
            elif number.count(".") >= 2:
                newNumber = int(numberList[2]) + 1
                finalTemp = temp[0] + numberList[0] + "." + numberList[1] + "." + str(newNumber)
                dataframe.iloc[index, 6] = finalTemp
                temp = finalTemp 
        
        #Transformação Nivel 6:
        if math.isnan(row[14]) == False:
             if number.count(".") == 2:
                 finalTemp = temp[0] + number + "." + "1"
                 dataframe.iloc[index, 6] = finalTemp
                 temp = finalTemp
             elif number.count(".") == 3:
                 newNumber = int(numberList[3]) + 1
                 finalTemp = temp[0] + numberList[0] + "." + numberList[1] + "." + numberList[2] + "." + str(newNumber)
                 dataframe.iloc[index, 6] = finalTemp
                 temp = finalTemp
             elif number.count(".") >= 3:
                 newNumber = int(numberList[3]) + 1
                 finalTemp = temp[0] + numberList[0] + "." + numberList[1] + "." + numberList[2] + "." + str(newNumber)
                 dataframe.iloc[index, 6] = finalTemp
                 temp = finalTemp
            
        #Transformação Nivel 7:    
        if math.isnan(row[15]) == False:
             if number.count(".") == 3:
                 finalTemp = temp[0] + number + "." + "1"
                 dataframe.iloc[index, 6] = finalTemp
                 temp = finalTemp
             elif number.count(".") == 4:
                 newNumber = int(numberList[4]) + 1
                 finalTemp = temp[0] + numberList[0] + "." + numberList[1] + "." + numberList[2] + "." + numberList[3] + "." + str(newNumber)
                 dataframe.iloc[index, 6] = finalTemp
                 temp = finalTemp
         
        index += 1
       

    # Cria Dicionários
    nivel2 = {}
    nivel3 = {}
    nivel4 = {}
    nivel5 = {}
    nivel6 = {}
    nivel7 = {}

    # Loop linha-a-linha
    for row in dataframe.itertuples():
        if str(row[7]).count(".") == 0 and len(str(row[7])) == 1:
            if math.isnan(row[10]) == False:
                nivel2.update({row[7] : [row[10], row[8]]})
        elif str(row[7]).count(".") == 0 and len(str(row[7])) > 1:
            if math.isnan(row[11]) == False:
                nivel3.update({row[7] : [row[11], row[8]]}) 
        elif str(row[7]).count(".") == 1:
            if math.isnan(row[12]) == False:
                nivel4.update({row[7] : [row[12], row[8]]}) 
        elif str(row[7]).count(".") == 2:
             if math.isnan(row[13]) == False:
                nivel5.update({row[7] : [row[13], row[8]]}) 
        elif str(row[7]).count(".") == 3:
             if math.isnan(row[14]) == False:
                nivel6.update({row[7] : [row[14], row[8]]}) 
        elif str(row[7]).count(".") == 4:
             if math.isnan(row[15]) == False:
                nivel7.update({row[7] : [row[15], row[8]]}) 

    #Outputs nivel 2
    resultados = {}

        
    if sum(nivel2[item][0] for item in nivel2) != 1:
         resultados.update({ " & ".join(nivel2.keys()) + " -- " + str(" & ".join(nivel2[item][1] for item in nivel2)): str(sum(nivel2[item][0] for item in nivel2)) + ", INCORRECT SUM"})
    else:
         resultados.update({ " & ".join(nivel2.keys()) + " -- " + str(" & ".join(nivel2[item][1] for item in nivel2)): str(sum(nivel2[item][0] for item in nivel2)) + ", CORRECT SUM"})
        
        
    #Outputs nivel 3

    key1 = list(nivel3.keys())[0][0]
    value1 = nivel3[list(nivel3.keys())[0]][0]
    index = 0

    for k, v in nivel3.items():
       if k[0] == key1 and index == 0:
            soma1 = float(value1)
            soma = soma1
            if 1.01 >= soma >= 0.99:
               resultados.update({k[0] + " -- " + v[1]: str(soma) + ", CORRECT SUM"})
            else:
               resultados.update({k[0] + " -- " + v[1]: str(soma) + ", INCORRECT SUM"})
       elif k[0] == key1:
            soma = soma + float(v[0])
       elif k[0] != key1:
            if 1.01 >= soma >= 0.99:
               resultados.update({k[0] + " -- " + v[1]: str(soma) + ", CORRECT SUM"})
            else:
               resultados.update({k[0] + " -- " + v[1]: str(soma) + ", INCORRECT SUM"})
            temp = k[0]
            key1 = temp      
            soma = float(v[0])
       index += 1
            

            
    #Outputs nivel 4

    key1 = list(nivel4.keys())[0].partition('.')[0]
    value1 = nivel4[list(nivel4.keys())[0]][0]
    index = 0


    for k, v in nivel4.items():
       if k.partition('.')[0] == key1 and index == 0:
            soma1 = float(value1)
            soma = soma1
       elif k.partition('.')[0] == key1:
            soma = soma + float(v[0])
            key0 = k.partition('.')[0]
       elif k.partition('.')[0] != key1:
            if 1.01 >= soma >= 0.99:
               resultados.update({key0 + " -- " + v[1]: str(soma) + ", CORRECT SUM"})        
            else:
               resultados.update({key0 + " -- " + v[1]: str(soma) + ", INCORRECT SUM"})
            temp = k.partition('.')[0]
            key1 = temp      
            soma = float(v[0])        
       index += 1
           


    #Outputs nivel 5

    key1 = '.'.join(list(nivel5.keys())[0].split(".",2)[:2])
    value1 = nivel5[list(nivel5.keys())[0]][0]
    index = 0
      
    for k, v in nivel5.items():
       if '.'.join(k.split(".",2)[:2]) == key1 and index == 0:
            soma1 = float(value1)
            soma = soma1
       elif '.'.join(k.split(".",2)[:2]) == key1:
            soma = soma + float(v[0])
            key0 = '.'.join(k.split(".",2)[:2])
       elif '.'.join(k.split(".",2)[:2]) != key1:
            if 1.01 >= soma >= 0.99:
               resultados.update({key0 + " -- " + v[1]: str(soma) + ", CORRECT SUM"}) 
            else:
               resultados.update({key0 + " -- " + v[1]: str(soma) + ", INCORRECT SUM"}) 
            temp = '.'.join(k.split(".",2)[:2])
            key1 = temp      
            soma = float(v[0])        
       index += 1
             
           
           
    #Outputs nivel 6

    key1 = '.'.join(list(nivel6.keys())[0].split(".",3)[:3])
    value1 = nivel6[list(nivel6.keys())[0]][0]
    index = 0
      
    for k, v in nivel6.items():
       if '.'.join(k.split(".",3)[:3]) == key1 and index == 0:
            soma1 = float(value1)
            soma = soma1
       elif '.'.join(k.split(".",3)[:3]) == key1:
            soma = soma + float(v[0])
            key0 = '.'.join(k.split(".",3)[:3])
       elif '.'.join(k.split(".",3)[:3]) != key1:
            if 1.01 >= soma >= 0.99:
               resultados.update({key0 + " -- " + v[1]: str(soma) + ", CORRECT SUM"}) 
            else:
               resultados.update({key0 + " -- " + v[1]: str(soma) + ", INCORRECT SUM"}) 
            temp = '.'.join(k.split(".",3)[:3])
            key1 = temp      
            soma = float(v[0])        
       index += 1
           

    #Outputs nivel 7    
           
    key1 = '.'.join(list(nivel7.keys())[0].split(".",4)[:4])
    value1 = nivel7[list(nivel7.keys())[0]][0]
    index = 0
      
    for k, v in nivel7.items():
       if '.'.join(k.split(".",4)[:4]) == key1 and index == 0:
            soma1 = float(value1)
            soma = soma1
       elif '.'.join(k.split(".",4)[:4]) == key1:
            soma = soma + float(v[0])
            key0 = '.'.join(k.split(".",4)[:4])
       elif '.'.join(k.split(".",4)[:4]) != key1:
            if 1.01 >= soma >= 0.99:
               resultados.update({key0 + " -- " + v[1]: str(soma) + ", CORRECT SUM"}) 
            else:
               resultados.update({key0 + " -- " + v[1]: str(soma) + ", INCORRECT SUM"}) 
            temp = '.'.join(k.split(".",4)[:4])
            key1 = temp      
            soma = float(v[0])        
       index += 1
       
    return resultados

# file1 = 'Exemplos\\file1.xlsx'
# roadmap(file1)

def color_background(val):
    color = 'rgba(203,44,48,1)' if val=='INCORRECT SUM' else 'white'
    return f'background-color: {color}'
