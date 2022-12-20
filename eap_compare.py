import openpyxl
import datetime




def write_to_excel(file1, file2, date):
    #Abrir arquivo Excel
    workbook = openpyxl.load_workbook(file1, data_only=True)
    workbook2 = openpyxl.load_workbook(file2, data_only=True)

    # Seleciona a guia
    ws = workbook.worksheets[0]
    ws2 = workbook2.worksheets[0]

    # Cor célula
    my_yellow = openpyxl.styles.colors.Color(rgb='FFFF00')
    my_fill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_yellow)

    # Obter data do primeiro dia do mês atual
    finalDate = datetime.datetime.strptime(date, "%b%Y").strftime("%Y-%m-%d")

    # Index da coluna
    for row in ws.iter_rows():
        for cell in row:          
            if str(finalDate) in str(cell.value):
               index = cell.column - 2
               indexMonth = cell.column - 1
            elif cell.value == 'Ref.':
               indexRef = cell.column - 1


    # Criar dicionário com Excel do mês anterior
    num = 1
    key = 'Forecast ' + str(num)
    forecast = {key: []}

    for row in ws2.iter_rows():
        if row[indexRef].value == 'Forecast':
            for cell in row:
                if cell.column > indexRef and cell.value != None and cell.value != 0:
                    key = 'Forecast ' + str(num)
                    if key not in forecast:
                        forecast[key] = []
                    forecast[key].append(cell.value)
            num += 1

    size = len(forecast)

    # Substituir valores das keys por números
    numForecast = {}

    for i, key in enumerate(forecast):
      # add the key and value to numbered_dict
      numForecast[i] = forecast[key] 
                   
    # Loop pintar célula   
    num2 = 0
    num3 = 2
    for row in ws.iter_rows():
        if row[indexRef].value == 'Baseline 1' and row[index].value != row[indexMonth].value: #verificar se é Baseline ou Actual
           row[indexMonth].fill = my_fill
        elif row[indexRef].value == 'Forecast':
           forecastList = numForecast[num2]                      
           for cell in row:               
               if num3 <= len(forecastList) - 1 and cell.column > indexMonth and cell.value != None and cell.value != 0 and cell.value != forecastList[num3]:                               
                   cell.fill = my_fill   
                   num3 += 1
        num2 += 1
        num3 = 2
        if num2 > size:
            break

    # Salvar o arquivo Excel
    #workbook.save(file1)
    return ws
    #return ws
    print("COMPARAÇÃO REALIZADA")
    

# file1 = 'Exemplos\\file1.xlsx'
# file2 = 'Exemplos\\file2.xlsx'
# date = 'Oct2022'

# write_to_excel(file1, file2, date)


  
