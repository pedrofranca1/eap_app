import pandas as pd
import warnings
import numpy as np
import math


def create_dataframe_levels(file):
    # Ignora avisos de compilação
    warnings.simplefilter("ignore")

    # Lê a tabela
    dataframe = pd.read_excel(file, 0)

    # Encontra o index da primeira linha da tabela
    columns_with_name = [
        col
        for col in dataframe.columns
        if "ESTRUTURA ANALÍTICA DE PROJETO" in col
    ]
    if len(columns_with_name) > 0:
        column_name = columns_with_name[0]
        valor_index = dataframe[dataframe[column_name] == "DESCRIÇÃO"].index[0]

    # Apaga as linhas de fora da tabela
    dataframe.drop(dataframe.index[0:valor_index], axis=0, inplace=True)

    # Renomeia as colunas
    dataframe.rename(columns=dataframe.iloc[0], inplace=True)

    # Reseta o index da tabela
    dataframe.reset_index(drop=True, inplace=True)

    # Apaga a primeira linha da tabela (já que ela está duplicada no cabeçalho)
    dataframe = dataframe.drop(dataframe.index[0])
    dataframe.reset_index(drop=True, inplace=True)

    # Preencher coluna 'ITEM'
    # dataframe["ITEM"] = dataframe["ITEM"].replace(np.nan, 0)

    # Quantidade de níveis
    index_level_1 = dataframe.columns.get_loc("NÍVEL")
    column_index = dataframe.iloc[0].eq("Valor").values.argmax()
    qtd_levels = int(column_index - index_level_1)
    index_level_item = dataframe.columns.get_loc("ITEM")
    dataframe = dataframe.drop(dataframe.index[0])
    dataframe.reset_index(drop=True, inplace=True)

    # Remove todas as linhas 'NaN' e deleta colunas não utilizadas
    for ind, row in dataframe.iterrows():
        is_interval_nan = row[index_level_item:column_index].isna().all()
        if is_interval_nan:
            dataframe = dataframe.drop(index=ind)
    dataframe = dataframe.drop(dataframe.index[0])
    dataframe.reset_index(drop=True, inplace=True)
    dataframe = dataframe.iloc[:, index_level_item:column_index]
    index_desc = dataframe.columns.get_loc("DESCRIÇÃO")
    new_columns = list(dataframe.columns[: index_desc + 1]) + [
        "ITEM " + str(i) for i in range(index_desc, len(dataframe.columns))
    ]
    dataframe = dataframe.rename(
        columns=dict(zip(dataframe.columns, new_columns))
    )
    dataframe["ITEM"] = dataframe["ITEM"].replace(np.nan, 0)

    # Criação de níveis no dataframe
    for index, row in dataframe.iterrows():
        if row["ITEM"] != 0:
            item = str(row["ITEM"])
        if len(item) > 1:
            number = item[1:]
            if item.find("."):
                number_list = number.split(".")

        # Transformação Nivel 4:
        if qtd_levels >= 4:
            if math.isnan(row["ITEM 4"]) == False and row["ITEM"] == 0:
                if number.count(".") == 0:
                    final_item = item + "." + "1"
                    dataframe.iat[
                        index, dataframe.columns.get_loc("ITEM")
                    ] = final_item
                    item = final_item
                elif number.count(".") == 1:
                    new_number = int(number_list[1]) + 1
                    final_item = (
                        item[0] + number_list[0] + "." + str(new_number)
                    )
                    dataframe.iat[
                        index, dataframe.columns.get_loc("ITEM")
                    ] = final_item
                    item = final_item
                elif number.count(".") == 2:
                    new_number = int(number_list[1]) + 1
                    final_item = (
                        item[0] + number_list[0] + "." + str(new_number)
                    )
                    dataframe.iat[
                        index, dataframe.columns.get_loc("ITEM")
                    ] = final_item
                    item = final_item

        # Transformação Nivel 5:
        if qtd_levels >= 5:
            if math.isnan(row["ITEM 5"]) == False and row["ITEM"] == 0:
                if number.count(".") == 1:
                    final_item = item + "." + "1"
                    dataframe.iat[
                        index, dataframe.columns.get_loc("ITEM")
                    ] = final_item
                    item = final_item
                elif number.count(".") == 2:
                    new_number = int(number_list[2]) + 1
                    final_item = (
                        item[0]
                        + number_list[0]
                        + "."
                        + number_list[1]
                        + "."
                        + str(new_number)
                    )
                    dataframe.iat[
                        index, dataframe.columns.get_loc("ITEM")
                    ] = final_item
                    item = final_item
    return dataframe


def sum_of_levels(dataframe):
    # Cria Dicionários
    nivel2 = {}
    nivel3 = {}
    nivel4 = {}
    nivel5 = {}

    # Loop linha-a-linha
    for index, row in dataframe.iterrows():
        if str(row["ITEM"]).count(".") == 0 and len(str(row["ITEM"])) == 1:
            if math.isnan(row["ITEM 2"]) == False:
                nivel2.update({row["ITEM"]: [row["ITEM 2"], row["DESCRIÇÃO"]]})
        elif str(row["ITEM"]).count(".") == 0 and len(str(row["ITEM"])) > 1:
            if math.isnan(row["ITEM 3"]) == False:
                nivel3.update({row["ITEM"]: [row["ITEM 3"], row["DESCRIÇÃO"]]})
        elif str(row["ITEM"]).count(".") == 1:
            if math.isnan(row["ITEM 4"]) == False:
                nivel4.update({row["ITEM"]: [row["ITEM 4"], row["DESCRIÇÃO"]]})
        elif str(row["ITEM"]).count(".") == 2:
            if math.isnan(row["ITEM 5"]) == False:
                nivel5.update({row["ITEM"]: [row["ITEM 5"], row["DESCRIÇÃO"]]})

    # Outputs nivel 2
    resultados = {}

    if sum(nivel2[item][0] for item in nivel2) != 1:
        resultados.update(
            {
                " & ".join(nivel2.keys())
                + " -- "
                + str(" & ".join(nivel2[item][1] for item in nivel2)): str(
                    sum(nivel2[item][0] for item in nivel2)
                )
                + ", INCORRECT SUM"
            }
        )
    else:
        resultados.update(
            {
                " & ".join(nivel2.keys())
                + " -- "
                + str(" & ".join(nivel2[item][1] for item in nivel2)): str(
                    sum(nivel2[item][0] for item in nivel2)
                )
                + ", CORRECT SUM"
            }
        )

    # Outputs nivel 3

    key1 = list(nivel3.keys())[0][0]
    value1 = nivel3[list(nivel3.keys())[0]][0]
    index = 0

    for k, v in nivel3.items():
        if k[0] == key1 and index == 0:
            soma1 = float(value1)
            soma = soma1
            if 1.01 >= soma >= 0.99:
                resultados.update(
                    {k[0] + " -- " + v[1]: str(soma) + ", CORRECT SUM"}
                )
            else:
                resultados.update(
                    {k[0] + " -- " + v[1]: str(soma) + ", INCORRECT SUM"}
                )
        elif k[0] == key1:
            soma = soma + float(v[0])
        elif k[0] != key1:
            if 1.01 >= soma >= 0.99:
                resultados.update(
                    {k[0] + " -- " + v[1]: str(soma) + ", CORRECT SUM"}
                )
            else:
                resultados.update(
                    {k[0] + " -- " + v[1]: str(soma) + ", INCORRECT SUM"}
                )
            temp = k[0]
            key1 = temp
            soma = float(v[0])
        index += 1

    # Outputs nivel 4

    key1 = list(nivel4.keys())[0].partition(".")[0]
    value1 = nivel4[list(nivel4.keys())[0]][0]
    index = 0

    for k, v in nivel4.items():
        if k.partition(".")[0] == key1 and index == 0:
            soma1 = float(value1)
            soma = soma1
        elif k.partition(".")[0] == key1:
            soma = soma + float(v[0])
            key0 = k.partition(".")[0]
        elif k.partition(".")[0] != key1:
            if 1.01 >= soma >= 0.99:
                resultados.update(
                    {key0 + " -- " + v[1]: str(soma) + ", CORRECT SUM"}
                )
            else:
                resultados.update(
                    {key0 + " -- " + v[1]: str(soma) + ", INCORRECT SUM"}
                )
            temp = k.partition(".")[0]
            key1 = temp
            soma = float(v[0])
        index += 1

    # Outputs nivel 5

    key1 = ".".join(list(nivel5.keys())[0].split(".", 2)[:2])
    value1 = nivel5[list(nivel5.keys())[0]][0]
    index = 0

    for k, v in nivel5.items():
        if ".".join(k.split(".", 2)[:2]) == key1 and index == 0:
            soma1 = float(value1)
            soma = soma1
        elif ".".join(k.split(".", 2)[:2]) == key1:
            soma = soma + float(v[0])
            key0 = ".".join(k.split(".", 2)[:2])
        elif ".".join(k.split(".", 2)[:2]) != key1:
            if 1.01 >= soma >= 0.99:
                resultados.update(
                    {key0 + " -- " + v[1]: str(soma) + ", CORRECT SUM"}
                )
            else:
                resultados.update(
                    {key0 + " -- " + v[1]: str(soma) + ", INCORRECT SUM"}
                )
            temp = ".".join(k.split(".", 2)[:2])
            key1 = temp
            soma = float(v[0])
        index += 1

    return resultados


# file = "data\Modelo_EAP - rev5.xlsx"
# dataframe = create_dataframe_levels(file)
# results = sum_of_levels(dataframe)
