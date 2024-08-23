import os
import pandas as pd


def find_code_sheet(path, thick, mat):
    ferro_path = os.path.join(path, "Ferro.xlsx")
    ferro_list = ["ferro", "S235JR", "S275JR", "S325JR"]
    ferro_bool = 0
    inox_path = os.path.join(path, "Inox.xlsx")
    inox_list = ["inox", "aisi"]
    inox_bool = 0
    zinc_path = os.path.join(path, "Zinc.xlsx")
    zinc_list = ["zinc"]
    zinc_bool = 0
    galv_path = os.path.join(path, "Galv.xlsx")
    galv_list = ["galv"]
    galv_bool = 0
    result = None

    def code_finder(path, thick, mat):
        df = pd.read_excel(path, header=1)
        result = None
        for index, row in df.iterrows():
            thick_data = row["ESPESSURA (MM)"]
            if float(thick) == float(thick_data):
                result = row["CÓDIGO"]
        return result

    for acro in ferro_list:
        if acro in mat:
            ferro_bool = 1
    for acro in inox_list:
        if acro.lower() in mat.lower():
            inox_bool = 1
            ferro_bool = 0
    for acro in zinc_list:
        if acro.lower() in mat.lower():
            inox_bool = 0
            ferro_bool = 0
            zinc_bool = 1
    for acro in galv_list:
        if acro.lower() in mat.lower():
            inox_bool = 0
            ferro_bool = 0
            zinc_bool = 0
            galv_bool = 1

    if ferro_bool == 1:
        result = code_finder(ferro_path, thick, mat)
    elif inox_bool == 1:
        result = code_finder(inox_path, thick, mat)
    elif zinc_bool == 1:
        result = code_finder(zinc_path, thick, mat)
    elif galv_bool == 1:
        result = code_finder(galv_path, thick, mat)

    return result


path = r""

result = find_code_sheet(path, "2", "S235JR Galv")

print(result)
