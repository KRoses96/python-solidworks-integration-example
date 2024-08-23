import os
import pandas as pd


def code_finder(string, material, path):
    upn_path = os.path.join(path, "UPN.xlsx")
    ipe_path = os.path.join(path, "IPE.xlsx")
    ipn_path = os.path.join(path, "IPN.xlsx")
    lpn_path = os.path.join(path, "LPN.xlsx")
    barra_path = os.path.join(path, "Barra.xlsx")
    tuborect_path = os.path.join(path, "TuboRect.xlsx")
    tubo_path = os.path.join(path, "TuboRed.xlsx")
    spiro_path = os.path.join(path, "Spiro.xlsx")
    tubomec_path = os.path.join(path, "TuboMec.xlsx")
    varao_path = os.path.join(path, "Varao.xlsx")

    # Setting up all search terms and variables
    aco_strings = ["s235jr", "s275jr", "ferro", "aço", "zinc"]
    inox_strings = ["inox 304", "aisi 304", "aisi304", "inox304", "inox", "aisi"]
    alum_strings = ["alloy", "alumínio", "aluminio"]
    inox_316_strings = ["inox 316", "aisi 316", "aisi316", "inox316"]
    inox_303_strings = ["inox 303", "aisi 303", "aisi303", "inox303"]
    ck45k_strings = ["ck4520", "ck45k", "ck45", "c4"]
    mat_dictionary = {
        "inox_303": inox_303_strings,
        "inox_316": inox_316_strings,
        "inox": inox_strings,
        "aco": aco_strings,
        "alum": alum_strings,
        "ck45k": ck45k_strings,
    }

    all_strings = []

    result = None

    for mat in mat_dictionary:
        for i in mat_dictionary[mat]:
            all_strings.append(i)

    for i in all_strings:
        if i in material.lower():
            keyword = i
            break
    try:
        for key, value in mat_dictionary.items():
            if keyword in value:
                parent_list = key
    except:
        return

    def search_result(df):
        df_material = pd.DataFrame()
        measure = ""
        measure_excel = ""
        df_measure = pd.DataFrame()
        result = None

        if "cost" not in string.lower():
            for index, row in df.iterrows():
                if "cost" in row["Design"].lower():
                    df = df.drop(index)

        for index, row in df.iterrows():
            row["Design"] = row["Design"].lower()
            for item in mat_dictionary[parent_list]:
                if item in row["Design"]:
                    df_material = df_material._append(row)
                    new_index = df_material.index[-1]
                    df_material.at[new_index, "Design"] = df_material.at[
                        new_index, "Design"
                    ].replace(item, "")
                    df_material.at[new_index, "Design"] = df_material.at[
                        new_index, "Design"
                    ].replace(",", ".")

        if len(df_material) > 1:
            df_material = df_material.sort_values(
                by="Design", key=lambda x: x.str.len()
            )
            df_material = df_material.drop_duplicates(subset="Ref", keep="first")

        def numeric_values(string_des):
            measure = ""
            for num in string_des:
                if num.isnumeric():
                    measure = measure + num
                elif num.lower() == "x":
                    measure = measure + num.lower()
                elif num == ".":
                    if len(measure) > 1 and (
                        measure[-1].isnumeric() or measure[-1] == "x"
                    ):
                        measure = measure + num
                else:
                    measure = measure + ""

            return measure

        measure = numeric_values(string)

        for index, row in df_material.iterrows():
            measure_excel = numeric_values(row["Design"])
            if len(measure_excel) == len(measure) and measure_excel in measure:
                if len(df_measure) == 0:
                    df_measure = df_measure._append(row, ignore_index=True)
            measure_excel = ""
        if len(df_measure) == 1:
            for index, row in df_measure.iterrows():
                result = row["Ref"]
        return result, parent_list, measure

    if "upn" in string.lower():
        df = pd.read_excel(upn_path)
        result = search_result(df)
    elif "ipe" in string.lower():
        df = pd.read_excel(ipe_path)
        result = search_result(df)
    elif "ipn" in string.lower():
        df = pd.read_excel(ipn_path)
        result = search_result(df)
    elif (
        "lpn" in string.lower()
        or "cantoneira" in string.lower()
        or "cant" in string.lower()
    ):
        df = pd.read_excel(lpn_path)
        result = search_result(df)
    elif "barra" in string.lower():
        df = pd.read_excel(barra_path)
        result = search_result(df)
    elif "tubo" in string.lower() and (
        "quad" in string.lower() or "rect" in string.lower() or "reta" in string.lower()
    ):
        df = pd.read_excel(tuborect_path)
        result = search_result(df)
    elif "tubo" in string.lower() and "spiro" in string.lower():
        df = pd.read_excel(spiro_path)
        result = search_result(df)
    elif "tubo" in string.lower() and (
        "mecanico" in string.lower() or "mecânico" in string.lower()
    ):
        df = pd.read_excel(tubomec_path)
        result = search_result(df)
    elif "varão" in string.lower() and (
        "quad" not in string.lower() or "rect" not in string.lower()
    ):
        df = pd.read_excel(varao_path)
        result = search_result(df)
    elif "tubo" in string.lower():
        df = pd.read_excel(tubo_path)
        result = search_result(df)
    if result != None:
        return result[0]
