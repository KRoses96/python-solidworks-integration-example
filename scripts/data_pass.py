"""

Author: Manuel Rosa
Description: Transfers all data inside a software made folder to the production folder, to then be read by their interface for further investigation

"""

# IMPORTS
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
import re
import sys
import shutil
import openpyxl
from datetime import datetime


def ask_for_directory():
    """Asks the user for the directory with the automated lists"""
    root = tk.Tk()
    root.withdraw()  
    directory_path = filedialog.askdirectory(title="Selecione a pasta com as listagens")
    if directory_path:
        files = [
            f
            for f in os.listdir(directory_path)
            if os.path.isfile(os.path.join(directory_path, f))
        ]
    else:
        show_error_message("Erro", "Nenhuma pasta selecionada")
        sys.exit()
    return directory_path, files


def show_error_message(title, message):
    """Error messagebox for error handling"""
    root = tk.Tk()
    root.withdraw() 
    messagebox.showerror(title, message)


def are_you_sure_message(user, version, assembly):
    """Confirmation box"""
    root = tk.Tk()
    root.withdraw()
    response = messagebox.askyesno(
        "Tem a certeza?",
        f"Deseja enviar o seguinte conjunto para produção: \n{assembly}_V{version}\nProjetista: {user}\n",
    )
    return response


def show_completed(version, assembly):
    """Completion box"""
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(
        "Enviado", f"Enviado para produção Conjunto: {assembly}_V{version}"
    )



pd.set_option("display.max_columns", None)

script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
parent_dir = os.path.dirname(script_dir)
options = os.path.join(parent_dir, "extras", "op.txt")
with open(options, "r") as file:
    lines = file.readlines()
for line in lines:
    if "User =" in line:
        user_value = line.split("User =", 1)[1].strip()
    if "phc =" in line:
        phc_path = line.split("=", 1)[1].strip()

print(phc_path)
dir_info = ask_for_directory()
dir_pt = dir_info[0]  
files = dir_info[1]  
index = dir_pt.find("Listas_") 

dxf_file = None
suffix_dxf = "DXF.xlsx"
weld_file = None
suffix_weld = "Perfis.xlsx"
weld_cut_file = None
suffix_weld_cut = "Perfis_Corte.xlsx"
buy_file = None
suffix_buy = "Compras.xlsx"


for file in files:
    if file.endswith(suffix_dxf):
        dxf_file = file
    if file.endswith(suffix_weld):
        weld_file = file
    if file.endswith(suffix_buy):
        buy_file = file
    if file.endswith(suffix_weld_cut):
        weld_cut_file = file
if (
    buy_file == None and weld_file == None and dxf_file == None
):  
    show_error_message(
        "Erro: Pasta Inválida", "Verifique se selecionou a pasta correta"
    )
    sys.exit()

if index != -1:
    folder_dir = dir_pt[index + len("Listas_") :]
folder_parts = folder_dir.split("-")
print(folder_parts)
main_dir = "-".join(folder_parts[:2])
main_dir = main_dir.strip()
print("main_dir:" + main_dir)
sub_dir = "-".join(folder_parts[2:])
sub_dir = sub_dir.strip()
print("sub_dir:" + sub_dir)

if sub_dir == "":
    show_error_message(
        "Erro Nomenclatura",
        f"Nomenclatura errada: {folder_dir}\n\nExemplo: F2024-0000-Conjunto",
    )
    sys.exit()


folder_final = os.path.join(phc_path, "Obras")  
folder_extract = os.path.join(folder_final, main_dir)  
subfolder_extract = os.path.join(folder_extract, sub_dir)
version = 1

folder_version = os.path.join(subfolder_extract, f"V{version}")


while os.path.exists(folder_version):
    version += 1
    folder_version = os.path.join(subfolder_extract, f"V{version}")


user_response = are_you_sure_message(
    user_value, version, folder_dir
)  


if user_response == False:
    sys.exit()

if not os.path.exists(folder_extract):
    os.makedirs(folder_extract)
if not os.path.exists(subfolder_extract):
    os.makedirs(subfolder_extract)


os.makedirs(folder_version)
csv_folder = os.path.join(folder_version, "CSV")
os.makedirs(csv_folder)
project_folder = os.path.join(folder_version, "Projeto")
os.makedirs(project_folder)
cut_folder = os.path.join(csv_folder, "Corte")
os.makedirs(cut_folder)


for file_name in files: 
    if not file_name.lower().endswith(".xlsx"):
        source_path = os.path.join(dir_pt, file_name)
        base_name, extension = os.path.splitext(file_name)
        new_file_name = f"{base_name}_V{version}{extension}"
        destination_path = os.path.join(project_folder, new_file_name)
        shutil.copy(source_path, destination_path)
for folder_name in os.listdir(dir_pt):
    source_path = os.path.join(dir_pt, folder_name)
    destination_path = os.path.join(project_folder, folder_name)
    if os.path.isdir(source_path):
        shutil.copytree(source_path, destination_path)

csv_dxf = "DXF.csv"
csv_nest_dxf = "NEST.csv"
csv_dxf_path = os.path.join(csv_folder, csv_dxf)
csv_nest_dxf_path = os.path.join(csv_folder, csv_nest_dxf)

if dxf_file != None:
    dxf_path = os.path.join(dir_pt, dxf_file)
    df = pd.read_excel(dxf_path, sheet_name="DXF", header=2)
    df = df.dropna(how="all")
    df = df[df.ne(df.columns, axis=1).any(axis=1)]
    df = df.ffill()
    df["Cortado"] = 0
    df.to_csv(csv_dxf_path, index=False)
    df_nest = pd.read_excel(dxf_path, sheet_name="DXF_Nest", header=1)
    new_rows = []
    for index, row in df_nest.iterrows():
        chapas_list = [x.strip() for x in row["Chapas"].split(",")]
        for chapas_entry in chapas_list:
            new_row = row.copy()
            new_row["Chapas"] = chapas_entry
            new_rows.append(new_row)
    df_nest = pd.DataFrame(new_rows)
    df_nest["Qt."] = df_nest["Chapas"].apply(
        lambda x: (
            int(re.search(r"\(x(\d+)\)", x).group(1))
            if re.search(r"\(x(\d+)\)", x)
            else 1
        )
    )

    df_nest["Chapas"] = df_nest["Chapas"].str.replace(r"\(x\d+\)", "", regex=True)
    df_nest = df_nest[["Codigo", "Material", "Esp.", "Chapas", "Qt."]]
    df_nest["Enc."] = 0
    df_nest["Stock"] = 0
    df_nest = df_nest.fillna("")
    df_nest.to_csv(csv_nest_dxf_path, index=False)


csv_weld = "Perfis.csv"
csv_weld_path = os.path.join(csv_folder, csv_weld)

if weld_file != None:
    weld_path = os.path.join(dir_pt, weld_file)
    df = pd.read_excel(weld_path, sheet_name="Perfis", header=1)
    df = df.fillna("")
    df = df[["Codigo", "Descrição", "Material", "Comp.(m)", "Comp.C.(m)", "Qt."]]
    df["Enc."] = 0
    df["Stock"] = 0
    df.to_csv(csv_weld_path, index=False)

    cut_weld_path = os.path.join(dir_pt, weld_cut_file)
    wb = openpyxl.load_workbook(cut_weld_path)
    sheet_names = wb.sheetnames
    for sheet_name in sheet_names:
        df_cut = pd.read_excel(cut_weld_path, sheet_name, header=4)
        df_cut = df_cut.iloc[:-2]
        columns_to_delete = ["Âng.1", "Âng.2", "Conjunto"]
        df_cut = df_cut.rename(columns={"Quant.": "Qt."})
        df_cut = df_cut.iloc[:, :-1]

        df_cut = df_cut.drop(columns=columns_to_delete)
        df_cut["Comprimento (mm)"] = df_cut["Comprimento (mm)"].astype(float) * 0.001
        df_cut = df_cut.rename(
            columns={
                "Designação": "Designação",
                "Material": "Material",
                "Comprimento (mm)": "Comp.(m)",
                "Qt.": "Qt.",
            }
        )
        df_cut_name = (
            df_cut["Designação"][0] + "_" + df_cut["Material"][0]
        )  
        cut_csv_file = f"{df_cut_name}.csv"
        no_mat_weld = 0
        cut_csv_file = "".join(i for i in cut_csv_file if i not in "\/:*?<>|")
        cut_csv_path = os.path.join(cut_folder, cut_csv_file)
        df_cut.to_csv(cut_csv_path, index=False)


csv_buy = "Compras.csv"
csv_buy_path = os.path.join(csv_folder, csv_buy)
if buy_file != None:
    buy_path = os.path.join(dir_pt, buy_file)
    df = pd.read_excel(buy_path, sheet_name="Material", header=5)
    columns_to_check = ["Designação", "Qt."]
    df = df.dropna(subset=columns_to_check)
    df.to_csv(csv_buy_path, index=False)


current_datetime = datetime.now()
formatted_datetime = current_datetime.strftime("%d/%m/%Y %H:%M")

log_txt = "log.txt"
log_txt_path = os.path.join(subfolder_extract, log_txt)
info_proje = f"V{version}: {user_value}, {formatted_datetime}"

if os.path.exists(log_txt_path):
    with open(log_txt_path, "a") as output_file:
        output_file.write(f"\n{info_proje}")
else:
    with open(log_txt_path, "w") as output_file:
        output_file.write(info_proje)

show_completed(version, folder_dir)
