import json
import os
import pathlib

# specify the path to the directory containing JSON files
path_to_input_fils  = "./mapper"
input_files = os.listdir(path_to_input_fils)

def read_json(file_path):
    try:
        with open(file_path, 'r') as file:
            data = json.load(file)
            return data
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return None
    except json.JSONDecodeError:
        print(f"Error decoding JSON from file: {file_path}")
        return None

def extract_input_file_type(filename):
    name, _ = os.path.splitext(filename)
    return name.split('_')[-1] 

def extract_mapper_fields(mapper_fields):
    fields = {}
    for key, value in mapper_fields.items():
        fields[key] = {
            "type": value.get("type", ""),
            "source": value.get("source", "")
        }
    return fields 

def create_file_name(filetype):
    return f"{filetype}_fields.xlsx"

def create_excel_sheet(fields, filename, file_label):
    import pandas as pd
    df_new = pd.DataFrame([
        {"Type": v["type"], "Field Name": k, file_label: v["source"]}
        for k, v in fields.items()
    ])

    if pathlib.Path(filename).exists():
        df_existing = pd.read_excel(filename)
        df_merged = pd.merge(df_existing, df_new[["Type", "Field Name", file_label]], on="Field Name", how="outer")
        df_merged["Type"] = df_merged["Type_x"].combine_first(df_merged["Type_y"])
        df_merged = df_merged.drop(columns=["Type_x", "Type_y"])
        cols = ["Type", "Field Name"] + [c for c in df_merged.columns if c not in ["Type", "Field Name"]]
        df_merged = df_merged[cols]
        df_merged.to_excel(filename, index=False)
        print(f"Appended new source column for {file_label} to Excel file: {filename}")
    else:
        df_new.to_excel(filename, index=False)
        print(f"Excel file created: {filename}")

if __name__ == "__main__":
    for file in input_files:
        filetype = extract_input_file_type(file)
        FILE_PATH = os.path.join(path_to_input_fils, file)
        data = read_json(FILE_PATH)
        if data:
            fields = extract_mapper_fields(data["fields"])
            file_label = os.path.splitext(file)[0] 
            create_excel_sheet(fields, create_file_name(filetype), file_label)