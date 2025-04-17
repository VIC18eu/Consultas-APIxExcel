import requests
import pandas as pd
import openpyxl
import os
import time
import psutil
from openpyxl import load_workbook
from tqdm import tqdm

def fechar_excel_consultas(ficheiro):
    if not os.path.exists(ficheiro):
        return

    try:
        os.rename(ficheiro, ficheiro)
    except PermissionError:
        print("⚙️ A fechar o consultas.xlsx...")
        for proc in psutil.process_iter(['name', 'cmdline']):
            try:
                if proc.info['name'] and 'EXCEL' in proc.info['name'].upper():
                    if any(ficheiro in str(arg) for arg in proc.info['cmdline']):
                        proc.terminate()
                        proc.wait(timeout=5)
                        print("✅ Excel fechado com sucesso!")
                        break
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        time.sleep(1)

def atualizar_excel(endpoint, ficheiro):

    if not os.path.exists(ficheiro):
        print("📄 Ficheiro não encontrado. A criar novo ficheiro...")
        criar_excel(endpoint, ficheiro)
        return
    
    fechar_excel_consultas(ficheiro)
    results = []

    offset = 0

    response = requests.get(endpoint + f"&offset={offset}")
    data = response.json()
    total_count = data['total_count']
    results.extend(data['results'])
    offset += 100

    if total_count > 9900:
        print("🏋️‍♂️ Demasiados para serem extraidos!")
        time.sleep(0.5)
        return

    with tqdm(total=total_count, desc="🔄 A descarregar dados", unit="registos") as pbar:
        pbar.update(len(data['results']))
        
        while offset < total_count:
            response = requests.get(endpoint + f"&offset={offset}")
            data = response.json()
            results.extend(data['results'])
            offset += 100
            pbar.update(len(data['results']))
    
    results_expandido = []
    for item in results:
        expanded_item = {}
        for key, value in item.items():
            if isinstance(value, list):
                expanded_item[key] = ", ".join(map(str, value))
            elif isinstance(value, dict):
                for subkey, subvalue in value.items():
                    expanded_item[f"{key}_{subkey}"] = subvalue
            else:
                expanded_item[key] = value
        results_expandido.append(expanded_item)

    df = pd.DataFrame(results_expandido)
    path = f"{ficheiro}.xlsx"

    with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=f'Dados API', index=False)

    print("✅ Excel atualizado com sucesso!")

    os.startfile(f'.\{ficheiro}.xlsx')

def criar_excel(endpoint, ficheiro):
    fechar_excel_consultas(ficheiro)
    results = []

    offset = 0

    response = requests.get(endpoint + f"&offset={offset}")
    data = response.json()
    total_count = data['total_count']
    results.extend(data['results'])
    offset += 100

    if total_count > 9900:
        print("🏋️‍♂️ Demasiados para serem extraidos!")
        time.sleep(0.5)
        return

    with tqdm(total=total_count, desc="🔄 A descarregar dados", unit="registos") as pbar:
        pbar.update(len(data['results']))
        
        while offset < total_count:
            response = requests.get(endpoint + f"&offset={offset}")
            data = response.json()
            results.extend(data['results'])
            offset += 100
            pbar.update(len(data['results']))
    
    results_expandido = []
    for item in results:
        expanded_item = {}
        for key, value in item.items():
            if isinstance(value, list):
                expanded_item[key] = ", ".join(map(str, value))
            elif isinstance(value, dict):
                for subkey, subvalue in value.items():
                    expanded_item[f"{key}_{subkey}"] = subvalue
            else:
                expanded_item[key] = value
        results_expandido.append(expanded_item)

    df = pd.DataFrame(results_expandido)
    path = f"{ficheiro}.xlsx"

    with pd.ExcelWriter(path, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, sheet_name='Dados API', index=False)

    print("✅ Excel criado com sucesso!")

    os.startfile(path)

