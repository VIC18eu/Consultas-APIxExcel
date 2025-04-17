import requests
import pandas as pd
import openpyxl
import os
import time
import psutil
from openpyxl import load_workbook
from tqdm import tqdm

def fechar_excel_consultas():
    ficheiro = 'consultas.xlsx'
    if not os.path.exists(ficheiro):
        return

    try:
        os.rename(ficheiro, ficheiro)
    except PermissionError:
        print("⚙️ A fechar o consultas.xlsx...")
        for proc in psutil.process_iter(['name', 'cmdline']):
            try:
                if proc.info['name'] and 'EXCEL' in proc.info['name'].upper():
                    if any('consultas.xlsx' in str(arg) for arg in proc.info['cmdline']):
                        proc.terminate()
                        proc.wait(timeout=5)
                        print("✅ Excel fechado com sucesso!")
                        break
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        time.sleep(1)

fechar_excel_consultas()

endpoint = "https://transparencia.sns.gov.pt/api/explore/v2.1/catalog/datasets/evolucao-das-consultas-medicas-nos-csp/records?order_by=tempo&limit=100"
consultas = []

offset = 0

# Primeiro pedido para obter o total_count
response = requests.get(endpoint + f"&offset={offset}")
data = response.json()
total_count = data['total_count']
consultas.extend(data['results'])
offset += 100

# Inicializa a barra de progresso
with tqdm(total=total_count, desc="🔄 A descarregar dados", unit="registos") as pbar:
    pbar.update(len(data['results']))
    
    while offset < total_count:
        response = requests.get(endpoint + f"&offset={offset}")
        data = response.json()
        consultas.extend(data['results'])
        offset += 100
        pbar.update(len(data['results']))

# Tratar os dados
for item in consultas:
    geo = item.pop('ponto_ou_localizacao_geografica', {})
    if geo is None:
        geo = {}
    item['longitude'] = geo.get('lon')
    item['latitude'] = geo.get('lat')

df = pd.DataFrame(consultas)

df = df.rename(columns={
    'tempo': 'Data',
    'regiao': 'Região',
    'entidade': 'Entidade',
    'no_de_consultas_medicas_presencias_qt': 'Nº de Consultas Presenciais',
    'no_de_consultas_medicas_nao_presenciais_ou_inespecificas_qt': 'Nº de Consultas Não Presenciais ou Inespecificas',
    'no_de_consultas_medicas_ao_domicilio_qt': 'Nº de Consultas ao Domicilio',
    'total_consultas': 'Total de Consultas',
    'longitude': 'Longitude',
    'latitude': 'Latitude',
})

path = 'consultas.xlsx'

# Usa o writer sem apagar
with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Consultas', index=False)

print("✅ Excel atualizado com sucesso!")

os.startfile('.\consultas.xlsx')
