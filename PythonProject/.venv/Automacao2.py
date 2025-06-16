import os
import pandas as pd

# MUDAR AQUI QUANDO QUISER TROCAR A PAGINA DO EXCEL
nome_aba = "Manufatura ok"
caminho_excel = os.path.expanduser("~/Downloads/Pasta.xlsx")

try:
    df = pd.read_excel(caminho_excel, sheet_name=nome_aba)
    print(f"Aba '{nome_aba}' carregada com sucesso!")
except FileNotFoundError:
    print("Arquivo Excel não encontrado.")
    exit()
except ValueError:
    print(f"Aba '{nome_aba}' não encontrada no arquivo.")
    exit()

# VALIDAR COLUNAS
col_nivel1 = df.columns[0]
col_nivel2 = df.columns[1]
col_nivel3 = df.columns[2]

df = df[[col_nivel1, col_nivel2, col_nivel3]].dropna().drop_duplicates()

# CAMINHO LOCAL DO SHAREPOINT SINCRONIZADO PELO ONEDRIVE
caminho_local_sharepoint = r"C:\Users\rodri\OneDrive - Stefanini\testeGuildasSAP - Documentos"

# CRIAR PASTAS LOCALMENTE
for index, row in df.iterrows():
    nivel1 = str(row[col_nivel1]).strip()
    nivel2 = str(row[col_nivel2]).strip()
    nivel3 = str(row[col_nivel3]).strip()

    caminho_nivel1 = os.path.join(caminho_local_sharepoint, nivel1)
    caminho_nivel2 = os.path.join(caminho_nivel1, nivel2)
    caminho_nivel3 = os.path.join(caminho_nivel2, nivel3)

    try:
        os.makedirs(caminho_nivel3, exist_ok=True)
        print(f"[OK] Pasta criada ou já existia: {caminho_nivel3}")
    except Exception as e:
        print(f"[ERRO] Falha ao criar pasta: {caminho_nivel3}")
        print("Detalhes:", e)
