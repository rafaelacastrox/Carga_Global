import requests
import pandas as pd

def obter_dados_api(url):
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Falha ao obter dados da API. URL: {url}, Status Code: {response.status_code}")
        return None

def criar_tabela():
    url = 'https://apicarga.ons.org.br/prd/cargaverificada'  # URL da API
    dados_api = obter_dados_api(url)
    
    if dados_api:
        # Criar um DataFrame pandas a partir dos dados da API
        df = pd.DataFrame(dados_api)
        
        # Salvar o DataFrame em um arquivo CSV
        df.to_csv('dados_api.csv', index=False)
        
        print("Tabela criada com sucesso.")
    else:
        print("Não foi possível obter dados da API.")

# Chamar a função para criar a tabela
criar_tabela()