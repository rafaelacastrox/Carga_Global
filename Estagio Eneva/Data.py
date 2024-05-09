# Importações necessárias
import requests
import logging
import tabulate
from openpyxl import load_workbook

# Carregar o arquivo Excel com os dados da API
workbook = load_workbook('dados_api.xlsx')

# Configuração do logger para salvar os logs em um arquivo chamado 'log.txt'
logging.basicConfig(filename='log.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Função para calcular a média de carga por estação e ano
def calcular_media_por_estacao_e_ano(url, ano):
    # Lista dos estados
    estados = ['SECO','S','NE','N','RJ','SP','MG','ES','MT','MS','DF','GO','AC','RO','PR','SC','RS','BASE','BAOE','ALPE','PBRN ','CE','PI','TON',
    'PA','MA ','AP','AM','RR','PESE','PES','PENE']
    
    # Iterar sobre os estados
    for estado in estados:
        # URLs para cada estação do ano
        url_estadoV = f'{url}?dat_inicio={ano}-01-01&dat_fim={ano}-03-31&cod_areacarga={estado}'
        url_estadoO = f'{url}?dat_inicio={ano}-04-01&dat_fim={ano}-06-31&cod_areacarga={estado}'
        url_estadoI = f'{url}?dat_inicio={ano}-07-01&dat_fim={ano}-09-31&cod_areacarga={estado}'
        url_estadoP = f'{url}?dat_inicio={ano}-10-01&dat_fim={ano}-12-31&cod_areacarga={estado}'
        
        # Obter dados da API para cada estação
        dados_apiV = obter_dados_api(url_estadoV)
        dados_apiI = obter_dados_api(url_estadoI)
        dados_apiO = obter_dados_api(url_estadoO)
        dados_apiP = obter_dados_api(url_estadoP)

        # Calcular a média de carga para cada estação
        media_por_estacao ={
            'Verão': calcular_media_por_estacao(dados_apiV),
            'Outono': calcular_media_por_estacao(dados_apiI),
            'Inverno': calcular_media_por_estacao(dados_apiO),
            'Primavera': calcular_media_por_estacao(dados_apiP)
        }

        # Exibir os resultados
        exibir_resultados(media_por_estacao, estado, ano)

# Função para obter dados da API
def obter_dados_api(url):
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    else:
        return None

# Função para calcular a média de carga
def calcular_media_por_estacao(dados):
    media_estacao = []
    
    for registro in dados:
        carga = registro.get('val_cargaglobal')
        media_estacao.append(carga)
    
    return media_estacao

# Função para calcular a estação do ano com base no mês
def calcular_estacao_por_mes(mes):
    if mes in [1, 2, 3]:
        return 'Verão'
    elif mes in [4, 5, 6]:
        return 'Outono'
    elif mes in [7, 8, 9]:
        return 'Inverno'
    elif mes in [10, 11, 12]:
        return 'Primavera'

# Função para exibir os resultados
def exibir_resultados(media_por_estacao, estado, ano):
    print(f"\nResultados para o estado {estado} no ano {ano}:")
    for estacao, cargas in media_por_estacao.items():
        if len(cargas) >= 3:
            media_trimestral = sum(cargas) / len(cargas)
            print(f"Média de carga para a estação {estacao}: {media_trimestral}")
        else:
            print(f"Não há dados suficientes para calcular a média para a estação {estacao}")

# Função para exibir a tabela comparativa
def exibir_tabela_comparativa(dados_gerais):
    tabela_dados = []
    for estado, dados_estado in dados_gerais.items():
        for ano, media_por_estacao in enumerate(dados_estado, 2010):
            tabela_dados.append([estado, ano, *media_por_estacao.values()])

    headers = ['Estado', 'Ano', 'Verão', 'Outono', 'Inverno', 'Primavera']
    print(tabulate(tabela_dados, headers=headers, tablefmt="grid"))

# Função para encontrar a estação com maior consumo de carga
def encontrar_maior_consumo(dados_gerais):
    maior_consumo = 0
    estacao_maior_consumo = ''
    ano_maior_consumo = 0

    for estado, dados_estado in dados_gerais.items():
        for ano, media_por_estacao in enumerate(dados_estado, 2010):
            for estacao, consumo in media_por_estacao.items():
                if consumo > maior_consumo:
                    maior_consumo = consumo
                    estacao_maior_consumo = estacao
                    ano_maior_consumo = ano

    print(f"A estação com maior consumo de carga foi {estacao_maior_consumo}, no ano de {ano_maior_consumo}, com consumo de {maior_consumo}.")

    return estacao_maior_consumo, ano_maior_consumo

# Função principal
def main():
    # URL da API
    url = 'https://apicarga.ons.org.br/prd/cargaverificada'
    
    # Ano desejado
    ano = input("Digite o ano desejado: ")
    
    # Calcular média por estação e ano
    calcular_media_por_estacao_e_ano(url, ano)

# Verificar se o código está sendo executado como script principal
if __name__ == "__main__":
    main()