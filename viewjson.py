import pandas as pd
import json

# Função para converter arquivo Excel em JSON
def excel_para_json(caminho_arquivo_excel):
    # Lê o arquivo Excel
    df = pd.read_excel(caminho_arquivo_excel)
    
    # Converte o DataFrame em JSON
    json_data = df.to_json(orient='records', date_format='iso')
    
    # Retorna o JSON formatado
    return json.loads(json_data)[0];

# Caminho para o arquivo Excel
caminho_arquivo = 'Relatório - Cadastro de Monitorados (2).xlsx'

# Converte o Excel para JSON
json_resultado = excel_para_json(caminho_arquivo)

# Exibe o resultado em formato JSON
print(json.dumps(json_resultado, indent=4))
