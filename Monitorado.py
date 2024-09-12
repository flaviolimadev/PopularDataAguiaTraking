import pandas as pd
import json
import requests

# Função para converter os dados de uma linha do Excel para o formato desejado
def converter_dados(entry):
    # Função auxiliar para formatar a data corretamente
    def formatar_data(data):
        if pd.isna(data):  # Verifica se a data é nula
            return "1980-01-01"  # Valor padrão
        if isinstance(data, pd.Timestamp):  # Verifica se é um Timestamp
            return data.strftime("%Y-%m-%d")  # Converte para string no formato desejado
        return str(data).split("T")[0]  # Caso seja uma string, faz o split

    # Mapear os campos do Excel para os campos do formato desejado
    monitorado_json = {
        "key_monitorado": str(entry.get("SIAPEN")),  # Usando SIAPEN como key_monitorado
        "matricula_monitorado": str(entry.get("SIAPEN")),  # Também usando SIAPEN como matrícula
        "nome_monitorado": entry.get("Nome", "Carlos Silva"),
        "dispositivo": "Dispositivo A",  # Valor fixo
        "agencia_id": 1,  # Valor fixo
        "estabelecimento_prisional_id": 2,  # Valor fixo
        "monitorado_vitima": 0,  # Valor fixo
        "perfil_id": 5,  # Valor fixo
        "nome_completo": entry.get("Nome", "Carlos Alberto Silva"),
        "nome_social": "",  # Não disponível no Excel, valor padrão
        "apelido": "",  # Não disponível no Excel, valor padrão
        "nome_mae": entry.get("Nome da Mãe", "Maria Silva"),
        "nome_pai": "José Silva",  # Valor fixo
        "genero": entry.get("Sexo", "Masculino"),
        "cpf": "12345678900",  # Valor fixo
        "rg": "MG1234567",  # Valor fixo
        "data_nascimento": formatar_data(entry.get("Data de Nascimento")),  # Formatando data
        "protocolo_monitoramento": "PRT2024",  # Valor fixo
        "regime_id": 2,  # Valor fixo
        "controle_prazo": "0",  # Valor fixo
        "inicio_medida": "2024-01-01",  # Valor fixo
        "dias_medida": 365,  # Valor fixo
        "prorrogacao": 30,  # Valor fixo
        "tipo_monitorado_id": 3,  # Valor fixo
        "periculosidade": 1,  # Valor fixo
        "faccao_id": 4,  # Valor fixo
        "religiao": "Católica",  # Valor fixo
        "estado_civil": "Casado",  # Valor fixo
        "situacao_trabalhista_id": 7,  # Valor fixo
        "escolaridade_id": 6,  # Valor fixo
        "fotos": "url_da_foto",  # Valor fixo
        "arquivos": "url_do_arquivo",  # Valor fixo
        "telefones": json.dumps([{"numero": entry.get("Telefones", ""), "whatsapp": 1}]),  # Formatando telefone
        "zonas": json.dumps([{"zona": "Zona Norte"}]),  # Valor fixo
        "processos": json.dumps([{"processo": "123456"}]),  # Valor fixo
        "agendamento_servicos": json.dumps([{"servico": "Consulta médica"}]),  # Valor fixo
        "comandos_dispositivo": json.dumps([{"comando": "Reset"}]),  # Valor fixo
        "notificacoes_observacoes": "Nenhuma",  # Valor fixo
        "historico_posicoes": "Histórico de posições"  # Valor fixo
    }

    return monitorado_json

# Função para ler o arquivo Excel e converter para o formato desejado
def excel_para_json(caminho_arquivo_excel):
    # Lê o arquivo Excel
    df = pd.read_excel(caminho_arquivo_excel, skiprows=2)

    # Renomeia as colunas para corresponder ao formato esperado
    df.columns = [
        'Seq', 'SIAPEN', 'Nome', 'Data de Nascimento', 'Sexo', 'Nome da Mãe',
        'Tipo Pessoa', 'Tipo Regime Cumprimento Pena', 'Unidade da Pessoa', 'Município do Endereço',
        'Logradouro do Endereço', 'Telefones', 'Data da Primeira Ativação', 'Situação'
    ]

    # Converte cada linha para o formato JSON esperado
    json_list = []
    for _, row in df.iterrows():
        json_entry = converter_dados(row)
        json_list.append(json_entry)

    return json_list

# Função para enviar o JSON para a URL via POST
def enviar_via_post(data, url):
    headers = {'Content-Type': 'application/json'}
    
    for entry in data:
        try:
            # Envia cada entrada via POST
            response = requests.post(url, data=json.dumps(entry), headers=headers)
            if response.status_code == 201:
                print("Dados enviados com sucesso!")
            else:
                print(f"Erro ao enviar dados: {response.status_code}")
                print(f"Resposta: {response.text}")
        except requests.RequestException as e:
            print(f"Erro ao fazer a requisição: {e}")

# Caminho para o arquivo Excel
caminho_arquivo = 'Relatório - Cadastro de Monitorados.xlsx'

# Converte o Excel para JSON
json_resultado = excel_para_json(caminho_arquivo)

# Envia os dados para a URL via POST
url = 'http://127.0.0.1:8000/monitorados/criar/'
enviar_via_post(json_resultado, url)
