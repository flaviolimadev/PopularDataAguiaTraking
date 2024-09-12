import pandas as pd
import json
import pymysql

# Função para converter arquivo Excel em JSON
def excel_para_json(caminho_arquivo_excel):
    # Lê o arquivo Excel
    df = pd.read_excel(caminho_arquivo_excel)
    
    # Converte o DataFrame em JSON
    json_data = df.to_json(orient='records', date_format='iso')
    
    # Retorna o JSON formatado
    return json.loads(json_data)

# Função para cadastrar dados no banco
def cadastrar_tracking(data):
    try:
        # Conectar ao banco de dados
        conn = pymysql.connect(
            host='localhost',
            user='root',
            password='123456789',
            database='aguia_tracker'
        )
        cursor = conn.cursor()

        # Preparar a query SQL para inserção
        query = """
            INSERT INTO traking_data 
            (coordenadas_geograficas, LBS, status_feixe_luz, deteccao_de_jamming, deteccao_de_violacao_de_caixa, 
            altura, velocidade, VDOP, HDOP, qualidade_satelite, nivel_bateria, e_sim_card, inercia, qualidade_GPRS, created_at, updated_at) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, NOW(), NOW())
        """

        # Inserir cada linha do JSON no banco de dados
        for entry in data:
            cursor.execute(query, (
                entry.get("Coordenadas"),
                str(entry.get("Número de Série")),
                "0",  # status_feixe_luz (não fornecido)
                "0",  # deteccao_de_jamming (não fornecido)
                "0",  # deteccao_de_violacao_de_caixa (não fornecido)
                str(entry.get("Altitude")),
                str(entry.get("Velocidade")),
                str(entry.get("VDOP")),
                str(entry.get("HDOP")),
                str(entry.get("Qtd. Satélites")),
                str(entry.get("Bateria (%)")),
                str(entry.get("Operadora")),
                "0",  # inercia (não fornecido)
                "0"   # qualidade_GPRS (não fornecido)
            ))

        # Confirmar as alterações no banco de dados
        conn.commit()
        print("Dados inseridos com sucesso!")

    except pymysql.MySQLError as err:
        print(f"Erro: {err}")
    
    finally:
        # Fechar a conexão
        if conn:
            cursor.close()
            conn.close()

# Caminho para o arquivo Excel
caminho_arquivo = 'POSIÇÕES 2.xlsx'

# Converte o Excel para JSON
json_resultado = excel_para_json(caminho_arquivo)

# Cadastra os dados no banco
cadastrar_tracking(json_resultado)
