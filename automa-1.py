import pandas as pd
import re

# Função para validar o formato do código
def validate_codigo(value):
    return bool(re.match(r'^\d{2}\.\d{2}\.\d{2}\.\d{10}$', str(value)))

# Função para processar os dados
def process_data(df):
    headers = ['ITENS', 'CODIGO', 'QND', 'DESCRIÇÃO', 'MASS', 'MATERIAL', 'LINK']
    error_log = []  # Lista para armazenar erros encontrados

    # Verificando e criando colunas necessárias
    for header in headers:
        if header not in df.columns:
            df[header] = None

    # Processar cada linha da tabela
    for index, row in df.iterrows():
        # Validar e formatar as colunas
        if pd.notnull(row['CODIGO']) and not validate_codigo(row['CODIGO']):
            error_log.append(f"Erro no código na linha {index + 2}: '{row['CODIGO']}' não está no formato correto.")

        if pd.isnull(row['DESCRIÇÃO']):
            error_log.append(f"Descrição ausente na linha {index + 2}.")

        df.at[index, 'MASS'] = 0.1  # Define "MASS" como 0.1
        df.at[index, 'MATERIAL'] = ''  # Garante que "MATERIAL" fique vazio
        df.at[index, 'LINK'] = ''  # Garante que "LINK" fique vazio

        # Se "ITENS" não estiver presente, preencher com números sequenciais
        if pd.isnull(row['ITENS']):
            df.at[index, 'ITENS'] = index + 1  # Preenche com número sequencial

    # Exibindo os erros encontrados
    if error_log:
        print("Erros encontrados:")
        for error in error_log:
            print(error)

    return df, error_log

# Função para salvar o dataframe em Excel
def save_to_excel(df, file_name='editado.xlsx'):
    df.to_excel(file_name, index=False)

# Função para salvar o dataframe em CSV
def save_to_csv(df, file_name='editado.csv'):
    df.to_csv(file_name, index=False)

# Função principal para importar, processar e exportar os dados
def main(input_file):
    # Carregar o arquivo Excel
    try:
        df = pd.read_excel(input_file)
        print("Planilha carregada com sucesso.")
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return

    # Processar os dados
    df_processed, error_log = process_data(df)

    # Mostrar a tabela processada
    print("\nTabela processada:")
    print(df_processed)

    # Salvar os dados processados em Excel e CSV
    save_to_excel(df_processed)
    save_to_csv(df_processed)

    print("\nOs dados foram exportados para 'editado.xlsx' e 'editado.csv'.")

# Exemplo de uso
if __name__ == "__main__":
    input_file = 'COMPRADOS - 560 - ANTONIO LENUDO DE OLIVEIRA EIRELI - CARROCERIA ELETRICIARIO EM AÇO - VW - 17.210.xlsx'  # Substitua pelo caminho do seu arquivo
    main(input_file)
