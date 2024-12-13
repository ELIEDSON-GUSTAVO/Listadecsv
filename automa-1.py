import pandas as pd
import re

# Função para validar o formato do código
def validate_codigo(value):
    return bool(re.match(r'^\d{2}\.\d{2}\.\d{2}\.\d{10}$', str(value)))

# Função para processar os dados
def process_data(df):
    # Definindo os títulos e a ordem das colunas para a planilha de exportação
    headers = ['ITENS', 'CODIGO', 'QND', 'DESCRIÇÃO', 'MASS', 'MATERIAL', 'LINK']
    error_log = []  # Lista para armazenar erros encontrados

    # Criar um DataFrame vazio com as colunas na ordem definida
    export_df = pd.DataFrame(columns=headers)

    # Processar cada linha da tabela de importação
    for index, row in df.iterrows():
        # Validar o formato do código
        if pd.notnull(row.get('CODIGO')) and not validate_codigo(row['CODIGO']):
            error_log.append(f"Erro no código na linha {index + 2}: '{row['CODIGO']}' não está no formato correto.")

        # Validar a presença de descrição
        if pd.isnull(row.get('DESCRIÇÃO')):
            error_log.append(f"Descrição ausente na linha {index + 2}.")

        # Criar uma linha para exportação com dados válidos
        new_row = {
            'ITENS': index + 1,  # Gerar sequencialmente, começando em 1
            'CODIGO': row.get('CODIGO', ''),
            'QND': row.get('QND', 0),
            'DESCRIÇÃO': row.get('DESCRIÇÃO', ''),
            'MASS': 0.1,  # Define "MASS" como 0.1
            'MATERIAL': '',  # "MATERIAL" fica vazio
            'LINK': ''  # "LINK" fica vazio
        }

        # Adicionar a nova linha ao DataFrame de exportação
        export_df = export_df.append(new_row, ignore_index=True)

    # Exibir os erros encontrados
    if error_log:
        print("Erros encontrados:")
        for error in error_log:
            print(error)

    return export_df, error_log

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
    export_df, error_log = process_data(df)

    # Mostrar a tabela processada
    print("\nTabela processada:")
    print(export_df)

    # Salvar os dados processados em Excel e CSV
    save_to_excel(export_df)
    save_to_csv(export_df)

    print("\nOs dados foram exportados para 'editado.xlsx' e 'editado.csv'.")

# Exemplo de uso
if __name__ == "__main__":
    input_file = 'COMPRADOS - 560 - ANTONIO LENUDO DE OLIVEIRA EIRELI - CARROCERIA ELETRICIARIO EM AÇO - VW - 17.210.xlsx'  # Substitua pelo caminho do seu arquivo
    main(input_file)
