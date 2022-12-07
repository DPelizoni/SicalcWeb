import pandas as pd


def read_excel(arquivo):
    """
    Abre o arquivo Excel, utilizando a biblíoteca pandas
    :param arquivo: Excel
    :return: list: CNPJ, Apuração, Código, Valor, quantidade(CNPJ)
    """
    # Abrir o arquivo em formato Excel
    df = pd.read_excel(arquivo)

    # Formata datetime em str. Ex: 12/2022
    df['Apuração'] = df['Apuração'].dt.strftime('%m/%Y')

    # Transforma float em str, substituindo o ponto pela virgula nas casas decimais
    df['Valor'] = df['Valor'].apply(str).replace('\\.', ',', regex=True)

    return list(df['Razão social']), list(df['CNPJ']), list(df['Apuração']), list(df['Código']), list(df['Valor']), len(
        list(df['CNPJ']))


def writer_excel(arquivo, dados):
    df = pd.DataFrame(dados)
    df.to_excel(arquivo, index=None, sheet_name='Concluído')


if __name__ == '__main__':
    pass
