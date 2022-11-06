import os
from datetime import datetime

from openpyxl.reader.excel import load_workbook

from lib.sicalc_web import SicalcWeb


class Planilha:

    def __init__(self, arquivo, planilha):
        self.arquivo = arquivo
        self.excel = load_workbook(self.arquivo)
        self.planilha = self.excel[planilha]

    @staticmethod
    def tempo():
        return datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    @staticmethod
    def transformar_str(numero):
        num_str = str(numero)
        return num_str.replace('.', ',')

    @staticmethod
    def transformar_data(data):
        return datetime.strftime(data, '%m/%Y')

    def selecionar_coluna(self, coluna):
        return self.planilha[coluna]

    def selecionar_valor_celula(self, coluna, linha):
        return self.planilha[f'{coluna}{linha}'].value

    def realizar_processo(self, coluna):
        return range(2, len(self.selecionar_coluna(coluna)) + 1)

    def salvar_arquivo(self):
        try:
            return self.excel.save(filename=self.arquivo)
        except PermissionError:
            return False

    def abrir_arquivo(self):
        os.startfile(self.arquivo)

    def preencher_celula(self, coluna, linha, valor):
        celula = self.planilha[f'{coluna}{linha}'] = valor
        return celula

    def nivel_progresso(self, coluna):
        celulas = len(self.selecionar_coluna(coluna)) - 1
        return 100 / celulas

    @staticmethod
    def selecionar_codigo_darf(codigo):
        codigo_0561 = 561
        codigo_0588 = 588
        codigo_8301 = 8301
        if codigo == codigo_0561:
            return '0561 - 07 - ME - a partir de 01/01/2008'
        elif codigo == codigo_0588:
            return '0588 - 06 - ME - a partir de 01/01/2008'
        elif codigo == codigo_8301:
            return '8301 - 02 - ME - a partir de 01/01/1997'

    def dados_planilha(self, col_cnpj, col_codigo, col_apuracao, col_valor, linha):
        cnpj = self.selecionar_valor_celula(col_cnpj, linha)
        codigo = self.selecionar_valor_celula(col_codigo, linha)
        apuracao = self.selecionar_valor_celula(col_apuracao, linha)
        valor = self.selecionar_valor_celula(col_valor, linha)
        return cnpj, codigo, apuracao, valor

    def status_planilha(self, col_status, col_ini, inicio, col_fim, linha, fim):
        self.preencher_celula(col_status, linha, 'Ok')
        self.preencher_celula(col_ini, linha, inicio)
        self.preencher_celula(col_fim, linha, fim)

    def main(self, col_cnpj, col_apuracao, col_valor, col_status, col_ini, col_fim, col_codigo, window):
        darf = SicalcWeb()
        nivel = self.nivel_progresso(col_cnpj)
        progresso_barra = self.nivel_progresso(col_cnpj)
        window['-STATUSBAR-'].update('Processando...', 'blue')

        for linha in self.realizar_processo(col_cnpj):
            # Marcador do tempo inicial do processo
            inicio = self.tempo()

            # Percorrer todas as células preenchidas
            cnpj, codigo, apuracao, valor = self.dados_planilha(col_cnpj, col_codigo, col_apuracao, col_valor, linha)

            # O processo de geração do DARF
            darf.gerar_darf(cnpj, self.selecionar_codigo_darf(codigo), self.transformar_data(apuracao),
                            self.transformar_str(valor))

            # Marcador do tempo final do processo
            fim = self.tempo()

            # Preencher planilha com o status, inicio e fim do processo
            self.status_planilha(col_status, col_ini, inicio, col_fim, linha, fim)

            window['-PROGBAR-'].update(progresso_barra % 101)
            progresso_barra += nivel

            self.salvar_arquivo()
        window['-STATUSBAR-'].update('Processo finalizado com sucesso!', 'green')
        self.abrir_arquivo()


if __name__ == '__main__':
    pass
