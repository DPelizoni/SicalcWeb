from datetime import datetime

import pyautogui

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

from lib.planilha_refatorada import read_excel, writer_excel


class SicalcWeb:

    def __init__(self):
        self.site = 'https://sicalc.receita.economia.gov.br/sicalc/rapido/contribuinte'
        self.navegador = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        self.navegador.implicitly_wait(20)
        self.navegador.get(self.site)
        self.SITE_MAP = {
            'buttons': {
                'continuar': {'xpath': '//*[@id="divBotoes"]/input[1]'},
                'calcular': {'xpath': '//*[@id="btnCalcular"]'},
                'darf': {'xpath': '//*[@id="btnDarf"]'},
                'retornar': {'xpath': '//*[@id="btnRetornar"]'}
            },
            'checkbox': {
                'selecionar_darf': {'xpath': '//*[@id="cts"]/tbody/tr/td[1]/input'}
            },
            'inputs': {
                'cod_receita': {'xpath': '//*[@id="codReceitaPrincipal"]'},
                'cnpj': {'xpath': '//*[@id="cnpjFormatado"]'}
            },
            'radio': {
                'pj': {'xpath': '//*[@id="optionPJ"]'}
            },
            'interacao': {
                'sou_humano': {'x': 77, 'y': 766},
                'info_darf': {'x': 269, 'y': 889},
                'tela_humano': {'x': 79, 'y': 880},
            },
        }

    @staticmethod
    def tempo_execucao():
        return datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    def sou_humano(self):
        pyautogui.click(self.SITE_MAP['interacao']['sou_humano']['x'], self.SITE_MAP['interacao']['sou_humano']['y'])

    @staticmethod
    def periodo_apuracao(apuracao):
        pyautogui.press('tab', presses=3, interval=0.2)
        pyautogui.write(apuracao)

    @staticmethod
    def valor_principal(valor):
        pyautogui.press('tab', presses=3, interval=0.2)
        pyautogui.write(valor)

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

    def xpath_click(self, path):
        self.navegador.find_elements(By.XPATH, path)[0].click()

    def xpath_escrever(self, path, texto):
        self.navegador.find_elements(By.XPATH, path)[0].send_keys(texto)

    def pessoa_juridica(self):
        self.xpath_click(self.SITE_MAP['radio']['pj']['xpath'])

    def informar_cnpj(self, cnpj):
        self.xpath_escrever(self.SITE_MAP['inputs']['cnpj']['xpath'], cnpj)

    def clicar_btn_continuar(self):
        self.xpath_click(self.SITE_MAP['buttons']['continuar']['xpath'])

    def informar_empresa(self, cnpj):
        pyautogui.sleep(1)
        self.pessoa_juridica()
        pyautogui.sleep(1)
        self.informar_cnpj(cnpj)
        pyautogui.sleep(1)
        self.sou_humano()
        self.clicar_btn_continuar()
        pyautogui.sleep(2)

    def informar_dados_darf(self, codigo):
        self.xpath_escrever(self.SITE_MAP['inputs']['cod_receita']['xpath'], codigo)
        pyautogui.click(self.SITE_MAP['interacao']['info_darf']['x'],
                        self.SITE_MAP['interacao']['info_darf']['y'],
                        duration=0.5)
        pyautogui.sleep(1)

    def clicar_btn_calcular(self):
        pyautogui.press('tab')
        pyautogui.sleep(1)
        self.xpath_click(self.SITE_MAP['buttons']['calcular']['xpath'])

    def selecionar_darf(self):
        pyautogui.sleep(1)
        self.xpath_click(self.SITE_MAP['checkbox']['selecionar_darf']['xpath'])

    def emitir_darf(self):
        pyautogui.sleep(1)
        self.xpath_click(self.SITE_MAP['buttons']['darf']['xpath'])

    def retornar(self):
        pyautogui.sleep(1)
        self.xpath_click(self.SITE_MAP['buttons']['retornar']['xpath'])

    def gerar_darf(self, arquivo, salvar, window):
        status = 'Ok'
        nome, cnpj, apuracao, codigo, valor, quantidade = read_excel(arquivo)
        nivel_progresso = 100 / quantidade
        barra_progresso = 100 / quantidade
        inicio = self.tempo_execucao()
        lista = []
        window['-STATUSBAR-'].update('Processando...', 'blue')

        for i in range(quantidade):
            self.informar_empresa(cnpj[i])
            self.informar_dados_darf(self.selecionar_codigo_darf(codigo[i]))
            self.periodo_apuracao(apuracao[i])
            self.valor_principal(valor[i])
            self.clicar_btn_calcular()
            self.selecionar_darf()
            self.emitir_darf()
            self.retornar()
            pyautogui.click(self.SITE_MAP['interacao']['tela_humano']['x'],
                            self.SITE_MAP['interacao']['tela_humano']['y'],
                            duration=1)
            fim = self.tempo_execucao()
            window['-PROGBAR-'].update(barra_progresso % 101)
            barra_progresso += nivel_progresso
            dados = {'Razão social': nome[i], 'CNPJ': cnpj[i], 'Apuração': apuracao[i], 'Código': codigo[i],
                     'Valor': valor[i], 'Status': status, 'Início': inicio, 'Fim': fim}
            lista.append(dados)
        writer_excel(salvar, lista)
        window['-STATUSBAR-'].update('Processo finalizado com sucesso!', 'green')


if __name__ == '__main__':
    pass
