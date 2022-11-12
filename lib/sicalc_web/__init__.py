import pyautogui as pag

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager


class SicalcWeb:

    def __init__(self):
        self.site = 'https://sicalc.receita.economia.gov.br/sicalc/rapido/contribuinte'
        self.navegador = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
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
            }
        }

    @staticmethod
    def clicar(x, y):
        pag.click(x, y)

    @staticmethod
    def teclar(tecla):
        pag.press(tecla)

    @staticmethod
    def escrever(texto):
        pag.typewrite(texto)

    def xpath_click(self, path):
        self.navegador.find_elements(By.XPATH, path)[0].click()

    def xpath_escrever(self, path, texto):
        self.navegador.find_elements(By.XPATH, path)[0].send_keys(texto)

    def pessoa_juridica(self):
        self.xpath_click(self.SITE_MAP['radio']['pj']['xpath'])

    def informar_cnpj(self, cnpj):
        self.xpath_escrever(self.SITE_MAP['inputs']['cnpj']['xpath'], cnpj)

    def sou_humano(self):
        self.clicar(77, 766)

    def clicar_btn_continuar(self):
        self.xpath_click(self.SITE_MAP['buttons']['continuar']['xpath'])

    def informar_empresa(self, cnpj):
        pag.sleep(1)
        self.pessoa_juridica()
        pag.sleep(1)
        self.informar_cnpj(cnpj)
        pag.sleep(1)
        self.sou_humano()
        pag.sleep(1)
        self.clicar_btn_continuar()
        pag.sleep(1)

    def informar_dados_darf(self, codigo):
        self.xpath_escrever(self.SITE_MAP['inputs']['cod_receita']['xpath'], codigo)
        pag.sleep(1)
        pag.click(269, 889)
        pag.sleep(1)

    def periodo_apuracao(self, apuracao):
        [self.teclar('tab') for _ in range(3)]
        self.escrever(apuracao)

    def valor_principal(self, valor):
        [self.teclar('tab') for _ in range(3)]
        self.escrever(valor)

    def clicar_btn_calcular(self):
        self.teclar('tab')
        pag.sleep(1)
        self.xpath_click(self.SITE_MAP['buttons']['calcular']['xpath'])

    def selecionar_darf(self):
        pag.sleep(1)
        self.xpath_click(self.SITE_MAP['checkbox']['selecionar_darf']['xpath'])

    def emitir_darf(self):
        pag.sleep(1)
        self.xpath_click(self.SITE_MAP['buttons']['darf']['xpath'])

    def retornar(self):
        pag.sleep(1)
        self.xpath_click(self.SITE_MAP['buttons']['retornar']['xpath'])

    def gerar_darf(self, cnpj, codigo, apuracao, valor):
        self.informar_empresa(cnpj)
        self.informar_dados_darf(codigo)
        self.periodo_apuracao(apuracao)
        self.valor_principal(valor)
        self.clicar_btn_calcular()
        self.selecionar_darf()
        self.emitir_darf()
        self.retornar()
        pag.sleep(2)
        self.clicar(79, 880)


if __name__ == '__main__':
    pass
