import PySimpleGUI as sg

from lib.planilha import Planilha


class Interface:

    def __init__(self):
        pass

    @staticmethod
    def layout_arquivo():
        layout = [[sg.Input(k='-ARQUIVO-'),
                   sg.FileBrowse('Excel', target='-ARQUIVO-', file_types=(('All files', '*.xlsx'),))]]
        return layout

    @staticmethod
    def layout_planilha():
        layout = [[sg.Input(k='-PLANILHA-', s=(30, 1), default_text='DARF')]]
        return layout

    @staticmethod
    def layout_colunas_dados():
        layout = [[sg.T('CNPJ'), sg.I(k='-CNPJ-', s=(3, 1), justification='c', default_text='C'),
                   sg.T('Apuração'), sg.I(k='-APURACAO-', s=(3, 1), justification='c', default_text='D'),
                   sg.T('Código'), sg.I(k='-CODIGO-', s=(3, 1), justification='c', default_text='E'),
                   sg.T('Valor'), sg.I(k='-VALOR-', s=(3, 1), justification='c', default_text='F')]]
        return layout

    @staticmethod
    def layout_colunas_status():
        layout = [[sg.T('Status'), sg.I(k='-STATUS-', s=(3, 1), justification='c', default_text='G'),
                   sg.T('Início'), sg.I(k='-INICIO-', s=(3, 1), justification='c', default_text='H'),
                   sg.T('Fim'), sg.I(k='-FIM-', s=(3, 1), justification='c', default_text='I')]]
        return layout

    @staticmethod
    def planilhas(values):
        try:
            planilha = Planilha(values['-ARQUIVO-'], values['-PLANILHA-'])
            return planilha
        except TypeError:
            return False

    @staticmethod
    def gerar_darf(planilha, values, window):
        planilha.main(values['-CNPJ-'], values['-APURACAO-'], values['-VALOR-'],
                      values['-STATUS-'], values['-INICIO-'], values['-FIM-'], values['-CODIGO-'], window)

    def layout_main(self):
        layout = [[sg.Frame('Arquivo', self.layout_arquivo())],
                  [sg.Frame('Planilha', self.layout_planilha())],
                  [sg.Frame('Colunas Dados', self.layout_colunas_dados())],
                  [sg.Frame('Colunas Status', self.layout_colunas_status())],
                  [sg.ProgressBar(100, k='-PROGBAR-', s=(35, 20))],
                  [sg.B('Gerar'), sg.B('Sair')],
                  [sg.StatusBar(' ' * 80, k='-STATUSBAR-')]]
        return layout

    def janela(self):
        return sg.Window('Recalculo DARF', self.layout_main(), keep_on_top=True, finalize=True, location=(1076, 372))

    def main(self):
        window = self.janela()
        while True:
            event, values = window.read()
            planilha = self.planilhas(values)
            if event in (sg.WINDOW_CLOSED, 'Sair'):
                break
            elif event == 'Gerar':
                self.gerar_darf(planilha, values, window)
        window.close()


if __name__ == '__main__':
    pass
