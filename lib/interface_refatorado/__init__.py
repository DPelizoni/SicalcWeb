import PySimpleGUI as sg

from lib.sicalc_web_refatorado import SicalcWeb


class Interface:

    def __init__(self):
        pass

    @staticmethod
    def layout_arquivo():
        layout = [[sg.Input(k='-ARQUIVO-'),
                   sg.FileBrowse('Excel', target='-ARQUIVO-', file_types=(('All files', '*.xlsx'),))]]
        return layout

    @staticmethod
    def layout_arquivo_save():
        layout = [[sg.Input(k='-SALVAR-'),
                   sg.FileSaveAs('Excel', target='-SALVAR-', file_types=(('All files', '*.xlsx'),))]]
        return layout

    def layout_main(self):
        layout = [[sg.Frame('Arquivo', self.layout_arquivo())],
                  [sg.Frame('Salvar', self.layout_arquivo_save())],
                  [sg.B('Gerar'), sg.B('Sair')],
                  [sg.ProgressBar(100, k='-PROGBAR-', s=(35, 20))],
                  [sg.StatusBar(' ' * 80, k='-STATUSBAR-')]]
        return layout

    def janela(self):
        return sg.Window('DARF Sicalc Web', self.layout_main(), keep_on_top=True, finalize=True, location=(1076, 372))

    def main(self):
        window = self.janela()
        while True:
            event, values = window.read()
            if event in (sg.WINDOW_CLOSED, 'Sair'):
                break
            elif event == 'Gerar':
                sicalc = SicalcWeb()
                sicalc.gerar_darf(values['-ARQUIVO-'], values['-SALVAR-'], window)
        window.close()


if __name__ == '__main__':
    a = Interface()
    a.main()
