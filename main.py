import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
from SiteUtils import SiteUtils
import os
import pygetwindow as gw


class ChatWhats:

    def __init__(self):
        self.table = 'clients.xlsx'
        self.page = 'Planilha1'
        self.siteutils = SiteUtils()

    def pegar_tabela_clients(self):
        workbook = openpyxl.load_workbook(self.table)
        return workbook[self.page]

    def mandar_mensagens(self):
        table_clients = self.pegar_tabela_clients()
        for linha in table_clients.iter_rows(min_row=2):
            nome = linha[0].value
            telefone = linha[1].value
            vencimento = linha[2].value
            try:
                self.siteutils.abrirMensagem(telefone, nome, vencimento)
                sleep(5)

                # Verifique a janela do WhatsApp
                janela = gw.getWindowsWithTitle('WhatsApp')
                if janela:
                    janela = janela[0]
                    x_janela, y_janela = janela.topleft
                    print(f'Janela encontrada: {
                          janela.title} - Coordenadas: {x_janela}, {y_janela}')
                else:
                    print('Janela do WhatsApp não encontrada.')
                    continue  # Passa para a próxima iteração
                screenshot = pyautogui.screenshot()
                # Salva uma captura de tela para análise
                screenshot.save('screenshot.png')
                print("Captura de tela salva como screenshot.png.")

                print('Tentando localizar a seta...')
                seta = pyautogui.locateCenterOnScreen(
                    'seta.png', confidence=0.8)
                print(f'Seta localizada: {seta}')  # Verifique as coordenadas

                sleep(2)  # Espera um pouco mais
                if seta:
                    pyautogui.click(x_janela + seta[0], y_janela + seta[1])
                else:
                    print(f'Seta não encontrada para {nome}.')

                sleep(2)
                pyautogui.hotkey('ctrl', 'w')
                sleep(2)
            except Exception as e:
                print(f'Não foi possível enviar mensagem para {nome}: {e}')
                with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                    arquivo.write(f'{nome},{telefone}{os.linesep}')


def main():
    chatWhats = ChatWhats()
    chatWhats.mandar_mensagens()  # Chame a coroutine com await


# Executa a função main
if __name__ == "__main__":
    main()
