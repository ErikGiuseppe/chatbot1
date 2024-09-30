import webbrowser
from urllib.parse import quote


class SiteUtils:
    def abirWhats(self):
        webbrowser.open('https://web.whatsapp.com/')

    def abrirMensagem(self, telefone, nome, vencimento):
        mensagem = f'Olá {nome} seu boleto vence no dia {vencimento}'
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={
            telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
