import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import pyperclip

cnt_total = 0
cnt_pag_ok = 0
cnt_pag_pendente = 0
cnt_pag_parcial = 0
cnt_cancelado = 0

valor_1lote = "R$ 66,00"
valor_inf_1lote = "R$ 40,50"

arte = "c:\\Users\\fpfpa\\Documents\\DIJKSTRA\\whatsapp_automation\\arte.jpg"
img_pix_1lote = "c:\\Users\\fpfpa\\Documents\\DIJKSTRA\\whatsapp_automation\\pix_parcelado_1lote.png"
img_pix_inf_1lote = "c:\\Users\\fpfpa\\Documents\\DIJKSTRA\\whatsapp_automation\\pix_inf_parcelado_1lote.png"

pix_cc_1lote = "00020126500014BR.GOV.BCB.PIX0128emamor.csmcuritiba@gmail.com520400005303986540566.005802BR5919Acampamento Em Amor6008Curitiba621605121LOTEPIXPARC6304703D"
pix_cc_inf_1lote = "00020126500014BR.GOV.BCB.PIX0128emamor.csmcuritiba@gmail.com520400005303986540540.505802BR5919Acampamento Em Amor6008Curitiba621905151LOTEPIXINFPARC63045F1B"

workbook = openpyxl.load_workbook('acamps-2024-csm-curitiba - inscricoes.xlsx')
pagina_clientes = workbook['inscricoes']

print(f'codigo - nome - cancelado - celular - pagamento')

for linha in pagina_clientes.iter_rows(min_row=2, max_col=43):
    nome = linha[0].value
    codigo = linha[4].value
    cancelado = linha[6].value
    celular = linha[21].value
    tipo = linha[32].value
    pagamento = linha[33].value

    cnt_total += 1
    if pagamento == "Pendente 1/5":
        cnt_pag_parcial += 1
    if pagamento == "Pendente":
        cnt_pag_pendente += 1
    if pagamento == None:
        cnt_pag_ok += 1
    if cancelado == "Sim":
        cnt_cancelado += 1

    if cnt_total > 7 and cancelado == "Não" and pagamento == "Pendente":
        
        print(f'{codigo} - {nome} - {cancelado} - {celular} - {tipo} - {pagamento}')

        if tipo == "Adulto (+13 anos) - R$ 330,00":
            valor = valor_1lote
            pix = pix_cc_1lote
            img = img_pix_1lote
        elif tipo == "Infantil (7 a 12 anos) - R$ 202,50":
            valor = valor_inf_1lote
            pix = pix_cc_inf_1lote
            img = img_pix_inf_1lote
        else:
            valor = "Isento"
            pix = "Isento"
            img = ""

        msg = \
f'''Olá acampante,

{nome} - {codigo}

O pagamento da primeira parcela da sua inscrição, no valor de {valor}, se encontra pendente. Realize o pagamento até o dia 25/05/2024 para evitar que sua inscrição seja cancelada. Caso isso aconteça você precisara refazer sua inscrição pelo site.

Chave pix email: emamor.csmcuritiba@gmail.com

Pix copia e cola:
{pix}
'''

        celular_formatado = "55" + celular.replace("(","").replace(")","").replace("-","").replace(" ","")
        link_mensagem_whatsapp = f'https://api.whatsapp.com/send?phone={celular_formatado}'

        # abre o link da conversa no whatsapp
        webbrowser.open(link_mensagem_whatsapp)
        sleep(3)

        try:
            # anexando a arte do evento
            # anexar
            botaoAnexarlocation = pyautogui.locateOnScreen('anexar.png')
            botaoAnexarpoint = pyautogui.center(botaoAnexarlocation)
            pyautogui.click(botaoAnexarpoint.x, botaoAnexarpoint.y)
            sleep(1)

            # fotos e videos
            botaoFotolocation = pyautogui.locateOnScreen('foto.png')
            botaoFotopoint = pyautogui.center(botaoFotolocation)
            pyautogui.moveTo(botaoFotopoint.x, botaoFotopoint.y)
            sleep(0.5)
            pyautogui.click()
            pyautogui.click()
            sleep(3)
            
            # colar o caminho da arte e confirma
            pyperclip.copy(arte)
            pyautogui.hotkey("ctrl", "v")
            pyautogui.press('enter')
            sleep(2)

            # cola a mensagem e envia
            pyperclip.copy(msg)
            pyautogui.hotkey("ctrl", "v")
            pyautogui.press('enter')
            sleep(0.5)

            # anexando o qr-code pix
            # anexar
            pyautogui.click(botaoAnexarpoint.x, botaoAnexarpoint.y)
            sleep(1)

            # fotos e videos
            pyautogui.moveTo(botaoFotopoint.x, botaoFotopoint.y)
            sleep(0.5)
            pyautogui.click()
            pyautogui.click()
            sleep(3)

            # colar o caminho da foto e confirma
            pyperclip.copy(img)
            pyautogui.hotkey("ctrl", "v")
            pyautogui.press('enter')
            sleep(0.5)

            # envia
            pyautogui.press('enter')
            sleep(0.5)
            
            # seleciona o campo de mensagem
            # pyautogui.click(botaoAnexarpoint.x+200, botaoAnexarpoint.y)
            # sleep(0.5)

        except Exception as ex:
            print(ex)
            
        # volta para o navegador
        pyautogui.hotkey("alt", "tab")
        sleep(0.5)

        # fecha a ultima aba
        pyautogui.hotkey("ctrl", "w")
        sleep(0.5)

        # ret = pyautogui.confirm(text="Deseja continuar", title='Confirmação', buttons=['OK', 'Cancelar'])
        # if ret == 'Cancelar':
        #     break
        


# mostra o numero total de inscritos e o numero de mensagens enviadas
confirm_text = \
f'''\nTotal de inscritos: {cnt_total}
- ok: {cnt_pag_ok}
- pendente: {cnt_pag_pendente}
- pendente 1/5: {cnt_pag_parcial}
'''
with pyautogui.hold("alt"):
    pyautogui.press(['tab', 'tab'])
sleep(1)
pyautogui.confirm(text=confirm_text, title='Macro pagamentos', buttons=['OK'])