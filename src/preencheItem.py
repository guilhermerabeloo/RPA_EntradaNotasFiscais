from sql import sqlPool
from pywinauto.application import Application
import pyautogui
import time

def preencheItem(idNota, TratamentoException):
    try:
        itens = sqlPool("SELECT", f"""
                        SELECT 
                            desenho,
                            unidade,
                            ZeevIntegration.dbo.ufPreencheZerosField7(pedido) as pedido,
                            unidade_divergente
                        FROM nfemaster.DWIN_entradaNFeProdutoXML_itens
                        WHERE
                            id_nota = {idNota}
                    """)
        
        for item in itens:
            itensXml = Application(backend="win32").connect(title="Importação XML Nota Fiscal de Entrada.", timeout=60)
            main_window = itensXml.top_window()
            main_window.set_focus()
            desenho =  item[0]
            unidade =  item[1]
            pedido =  item[2]
            unidade_divergente = item[3]

            time.sleep(1)
            pyautogui.press('TAB')
            time.sleep(1)
            for i in range(7):
                pyautogui.press('DELETE')
            pyautogui.write(desenho)
            time.sleep(.5)

            quantidade_tab = 1 if unidade_divergente == '0' else 2

            for i in range(quantidade_tab):
                time.sleep(1)
                pyautogui.press('TAB')

            time.sleep(.5)
            pyautogui.write(unidade)
            pyautogui.press('TAB')
            time.sleep(1)
            pyautogui.write(pedido)

            janelaAtencaoVisivel = False
            cont = 0
            while janelaAtencaoVisivel==False:
                time.sleep(1)
                cont+=1
                if cont > 5:
                    break

                try:
                    Application(backend="win32").connect(title="Atenção")
                    janelaAtencaoVisivel = True
                except:
                    pass

            if janelaAtencaoVisivel:
                atencao_app = Application(backend="win32").connect(title="Atenção")
                descricao = atencao_app.Atencao.children()[0].window_text()
                if "Produto não está presente no pedido" in descricao:
                    raise TratamentoException(f'Pedido inválido')

    except TratamentoException as err:
        raise TratamentoException(f'Erro ao preencher itens: {err}')