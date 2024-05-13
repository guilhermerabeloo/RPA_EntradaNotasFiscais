from sql import sqlPool
from pywinauto.application import Application
import pyautogui
import time

def preencheItem(idNota):
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

            quantidade_tab = 2 if unidade_divergente == '0' else 3

            for i in range(quantidade_tab):
                pyautogui.press('TAB')

            time.sleep(.5)
            pyautogui.write(pedido)

            try: 
                atencao = Application(backend="win32").connect(title="Atenção", timeout=3)
                texto = atencao.Atencao.children()[0].window_text()

                if "Pedido" in texto:
                    print('erro no pedido')

                    raise Exception(f'Pedido {pedido} inválido')
            except:
                pass

    except Exception as err:
        raise Exception(f'Erro ao preencher itens: {err}')