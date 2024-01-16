from sql import sqlPool
import pyautogui

def preencheItem(idNota):
    try:
        itens = sqlPool("SELECT", f"""
                        SELECT 
                            desenho,
                            unidade,
                            pedido
                        FROM nfemaster.DWIN_entradaNFeProdutoXML_itens
                        WHERE
                            id_nota = {idNota}
                    """)
        
        for item in itens:
            desenho =  item[0]
            unidade =  item[1]
            pedido =  item[2]
            
            pyautogui.press('TAB')

            # preenchimento de desenho
            for i in range(7):
                pyautogui.press('DELETE')
            pyautogui.write(desenho)
            pyautogui.press('TAB')

            # preenchimento de unidade
            for i in range(30):
                pyautogui.press('up')
            pyautogui.write(unidade)
            pyautogui.press('TAB')

            # preenchimento de pedido
            # pyautogui.write(pedido)
            pyautogui.write('0000000')
    except Exception as err:
        raise Exception(f'Erro ao preencher itens: {err}')