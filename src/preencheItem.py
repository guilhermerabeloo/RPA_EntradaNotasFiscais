from sql import sqlPool
from pywinauto.application import Application
from util import clicarEmImagem
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
        
        itensXml = Application(backend="win32").connect(title="Importação XML Nota Fiscal de Entrada.", timeout=60)
        main_window = itensXml.top_window()
        main_window.set_focus()

        time.sleep(1)
        pyautogui.press('TAB')
        time.sleep(1)

        for desenho in itens:
            desenhoCurrent = desenho[0]

            for _ in range(7):
                pyautogui.press('DELETE')
            pyautogui.write(desenhoCurrent)
            time.sleep(.5)

            try:
                atencao_app = Application(backend="win32").connect(title="Atenção", timeout=3)
                descricao = atencao_app.Atencao.children()[0].window_text()

                if 'Item Obsoleto - Política Especial' in descricao:
                    pyautogui.press('ENTER')
            except:
                pass

            pyautogui.press('down')
            time.sleep(.5)

        imagemUnidade = 'C:\\Users\\automacao\\Documents\\RPA_python\\RPA_EntradaNotasFiscais\\assets\\unidade.png'
        clicarEmImagem(imagemUnidade, 0)
        time.sleep(.5)
        pyautogui.press('TAB')
        time.sleep(.5)

        for pedido in itens:
            pedidoCurrent = pedido[2]

            pyautogui.write(pedidoCurrent)
            time.sleep(.5)
            pyautogui.press('down')
            time.sleep(.5)

            janelaAtencaoVisivel = False
            cont = 0
            while janelaAtencaoVisivel==False:
                time.sleep(1)
                cont+=1
                if cont > 3:
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

        imagemSistema = 'C:\\Users\\automacao\\Documents\\RPA_python\\RPA_EntradaNotasFiscais\\assets\\sistema.png'
        clicarEmImagem(imagemSistema, 0)
        time.sleep(.5)
        clicarEmImagem(imagemUnidade, 0)
        time.sleep(.5)

        for unidade in itens:
            unidadeCurrent = unidade[1]

            pyautogui.write(unidadeCurrent)
            time.sleep(.5)
            pyautogui.press('enter')
            time.sleep(.5)

        time.sleep(1)

    except TratamentoException as err:
        raise TratamentoException(f'Erro ao preencher itens: {err}')