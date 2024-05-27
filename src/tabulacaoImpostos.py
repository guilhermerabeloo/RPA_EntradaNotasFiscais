from pywinauto.application import Application
from sql import sqlPool
import pyautogui
import time

def tabulaItens(idNota, logger, TratamentoException):
    try:
        app = Application(backend="win32").connect(class_name="FNWND3115", timeout=60)
        main_window = app.top_window()
        main_window.set_focus()
        time.sleep(1)

        app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="&C", class_name="Button", found_index=1).click_input()
        time.sleep(3)

        itens = sqlPool("SELECT", F"""
                SELECT 
                    desenho
                FROM nfemaster.DWIN_entradaNFeProdutoXML_itens
                WHERE
                    id_nota = {idNota}
            """)

        notaAtual = ""
        cont = 0 
        for item in itens:
            desenho =  item[0]

            app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title=f"{notaAtual}", class_name="PBEDIT115", found_index=0).double_click()
            time.sleep(.1)
            pyautogui.press('DELETE')
            time.sleep(.1)
            pyautogui.write(desenho)

            for _ in range(3):
                time.sleep(.1)
                pyautogui.press('TAB')
            pyautogui.press('SPACE')

            time.sleep(3)
            pyautogui.hotkey('alt', 'I')
            time.sleep(2)

            for i in range(50):
                time.sleep(.2)
                pyautogui.press('TAB')

            time.sleep(2)
            pyautogui.hotkey('alt', 'O')

            try:
                telaAtencao = Application(backend="win32").connect(title="Atenção", timeout=5)
                time.sleep(1)

                telaAtencao.Atencao.child_window(title="Não existe o NBM cadastrado.", class_name="Edit").click_input()
                pyautogui.press('ENTER')
                time.sleep(2)
                pyautogui.hotkey('alt', 'c')

                time.sleep(1)
            except Exception as e:
                logger.error(f'NBM não cadastrado para o produto {desenho}')
                pass
            time.sleep(2)

            janelaAtencaoVisivel = False
            contAux = 0
            while janelaAtencaoVisivel==False:
                time.sleep(1)
                contAux+=1
                if contAux > 3:
                    break

                try:
                    Application(backend="win32").connect(title="Redução Base Cálculo")
                    janelaAtencaoVisivel = True
                except:
                    pass

            if janelaAtencaoVisivel:
                atencao_app = Application(backend="win32").connect(title="Redução Base Cálculo")
                descricao = atencao_app.ReducaoBaseCalculo.children()[0].window_text()
                if "Valor de Redução da Base de Cálculo do ICMS inválido" in descricao:
                    raise TratamentoException(f'Erro na natureza de operação: Redução da BC do ICMS inválida')

            pyautogui.hotkey('alt', 'O')

            try:
                telaAtencao = Application(backend="win32").connect(title="Atenção", timeout=5)
                time.sleep(2)

                infoTelaAtencao = telaAtencao.Atencao.child_window(title="Valor digitado do ICMS não retido pelo Fornecedor difere do Valor calculado. Confirma o valor digitado?", class_name="Edit")
                if infoTelaAtencao.is_visible():
                    telaAtencao.Atencao.child_window(title="&Sim", class_name="Button").click_input()

            except Exception as e:
                pass
            notaAtual = desenho
            cont+=1

        rangeTab = 7 if cont > 1 else 6
        for _ in range(rangeTab):
            time.sleep(.1)
            pyautogui.press('TAB')
        pyautogui.press('SPACE')
    except TratamentoException as err:
        raise TratamentoException(err)