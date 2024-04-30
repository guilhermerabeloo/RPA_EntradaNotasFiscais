from pywinauto.application import Application
from sql import sqlPool
import pyautogui
import time

def tabulaItens(idNota, logger):
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
            pyautogui.hotkey('alt', 'O') #ATIVAR ESSA LINHA EM PRODUCAO

            try:
                telaAtencao = Application(backend="win32").connect(title="Atenção", timeout=5)
                time.sleep(2)

                infoTelaAtencao = telaAtencao.Atencao.child_window(title="Valor digitado do ICMS não retido pelo Fornecedor difere do Valor calculado. Confirma o valor digitado?", class_name="Edit")
                if infoTelaAtencao.is_visible():
                    telaAtencao.Atencao.child_window(title="&Sim", class_name="Button").click_input()

            except Exception as e:
                pass
            notaAtual = desenho

        for _ in range(7):
            time.sleep(.1)
            pyautogui.press('TAB')
        pyautogui.press('SPACE')
    except Exception as err:
        raise Exception(f'Erro ao tabular impostos: {err}')