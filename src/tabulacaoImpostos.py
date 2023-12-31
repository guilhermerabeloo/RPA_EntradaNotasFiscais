from pywinauto.application import Application
from sql import sqlPool
import pyautogui
import time

def tabulaItens(idNota):
    app = Application(backend="win32").connect(class_name="FNWND3115")
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

        app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title=f"{notaAtual}", class_name="PBEDIT115", found_index=0).click_input()
        time.sleep(.3)
        for i in range(15):
            time.sleep(.1)
            pyautogui.press('BACKSPACE')
        pyautogui.write(desenho)

        for i in range(3):
            time.sleep(.3)
            pyautogui.press('TAB')
        pyautogui.press('SPACE')

        time.sleep(3)
        pyautogui.hotkey('alt', 'I')
        time.sleep(2)

        for i in range(50):
            time.sleep(.1)
            pyautogui.press('TAB')

        pyautogui.hotkey('alt', 'O')
        time.sleep(1)
        pyautogui.hotkey('alt', 'C')
        notaAtual = desenho

    app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="&V", class_name="Button", found_index=0).click_input()
