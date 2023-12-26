from pywinauto.application import Application
import time
import pyautogui

def selecionarEmpresa(empresa):
    app = Application(backend="win32").connect(class_name="FNWND3115")
    main_window = app.top_window()
    main_window.set_focus()
    time.sleep(5)

    pyautogui.hotkey('alt', 'p')
    time.sleep(.5)
    for i in range(6):
        pyautogui.press('up')
        time.sleep(.01)
    time.sleep(1)
    pyautogui.press('ENTER')
    time.sleep(5)

    telaEmpresa = Application(backend="win32").connect(title='Segurança')
    time.sleep(.5)
    telaEmpresa.Seguranca.set_focus()
    time.sleep(2)
    empresaEncontrada = False
    cont = 0
    while empresaEncontrada==False:
        empresaSelecionada = telaEmpresa.Seguranca.children()[2].window_text()
        empresaFormatada = empresaSelecionada[:-2].strip()

        if empresaFormatada==empresa:
            empresaEncontrada = True
        else:
            if cont <= 25:
                pyautogui.press('down')
            else: 
                pyautogui.press('up')
        time.sleep(.5)
        cont+=1
        if cont == 50:
            break
    telaEmpresa.Seguranca.child_window(title="&OK", class_name="Button").wrapper_object().click_input()
    time.sleep(5)
                      