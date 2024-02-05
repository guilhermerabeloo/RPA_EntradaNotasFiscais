from pywinauto.application import Application
import pyautogui
import time
import subprocess

def loginDealernet(modulo, empresa, senha):
    try:
        comando_powershell = f"start {modulo}"
        subprocess.run(["powershell", "-Command", comando_powershell], capture_output=True, text=True)

        telaLogin = Application(backend="win32").connect(title='Segurança', timeout=60)
        time.sleep(2)
        telaLogin.Seguranca.set_focus()

        telaLogin.Seguranca.child_window(class_name="Edit").wrapper_object().click_input()
        pyautogui.write(senha)
        time.sleep(2)

        pyautogui.press('TAB')
        time.sleep(.1)
        pyautogui.press('TAB')
        time.sleep(.1)
        pyautogui.write(empresa)

        time.sleep(1)
        telaLogin.Seguranca.child_window(title="&OK", class_name="Button").wrapper_object().click_input()
        time.sleep(10)
    except Exception as err:
        raise Exception(f'Erro ao fazer login: {err}')