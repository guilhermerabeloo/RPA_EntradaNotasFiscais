from pywinauto.application import Application
import pyautogui
import time

def preenchimentoModoPgt(infoPagamento):
    rodape = Application(backend="win32").connect(class_name="FNWND3115", timeout=5)
    campos = rodape.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()

    pyautogui.write(infoPagamento['modPagamento'])
    time.sleep(.3) 
    pyautogui.press('TAB')

    if infoPagamento['tipoPrenchimento'] == 1: # Conta bancaria
        pyautogui.write(infoPagamento['tipoConta'])
        time.sleep(.3) 
        pyautogui.press('TAB')

        pyautogui.write(infoPagamento['banco'])
        time.sleep(.3) 
        pyautogui.press('TAB')

        pyautogui.write(infoPagamento['agencia'])
        time.sleep(.3) 
        pyautogui.press('TAB')

        pyautogui.write(infoPagamento['conta'])
        time.sleep(.3) 
        pyautogui.press('TAB')

        pyautogui.write(infoPagamento['indent'])
        time.sleep(.3) 
        pyautogui.press('TAB')

    elif infoPagamento['tipoPrenchimento'] == 2:  # Pix QR code
        campos[20].click_input()
        time.sleep(3)
        pyautogui.write(infoPagamento['tipoPixQrcode'])
        time.sleep(.3) 

        campos[25].double_click()
        time.sleep(.3) 
        pyautogui.press('DELETE')
        time.sleep(.3) 
        pyautogui.write(infoPagamento['qrcoodePix'])
        time.sleep(.3) 

    elif infoPagamento['tipoPrenchimento'] == 3:  # Pix chave
        campos[20].click_input()
        pyautogui.write(infoPagamento['tipoPix'])
        time.sleep(.3) 

        campos[24].click_input()
        pyautogui.write(infoPagamento['chavePix'])
        time.sleep(.3)
        
        campos[27].click_input()
        pyautogui.write(infoPagamento['transfPix'])
        time.sleep(.3) 

    elif infoPagamento['tipoPrenchimento'] == 3:  # Codigo de barras
        pyautogui.write(infoPagamento['codigoBarras'])
        time.sleep(.3) 
        pyautogui.press('TAB')

    campos[43].click_input()

# info = {'tipoPrenchimento': 2, 'modPagamento': 'PIX QR-CODE OUTROS BANCOS', 'tipoConta': None, 'banco': None, 'agencia': None, 'conta': None, 'indent': None, 'codigoBarras': None, 'tipoPix': None, 'chavePix': None, 'transfPix': None, 'tipoPixQrcode': 'Estático', 'qrcoodePix': '475514145c54c4sa1sc'}

# preenchimentoModoPgt(info)