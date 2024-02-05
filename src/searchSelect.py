import time
import pyautogui

def selecionaOptionSelect(campo, valorProcurado):
    valorEncontrado = False
    cont = 0
    campo.click_input()
    time.sleep(1)
    pyautogui.write(valorProcurado[:1])
    while valorEncontrado==False:
        valorSelecionado = campo.window_text()
        valorFormatado = valorSelecionado[:-2].strip()

        if valorFormatado==valorProcurado:
            time.sleep(1)
            pyautogui.press('TAB')

            valorEncontrado = True
        else:
            pyautogui.press('down')
        time.sleep(.5)
        cont+=1
        if cont == 150:
            break


