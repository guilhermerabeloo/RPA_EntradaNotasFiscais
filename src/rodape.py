from pywinauto.application import Application
from sql import sqlPool
import pyautogui
import time

def preencheRodape(idNota):
    parcelas = sqlPool("SELECT", f"""
            SELECT 
                numero_parcela,
                REPLACE(CONVERT(VARCHAR(50), valor_parcela), '.', ',') as valor,
                CONVERT(VARCHAR(10), vencimento, 103) as vencimento,
                tipo_obrigacao,
                moeda_correcao,
                moeda_juros,
                agente_portador
            FROM nfemaster.DWIN_entradaNFeProdutoXML_rodape
            WHERE
                id_nota = {idNota}
        """)

    app = Application(backend="win32").connect(class_name="FNWND3115", timeout=60)
    main_window = app.top_window()
    main_window.set_focus()
    time.sleep(1)

    app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="&R", class_name="Button").wrapper_object().click_input()
    time.sleep(1)
    app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="Pedido OK", class_name="Button").click_input()
    time.sleep(.01)
    app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="Pedido OK", class_name="Button").click_input()
    time.sleep(.1) 
    pyautogui.press('TAB')
    time.sleep(.01)
    pyautogui.press('SPACE')
    time.sleep(3)
    app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="&E", class_name="Button").click_input()
    
    time.sleep(.01)
    pyautogui.press('TAB')
    time.sleep(.01)
    pyautogui.press('SPACE')

    time.sleep(5)

    for parcela in parcelas:
        valor = parcela[1]
        vencimento = parcela[2]
        tipoObrigacao = parcela[3]
        moedaCorrecao = parcela[4]
        moedaJuros = parcela[5]
        agentePortador = parcela[6]

        time.sleep(.1) 
        pyautogui.press('TAB')
        time.sleep(.1) 
        pyautogui.write(tipoObrigacao)
        time.sleep(.1) 
        pyautogui.press('TAB')
        time.sleep(.1) 
        pyautogui.press('TAB')
        time.sleep(.1) 
        pyautogui.write(vencimento)
        time.sleep(.1) 
        pyautogui.press('TAB')
        time.sleep(.1) 
        pyautogui.write(valor)
        for _ in range(4):
            time.sleep(.1)
            pyautogui.press('TAB')
        time.sleep(.1) 
        pyautogui.write(moedaCorrecao)
        time.sleep(.1) 
        pyautogui.press('TAB')
        time.sleep(.1) 
        pyautogui.press('TAB')
        time.sleep(.1) 
        pyautogui.write(moedaJuros)
        time.sleep(.1) 
        pyautogui.press('TAB')
        time.sleep(.1) 
        print(agentePortador)
        pyautogui.write(agentePortador)
        time.sleep(.1) 
        pyautogui.hotkey('alt', 'O')
        time.sleep(.1) 
        pyautogui.press('TAB')
        time.sleep(.1) 
        pyautogui.press('TAB')

    time.sleep(.1) 
    pyautogui.hotkey('alt', 'v')

    time.sleep(1)
    app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[29].click_input()
    time.sleep(1)
    pyautogui.hotkey('alt', 'v')
