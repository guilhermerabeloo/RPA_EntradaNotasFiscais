from pywinauto.application import Application
from preencheItem import preencheItem
import pyautogui
import time

def preenchimentoCapaNota(caminhoDoArquivo, idNota, naturezaOperacao, documento, departamento, almoxarifado):
    # Colocando janela em foco
    app = Application(backend="win32").connect(class_name="FNWND3115")
    main_window = app.top_window()
    main_window.set_focus()
    time.sleep(1)

    # Entrando no módulo de emissao da nota
    pyautogui.hotkey('alt', 'm')
    time.sleep(.5)
    pyautogui.press('down')
    time.sleep(.5)
    pyautogui.press('ENTER')
    time.sleep(1)
    btnSair = app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="&S", class_name="Button", found_index=0)
    btnSair.wait('visible', timeout=10)

    # Entrando na tela de importacao do XML
    for i in range(2):
        pyautogui.press('TAB')
        time.sleep(.01)
    pyautogui.press('ENTER')
    
    time.sleep(3)
    telaImportar = Application(backend="win32").connect(title="Exportar/Importar .XML NFE(w_imp_exp_nfe)")
    telaImportar.ExportarImportarXmlNfe.child_window(title="Importar Nota de Produtos para Janela de Cadastro", class_name="Button").wrapper_object().click_input()
    time.sleep(.5)
    telaImportar.ExportarImportarXmlNfe.child_window(class_name="Edit").type_keys(caminhoDoArquivo)
    time.sleep(1)
    telaImportar.ExportarImportarXmlNfe.child_window(title="&Carregar", class_name="Button").wrapper_object().click_input()
    time.sleep(7)
    elementosExportarImportar = telaImportar.ExportarImportarXmlNfe.children()
    elementosExportarImportar[8].click_input()
    for i in range(3):
        time.sleep(.01)
        pyautogui.press('BACKSPACE')
    time.sleep(.01)
    pyautogui.press('DELETE')
    time.sleep(.01)
    pyautogui.write(naturezaOperacao)
    time.sleep(1)
    pyautogui.press('TAB')
    time.sleep(.5)
    pyautogui.write(documento)
    time.sleep(.5)
    pyautogui.press('TAB')
    time.sleep(.5)
    pyautogui.write(departamento)
    time.sleep(.5)
    pyautogui.press('TAB')
    time.sleep(.5)
    pyautogui.write(almoxarifado)
    time.sleep(.5)

    telaImportar.ExportarImportarXmlNfe.child_window(title="&Importar", class_name="Button").wrapper_object().click_input()

    time.sleep(8)
    importar = Application(backend="win32").connect(title="Importação XML Nota Fiscal de Entrada.")
    importar_window = importar.top_window()
    importar_window.set_focus()
    time.sleep(1)

    # preenchendo informacao do item
    # preencheItem(idNota)
    time.sleep(1)
    importar.importacaoXmlNotaFiscalDeEntrada.child_window(title="OK", class_name="Button").wrapper_object().click_input()
    
    btnSair.wait('visible', timeout=10)

    app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="Conta Gerencial:", class_name="Button").click_input()
    time.sleep(.01)
    app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="2.01.02.   ", class_name="PBEDIT115").click_input()
    time.sleep(.01)
    pyautogui.press('TAB')
    time.sleep(.01)
    pyautogui.hotkey('ctrl', 'end')
    time.sleep(.01)
    pyautogui.write(' - REALIZADO POR RPA')

