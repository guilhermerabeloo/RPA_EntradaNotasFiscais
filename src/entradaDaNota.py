from pywinauto.application import Application
import pyautogui
import time

def preenchimentoCapaNota(caminhoDoArquivo):
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
    time.sleep(1)
    telaImportar.ExportarImportarXmlNfe.child_window(title="&Importar", class_name="Button").wrapper_object().click_input()
    time.sleep(3)
    importar = Application(backend="win32").connect(title="Importação XML Nota Fiscal de Entrada.")
    importar_window = importar.top_window()
    importar_window.set_focus()
    time.sleep(1)
    importar.ImportacaoXmlNotaFiscalDeEntrada.child_window(class_name="Edit").wrapper_object().click_input()


# preenchimentoCapaNota("C:\\Users\\guilherme.rabelo\\Documents\\RPA_docs\\EntradaDaNota\\arquivo.xml")