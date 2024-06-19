from pywinauto.application import Application
from preencheItem import preencheItem
from searchSelect import selecionaOptionSelect
from util import formataDepartamento
import pyautogui
import time


def preenchimentoCapaNota(caminhoDoArquivo, idNota, naturezaOperacao, documento, departamento, almoxarifado, TratamentoException):
    try:
        # Colocando janela em foco
        app = Application(backend="win32").connect(class_name="FNWND3115", timeout=60)
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
        app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="&S", class_name="Button", found_index=0, timeout=60)
   
        # Entrando na tela de importacao do XML
        for i in range(2):
            pyautogui.press('TAB')
            time.sleep(.01)
        pyautogui.press('ENTER')
        
        telaImportar = Application(backend="win32").connect(title="Exportar/Importar .XML NFE(w_imp_exp_nfe)", timeout=60)
        telaImportar.ExportarImportarXmlNfe.child_window(title="Valores do XML", class_name="Button").wrapper_object().click_input()
        time.sleep(.5)
        telaImportar.ExportarImportarXmlNfe.child_window(title="Importar Nota de Produtos para Janela de Cadastro", class_name="Button").wrapper_object().click_input()
        time.sleep(.5)
        telaImportar.ExportarImportarXmlNfe.child_window(class_name="Edit").type_keys(caminhoDoArquivo)
        time.sleep(1)
        telaImportar.ExportarImportarXmlNfe.child_window(title="&Carregar", class_name="Button").wrapper_object().click_input()
        time.sleep(15)
        elementosExportarImportar = telaImportar.ExportarImportarXmlNfe.children()
        elementosExportarImportar[8].double_click()
        time.sleep(.1) 
        pyautogui.press('DELETE')
        time.sleep(.1) 
        
        time.sleep(.01)
        pyautogui.write(naturezaOperacao)
        time.sleep(1)
        pyautogui.press('TAB')
        time.sleep(.5)
        pyautogui.write(documento)
        time.sleep(.5)
        pyautogui.press('TAB')
        time.sleep(.5)
        campoDepartamento = telaImportar.ExportarImportarXmlNfe.children()[13]
        # selecionaOptionSelect(campoDepartamento, formataDepartamento(departamento))
        selecionaOptionSelect(campoDepartamento, departamento)
        time.sleep(.5)
        pyautogui.write(almoxarifado)
        time.sleep(.5)

        telaImportar.ExportarImportarXmlNfe.child_window(title="&Importar", class_name="Button").wrapper_object().click_input()

        try:
            atencao_app1 = Application(backend="win32").connect(title="Atenção", timeout=3)
            descricao1 = atencao_app1.Atencao.children()[0].window_text()
        
            if "não disponível para a Empresa" in descricao1:
                raise TratamentoException(f'Natureza de operação não disponivel para a empresa')
        except:
            pass

        importar = Application(backend="win32").connect(title="Importação XML Nota Fiscal de Entrada.", timeout=60)
        importar_window = importar.top_window()
        importar_window.set_focus()
        time.sleep(1)

        # preenchendo informacao do item
        preencheItem(idNota, TratamentoException)          
        time.sleep(1)
        importar.importacaoXmlNotaFiscalDeEntrada.child_window(title="OK", class_name="Button").wrapper_object().click_input()
        app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="&S", class_name="Button", found_index=0, timeout=60)

        janelaAtencaoAtiva = False
        while janelaAtencaoAtiva==False:
            time.sleep(3)

            try:
                atencao_app = Application(backend="win32").connect(title="Atenção", timeout=2)
                descricao = atencao_app.Atencao.children()[0].window_text()
                if "Será considerado o NBM do XML?" in descricao:
                    pyautogui.press('ENTER')
                elif "Existem mais de um fornecedor cadastrado no sistema com o CNPJ" in descricao:
                    pyautogui.press('ENTER')
            except:
                janelaAtencaoAtiva = True

        app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="Conta Gerencial:", class_name="Button").click_input()
        time.sleep(.01)
        app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[40].click_input()
        time.sleep(1)
        pyautogui.press('TAB')
        time.sleep(.01)
        pyautogui.hotkey('ctrl', 'end')
        time.sleep(.01)
        pyautogui.write(' - REALIZADO POR RPA')
    except Exception as err:
        raise Exception(f'Erro preencher a capa da nota: {err}')
