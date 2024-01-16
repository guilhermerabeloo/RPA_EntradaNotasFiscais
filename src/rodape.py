from pywinauto.application import Application
from preencheModoPgt import preenchimentoModoPgt
from sql import sqlPool
import pyautogui
import time

def preencheRodape(idNota):
    try:
        parcelas = sqlPool("SELECT", f"""
                SELECT 
                    numero_parcela,
                    REPLACE(CONVERT(VARCHAR(50), valor_parcela), '.', ',') as valor,
                    CONVERT(VARCHAR(10), vencimento, 103) as vencimento,
                    tipo_obrigacao,
                    moeda_correcao,
                    moeda_juros,
                    agente_portador,
                    tipo_preenchimento,
                    mod_pagamento,
                    tipo_cc,
                    banco,
                    agencia,
                    conta,
                    ident,
                    codigo_barras,
                    tipo_pix,
                    chave_pix,
                    transf_pix,
                    tipo_pix_qrcode,
                    qrcode_pix
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
            infoPagamento = {
                'tipoPrenchimento': parcela[7],
                'modPagamento': parcela[8],
                'tipoConta': parcela[9],
                'banco': parcela[10],
                'agencia': parcela[11],
                'conta': parcela[12],
                'indent': parcela[13],
                'codigoBarras': parcela[14],
                'tipoPix': parcela[15],
                'chavePix': parcela[16],
                'transfPix': parcela[17],
                'tipoPixQrcode': parcela[18],
                'qrcoodePix': parcela[19]
            }

            time.sleep(.3) 
            pyautogui.press('TAB')
            time.sleep(.3)      
            pyautogui.write(tipoObrigacao)
            time.sleep(.3) 
            pyautogui.press('TAB')
            time.sleep(.3) 
            pyautogui.press('TAB')
            time.sleep(.3) 
            pyautogui.write(vencimento)
            time.sleep(.3) 
            pyautogui.press('TAB')
            time.sleep(.3) 
            pyautogui.write(valor)
            time.sleep(.3) 
            pyautogui.press('TAB')
            time.sleep(.3) 
            pyautogui.press('TAB')
            time.sleep(.3) 
            pyautogui.press('TAB')
            time.sleep(.3) 

            preenchimentoModoPgt(infoPagamento)

            pyautogui.write(moedaCorrecao)
            pyautogui.press('TAB')
            time.sleep(2)
            app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[38].click_input()
            time.sleep(2) 
            pyautogui.write(moedaJuros) 
            pyautogui.press('TAB')
            time.sleep(2)                   
            app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[50].click_input()
            time.sleep(2) 
            pyautogui.write(agentePortador)
            pyautogui.press('TAB')
            time.sleep(2) 
            pyautogui.hotkey('alt', 'O')
            time.sleep(2) 
            pyautogui.press('TAB')
            time.sleep(2)  
            pyautogui.press('TAB')

        time.sleep(2)
        pyautogui.hotkey('alt', 'v')

        time.sleep(2)
        app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[29].click_input()
        time.sleep(2)
        pyautogui.hotkey('alt', 'v')
    except Exception as err:
        raise Exception(f'Erro ao preencher rodapé da nota: {err}')
    