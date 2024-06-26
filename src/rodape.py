from pywinauto.application import Application
from preencheModoPgt import preenchimentoModoPgt
from sql import sqlPool
import pyautogui
import time

def preencheRodape(idNota, TratamentoException):
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
        
        time.sleep(.5)
        pyautogui.press('TAB')
        time.sleep(.5)
        pyautogui.press('SPACE')
        time.sleep(2)

        janelaRodapeVisivel = True
        while janelaRodapeVisivel:
            time.sleep(.5)
            pyautogui.press('E')
            try:
                atencao_app1 = Application(backend="win32").connect(title="Atenção", timeout=3)
                descricao1 = atencao_app1.Atencao.children()[0].window_text()

                if "Confirma Exclusão" in descricao1:
                    time.sleep(.5)
                    pyautogui.press('ENTER')

            except:
                janelaRodapeVisivel = False

        elementsRodape = app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()
        valorTotalWindow = elementsRodape[9].window_text().replace('.', '')

        pyautogui.hotkey('ALT', 'I')
        time.sleep(5)

        for parcela in parcelas:
            # valor = parcela[1]
            valor = valorTotalWindow
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
            
            app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[53].click_input()
            time.sleep(2)
            pyautogui.write(tipoObrigacao)
            time.sleep(.5) 
            pyautogui.press('TAB')
            time.sleep(1.5) 
            app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[56].double_click()
            time.sleep(1.5) 
            pyautogui.press('DELETE')
            time.sleep(1.5) 
            pyautogui.write(vencimento)
            time.sleep(1.5) 
            pyautogui.press('TAB')
            time.sleep(.5) 
            pyautogui.write(valor)
            time.sleep(.5) 
            pyautogui.press('TAB')

            # prevendo janela de atancao com valor da obrigacao divergente
            janelaAtencaoVisivel = False
            cont = 0
            while janelaAtencaoVisivel==False:
                time.sleep(1)
                cont+=1
                if cont > 5:
                    break

                try:
                    Application(backend="win32").connect(title="Atenção")
                    janelaAtencaoVisivel = True
                except:
                    pass

            if janelaAtencaoVisivel:
                atencao_app = Application(backend="win32").connect(title="Atenção")
                descricao = atencao_app.Atencao.children()[0].window_text()
                if "Valor da Obrigação é maior do que" in descricao:
                    raise TratamentoException(f'Valor da obrigação divergente com o valor da nota')
                
            time.sleep(.5) 
            pyautogui.press('TAB')
            time.sleep(.5) 
            pyautogui.press('TAB')
            time.sleep(.5) 

            # preenchimentoModoPgt(infoPagamento)

            time.sleep(2)
            app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[43].click_input()
            time.sleep(2)
            pyautogui.write(moedaCorrecao)
            time.sleep(.5) 
            pyautogui.press('TAB')
            time.sleep(2)
            app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[38].click_input()
            time.sleep(2) 
            pyautogui.write(moedaJuros) 
            time.sleep(.5) 
            pyautogui.press('TAB')
            time.sleep(2)                   
            app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[50].click_input()
            time.sleep(2) 
            pyautogui.write(agentePortador)
            time.sleep(.5) 
            pyautogui.press('TAB')
            time.sleep(2) 
            pyautogui.hotkey('alt', 'O')

        time.sleep(2)
        pyautogui.hotkey('alt', 'v')

        time.sleep(2)
        app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[29].click_input()
        time.sleep(2)
        pyautogui.hotkey('alt', 'v')
    except TratamentoException as err:
        raise TratamentoException(err)
