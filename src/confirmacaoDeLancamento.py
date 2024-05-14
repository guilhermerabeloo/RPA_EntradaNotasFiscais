from pywinauto.application import Application
import time

def confirmarLancamento(TratamentoException):
    try:
        app = Application(backend="win32").connect(class_name="FNWND3115", timeout=60)
        main_window = app.top_window()
        main_window.set_focus()
        time.sleep(1)
        app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.child_window(title="&O", class_name="Button").click_input()
        time.sleep(90)
        
        texto = app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[239].window_text()

        # nePreenchida = False
        # cont = 0 
        # neLancada = ''

        # while nePreenchida == False:
        #     texto = app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[239].window_text()
        #     if texto != '':
        #         nePreenchida = True
        #         neLancada = texto

        #     try:
        #         Application(backend="win32").connect(title="Atenção", timeout=0.5)
        #         erro = 'Erro ao confirmar o lancamento'
        #         neLancada = ''
        #         nePreenchida = True
        #     except:
        #         pass
            
        #     time.sleep(5)
        #     cont+=1

        # if erro != '':
        #     raise(f'Erro ao confirmar lancamento: {erro}')

        atencao_app = Application(backend="win32")
        if atencao_app.connect(title="Atenção", timeout=3, found_index=0):
            descricao = atencao_app.Atencao.children()[0].window_text()

            if "Difere do Total da Nota" in descricao:
                raise Exception(f'Total das obrigações difere do total da nota')
            else:
                raise TratamentoException(f'Erro ao confirmar lançamento')
        
        return texto
    except TratamentoException as err:
        print(f'Erro {err}')
        raise TratamentoException(f'Erro ao confirmar lancamento')
