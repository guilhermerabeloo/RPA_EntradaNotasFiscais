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
            if "Difere do Total da Nota" in descricao:
                raise TratamentoException(f'Total das obrigações difere do total da nota')
            elif "TABELA DE TIPOS DE PRODUTOS" in descricao:
                raise TratamentoException(f'Há erros no cadastro de um ou mais produtos')
            elif "TABELA DE GRUPO DE MATERIAIS" in descricao:
                raise TratamentoException(f'Há erros no cadastro de grupos de materiais')
            else:
                raise TratamentoException(f'Erro ao confirmar lançamento')
        
        texto = app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[239].window_text()
        return texto
    except TratamentoException as err:
        print(f'Erro {err}')
        raise TratamentoException(err)
