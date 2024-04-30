from pywinauto.application import Application
import time

def confirmarLancamento():
    try:
        nePreenchida = False
        cont = 0 
        neLancada = ''
        erro = ''
        while nePreenchida == False:
            texto = app.AdministracaoDeEstoqueEmpresaUsuarioAutomacao.children()[238].window_text()
            if texto != '':
                nePreenchida = True
                neLancada = texto

            try:
                app = Application(backend="win32").connect(tittle="Atenção", timeout=0.5)
                erro = 'Erro ao confirmar o lancamento'
                nePreenchida = True

            except:
                pass
            
            time.sleep(1)
            cont+=1

        if erro != '':
            raise(f'Erro ao confirmar lancamento: {erro}')
        
        return neLancada
    except Exception as err:
        raise(f'Erro ao confirmar lancamento: {erro}')

