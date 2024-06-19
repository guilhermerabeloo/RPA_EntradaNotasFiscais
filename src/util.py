import pyautogui    

def formataDepartamento(departamento):
    index = departamento.find('>')
    formatado = f'{departamento[:index+1]}{departamento[index+2:]}'

    return formatado


def clicarEmImagem(imagem, i):
    ocorrencias = list(pyautogui.locateAllOnScreen(imagem))
    posicao = ocorrencias[i]
    left, top, width, height = posicao
    centro_x = left + width / 2
    centro_y = top + height / 2
    pyautogui.click(centro_x, centro_y)