def formataDepartamento(departamento):
    index = departamento.find('>')
    formatado = f'{departamento[:index+1]}{departamento[index+2:]}'

    return formatado