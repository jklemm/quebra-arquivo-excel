import xlrd
import xlwt


def abrir_arquivo_excel():
    return xlrd.open_workbook("planilha_leno_28k.xls")


def tranformar_conteudo_em_uma_lista(workbook):
    sheet_values = []
    sheet = workbook.sheet_by_index(0)
    for row in range(sheet.nrows):
        row_values = []
        for col in range(sheet.ncols):
            value = sheet.cell_value(row, col)
            row_values.append(value)
        sheet_values.append(row_values)
    return sheet_values


def criar_novo_arquivo_e_planilha():
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Planilha1")
    return workbook, sheet


def imprime_cabecalho(cabecalho, planilha):
    for indice_coluna, valor_coluna in enumerate(cabecalho):
        planilha.write(0, indice_coluna, valor_coluna)


if __name__ == '__main__':
    arquivo_excel_grande = abrir_arquivo_excel()

    lista = tranformar_conteudo_em_uma_lista(arquivo_excel_grande)

    altura_da_lista = len(lista)
    comprimento_da_lista = len(lista[0])
    cabecalho = lista[0]

    # TODO: Ao invés de criar novo arquivo e planilha, deveria utilizar o modelo pronto e apenas popular
    # TODO: Não esquecer de validar qual é o modelo da planilha de entrada, para utilizar o modelo correto
    # TODO: Como existe importação de Clientes e Produtos, vale a pena avaliar o conteúdo da primeira coluna da planilha 
    # (Código do Produto ou Razão Social) para avaliar qual é o modelo a ser utilizado.

    arquivo_excel_pequeno, planilha = criar_novo_arquivo_e_planilha()
    contador_de_arquivos = 1
    quebra_em_linhas = 3000

    destino_x = 0
    destino_y = 0

    for x in range(1, altura_da_lista):

        if destino_x == 0:
            imprime_cabecalho(cabecalho, planilha)
            destino_x += 1

        for y in range(comprimento_da_lista):
            planilha.write(destino_x, destino_y, lista[x][y])
            destino_y += 1

        destino_y = 0
        destino_x += 1

        if x > 0 and x % quebra_em_linhas == 0 or x == altura_da_lista - 1:
            contador = str(contador_de_arquivos).zfill(2)
            arquivo_excel_pequeno.save('planilha_leno_{}.xls'.format(contador))
            contador_de_arquivos += 1
            arquivo_excel_pequeno, planilha = criar_novo_arquivo_e_planilha()
            destino_x = 0
            destino_y = 0
