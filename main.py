import xlrd
import xlwt


def abrir_arquivo_excel():
    return xlrd.open_workbook("modelo.xls")


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


if __name__ == '__main__':
    arquivo_excel_grande = abrir_arquivo_excel()

    lista = tranformar_conteudo_em_uma_lista(arquivo_excel_grande)

    altura_da_lista = len(lista)
    comprimento_da_lista = len(lista[0])

    arquivo_excel_pequeno, planilha = criar_novo_arquivo_e_planilha()
    contador_de_arquivos = 1
    quebra_em_linhas = 40

    # FIXME: Apenas o primeiro arquivo excel está sendo gerado com cabeçalho

    destino_x = 0
    destino_y = 0

    for x in range(altura_da_lista):

        for y in range(comprimento_da_lista):
            planilha.write(destino_x, destino_y, lista[x][y])
            destino_y += 1

        destino_y = 0
        destino_x += 1

        if x > 0 and x % quebra_em_linhas == 0 or x == altura_da_lista - 1:
            contador = str(contador_de_arquivos).zfill(2)
            arquivo_excel_pequeno.save('novo_arquivo_{}.xls'.format(contador))
            contador_de_arquivos += 1
            arquivo_excel_pequeno, planilha = criar_novo_arquivo_e_planilha()
            destino_x = 0
            destino_y = 0
