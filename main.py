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

    # FIXME: Apenas o primeiro arquivo excel está sendo gerado com cabeçalho

    for x in range(altura_da_lista):
        for y in range(comprimento_da_lista):
            planilha.write(x, y, lista[x][y])
            
        if x > 0 and x % 40 == 0:
            arquivo_excel_pequeno.save('novo_arquivo_{}.xls'.format(contador_de_arquivos))
            contador_de_arquivos += 1
            arquivo_excel_pequeno, planilha = criar_novo_arquivo_e_planilha()
