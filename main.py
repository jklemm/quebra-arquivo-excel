import xlrd

# TODO: Obter uma planilha real com muitas linhas para importação
# TODO: Como se geram arquivos excel a partir do python


def abrir_arquivo_excel():
    book = xlrd.open_workbook("modelo.xls")
    print("The number of worksheets is {0}".format(book.nsheets))
    print("Worksheet name(s): {0}".format(book.sheet_names()))
    sh = book.sheet_by_index(0)
    print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
    print("Cell D30 is {0}".format(sh.cell_value(rowx=0, colx=0)))

    for rx in range(sh.nrows):
        print(sh.row(rx))

if __name__ == '__main__':
    abrir_arquivo_excel()
