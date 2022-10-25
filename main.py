import xlsxwriter



def escribirExcel(series):
    workbook = xlsxwriter.Workbook('series.xlsx')
    worksheet = workbook.add_worksheet()
    valores_digitar_excel = []
    reg = 0
    dup = 'No'
    print(f'Encontradas {len(series)} series')
    for serie in series:
        if serie.count('AL') == 1:
            al = indexof(serie, 'AL')
            if 0 < al < len(serie) - 1:
                delimitadores = serie.split('AL')
                try:
                    i = int(delimitadores[0])
                    j = int(delimitadores[1]) + 1
                    if i < j:
                        for l in range(i, j):
                            if indexof(valores_digitar_excel, l) == -1:
                                valores_digitar_excel.append(l)
                            else:
                                dup = 'Si'
                        print(f'Serie {str(series.index(serie) + 1).center(20)}  |  {serie.center(35)}   |   {str(j - i).center(10)} valores')
                        reg += j - i
                    else:
                        print(f'Error en serie {str(series.index(serie) + 1).center(11)}  |  {serie.center(35)}   |   {str(0).center(10)} valores')
                except Exception:
                    print(f'Error en serie {str(series.index(serie) + 1).center(11)}  |  {serie.center(35)}   |   {str(0).center(10)} valores')
            else:
                print(f'Error en serie {str(series.index(serie) + 1).center(11)}  |  {serie.center(35)}   |   {str(0).center(10)} valores')
        else:
            try:
                if indexof(valores_digitar_excel, int(serie)) == -1:
                    valores_digitar_excel.append(int(serie))
                else: dup='Si'
                print(f'Serie {str(series.index(serie) + 1).center(20)}  |  {serie.center(35)}   |   {str(1).center(10)} valor')
                reg += 1
            except Exception:
                print(f'Error en serie {str(series.index(serie) + 1).center(11)}  |  {serie.center(35)}   |   {str(0).center(10)} valores')

    valores_digitar_excel.sort()

    worksheet.write('A1', 'ITEM')
    worksheet.write('B1', 'NUMERACION')
    for mynumber in valores_digitar_excel:
        celda_item = valores_digitar_excel.index(mynumber) + 1
        worksheet.write(celda_item, 0, celda_item)
        worksheet.write(celda_item, 1, mynumber)

    workbook.close()

    print(f'Cantidad de Numeraciones subidas a Excel: {len(valores_digitar_excel)}')
    print(f'Cantidad de numeraciones digitadas: {reg}')

    print(f'Existen rangos con duplicidades: {dup}')


def indexof(valor, encontrar):
    try:
        return valor.index(encontrar)
    except ValueError:
        return -1


if __name__ == '__main__':
    txt = input('Escribe la serie (-)=>Separa Series  (AL)=>Especifica Rango :')
    txt = txt.replace('\n', '')
    txt = txt.replace(' ', '')
    txt = txt.replace('al', 'AL')

    if indexof(txt, '-') > 0:
        series = txt.split('-')
    else:
        series = [txt]

    escribirExcel(series)
