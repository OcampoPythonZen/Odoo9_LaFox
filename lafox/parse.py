#! -*- coding: utf-8 -*-
import psycopg2
import xlrd


def cast_str_to_int( s ):
    mto = str(s).split('.')[0]
    l = len(mto) - 1
    suma = 0
    for letra in mto:
        n = (10**l)*(ord(letra) - 48)
        l -= 1
        suma += n
    return suma


def cargar_MA( nombre_archivo ):

    archivo_excel_nombre = str(nombre_archivo)
    archivo = xlrd.open_workbook( archivo_excel_nombre )
    hoja = archivo.sheet_by_index(0)

    numero_registros = hoja.nrows
    #print "Numero de registros en la hoja: %d" % numero_registros
    r = 1
    lista_registros = []

    lista_dicc = []

    while r < numero_registros:

        v0 = hoja.cell_value( rowx=r, colx=0 )
        v1 = hoja.cell_value( rowx=r, colx=1 )
        v2 = hoja.cell_value( rowx=r, colx=2 )
        v3 = hoja.cell_value( rowx=r, colx=3 )
        v4 = hoja.cell_value( rowx=r, colx=4 )
        v5 = hoja.cell_value( rowx=r, colx=5 )        
        v6 = hoja.cell_value( rowx=r, colx=6 )        
        v7 = hoja.cell_value( rowx=r, colx=7 )        
        v8 = hoja.cell_value( rowx=r, colx=8 )        
        v9 = hoja.cell_value( rowx=r, colx=9 )        
        v10 = hoja.cell_value( rowx=r, colx=10 )        
        v11 = hoja.cell_value( rowx=r, colx=11 )        
        v12 = hoja.cell_value( rowx=r, colx=12 )        
        v13 = hoja.cell_value( rowx=r, colx=13 )        

        d = {
            'CLAVE':v0,
            'CODE_BARRAS': v1,
            'DESCRIPCION': v2,
            'UNIDDAD': v3,
            'GPO_PRECIO': v5,
            'REFERENCIA': v6,
            'GPO_INVENTARIO': v7,
            'MONEDA': v8,
            'IVA': v9,
            'PRECIO_BASE': v10,
            'OBSERVACION': v11,
            'ACTIVO': v12,
            'PROVEEDOR': v13,
            }
        
        lista_dicc.append( d )
        r += 1
    	
    return lista_dicc

def cargar_PT( nombre_archivo ):

    archivo_excel_nombre = str(nombre_archivo)
    archivo = xlrd.open_workbook( archivo_excel_nombre )
    hoja = archivo.sheet_by_index(0)

    numero_registros = hoja.nrows
    #print "Numero de registros en la hoja: %d" % numero_registros
    r = 1
    lista_registros = []

    lista_dicc = []

    while r < numero_registros:

        v0 = hoja.cell_value( rowx=r, colx=0 )
        v1 = hoja.cell_value( rowx=r, colx=1 )
        v2 = hoja.cell_value( rowx=r, colx=2 )
        v3 = hoja.cell_value( rowx=r, colx=3 )
        v4 = hoja.cell_value( rowx=r, colx=4 )
        v5 = hoja.cell_value( rowx=r, colx=5 )        
        v6 = hoja.cell_value( rowx=r, colx=6 )        
        v7 = hoja.cell_value( rowx=r, colx=7 )        
        v8 = hoja.cell_value( rowx=r, colx=8 )        
        v9 = hoja.cell_value( rowx=r, colx=9 )        
        v10 = hoja.cell_value( rowx=r, colx=10 )        
        v11 = hoja.cell_value( rowx=r, colx=11 )        
        v12 = hoja.cell_value( rowx=r, colx=12 )        
        v13 = hoja.cell_value( rowx=r, colx=13 )        
        v14 = hoja.cell_value( rowx=r, colx=14 )        
        v15 = hoja.cell_value( rowx=r, colx=15 )        
        v16 = hoja.cell_value( rowx=r, colx=16 )        
        v17 = hoja.cell_value( rowx=r, colx=17 )        
        v18 = hoja.cell_value( rowx=r, colx=18 )        
        v19 = hoja.cell_value( rowx=r, colx=19 )        
        v20 = hoja.cell_value( rowx=r, colx=20 )        
        v21 = hoja.cell_value( rowx=r, colx=21 )        
        v22 = hoja.cell_value( rowx=r, colx=22 )        
        v23 = hoja.cell_value( rowx=r, colx=23 )        
        v24 = hoja.cell_value( rowx=r, colx=24 )        
        v25 = hoja.cell_value( rowx=r, colx=25 )        
        v26 = hoja.cell_value( rowx=r, colx=26 )        
        v27 = hoja.cell_value( rowx=r, colx=27 )        
    
        d = {
            'CLAVE':v0,
            'NOMBRE': v1,
            'DIRECCION': v2,
            'COLONIA': v3,
            'CIUDAD_DELEG': v4,
            'ZIP': v5,
            'ESTADO': v6,
            'PAIS': v7,
            'LADA': v8,
            'TELEFONO': v9,
            'MOBILE': v10,
            'EMAIL': v11,
            'VENDEDOR': v12,
            'MEDIO_CONTACTO': v13,
            'REFERENCIA': v14,
            'OBSERVACION': v15,
            'RFC': v16,
            'METODO_PAGO': v17,
            'FISICA_MORAL': v18,
            'MONEDA': v19,
            'LIMITE_CREDITO': v20,
            'DIAS_CREDITO': v21,
            'ESCALA_PRECIOS': v22,
            'SALDO_PENDIENTE': v23,
            'PROMEDIO_COMPRA_M': v24,
            'COMPRA_MAYOR': v25,
            'ULT_COMPRA': v26,
            'PUNTUALIDAD_PAGO': v27,
            }
        
        lista_dicc.append( d )
        r += 1
        
    return lista_dicc
# print cargar_PT( "lAYOUT CLIENTES  15.xlsx" )[0]

