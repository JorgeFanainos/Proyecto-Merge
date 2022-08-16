from xlrd import open_workbook

lista_apellidos = []
wb = open_workbook('CUENTA-NOMINA.xls')
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)
column_index = 3
column = sheet.cell_value(0, column_index)


for row in range(1, sheet.nrows):
    lista_apellidos.append(sheet.cell_value(row, column_index))

print("""
█▀ █▀█ █   █ █▀▀ █ ▀█▀ █ █ █▀▄   █▀▄ █▀▀   █▀▄ ▄▀█ ▀█▀ █▀█ █▀
▄█ █▄█ █▄▄ █ █▄▄ █  █  █▄█ █▄▀   █▄▀ ██▄   █▄▀ █▀█  █  █▄█ ▄█
""")

aux2 = True
while aux2 == True:
    nombre_analizar = input("Ingrese el nombre del documento que desea analizar:  ")
    try:
        f=open(nombre_analizar+".TXT","r")
        aux2 = False
    except:
        print("\nEl documento que ingreso no se encuentra en la carpeta o ha colocado el nombre incorrectamente!\n")
        aux2 = True  


completo_por_linea=f.readlines()

aux = True
while aux == True:
    print("""

█▀▀ █▀▀ █▄ █ █▀▀ █▀█ ▄▀█ █▀▀ █ █▀█ █▄ █   █▀▄ █▀▀   █▀█ █▀▀ █▀█ █▀█ █▀█ ▀█▀ █▀▀
█▄█ ██▄ █ ▀█ ██▄ █▀▄ █▀█ █▄▄ █ █▄█ █ ▀█   █▄▀ ██▄   █▀▄ ██▄ █▀▀ █▄█ █▀▄  █  ██▄

""")
    nombre_nuevo = input("\nIngrese el nombre del documento que desea generar:  ")
    with open(nombre_nuevo+'.txt', 'w') as temp_file:
        temp_file.write("FECHA	    H O R A	   PANT	 CEDULA	  NROEMP	 NOMBRE DEL USUARIO	                   NUMERO DE CUENTA	                TITULAR DE LA CUENTA"+"\n\n") 
        for apellidos in lista_apellidos:            
            for completo in completo_por_linea:
                if apellidos in completo:
                    temp_file.write("%s\n" % completo)            
                    aux = False
       
