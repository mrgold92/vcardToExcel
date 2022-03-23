import vobject
import openpyxl
import xlrd
import pandas as pd
import os

d = {}
def read_vcf():
    with open("john-doe.vcf") as inf:


        indata = inf.read()
        vc = vobject.readComponents(indata)
        vo = next(vc, None)
        while vo is not None:
           
            d['Nombre'] = vo.fn.value if vo.fn.value is not None else ''
            d['Telefono'] = vo.tel.value if vo.tel.value is not None else ''
            d['Email'] = vo.email.value if vo.email.value is not None else ''
            # ContentLine to dict
            d['Direccion'] = vo.adr.value if vo.adr.value is not None else ''
            d['Direccion'] = str(d['Direccion'])
            d['Direccion'] = d['Direccion'].split('\n')[0]
            
            d["Ciudad"] = vo.adr.value.city if vo.adr.value.city is not None else ''
            d["Pais"] = vo.adr.value.country if vo.adr.value.country is not None else ''
            d["Codigo Postal"] = vo.adr.value.code if vo.adr.value.code is not None else ''

            d['URL'] = vo.url.value if vo.url.value is not None else ''
            d['URL'] = d['URL'].replace('\\', '')

            vo = next(vc, None)

def save_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    # Dict_keys not supported in openpyxl
    ws.append(list(d.keys()))
    # Los valores pueden ser un array
    valores = list(d.values())
    if isinstance(valores[0], list):
        for val in valores:
            ws.append(val)
    else:
        ws.append(valores)

    wb.save("john-doe.xlsx")


def write_vcf(excel_file):
    # '''
    # Vamos a leer los datos del excel y crear un .vcf a partir de ahí
    # '''
    # # Leemos el excel
    # wb = openpyxl.load_workbook(excel_file)
    # ws = wb.active
    # # Leemos las filas
    # rows = ws.rows
    # # Creamos un vcf
    # vcf = vobject.vCard()
    # # Iteramos sobre las filas
    # # Ignoramos la primera fila, que es la cabecera
   
    # i = 0
    # for row in rows: 
    #     # Si estamos en la primera fila, salimos
    #     if i == 0:
    #         i += 1
    #         continue
    #     # Leemos los valores de cada fila
    #     values = [cell.value for cell in row]
    #     # Creamos todas las propiedades
    #     vcf.add('fn')
    #     vcf.add('tel')
    #     vcf.add('email')
    #     vcf.add('adr')
       
       
    #     vcf.add('url')
    #     # Asignamos los valores a las propiedades
    #     vcf.fn.value = values[0]
    #     vcf.tel.value = values[1]
    #     vcf.email.value = values[2]
    #     vcf.adr.value = values[3]
    #     # vcf.adr.value.city = values[4]
    #     # vcf.adr.value.country = values[5]
    #     # vcf.adr.value.code = values[6]
    #     vcf.url.value = values[7]
    #     # Guardamos el vcf en un archivo
    #     with open("john-doe2.vcf", "w") as outf:
          
    #         c = vcf.serialize()
    #         # Serialize is not working
    #         outf.write(c)

    excelfile = pd.ExcelFile(excel_file)
    column = excelfile.parse(excelfile.sheet_names[0])
    s = ""
    begin = """BEGIN:VCARD\nVERSION:3.0"""

    for i in range(len(column)):
        fName=""
        sName=""
        mName= ""
        prefix =""
        suffix =""

        secMail=""
        secOrg=""
        secTit=""
        if(str(column["Telefono"][i])!="nan"):
            if(str(column["Nombre"][i])!="nan"):
                fName=str(column["Nombre"][i])
                # if key "apellido" exists
            if "Apellido" in column.keys():
                if(str(column["Apellido"][i])!="nan"):
                    sName=str(column["Apellido"][i])
            if "Segundo Nombre" in column.keys():
                if(str(column["Segundo Nombre"][i])!="nan"):
                    mName=str(column["Segundo Nombre"][i])
            if "Prefijo" in column.keys():
                if(str(column["Prefijo"][i])!="nan"):
                    prefix=str(column["Prefijo"][i])
            if "Sufijo" in column.keys():
                if(str(column["Sufijo"][i])!="nan"):
                    suffix=str(column["Sufijo"][i])
                    
            secN= "\nN:"+fName+" "+sName+" "+mName+" "+prefix+" "+suffix
            secFN="\nFN:"+fName

            
            # secPhone="\nTEL;CELL:+"+str(column["Phone"][i]).split(".")[0] #v1
            secPhone="\nTEL;CELL:"+str(column["Telefono"][i]) #v2
            # print("Phone ",secPhone) #For testing purposes
            if("Email" in column.columns.values):
                secMail=""
                if(str(column["Email"][i]) != "nan"):
                    secMail="\nEMAIL;HOME:"+str(column["Email"][i])
            if("Organization" in column.columns.values):
                secOrg=""
                if(str(column["Organization"][i]) != "nan"):
                    secOrg="\nORG:" + str(column["Organization"][i])
            if("Title" in column.columns.values):
                secTit=""
                if(str(column["Title"][i]) != "nan"):
                    secTit="\nTITLE:" + str(column["Title"][i])
            if("URL" in column.columns.values):
                secUrl=""
                if(str(column["URL"][i]) != "nan"):
                    secUrl="\nURL:" + str(column["URL"][i])
            if("Direccion" in column.columns.values):
                secAdr=""
                if(str(column["Direccion"][i]) != "nan"):
                    secAdr="\nADR;HOME:;;"+str(column["Direccion"][i])+";"+str(column["Ciudad"][i])+";"+str(column["Pais"][i])+";"+str(column["Codigo Postal"][i])
            s += begin + secN + secFN +secPhone + secMail + secOrg + secTit + secUrl + secAdr + "\nEND:VCARD\n"
            
    text_file = open("Exported.vcf", "w",encoding="utf-8") #Encoding utf-8 added
    text_file.write(s)
    text_file.close()

    
if __name__ == '__main__':
    read_vcf()
    # save_excel()
    write_vcf("john-doe.xlsx")
