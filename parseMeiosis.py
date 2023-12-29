 # bloque de datos sacado capturando con debug la llamada dentro del jquery. Cambiamos el número de filas mostradas al máximo y luego llamamos con Postman para obtener el XML.
 # Data obtained thought a jquery call. Number of rows shown changed in the call and Postman used for XML.
 # Link of th call: https://mcg.ustc.edu.cn/bsc/meiosis/browse_data.php?search_tag=Stage&stage=%20&_search=false&nd=1703200304431&rows=22653&page=1&sidx=web_id&sord=asc
import re
import xlsxwriter

# Dictionary key = number id, value = data in row
data = dict()

acc = 0
with open("E:/Mis cosas/Programacion/Py_programs/tableMeiosis/meiosis.txt", "r") as file:
    for line in file:
        
        # Row change in the data
        if re.match("\s+<row id=.+", line):
            acc = acc + 1   # Data will always start in 1
            data[acc] = []
        
        # Match with these type: <![CDATA[<a target='_blank' href='content.php?mg_id=mg0000001&id=1'>mg0000001</a>]]>
        if re.match("\s+.+CDATA.<.+", line):
            aux = re.sub("\s+.+'>|\n|</.+", "", line)
            data[acc].append(aux)
        
        # Match with this type: <![CDATA[Male&nbsp;&nbsp;]]> or <![CDATA[-]]>   
        if re.match("\s+.+CDATA.\w.+|\s+.+CDATA.-..", line):
            # Split the line to get the content
            auxlist = re.split("\[|\]", line)
            # Clean the content
            if (re.match("\w+&.+;\w", auxlist[2])):
                aux = re.sub("&nbsp;", "__", auxlist[2], 2)
                aux = re.sub("&nbsp;", "", aux)
            else:
                aux = re.sub("&.+|<br/>", "", auxlist[2])   # If ends only in strange characters
            data[acc].append(aux)
        
        # Match this type: <cell>Csmd1</cell>    
        if re.match("\s+<cell>.+", line):
            aux = re.sub("\s+<cell>|\n|</cell>", "", line)
            data[acc].append(aux)

# Excel creation
workbook = xlsxwriter.Workbook("meiosis.xlsx")
worksheet = workbook.add_worksheet()

worksheet.write("A1", "MG ID")
worksheet.write("B1", "Status")
worksheet.write("C1", "Gene Names")
worksheet.write("D1", "Uniprot ID")
worksheet.write("E1", "Species")
worksheet.write("F1", "Gender")
worksheet.write("G1", "Development Stage")
worksheet.write("H1", "Fecundity")
worksheet.write("I1", "Experiment methods")

for keys in data:
    worksheet.write_row("A"+str(keys+1), data[keys])

worksheet.autofit()

workbook.close()
