import sys
import pandas as pd
import xml.etree.ElementTree as ET 

def excel_to_xml(file):
    # Загружаем spreadsheet в объект pandas
    df = pd.read_excel(file, sheet_name='Sheet1', header=4, 
                    names=['CERTNO', 'CERTDATE', 'STATUS', 'IEC', 'EXPNAME', 'BILLID','SDATE', 'SCC', 'SVALUE'],
                        converters={} )
    head = pd.read_excel(file, nrows=3, names=['name', 'CERTDATA'], header=None, index_col=0, usecols="A:B")

    #df=xl.copy()
    #print(df.at['CERTDATE', 2])
    root = ET.Element('CERTDATA') # Root элемент
    tree=ET.ElementTree(root)

    filename=ET.SubElement(root, 'FILENAME')
    filename.text = str(head.at['FileName', 'CERTDATA'])
    envelope=ET.SubElement(root, 'ENVELOPE')
    for row in df.index:
        ecert = ET.SubElement(envelope, 'ECERT')
        for column in df.columns: 
            entry = ET.SubElement(ecert, column)
            data=df.at[row, column]
            if column=='CERTDATE'or column=='SDATE': data=data.date()
            if column=='IEC': data = f'{data:010d}'
            if column=='EXPNAME':data='\"'+f'{data}'+'\"'
            if column=='SVALUE' and data%10==0: data=int(data)
            entry.text = str(data)

    ET.indent(tree, space="\t", level=0)
    tree.write("test.xml", encoding="UTF-8", xml_declaration=True)

if __name__ == '__main__':
    file_name=sys.argv[1]
    excel_to_xml(file_name)