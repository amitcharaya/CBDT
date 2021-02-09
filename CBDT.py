import traceback
import datetime
import re,os.path
from tkinter import Tk, Menu, mainloop, filedialog, messagebox
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, Color
from openpyxl.styles import colors
from openpyxl.comments import Comment

import openpyxl as xl


def createExcel():
    filePath = filedialog.askopenfilename(initialdir="./", title="Select Input File",
                                      filetypes=(("Text Files", "*.txt"), ("all files", "*.*")))
    file=open(filePath)
    header=file.readline()
    headerIdentifier=header[:2]
    orignatorCode=header[2:13]
    responderCode=header[13:24]
    fileReferenceNumber=header[24:34]
    totalRecords=header[34:40]
    wb = xl.Workbook()
    ws = wb.active
    ws.cell(1, 1).value = "Header Identifier"
    ws.cell(2, 1).value = "Orignator Code"
    ws.cell(3, 1).value = "Reponder Code"
    ws.cell(4, 1).value = "File Ref No"
    ws.cell(5, 1).value = "Total Records in File"

    ws.cell(1, 2).value=headerIdentifier
    ws.cell(2, 2).value = orignatorCode
    ws.cell(3, 2).value = responderCode
    ws.cell(4, 2).value = fileReferenceNumber
    ws.cell(5, 2).value = totalRecords
    r=6
    c=1
    ws.cell(r, c).value = "Account Valid Flag"
    c=c+1
    ws.cell(r, c).value = "Joint Account Flag"
    c = c + 1
    ws.cell(r, c).value = "Primary A/C Holder's PAN No"
    c = c + 1
    ws.cell(r, c).value = "Secondary A/C Holder's PAN No"
    c = c + 1
    ws.cell(r, c).value = "Primary Account Holder's Name"
    c = c + 1
    ws.cell(r, c).value = "Secondary Account Holder's Name"
    c = c + 1
    ws.cell(r, c).value = "Account type"
    c = c + 1
    ws.cell(r, c).value = "Record  Reference Number"
    c = c + 1
    ws.cell(r, c).value = "IFSC code of the bank branch at which customer account is maintained"
    c = c + 1
    ws.cell(r, c).value = "Destination Bank Account Number"
    c = c + 1
    ws.cell(r, c).value = "Filler 1"
    c = c + 1
    ws.cell(r, c).value = "Filler 2"
    c = c + 1
    ws.cell(r, c).value = "Filler 3"
    c = c + 1
    ws.cell(r, c).value = "Filler 4"
    c = c + 1
    ws.cell(r, c).value = "Filler 5"
    c = c + 1
    ws.cell(r, c).value = "Filler 6"
    c = c + 1
    ws.cell(r, c).value = "Filler 7"
    c = c + 1
    ws.cell(r, c).value = "Filler 8"
    c = c + 1
    ws.cell(r, c).value = "Filler 9"

    r=7
    c=1
    avf = DataValidation(type="list", formula1='"00,01,02"', allow_blank=True)
    ws.add_data_validation(avf)
    javf = DataValidation(type="list", formula1='"00,01"', allow_blank=True)
    ws.add_data_validation(javf)
    at = DataValidation(type="list", formula1='"SB,CA,CC,OD,TD,LN,SG,CG,OT,NR,PP,NO"', allow_blank=True)
    ws.add_data_validation(javf)
    for x in range(int(totalRecords)):
        line=file.readline()
        recordIdentifier=line[:2]
        recordRefNo=line[2:17]
        ifscCode=line[17:28]
        destinationBankAccountNo=line[28:63]
        filler1=line[67:87]
        filler2=line[87:137]
        filler3 = line[137:187]
        filler4 = line[187:237]
        filler5 = line[237:287]
        filler6 = line[287:337]
        filler7 = line[337:387]
        filler8 = line[387:437]
        filler9 = line[437:500]

        avf.add(ws.cell(r, c))
        ws.cell(r,c).number_format = '@'
        c = c + 1
        javf.add(ws.cell(r, c))
        ws.cell(r, c).number_format = '@'
        c = c + 1
        ws.cell(r, c).number_format = '@'
        c = c + 1
        ws.cell(r, c).number_format = '@'
        c = c + 1
        ws.cell(r, c).number_format = '@'
        c = c + 1
        ws.cell(r, c).number_format = '@'
        c = c + 1
        at.add(ws.cell(r, c))
        ws.cell(r, c).value = recordIdentifier
        c=c+1
        ws.cell(r, c).value = recordRefNo
        c = c + 1
        ws.cell(r, c).value = ifscCode
        c = c + 1
        ws.cell(r, c).value = destinationBankAccountNo
        c = c + 1
        ws.cell(r, c).value = filler1
        c = c + 1
        ws.cell(r, c).value = filler2
        c = c + 1
        ws.cell(r, c).value = filler3
        c = c + 1
        ws.cell(r, c).value = filler4
        c = c + 1
        ws.cell(r, c).value = filler5
        c = c + 1
        ws.cell(r, c).value = filler6
        c = c + 1
        ws.cell(r, c).value = filler7
        c = c + 1
        ws.cell(r, c).value = filler8
        c = c + 1
        ws.cell(r, c).value = filler9
        r=r+1
        c=1
    fileName = os.path.basename(filePath)
    ws.cell(3, 3).value = "FCBX68"
    ws.cell(3, 4).value = fileName[23:31]
    ws.cell(3, 5).value = fileName[32:37]
    wb.save(filePath+".xlsx")


def createResponse():
    filePath = filedialog.askopenfilename(initialdir="./", title="Select Input File",
                                          filetypes=(("Text Files", "*.xlsx"), ("all files", "*.*")))

    wb=xl.load_workbook(filePath)

    sheet=wb.active
    cell=sheet.cell(row=1,column=2)
    fileName="AV-"+ sheet.cell(row=2, column=2).value[:4] + "-" + sheet.cell(row=3, column=2).value[:4] + "-" + sheet.cell(row=3, column=3).value + "-" + sheet.cell(row=3, column=4).value + "-" +  str(sheet.cell(row=3, column=5).value) + "-RES.txt"
    f = open(fileName, "w")
    f.write(cell.value)
    cell = sheet.cell(row=2, column=2)
    txt=cell.value

    f.write(cell.value)
    cell = sheet.cell(row=3, column=2)
    f.write(cell.value)
    cell = sheet.cell(row=4, column=2)
    f.write(cell.value)
    cell = sheet.cell(row=5, column=2)
    f.write(cell.value)
    data = "                                                  "
    f.write(data)
    f.write(data)
    f.write(data)
    f.write(data)
    f.write(data)
    f.write(data)
    f.write(data)
    f.write(data)
    f.write(data)
    data = "          "
    f.write(data)
    r=7


    for x in range(int(cell.value)):
        f.write("\n")
        cell = sheet.cell(row=r, column=7)
        f.write(cell.value)
        cell = sheet.cell(row=r, column=8)
        spaces = " "
        if len(cell.value) < 15:
            spaces = spaces * (15 - len(cell.value))
            cell.value = cell.value + spaces
        f.write(cell.value)
        cell = sheet.cell(row=r, column=9)
        spaces = " "
        if len(cell.value) < 11:
            spaces = spaces * (11 - len(cell.value))
            cell.value = cell.value + spaces
        f.write(cell.value)
        cell = sheet.cell(row=r, column=10)
        spaces = " "
        if len(cell.value) < 35:
            spaces = spaces * (35 - len(cell.value))
            cell.value = cell.value + spaces
        f.write(cell.value)

        cell = sheet.cell(row=r, column=1)
        f.write(cell.value)

        cell = sheet.cell(row=r, column=2)
        if cell.value==None:
            cell.value="  "

        f.write(cell.value)
        cell = sheet.cell(row=r, column=3)
        if cell.value==None:
            cell.value="          "
        f.write(cell.value)

        cell = sheet.cell(row=r, column=4)
        if cell.value==None:
            cell.value="          "
        f.write(cell.value)
        cell = sheet.cell(row=r, column=5)
        if cell.value==None:
            cell.value="                                                  "
        if len(cell.value) < 50:
            spaces = spaces * (50 - len(cell.value))
            cell.value = cell.value + spaces

        f.write(cell.value)

        cell = sheet.cell(row=r, column=6)

        if cell.value == None:
            cell.value = "                                                  "
        if len(cell.value) < 50:
            spaces = spaces * (50 - len(cell.value))
            cell.value = cell.value + spaces
        f.write(cell.value)



        cell = sheet.cell(row=r, column=11)
        if cell.value==None:
            cell.value="                                                  "
        f.write(cell.value)
        cell = sheet.cell(row=r, column=12)
        if cell.value==None:
            cell.value="                                                  "
        f.write(cell.value)
        cell = sheet.cell(row=r, column=13)
        if cell.value==None:
            cell.value="                                                  "
        f.write(cell.value)
        cell = sheet.cell(row=r, column=14)
        if cell.value==None:
            cell.value="                                                  "
        f.write(cell.value)
        cell = sheet.cell(row=r, column=15)
        if cell.value==None:
            cell.value="                                                  "
        f.write(cell.value)
        cell = sheet.cell(row=r, column=16)
        if cell.value==None:
            cell.value="             "
        f.write(cell.value)

        r=r+1

    f.close()

main_window=Tk()
main_window.title="Utility for MIcrosoft Excel"
main_window.attributes("-fullscreen",True)



menu=Menu(main_window)
main_window.config(menu=menu)
splitMenu=Menu(menu)
menu.add_cascade(label="Cibil", menu=splitMenu)
splitMenu.add_command(label="Create Excel",command=createExcel)
splitMenu.add_command(label="Create Response",command=createResponse)
mainloop()