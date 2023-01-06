import shutil
import os
from os import listdir
import xlwings as xw
import PySimpleGUI as sg
import pandas as pd
import numpy as np

sg.theme('DarkBlue2')
   
path=""
address_file=""
list_file = ""
list_file_r = ""

directory = os.getcwd()
layout=[
        [sg.Text("Auto arrange file",background_color="Green",text_color="Yellow",justification="Left")],
        [sg.Text("Danh sách",size=(15,1)),sg.Combo(list_file,key="Bộ",size=(60,5))],
        [sg.Text("Đầu vào file",size=(15,1)),sg.InputText(path,key="Choose File",size=(60,2)),sg.FolderBrowse("Đường dẫn",key="Choose File")],
        [sg.Text("Đầu ra",size=(15,1)),sg.InputText(address_file,key="Choose link",size=(60,2)),sg.FolderBrowse("Đường dẫn",key="Choose link")],
        [sg.Button('Copy',key='Save',size=10),sg.Button('Move',key='Move',size=10),sg.Button('Updates',key='Updates'),sg.Button('Quit',key='Exit')],
        [sg.Table(values=list_file,
            headings=["Tên File","Extension"],
            key="table",row_height=20,justification="Center",expand_x=True,expand_y=True),
         sg.Table(values=list_file_r,
            headings=["Tên File","Extension"],
            key="table1",row_height=20,justification="Center",expand_x=True,expand_y=True)]
    
        ]
window=sg.Window("Automation arrange",layout)

            
while True:
    event,values=window.read()
    if event== sg.WINDOW_CLOSED or event == "Exit":
        break
    
    #các trường hợp của update
    elif event =="Updates":
        if  values["Choose File"] == "":
            sg.popup("Hãy chọn đầu ra")
            R1 = values["Choose link"]
            list_file_R1 = os.listdir(R1)
            window["table1"].Update(values=list_file_R1) 
            
        elif values["Choose link"] == "":
            sg.popup("Hãy chọn đầu ra")
            F1 = values["Choose File"]
            list_file1 = os.listdir(F1)
            window["Bộ"].Update(values=list_file1)
            window["table"].Update(values=list_file1)
        else:
             F1 = values["Choose File"]
             list_file1 = os.listdir(F1)
             window["Bộ"].Update(values=list_file1)
             window["table"].Update(values=list_file1)
            
             R1 = values["Choose link"]
             list_file_R1 = os.listdir(R1)
             window["table1"].Update(values=list_file_R1)
            
    # cập nhật thông tin danh sách trên bảng với các điều kiện đặt ra khi copy
    elif event =="Save":
        if values["Bộ"] == "" and values["Choose link"] == "" and values["Choose File"] =="":
            sg.popup("Hãy chọn đường dẫn & danh sách")
        elif values["Bộ"] == "" and values["Choose link"] == "":
            sg.popup("Hãy chọn đầu ra và danh sách")
            F1 = values["Choose File"]
            list_file1 = os.listdir(F1)
            window["Bộ"].Update(values=list_file1)
            window["table"].Update(values=list_file1)
            
        elif values["Bộ"] == "" and values["Choose File"] == "":
            sg.popup("Hãy chọn đầu vào và danh sách")
            R1 = values["Choose link"]
            list_file_R1 = os.listdir(R1)
            window["table1"].Update(values=list_file_R1)
            
        elif values["Choose link"] == "" and values["Choose File"] == "":
            sg.popup("Hãy chọn đầu vào và đầu ra")
            
        elif values["Choose link"] == "":
            sg.popup("Hãy chọn đầu ra")
            F1 = values["Choose File"]
            list_file1 = os.listdir(F1)
            window["Bộ"].Update(values=list_file1)
            window["table"].Update(values=list_file1)
            
        elif values["Choose File"] == "":
            sg.popup("Hãy chọn đầu ra")
            R1 = values["Choose link"]
            list_file_R1 = os.listdir(R1)
            window["table1"].Update(values=list_file_R1)
            
        elif values["Bộ"] == "":
            sg.popup("Hãy chọn danh sách")
            F1 = values["Choose File"]
            list_file1 = os.listdir(F1)
            window["Bộ"].Update(values=list_file1)
            window["table"].Update(values=list_file1)
            
            R1 = values["Choose link"]
            list_file_R1 = os.listdir(R1)
            window["table1"].Update(values=list_file_R1)
        else:
            path=values["Choose File"]
            src = values["Bộ"]
            print(src)
            src1 = os.path.join(path,src)
            print(src1)
            dst=values["Choose link"]
            dst1=os.path.join(dst,src)
            shutil.copyfile(src1,dst1)
            sg.popup("Đã lưu file")
            # update lên table
            
            list_file1 = os.listdir(path)
            window["Bộ"].Update(values=list_file1)
            window["table"].Update(values=list_file1)
            
            list_file_R1 = os.listdir(dst)
            window["table1"].Update(values=list_file_R1)
            continue

    # cập nhật thông tin danh sách trên bảng với các điều kiện đặt ra khi move
    elif event =="Move":
        if values["Bộ"] == "" and values["Choose link"] == "" and values["Choose File"] =="":
            sg.popup("Hãy chọn đường dẫn & danh sách")
        elif values["Bộ"] == "" and values["Choose link"] == "":
            sg.popup("Hãy chọn đầu ra và danh sách")
            F1 = values["Choose File"]
            list_file1 = os.listdir(F1)
            window["Bộ"].Update(values=list_file1)
            window["table"].Update(values=list_file1)
            
        elif values["Bộ"] == "" and values["Choose File"] == "":
            sg.popup("Hãy chọn đầu vào và danh sách")
            R1 = values["Choose link"]
            list_file_R1 = os.listdir(R1)
            window["table1"].Update(values=list_file_R1)
            
        elif values["Choose link"] == "" and values["Choose File"] == "":
            sg.popup("Hãy chọn đầu vào và đầu ra")
            
        elif values["Choose link"] == "":
            sg.popup("Hãy chọn đầu ra")
            F1 = values["Choose File"]
            list_file1 = os.listdir(F1)
            window["Bộ"].Update(values=list_file1)
            window["table"].Update(values=list_file1)
            
        elif values["Choose File"] == "":
            sg.popup("Hãy chọn đầu ra")
            R1 = values["Choose link"]
            list_file_R1 = os.listdir(R1)
            window["table1"].Update(values=list_file_R1)
            
        elif values["Bộ"] == "":
            sg.popup("Hãy chọn danh sách")
            F1 = values["Choose File"]
            list_file1 = os.listdir(F1)
            window["Bộ"].Update(values=list_file1)
            window["table"].Update(values=list_file1)
            
            R1 = values["Choose link"]
            list_file_R1 = os.listdir(R1)
            window["table1"].Update(values=list_file_R1)
        else:
            path=values["Choose File"]
            src = values["Bộ"]
            print(src)
            src1 = os.path.join(path,src)
            print(src1)
            dst=values["Choose link"]
            dst1=os.path.join(dst,src)
            shutil.move(src1,dst1)
            sg.popup("Đã lưu file")
            # update lên table
            
            list_file1 = os.listdir(path)
            window["Bộ"].Update(values=list_file1)
            window["table"].Update(values=list_file1)
            
            list_file_R1 = os.listdir(dst)
            window["table1"].Update(values=list_file_R1)
            continue
        

        
        
            
                
                

