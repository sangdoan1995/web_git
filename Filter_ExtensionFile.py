import shutil
import os
from os import listdir
import xlwings as xw
import PySimpleGUI as sg

sg.theme('DarkBlue2')
path = ""
address_file = ""
list_add=""
list_add_1=""
list_combo =['.docx','.doc','.xlsx','.xlsm','.xls','.pdf','.csv','.dwg','.rvt','.dwl','.jpg','.png','.jfif','.exe','.zip','.rar','.mp4','.mp3']
#Create table 
layout=[
        [sg.Text("Auto Filter File",background_color="grey",text_color="Black",justification="Left")],
        [sg.Text("Entrance Link",size=(15,1)),sg.InputText(path,key="Choose File",size=(60,2)),sg.FolderBrowse("Đường dẫn",key="Choose File")],
        [sg.Text("Filter Link",size=(15,1)),sg.InputText(address_file,key="Choose link",size=(60,2)),sg.FolderBrowse("Đường dẫn",key="Choose link")],
        [sg.Text("Extension",size=(15,1)),sg.Combo(list_combo,key="Extension",size=(60,5))],
        [sg.Button('Copy & Filter',key='Copy'),sg.Button('Move & Filter',key='Move'),sg.Button('Updates',key='Updates'),sg.Button('Quit',key='Quit')],
        [sg.Table(values=list_add_1,
            headings=["Tên File","Extension"],
            key="table1",row_height=20,justification="Center",expand_x=True,expand_y=True),
         sg.Table(values=list_add,
            headings=["Tên File","Extension"],
            key="table",row_height=20,justification="Center",expand_x=True,expand_y=True)]
        ]
window=sg.Window("Automation Development of DUSA",layout)

while True:
    event,values = window.read()
    if event in (sg.WIN_CLOSED,"Quit"):
        break
    #các trường hợp khi chọn update thì sẽ cập nhật danh sách table
    elif event =="Updates":
        if  values["Choose File"] == "" and values["Choose link"] == "":
            sg.popup("please chọn Links Entrance & Filter")
            
        elif values["Choose File"] == "":
            sg.popup("please chọn Entrance link")
            address_file = values["Choose link"]
            list_add = os.listdir(address_file)
            print(list_add)
            window["table"].Update(values=list_add)
            
        elif values["Choose link"] == "":
            sg.popup("please chọn Filter link")
            path = values["Choose File"]
            list_add_1 = os.listdir(path)
            print(list_add_1)
            window["table1"].Update(values=list_add_1)
        else:
            sg.popup("Hãy thực hiện Copy or Move")
            address_file = values["Choose link"]
            list_add = os.listdir(address_file)
            print(list_add)
            window["table"].Update(values=list_add)
            
            path = values["Choose File"]
            list_add_1 = os.listdir(path)
            print(list_add_1)
            window["table1"].Update(values=list_add_1)
        
    #thêm các trường hợp rỗng và copy, move filter
    elif values["Choose File"] == "" and values["Choose link"] == "" and values["Extension"] =="":
        sg.popup("please choose links")
    elif values["Choose File"] == "" and values["Choose link"] == "":
        sg.popup("please chọn Links Entrance & Filter")
    elif values["Choose link"] == "" and values["Extension"] =="":
        sg.popup("please chọn link Filter & Extension")
        path = values["Choose File"]
        list_add_1 = os.listdir(path)
        print(list_add_1)
        window["table1"].Update(values=list_add_1)
        
    elif values["Choose File"] == "" and values["Extension"] =="":
        sg.popup("please chọn link Entrance & Extension")
        address_file = values["Choose link"]
        list_add = os.listdir(address_file)
        print(list_add)
        window["table"].Update(values=list_add)
        
    elif values["Choose File"] == "":
        sg.popup("please chọn Entrance link")
        address_file = values["Choose link"]
        list_add = os.listdir(address_file)
        print(list_add)
        window["table"].Update(values=list_add)
        
    elif values["Choose link"] == "":
        sg.popup("please chọn Filter link")
        path = values["Choose File"]
        list_add_1 = os.listdir(path)
        print(list_add_1)
        window["table1"].Update(values=list_add_1)
        
    elif values["Extension"] == "":
        sg.popup("please chọn Extension")
        address_file = values["Choose link"]
        list_add = os.listdir(address_file)
        print(list_add)
        window["table"].Update(values=list_add)
        
        path = values["Choose File"]
        list_add_1 = os.listdir(path)
        print(list_add_1)
        window["table1"].Update(values=list_add_1)
    
    #tạo điều kiện khi chọn copy
    elif event == "Copy":
        path = values["Choose File"]
        path_name = os.listdir(path)
        print(path_name)
        # chuyển đổi đường dẫn trước khi ghép nối
        path1=os.path.normpath(path)
        dst=values["Choose link"]
        dst1=os.path.normpath(dst)
        
        for list_type in path_name:
            path2 = os.path.join(path1,list_type)
    
            list_name, list_extension = os.path.splitext(list_type)

            #tạo điều kiện cho Extension
            if  values["Extension"] ==list_extension:
                dst2 = os.path.join(dst1,list_type)
                #copy file
                shutil.copyfile(path2,dst2)
                df = os.listdir(dst)
                window["table"].Update(values=df)
                df1 = os.listdir(path)
                window["table1"].Update(values=df1)
                continue
              
    #tạo điều kiện khi chọn Move
    elif event == "Move":
        path = values["Choose File"]
        path_name = os.listdir(path)
        print(path_name)
        # chuyển đổi đường dẫn trước khi ghép nối
        path1=os.path.normpath(path)
        dst=values["Choose link"]
        dst1=os.path.normpath(dst)
        
        for list_type in path_name:
            path2 = os.path.join(path1,list_type)
    
            list_name, list_extension = os.path.splitext(list_type)
            
            #tạo điều kiện cho Extension
            if  values["Extension"] ==list_extension:
                dst2 = os.path.join(dst1,list_type)
                #move file
                shutil.move(path2,dst2)
                df = os.listdir(dst)
                window["table"].Update(values=df)
                df1 = os.listdir(path)
                window["table1"].Update(values=df1)
                continue
            



    