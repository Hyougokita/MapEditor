import tkinter
import tkinter as tk
import tkinter.ttk as ttk
import  csv
import  openpyxl
from openpyxl import  Workbook
from tkinter import  filedialog
from tkinter import simpledialog



class Mapedit(tk.Frame):
    def __init__(self, master = None):
        super().__init__(master)

        #ウィンドウズのサイズ
        self.editor_size_x = 960
        self.editor_size_y = 540


        self.master.title("タイトル")    # ウィンドウタイトル
        self.master.geometry("960x540") # ウィンドウサイズ(幅x高さ)
        # Canvasの作成
        self.canvas = tk.Canvas(self.master,width = self.editor_size_x,height = self.editor_size_y,bg = "white")
        # Canvasを配置

        # 初めてのマス位置
        self.startX = 10
        self.startY = 20

        # マス
        self.scale = 1.0
        self.masuWidth = 20 * self.scale
        self.masuHeight = 20 * self.scale
        self.padding = 10 * self.scale

        self.masu_x = -1
        self.masu_y = -1

        #建物のフラグを保存するリスト
        self.map_chip = []
        #マス描画用四角形保存リスト
        self.masu_rect_list = []

        #Mouse
        self.mousePosX = 0
        self.mousePosY = 0

        #現在の建築フラグ
        self.constriction_flag = -1

        #サブメニュー関連
        self.sub_menu_active = True
        self.sub_menu_item_number = 0
        self.sub_menu_item_list = []
        self.sub_menu_item_color = ["LightPink","Purple","Gold","Aqua","Azure",
                                    "Brown","Cyan","Green","IndianRed","LightBlue",
                                    "Maroon","Navy","Orange","Salmon","SkyBlue",
                                    "Tan","Tomato","Violet","Wheat","Yellow"]
        self.sub_menu_item_color_can_be_choice = []
        for i in self.sub_menu_item_color:
            self.sub_menu_item_color_can_be_choice.append(i)
        self.sub_menu_button_list = []
        self.sub_menu_button_start_position_x = 820
        self.sub_menu_button_start_position_y = 35
        self.sub_menu_button_padding_y = 45
        self.sub_menu_item_dict = []


        #ファイル名前関連
        self.map_excel_file_full_name = ""
        self.map_excel_file_name = ""


    def DrawMasu(self,x,y,color,list):
        self.masu = self.canvas.create_rectangle(x,y,x + self.masuWidth,y + self.masuHeight,fill=color)
        print("type",type(self.masu))
        list.append(self.masu)

    def SetMasu(self,yoko,tate):
        for i in range(0,yoko):
            temp_list = []
            for j in range(0,tate):
                print("map_chip:",int(self.map_chip[i][j]))
                self.DrawMasu(self.startX + (self.masuWidth + self.padding) * i,
                              self.startY + (self.masuHeight + self.padding) * j,
                              self.sub_menu_item_color[int(self.map_chip[i][j])],temp_list)
            self.masu_rect_list.append(temp_list)
        print(self.masu_rect_list)
        self.canvas.pack()

    def CheckMasu(self):
        self.tempMasuX = (int)(self.mousePosX / (self.masuWidth + self.padding))
        self.tempMasuY = (int)(self.mousePosY / (self.masuHeight + self.padding))
        print("(",self.tempMasuX,",",self.tempMasuY,")")

        return  (self.tempMasuX,self.tempMasuY)

    def CheckIn(self,x,y):
        self.tempMasuX,self.tempMasuY = self.CheckMasu()
        if( x > self.startX + (self.masuWidth + self.padding) * self.tempMasuX and
            x < self.startX + (self.masuWidth + self.padding) * self.tempMasuX + self.masuWidth and
            y > self.startY + (self.masuHeight + self.padding) * self.tempMasuY and
            y < self.startY + (self.masuHeight + self.padding) * self.tempMasuY + self.masuHeight):
            print("in")
            return True
        print("no")
        return False

    def ChangeColor(self,x,y):
        if __debug__:
            print("sub_menu_item_color",self.sub_menu_item_color)
            print("sub_menu_item_dict",self.sub_menu_item_dict)
        self.canvas.itemconfig(self.masu_rect_list[x][y],fill = self.sub_menu_item_color[self.sub_menu_item_dict[self.constriction_flag]['obj_flag'] - 1])


    def WriteCsv(self,col,cow,constriction_flag):
        wb = Workbook()
        ws = wb.active
        ws.cell(col + 1,cow + 1,constriction_flag)
        wb.save(r'map1.xlsx')
        pass

    def SaveMapChipToExcel(self):
        if (self.map_excel_file_full_name != ""):
            book = openpyxl.load_workbook(self.map_excel_file_full_name)
            sheet = book.get_sheet_by_name("Map")
            #sheet = self.OpenExcel("Map")['sheet']
            #book = self.OpenExcel("Map")['book']
            for i in range(self.masu_x):
                for j in range(self.masu_y):
                    sheet.cell(row = j + 1, column = i + 1).value = self.map_chip[i][j]
            book.save(self.map_excel_file_full_name)

    def LeftClickEvent(self,event):
        print("X:",event.x," Y:",event.y)
        self.mousePosX = event.x
        self.mousePosY = event.y
        if(self.CheckIn(event.x,event.y)):
            self.ChangeColor(self.tempMasuX,self.tempMasuY)
            self.map_chip[self.tempMasuX][self.tempMasuY] = self.constriction_flag
            #self.WriteCsv(self.tempMasuX,self.tempMasuY,self.constriction_flag)

        pass



    def SetButton(self,text,x,y,func):
        self.btn = tk.Button(self.canvas,text = text, width = 4, height = 2, bd=1, padx=0, pady=0, relief='ridge',command=func)
        self.btn.place(x=x, y=y)
        #self.btn.pack()
        self.sub_menu_button_list.append(self.btn)

    def SetButton(self,text,x,y,func,width = 4,height = 2):
        self.btn = tk.Button(self.canvas,text = text, width = width, height = height, bd=1, padx=0, pady=0, relief='ridge',command=func)
        self.btn.place(x=x, y=y)
        #self.btn.pack()
        self.sub_menu_button_list.append(self.btn)

    def SetButton(self, text, x, y, func, width=4, height=2, bg = "white"):
        self.btn = tk.Button(self.canvas, text=text, width=width, height=height, bd=1, padx=0, pady=0, relief='ridge',
                             command=func, bg = bg)
        self.btn.place(x=x, y=y)
        #self.btn.pack()
        self.sub_menu_button_list.append(self.btn)

    def ButtonEventTest(self):
        print("Button")


    def DrawSubMenuSingleItem(self,text,color = "red"):
        self.SetButton(text,x = self.sub_menu_button_start_position_x,y = self.sub_menu_button_start_position_y,
                       bg = color,
                       width = 8,
                       func = lambda f = text:self.CheckSubMenuItemNumber(f))

    def DrawSubMenuSingleItem(self,text):
        color = self.sub_menu_item_color[len(self.sub_menu_item_list)]
        self.SetButton(text,x = self.sub_menu_button_start_position_x,y = self.sub_menu_button_start_position_y,
                       bg = color,
                       width = 8,
                       func = lambda f = text:self.CheckSubMenuItemNumber(f))

    def OpenExcel(self,sheet_name):
        if(self.map_excel_file_full_name != ""):
            book = openpyxl.load_workbook(self.map_excel_file_full_name)
            sheet = book.get_sheet_by_name(sheet_name)
            return  {'sheet':sheet,'book':book}

    #Excelに登録したconfigを辞書リストに保存する
    def ReadSubMenuConfigFromExcel(self):
            sheet = self.OpenExcel("Submenu")['sheet']
            for i in range(2,30):
                obj_name = sheet.cell(row = i, column = 1).value
                obj_color = sheet.cell(row = i, column = 2).value
                obj_flag = sheet.cell(row = i, column = 3).value
                if(obj_name != None):
                    self.sub_menu_item_dict.append({'obj_name':obj_name,'obj_color':obj_color,'obj_flag':obj_flag})
                    self.sub_menu_item_color_can_be_choice.remove(obj_color)
                else:
                    break

            if __debug__:
                for i in range(len(self.sub_menu_item_dict)):
                    print(self.sub_menu_item_dict[i])


    def DrawSubMenuItem(self,x,y):
        if __debug__:
            print("length of sub_menu_item_dict",len(self.sub_menu_item_dict))
        for i in range(len(self.sub_menu_item_dict)):
            temp_obj_name = self.sub_menu_item_dict[i]['obj_name']
            temp_obj_color = self.sub_menu_item_dict[i]['obj_color']
            self.SetButton(temp_obj_name,x = x , y = y + self.sub_menu_button_padding_y * i,
                           bg = temp_obj_color,width = 8,
                           func = lambda f = temp_obj_name : self.CheckSubMenuItemNumber(f))




    def CheckSubMenuItemNumber(self, text):
        if __debug__:
            print("length of sub_men_button_list:",len(self.sub_menu_button_list))
        for i in range(len(self.sub_menu_button_list)):
            if(self.sub_menu_button_list[i]['text'] == text):
                self.constriction_flag = i
                if __debug__:
                    print("constriction_flag",i)
                    #print(self.sub_menu_item_color)
                break

    def SubMenuEvent(self):
        if __debug__:
            print("sub_menu_active",self.sub_menu_active)
            print("sub_button_list",self.sub_menu_button_list)
        if self.sub_menu_active == True:
            #self.canvas.delete(self.rect)
            #for i in (self.sub_menu_button_list):
                #print("type",type(i))
                #i.pack_forget()
            self.sub_menu_button_list = []
            self.sub_menu_active = False
        else:
            #self.DrawSubMenu()
            self.sub_menu_button_list = []
            self.DrawSubMenuItem(self.sub_menu_button_start_position_x,self.sub_menu_button_start_position_y)

            self.sub_menu_active = True

    def DestoryButton(self):
        for i in self.sub_menu_button_list:
            print(i)
            i.pack_forget()

    def Set(self):
        if __debug__:
            self.test_btn = tk.Button(text = "テスト",command=self.DestoryButton)
            self.test_btn.place( x = 300, y = 200,)
        self.canvas.bind('<ButtonPress-1>',self.LeftClickEvent)


    def DrawSubMenu(self):
        if self.sub_menu_active:
            self.rect = self.canvas.create_rectangle(800,20,900,400,fill = 'white')
            if __debug__:
                print("self.rect番号：",self.rect)
        else:
            pass

    def ReadMapInfoFromExcel(self):
        if(self.map_excel_file_full_name != ""):
            book = openpyxl.load_workbook(self.map_excel_file_full_name)
            active_sheet = book.active
            sheet = book.get_sheet_by_name("MapInfo")
            if __debug__:
                print("MapInfo Size X:",sheet.cell(row = 1, column = 2).value)
                print("MapInfo Size Y:",sheet.cell(row = 2, column = 2).value)

            self.masu_x = sheet.cell(row = 1, column= 2).value
            self.masu_y = sheet.cell(row = 2, column= 2).value
            #return (sheet.cell(row = 1,column=2).value,sheet.cell(row = 2,column=2).value)

    def AddNewSubMenuItem(self):
        pop_up_window = tk.Tk()
        pop_up_window.title("建物新規作成")
        pop_up_window.geometry("400x120")

        label_construction_name = tkinter.Label(pop_up_window,text = "建物の名前")
        label_construction_color = tkinter.Label(pop_up_window,text = "記号の色")
        label_construction_name.place(x = 20, y = 20)
        label_construction_color.place(x = 20, y = 50)

        construction_entry_box = tkinter.Entry(pop_up_window,width=40)
        construction_entry_box.insert(tkinter.END,"建物")
        construction_entry_box.place(x = 100, y = 20)

        construction_color_combobox = ttk.Combobox(pop_up_window,
                                                   height = len(self.sub_menu_item_color_can_be_choice),
                                                   values = self.sub_menu_item_color_can_be_choice)
        construction_color_combobox.place(x=100, y=50)

        construction_color_combobox.bind('<<ComboboxSelected>>',func = lambda f : print(construction_color_combobox.get()))

        def GetPopUpWindowData():
            print("color:",construction_color_combobox.get())
            print("construction_name:",construction_entry_box.get())
            self.WriteSubMenuItemConfigToExcel(construction_entry_box.get(),construction_color_combobox.get())
            pop_up_window.destroy()

            if __debug__:
                print("length of button list", len(self.sub_menu_button_list))

            if(self.sub_menu_active):
                #self.DrawSubMenu()
                self.sub_menu_button_list = []
                self.DrawSubMenuItem(self.sub_menu_button_start_position_x, self.sub_menu_button_start_position_y)


        pop_up_window_button = tkinter.Button(pop_up_window,text = "新規", command = GetPopUpWindowData)
        pop_up_window_button.pack()
        pop_up_window_button.place(x = 30, y = 80)

    def WriteSubMenuItemConfigToExcel(self,obj_name,obj_color):
        book = openpyxl.load_workbook(self.map_excel_file_full_name)
        sheet = book.get_sheet_by_name("Submenu")
        row = len(self.sub_menu_item_dict) + 2
        print("row:",row)
        sheet.cell(row = row, column = 1).value = obj_name
        sheet.cell(row = row, column = 2).value = obj_color
        sheet.cell(row = row, column = 3).value = self.SearchList(self.sub_menu_item_color,obj_color) + 1
        self.sub_menu_item_dict.append({'obj_name': obj_name, 'obj_color': obj_color, 'obj_flag': self.SearchList(self.sub_menu_item_color,obj_color) + 1})
        self.sub_menu_item_color_can_be_choice.remove(obj_color)
        book.save(self.map_excel_file_full_name)

    def SearchList(self,list,obj):
        print("Search List")
        print(list)
        print(obj)
        for i in range(len(list)):
            print(list[i],obj)
            if list[i] == obj:
                print("i:",i)
                return i

    def DrawMenu(self):
        self.menu_bar = tk.Menu(self.canvas)

        self.file_menu = tk.Menu(self.menu_bar, tearoff = False)
        self.menu_bar.add_cascade(label = "ファイル", menu = self.file_menu)

        self.file_menu.add_command(label = "新規",accelerator = "Ctrl+N")
        self.file_menu.add_command(label = "開く", command = self.OpenMapExcelFile, accelerator = "Ctrl+O")
        self.file_menu.add_command(label = "上書き保存", accelerator = "Ctrl+S",command = self.SaveMapChipToExcel)
        self.file_menu.add_command(label="名前を付けて保存", accelerator = "Ctrl+Alt+S")

        self.sub_menu = tk.Menu(self.menu_bar, tearoff = False)
        self.menu_bar.add_cascade(label = "サブメニュー", menu = self.sub_menu)

        self.sub_menu.add_command(label = "表示/隠す", command = self.SubMenuEvent)
        self.sub_menu.add_command(label = "新規", command = self.AddNewSubMenuItem)

        self.canvas.bind('<Control-Key-O>', self.OpenMapExcelFile)
        self.canvas.bind('<Control-Key-o>',self.OpenMapExcelFile)

    def OpenExplorer(self,event = 0):
        self.map_excel_file_full_name = filedialog.askopenfilename(defaultextension = ".xlsx")
        name_list = self.map_excel_file_full_name.split('/')
        self.map_excel_file_name = name_list[-1]
        #print(self.map_excel_file_name)

    def OpenMapExcelFile(self):
        self.OpenExplorer()
        self.ReadMapInfoFromExcel()
        self.ReadMapFromExcel()
        if(self.masu_x != -1 and self.masu_y != -1):
            self.SetMasu(self.masu_x, self.masu_y)
        self.ReadSubMenuConfigFromExcel()


    def ReadMapFromExcel(self):
        if (self.map_excel_file_full_name != ""):
            book = openpyxl.load_workbook(self.map_excel_file_full_name)
            active_sheet = book.active
            sheet = book.get_sheet_by_name("Map")
            for i in range(1,self.masu_x + 1):
                temp_list = []
                for j in range(1,self.masu_y + 1):
                    temp_value = sheet.cell(row = j, column= i).value
                    if __debug__:
                        print("cell value:",(j,i),":",temp_value)
                    if(temp_value == None):
                        temp_list.append(0)
                    else:
                        temp_list.append(temp_value)
                self.map_chip.append(temp_list)
            if __debug__:
                print(self.map_chip)






if __name__ == "__main__":
    root = tk.Tk()
    app = Mapedit(master = root)
    #app.SetMasu(3,3)
    app.DrawMenu()
    #app.DrawSubMenu()
    app.Set()

    root.config(menu = app.menu_bar)
    app.mainloop()