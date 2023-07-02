import tkinter as tk
import  csv
import  openpyxl
from openpyxl import  Workbook
from tkinter import  filedialog



class Mapedit(tk.Frame):
    def __init__(self, master = None):
        super().__init__(master)

        self.editor_size_x = 960
        self.editor_size_y = 540


        self.master.title("タイトル")    # ウィンドウタイトル
        self.master.geometry("960x540") # ウィンドウサイズ(幅x高さ)
        # Canvasの作成
        self.canvas = tk.Canvas(self.master,width = self.editor_size_x,height = self.editor_size_y,bg = "white")
        # Canvasを配置


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
        self.sub_menu_active = False
        self.sub_menu_item_number = 0
        self.sub_menu_item_list = []
        self.sub_menu_item_color = ["LightPink","Purple","Gold","Aqua"]
        self.sub_menu_button_list = []


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
        self.canvas.itemconfig(self.masu_rect_list[x][y],fill = self.sub_menu_item_color[self.constriction_flag])


    def WriteCsv(self,col,cow,constriction_flag):
        wb = Workbook()
        ws = wb.active
        ws.cell(col + 1,cow + 1,constriction_flag)
        wb.save(r'map1.xlsx')
        pass

    def SaveMapChipToExcel(self):
        if (self.map_excel_file_full_name != ""):
            book = openpyxl.load_workbook(self.map_excel_file_full_name)
            active_sheet = book.active
            sheet = book.get_sheet_by_name("Map")
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
        self.btn.pack()
        self.btn.place(x=x, y=y)
        self.sub_menu_button_list.append(self.btn)

    def SetButton(self,text,x,y,func,width = 4,height = 2):
        self.btn = tk.Button(self.canvas,text = text, width = width, height = height, bd=1, padx=0, pady=0, relief='ridge',command=func)
        self.btn.pack()
        self.btn.place(x=x, y=y)
        self.sub_menu_button_list.append(self.btn)

    def SetButton(self, text, x, y, func, width=4, height=2, bg = "white"):
        self.btn = tk.Button(self.canvas, text=text, width=width, height=height, bd=1, padx=0, pady=0, relief='ridge',
                             command=func, bg = bg)
        self.btn.pack()
        self.btn.place(x=x, y=y)
        self.sub_menu_button_list.append(self.btn)

    def ButtonEventTest(self):
        print("Button")

    def DrawSubMenuItem(self,x,y):
        """  TEST  """
        self.SetButton("道",x=x,y=y,bg=self.sub_menu_item_color[0],width=8,
                       func=lambda f = "道":self.CheckSubMenuItemNumber(f))
        self.SetButton("建物",x=x,y=y+40,bg=self.sub_menu_item_color[1],width=8,
                       func=lambda f = "建物":self.CheckSubMenuItemNumber(f))
        print(self.sub_menu_button_list[0]['text'],self.sub_menu_button_list[1]['text'],)
        """TEST END"""

    def CheckSubMenuItemNumber(self, text):
        for i in range(len(self.sub_menu_button_list)):
            if(self.sub_menu_button_list[i]['text'] == text):
                self.constriction_flag = i
                print(i)
                break

    def SubMenuEvent(self):
        if self.sub_menu_active:
            self.canvas.delete(self.rect)
            self.sub_menu_active = False
        else:
            self.sub_menu_active = True
            self.DrawSubMenu()
            self.DrawSubMenuItem(820,35)



    def Set(self):
        #self.btn = tk.Button(self.canvas, text='text', width=300, height=150, state='active')
        #self.SetButton("サブメニュー",920,20,self.SubMenuEvent)
        self.canvas.bind('<ButtonPress-1>',self.LeftClickEvent)


    def DrawSubMenu(self):
        if self.sub_menu_active:
            self.rect = self.canvas.create_rectangle(800,20,900,400,fill = 'white')
            print("self.rect",self.rect)
        else:
            pass

    def ReadSubMenuItemFromExcel(self):
        pass

    def ReadMapInfoFromExcel(self):
        if(self.map_excel_file_full_name != ""):
            book = openpyxl.load_workbook(self.map_excel_file_full_name)
            active_sheet = book.active
            sheet = book.get_sheet_by_name("MapInfo")
            print(sheet.cell(row = 1, column = 2).value)
            print(sheet.cell(row = 2, column = 2).value)
            print(sheet.cell(row = 3, column = 3).value)

            self.masu_x = sheet.cell(row = 1, column= 2).value
            self.masu_y = sheet.cell(row = 2, column= 2).value
            #return (sheet.cell(row = 1,column=2).value,sheet.cell(row = 2,column=2).value)


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

        self.canvas.bind('<Control-O>', self.OpenExplorer)
        self.canvas.bind('<Control-o>', self.OpenExplorer)

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


    def ReadMapFromExcel(self):
        if (self.map_excel_file_full_name != ""):
            book = openpyxl.load_workbook(self.map_excel_file_full_name)
            active_sheet = book.active
            sheet = book.get_sheet_by_name("Map")
            for i in range(1,self.masu_x + 1):
                temp_list = []
                for j in range(1,self.masu_y + 1):
                    temp_value = sheet.cell(row = j, column= i).value
                    print(temp_value)
                    if(temp_value == None):
                        temp_list.append(0)
                    else:
                        temp_list.append(temp_value)
                self.map_chip.append(temp_list)
            print(self.map_chip)






if __name__ == "__main__":
    root = tk.Tk()
    app = Mapedit(master = root)
    #app.SetMasu(3,3)
    app.DrawMenu()
    app.DrawSubMenu()
    app.Set()

    root.config(menu = app.menu_bar)
    app.mainloop()