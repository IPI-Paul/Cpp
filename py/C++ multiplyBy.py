from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from tkinter.scrolledtext import ScrolledText
from tkinter.colorchooser import askcolor
from ctypes import *
import _ctypes, sys, os
import win32com.client as win32

class MultiplyBy(Frame):
    bgRGB = (0, 0, 108)
    fgRGB = (255, 255, 255)
    
    def __init__(self, parent=None):
        Frame.__init__(self, parent)
        self.dll = cdll.LoadLibrary('../dll/multiplyBy x64.dll')
        self.dll.multiplyBy.argtypes = [POINTER(c_double), POINTER(c_double)]
        self.var1 = StringVar(self)
        self.var2 = StringVar(self)
        self.numrow = 0
        self.numcol = 3
        self.choices = {"General": ",f", "#,###,##0.00": ",.2f",
                    "#,###,##0": ",.0f"}
        self.bgColour = '#%02x%02x%02x' % self.bgRGB
        self.fgColour = '#%02x%02x%02x' % self.fgRGB
        main = Frame(self)
        nb = ttk.Notebook(main, height=500, width=670)
        frame_canvas = Frame(nb)
        self.canvas = Canvas(frame_canvas, width=650, height=500)
        sbar = Scrollbar(frame_canvas, orient="vertical")
        sbar.config(command=self.canvas.yview)
        sbar.grid(row=0, column=1, sticky="ns")
        self.canvas.grid(row=0, column=0, sticky="ns")
        self.canvas.configure(yscrollcommand=sbar.set)
        self.calc = Frame(self.canvas, width=650, height=1000)
        self.canvas.create_window((0, 0), window=self.calc, anchor="nw")
        self.makeWidgets(self.numrow, self.numcol)
        self.var1.trace('w', self.frmtChange)
        self.var2.trace('w', self.runFunction)
        nb.add(frame_canvas, text="Calculate")
        nb.pack(side=LEFT, expand=YES, fill=BOTH)
        main.grid(row=0, columnspan=5)
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def addRow(self):
        self.numrow += 1            
        cols = []
        for j in range(self.numcol):
            if j < self.numcol - 1:
                ent = Entry(self.calc, relief=RIDGE, justify="right")
                ent.grid(row=self.numrow, column=j, sticky=NSEW)
                ent.insert(END, "")
                ent.bind('<Key-Return>', lambda event, x=self.numrow-1: self.onChange(event, row=x))
                ent.bind('<Key-Tab>', lambda event, x=self.numrow-1: self.onChange(event, row=x))
                ent.bind('<Control-c>', self.toClipboard)
                cols.append(ent)
            else:
                lab = Label(self.calc, text="0", relief=RIDGE, anchor="e")
                lab.grid(row=self.numrow, column=self.numcol-1, sticky=NSEW)                
                cols.append(lab)
        self.rows.append(cols)
        
    def centimetersToPoints(self, centimeter):
        return centimeter * 28.3464567
    
    def frmtChange(self, *args):
        for i in range(self.numrow):
            for j in range(self.numcol):
                if j < self.numcol - 1:
                    num1 = self.restyle(1, self.normVal(str(self.rows[i][j].get())))
                    self.rows[i][j].delete(0, END)
                    self.rows[i][j].insert(0, num1)
                else:
                    num1 = self.restyle(1, self.normVal(str(self.rows[i][j].cget("text"))))
                    self.rows[i][j].config(text = num1)
                num2 = self.restyle(2, self.normVal(self.sums[j].cget("text")))
                self.sums[j].config(text = num2)
    
    def makeWidgets(self, numrow, numcol):
        hdrs = ["Number1", "Number2", "Multiplied Result"]
        self.hdrs = []
        for i in range(len(hdrs)):
            lab = Label(self.calc, text=hdrs[i], relief=RIDGE, width=30)
            lab.grid(row=0, column=i, sticky=NSEW)
            lab.config(bg=self.bgColour, fg=self.fgColour)
            self.hdrs.append(lab)
            
        self.rows = []
        self.addRow()
            
        self.sums = []
        for i in range(numcol):
            lab = Label(self.calc, text="0", relief=SUNKEN, anchor="e")
            lab.grid(row=40, column=i, sticky=NSEW)
            self.sums.append(lab)
            
        formats = OptionMenu(self, self.var1, *self.choices)
        formats.grid(row=1, column=3)
        self.var1.set("General")
        
        choices = ["", "Add Row", 
                   "Set Header Back Colour", "Set Header Fore Colour",
                   "Send to New Email", "Send to Open Email", 
                   "Send to New Word Document", "Send to Open Word Document" 
                   ]
        functions = OptionMenu(self, self.var2, *choices)
        functions.grid(row=1, column=4)
        self.var2.set("")
        
    def multiplyBy(self, num, num1):
        num = self.normVal(num)
        num1 = self.normVal(num1)
        
        outdata = (c_double * 1)(num1)
        self.dll.multiplyBy(c_double(num), outdata)
        
        #libHandle = self.dll._handle
        #del self.dll
        #_ctypes.FreeLibrary(libHandle)
        
        return outdata[0]
    
    def normVal(self, num):
        num = str.replace(str.replace(num, ",", ""), "'", "")
        if num == "": num = 0
        return float(num)
    
    def onChange(self, *args, row):
        if row >= 0:
            self.rows[row][2].config(
                    text = self.multiplyBy(self.rows[row][0].get(), 
                                         self.rows[row][1].get()))
        tots = [0] * self.numcol
        for i in range(self.numrow):
            num1 , num2 = 0, 0
            for j in range(self.numcol):
                if j < self.numcol - 1:
                    num1 = self.normVal(str(self.rows[i][j].get()))
                else:
                    num1 = self.normVal(str(self.rows[i][j].cget("text")))
                tots[j] += num1
        for i in range(self.numcol):
            self.sums[i].config(text = self.restyle(2, tots[i]))
        if self.rows[row][0].get() != "" and self.rows[row][1].get() != "":
            self.frmtChange()
            
    def restyle(self, typ, num):
        if typ == 1:
            return format(num, ",f")
        else:
            return format(num, self.choices[self.var1.get()])
    
    def rgb(self, col):
        return eval('+'.join(str(col[x] * (1, 256, 65536)[x]) for x in range(3)))
    
    def runFunction(self, *args):
        if self.var2.get() != "":
            if self.var2.get() == "Add Row":
                self.addRow()
            if self.var2.get() == "Set Header Back Colour":
                self.bgRGB, self.bgColour = askcolor(initialcolor=self.bgColour)
                self.bgRGB = tuple([int(i) for i in self.bgRGB])
                self.updColours()
            if self.var2.get() == "Set Header Fore Colour":
                self.fgRGB, self.fgColour = askcolor(initialcolor=self.fgColour)
                self.fgRGB = tuple([int(i) for i in self.fgRGB])
                self.updColours()
            if self.var2.get() == "Send to New Email":
                tbl = self.tblHtml()
                self.sendToEmail("New", tbl)
            if self.var2.get() == "Send to Open Email":
                tbl = self.tblHtml()
                self.sendToEmail("Open", tbl)
            if self.var2.get() == "Send to New Word Document":
                tbl = "" #self.tblHtml()
                self.sendToWord("New", tbl)
            if self.var2.get() == "Send to Open Word Document":
                tbl = "" #self.tblHtml()
                self.sendToWord("Open", tbl)
                
            self.var2.set("")
    
    def sendToEmail(self, typ, htm):
        outlook = win32.Dispatch("outlook.application")
        if typ == "New":
            oMail = outlook.CreateItem(0)
            oMail.Subject = "C++ multiplyBy Python Test"
            oMail.Display()
            oMail.HtmlBody = htm
        else:
            oMail = outlook.ActiveInspector().CurrentItem
            outlook.ActiveInspector().WordEditor.Application.Selection = "placeHere"
            oMail.HtmlBody = oMail.HtmlBody.replace("placeHere", htm)

    def sendToWord(self, typ, htm):
        wdOrientLandscape = 1; 
        word = win32.Dispatch("word.application")
        if typ == "New":
            word.Visible = True
            doc = word.Documents.Add()
            doc.PageSetup.Orientation = wdOrientLandscape;
            cent = self.centimetersToPoints(1.75)
            _ = doc.PageSetup
            _.TopMargin = _.LeftMargin = _.BottomMargin = _.RightMargin = cent
        else:
            doc = word.ActiveDocument
        self.tblWord(doc)

    def tblHtml(self):
        stl = "<style>\n"
        stl += "table {\n"
        stl += "\tborder-collapse: collapse;\n"
        stl += "}\n"
        stl += "table, th, td {\n"
        stl += "\tpadding: 0px 5px 0px 5px;\n"
        stl += "\tborder: 1 solid black;\n"
        stl += "\tfont-size: 1em;\n"
        stl += "}\n"
        stl += "td {\n"
        stl += "\tcolor: black;\n"
        stl += "\ttext-align: right;\n"
        stl += "}\n"
        stl += "th {\n"
        stl += "\tcolor: rgb(%i, %i, %i);\n" % self.fgRGB
        stl += "\tbackground-color: rgb(%i, %i, %i);\n" % self.bgRGB
        stl += "}\n"
        stl += "</style>\n"
        tbl = "<table id=\"mulitplyBy\">\n"
        trd = "\t<tr>\n"
        tre = "\n\t</tr>\n"
        thd = "\t\t<th>\n\t\t\t"
        the = "\n\t\t</th>\n"
        tdd = thd.replace("th", "td")
        tde = the.replace("th", "td")
        th = tr = tt = ""
        for col in range(self.numcol):
            th += thd + self.hdrs[col].cget("text") + the
        tbl += th
        for row in range(self.numrow):
            td = ""
            for col in range(self.numcol - 1):
                td += tdd + self.restyle(2, self.normVal(self.rows[row][col].get())) + tde
            td += tdd + self.restyle(2, self.normVal(self.rows[row][self.numcol - 1].cget("text"))) + tde
            tr += trd + td + tre
            
        tr += trd + "<td colspan=\"3\" style=\"border-left: none; border-right: none;\">  &nbsp; </td>" + tre            
        for col in range(self.numcol):
            tt += tdd + self.sums[col].cget("text") + tde
        tr += trd + tt + tre
        tbl += tr + "</table>"
        stl += tbl
        return stl
            
    def tblWord(self, doc):
        wdActiveEndPageNumber = 3; wdBorderLeft = -2
        wdBorderRight = -4; wdBorderVertical = -6; wdLineStyleNone = 0
        wdLineStyleSingle = 1; wdAlignParagraphRight = 2
        wdAlignParagraphLeft = 0   
        tbl = doc.Application.Selection.Tables.Add(
            doc.Application.Selection.Range, self.numrow + 3, self.numcol)
        _ = tbl.Borders
        _.InsideLineStyle = wdLineStyleSingle
        _.OutsideLineStyle = wdLineStyleSingle
        _.InsideColor = self.rgb(self.bgRGB)
        _.OutsideColor = self.rgb(self.bgRGB)
        pg = tbl.Rows[1].Range.Information(wdActiveEndPageNumber)
        i = 0
        for j in range(self.numcol):
            _ = tbl.Rows[i].Cells[j].range
            _.text = self.hdrs[j].cget("text")
            _.Shading.BackgroundPatternColor = self.rgb(self.bgRGB)
            _.Font.Color = self.rgb(self.fgRGB)
            _.Font.bold = True
        i += 1
        for row in range(self.numrow):
            if pg < tbl.Rows[i].Range.Information(wdActiveEndPageNumber):
                pg = tbl.Rows[i].Range.Information(wdActiveEndPageNumber)
                tbl.Rows.Add(tbl.Rows[i])
                for j in range(self.numcol):
                    _ = tbl.Rows[i].Cells[j].range
                    _.text = self.hdrs[j].cget("text")
                    _.Shading.BackgroundPatternColor = self.rgb(self.bgRGB)
                    _.Font.Color = self.rgb(self.fgRGB)
                    _.Font.bold = True
                i += 1
            for j in range(self.numcol):
                _ = tbl.Rows[i].Cells[j].range
                if j < self.numcol - 1:
                    num1 = self.normVal(self.rows[row][j].get())
                else:
                    num1 = self.normVal(self.rows[row][j].cget("text"))
                _.text = self.restyle(2, num1)
                _.ParagraphFormat.Alignment = wdAlignParagraphRight
            i += 1        
        if pg < tbl.Rows[i].Range.Information(wdActiveEndPageNumber):
            pg = tbl.Rows[i].Range.Information(wdActiveEndPageNumber)
            tbl.Rows.Add(tbl.Rows[i])
            for j in range(self.numcol):
                _ = tbl.Rows[i].Cells[j].range
                _.text = self.hdrs[j].cget("text")
                _.Shading.BackgroundPatternColor = self.rgb(self.bgRGB)
                _.Font.Color = self.rgb(self.fgRGB)
                _.Font.bold = True
            i += 1        
        _ = tbl.Rows[i].Cells
        _.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        _.Borders(wdBorderRight).LineStyle = wdLineStyleNone
        _.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        i += 1
        if pg < tbl.Rows[i].Range.Information(wdActiveEndPageNumber):
            pg = tbl.Rows[i].Range.Information(wdActiveEndPageNumber)
            tbl.Rows.Add(tbl.Rows[i])
            for j in range(self.numcol):
                _ = tbl.Rows[i].Cells[j].range
                _.text = self.hdrs[j].cget("text")
                _.Shading.BackgroundPatternColor = self.rgb(self.bgRGB)
                _.Font.Color = self.rgb(self.fgRGB)
                _.Font.bold = True
            i += 1
        for j in range(self.numcol):
            _ = tbl.Rows[i].Cells[j].range
            _.text = self.sums[j].cget("text")
            _.ParagraphFormat.Alignment = wdAlignParagraphRight
    
    def toClipboard(self, *args):
        clip = ""
        for col in range(self.numcol):
            clip += "%s" % self.hdrs[col].cget("text")
            if col < self.numcol - 1:
                clip += "\t"
        clip += "\n"
        for row in range(self.numrow):
            for col in range(self.numcol):
                if col < self.numcol - 1:
                    num1 = self.rows[row][col].get()
                else:
                    num1 = self.rows[row][col].cget("text")
                clip += "%s" % self.restyle(2, self.normVal(str(num1))) 
                if col < self.numcol - 1:
                    clip += "\t"
            clip += "\n"
        clip += "\n"
        for col in range(self.numcol):
            clip += "%s" % self.sums[col].cget("text")
            if col < self.numcol - 1:
                clip += "\t"
        objData = Tk()
        objData.withdraw()
        objData.clipboard_clear()
        objData.clipboard_append(clip)
        
    def updColours(self):
        for i in range(self.numcol):
            self.hdrs[i].config(bg = self.bgColour, fg = self.fgColour)
    

if __name__ == '__main__':
    import sys
    if len(sys.argv) == 1:
        root = Tk()
        root.title('IPI Paul - C++ multiplyBy')
        MultiplyBy(root).pack()  
        mainloop()
    else: 
        if len(sys.argv) > 2:
            num, num1 = sys.argv[1:]
            self = MultiplyBy()
            print(MultiplyBy.multiplyBy(self, num, num1))
        else:
            num, num1 = str.split(input('Please give numbers to mulitply by: '))
            self = MultiplyBy()
            print(MultiplyBy.multiplyBy(self, num, num1))
        input('Press any key to exit!')