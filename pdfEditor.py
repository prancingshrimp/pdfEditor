import os, subprocess
import tempfile
import tkinter as tk
import tkinter.filedialog as fd
from tkinter import messagebox
from PIL import ImageTk, Image
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger


class ListFrame(tk.Frame):
    def __init__(self, master, items=[]):
        super().__init__(master)
        # self.list = tk.Listbox(self, selectmode='BROWSE', width=30, height=32)
        self.list = tk.Listbox(self, width=30, height=32)
        self.scroll = tk.Scrollbar(self, orient = tk.VERTICAL, command = self.list.yview)
        self.list.config(yscrollcommand = self.scroll.set)
        self.list.pack(side = tk.LEFT)
        self.scroll.pack(side = tk.LEFT, fill = tk.Y)


    def fillList(self, displayListNames):
        self.list.delete(0, tk.END)
        for i in range(0, len(displayListNames)):
            self.list.insert(tk.END, displayListNames[i])


class pdfFile:
    def __init__(self):
        self.filePath = ""
        self.fileName = ""
        self.pageCount = 0
        self.pageList = []

    def readFilePath(self, filePath):
        self.filePath = filePath
        self.fileName = os.path.basename(self.filePath)
        self.tmpDir = tempfile.TemporaryDirectory()
        self.tmpDirName = self.tmpDir.name

        print('created temporary directory', self.tmpDirName)

        self.prefix = "pdf"
        self.thumbnail = self.prefix + "%d"
        self.extension = ".jpg"
    
        self.thumbnailDir = self.tmpDirName + "\\" + "thumbnails"
        os.mkdir(self.thumbnailDir)
        self.thumbnailFiles = self.thumbnailDir + "\\" + self.thumbnail + self.extension

        self.ghostScript = ".\\Ghost\\gswin32c.exe"
        self.options = (" -dNumRenderingThreads=4 -dNOPAUSE -sDEVICE=jpeg -g200x150 -dPDFFitPage -dORIENT1=true -sOutputFile=" +
            self.thumbnailFiles  + " -dJPEGQ=100 -r300 -q " + self.filePath + " -c quit")

        subprocess.run(self.ghostScript + self.options)

        self.pdfFile = PdfFileReader(open(self.filePath, "rb"))
        self.pageCount = self.pdfFile.getNumPages()
        self.pageList = self.createPageList()

    def createPageList(self):
        tmpList = []
        displayListName = self.fileName.replace(".pdf","")[-20:]
        for i in range(0, self.pageCount):
            tmpList.append([i+1, self.filePath, self.thumbnailDir + "\\" + self.prefix + str(i+1) + self.extension, displayListName + "\\Seite " + str(i+1), "Seite " + str(i+1)])
        return tmpList

    def getPageListNames(self):
        tmpList = []
        for i in range(0, self.pageCount):
            tmpList.append(self.pageList[i][3])
        return tmpList

    def displayFileName(self):
        return self.fileName.replace(".pdf","")[-60:]




class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("pdfEditor")
        self.iconbitmap("pdfEditor.ico")
        self.geometry("950x695+20+20")
        self.resizable(0, 0)

        self.labelFrameLeft  = tk.LabelFrame(self, text="1 - pdf-Datei Zeichnungssatz zur Bearbeitung",     width=450, height=610)
        self.labelFrameLeft.place(x=10,  y=10)
        self.labelFrameRight = tk.LabelFrame(self, text="2 - pdf-Datei zur Ergänzung des Zeichnungssatzes", width=450, height=610)
        self.labelFrameRight.place(x=490, y=10)
        self.frameLeft  = ListFrame(self)
        self.frameLeft.place(x=250, y=90)
        self.frameRight = ListFrame(self)
        self.frameRight.place(x=500, y=90)

        self.btnOeffnenLeft = tk.Button(self, text="Öffnen", command=self.OeffnenLeft)
        self.btnOeffnenLeft.place(x=30, y=45)
        self.labelOeffnenLeft = tk.Label(self, text="-")
        self.labelOeffnenLeft.place(x=85, y=45)

        self.labelLeftVorgaenger = tk.Label(self, text="Vorgänger:")
        self.labelLeftVorgaenger.place(x=30, y=85)
        self.labelLeftVorgaengerCur = tk.Label(self, text="-")
        self.labelLeftVorgaengerCur.place(x=100, y=85)
        self.imgLeftVorgaenger = []
        self.panelLeftVorgaenger = tk.Label(self, image = self.imgLeftVorgaenger)
        self.panelLeftVorgaenger.place(x=30, y=105)

        self.labelLeftAktuell = tk.Label(self, text="Aktuell:")
        self.labelLeftAktuell.place(x=30, y=260)
        self.labelLeftAktuellCur = tk.Label(self, text="-")
        self.labelLeftAktuellCur.place(x=100, y=260)
        self.imgLeftAktuell = []
        self.panelLeftAktuell = tk.Label(self, image = self.imgLeftAktuell)
        self.panelLeftAktuell.place(x=30, y=280)

        self.labelLeftNachfolger = tk.Label(self, text="Nachfolger:")
        self.labelLeftNachfolger.place(x=30, y=435)
        self.labelLeftNachfolgerCur = tk.Label(self, text="-")
        self.labelLeftNachfolgerCur.place(x=100, y=435)
        self.imgLeftNachfolger = []
        self.panelLeftNachfolger = tk.Label(self, image = self.imgLeftNachfolger)
        self.panelLeftNachfolger.place(x=30, y=455)

        self.btnOeffnenRight  = tk.Button(self, text="Öffnen", command=self.OeffnenRight)
        self.btnOeffnenRight.place(x=500, y=45)
        self.labelOeffnenRight = tk.Label(self, text="-")
        self.labelOeffnenRight.place(x=555, y=45)

        self.labelRightVorgaenger = tk.Label(self, text="Vorgänger:")
        self.labelRightVorgaenger.place(x=715, y=85)
        self.labelRightVorgaengerCur = tk.Label(self, text="-")
        self.labelRightVorgaengerCur.place(x=785, y=85)
        self.imgRightVorgaenger = []
        self.panelRightVorgaenger = tk.Label(self, image = self.imgRightVorgaenger)
        self.panelRightVorgaenger.place(x=715, y=105)

        self.labelRightAktuell = tk.Label(self, text="Aktuell:")
        self.labelRightAktuell.place(x=715, y=260)
        self.labelRightAktuellCur = tk.Label(self, text="-")
        self.labelRightAktuellCur.place(x=785, y=260)
        self.imgRightAktuell = []
        self.panelRightAktuell = tk.Label(self, image = self.imgRightAktuell)
        self.panelRightAktuell.place(x=715, y=280)

        self.labelRightNachfolger = tk.Label(self, text="Nachfolger:")
        self.labelRightNachfolger.place(x=715, y=435)
        self.labelRightNachfolgerCur = tk.Label(self, text="-")
        self.labelRightNachfolgerCur.place(x=785, y=435)
        self.imgRightNachfolger = []
        self.panelRightNachfolger = tk.Label(self, image = self.imgRightNachfolger)
        self.panelRightNachfolger.place(x=715, y=455)

        self.btnX      = tk.Button(self, text="X", command=self.Remove)
        self.btnX.place(x=285, y=630)
        self.btnUp     = tk.Button(self, text="˄", command=self.Up)
        self.btnUp.place(x=315, y=630)
        self.btnDown   = tk.Button(self, text="˅", command=self.Down)
        self.btnDown.place(x=345, y=630)
        self.btnInsert = tk.Button(self, text="<", command=self.Insert)
        self.btnInsert.place(x=375, y=630)
        self.btnSave   = tk.Button(self, text="Speichern", command=self.Save)
        self.btnSave.place(x=808, y=660)
        self.btnClose  = tk.Button(self, text="Schließen", command=self.Close)
        self.btnClose.place(x=878, y=660)

        self.labelHint.place(x=5, y=675)

        self.pdfFileLeft  = pdfFile()
        self.pdfFileRightList = []

        self.frameLeft.list.bind( '<<ListboxSelect>>', self.SelectItem)
        self.frameRight.list.bind('<<ListboxSelect>>', self.SelectItemRight)


    def OeffnenLeft(self):
        fileTypes = (("pdf-Datei", "*.pdf"), ("Alle Dateien", "*"))
        filePath = fd.askopenfilename(title="Öffne pdf-Datei", initialdir="/", filetypes=fileTypes)
        if os.path.isfile(filePath):
            self.pdfFileLeft.readFilePath(filePath)
            self.labelOeffnenLeft['text'] = self.pdfFileLeft.displayFileName()
            self.frameLeft.fillList(self.pdfFileLeft.getPageListNames())
            self.frameLeft.list.selection_set(0)
            self.frameLeft.list.activate(0)
            self.SelectItemLeft(0)

    def OeffnenRight(self):
        fileTypes = (("pdf-Datei", "*.pdf"), ("Alle Dateien", "*"))
        filePath = fd.askopenfilename(title="Öffne pdf-Datei", initialdir="/", filetypes=fileTypes)
        if os.path.isfile(filePath):
            tmp = pdfFile()
            tmp.readFilePath(filePath)
            self.labelOeffnenRight['text'] = tmp.displayFileName()
            self.frameRight.fillList(tmp.getPageListNames())
            self.pdfFileRightList.append(tmp)
            self.SelectItemRight(-1)

    def Close(self):
        self.after(50, self.destroy)

    def Remove(self):
        if self.pdfFileLeft.pageCount == 1:
            return
        active = self.frameLeft.list.get(tk.ACTIVE)
        self.frameLeft.list.delete(tk.ACTIVE)
        new = []
        pageNumber = 0
        for i in range(0, len(self.pdfFileLeft.pageList)):
            if self.pdfFileLeft.pageList[i][3] != active:
                new.append(self.pdfFileLeft.pageList[i])
            else:
                pageNumber = i
        self.pdfFileLeft.pageList = new
        self.pdfFileLeft.pageCount = self.pdfFileLeft.pageCount-1
        if pageNumber == 0:
            pageNumber = 1
        self.frameLeft.list.selection_set(pageNumber-1)
        self.frameLeft.list.activate(pageNumber-1)
        self.SelectItemLeft(pageNumber-1)

    def Up(self):
        pageNumber = self.frameLeft.list.index(tk.ACTIVE)
        if pageNumber == 0:
            return
        pre = self.pdfFileLeft.pageList[pageNumber-1]
        current = self.pdfFileLeft.pageList[pageNumber]
        self.pdfFileLeft.pageList[pageNumber] = pre
        self.pdfFileLeft.pageList[pageNumber-1] = current
        self.frameLeft.fillList(self.pdfFileLeft.getPageListNames())
        self.frameLeft.list.selection_set(pageNumber-1)
        self.frameLeft.list.activate(pageNumber-1)
        self.SelectItemLeft(pageNumber-1)

    def Down(self):
        pageNumber = self.frameLeft.list.index(tk.ACTIVE)
        if pageNumber == self.pdfFileLeft.pageCount-1:
            return
        suc = self.pdfFileLeft.pageList[pageNumber+1]
        current = self.pdfFileLeft.pageList[pageNumber]
        self.pdfFileLeft.pageList[pageNumber] = suc
        self.pdfFileLeft.pageList[pageNumber+1] = current
        self.frameLeft.fillList(self.pdfFileLeft.getPageListNames())
        self.frameLeft.list.selection_set(pageNumber+1)
        self.frameLeft.list.activate(pageNumber+1)
        self.SelectItemLeft(pageNumber+1)

    def SelectItemLeft(self, selection):
        if selection != 0:
            self.imgLeftVorgaenger = ImageTk.PhotoImage(Image.open(self.pdfFileLeft.pageList[selection-1][2]))
            self.panelLeftVorgaenger['image'] = self.imgLeftVorgaenger
            self.labelLeftVorgaengerCur['text'] = self.pdfFileLeft.pageList[selection-1][4]
        else:
            self.imgLeftVorgaenger = []
            self.panelLeftVorgaenger['image'] = self.imgLeftVorgaenger
            self.labelLeftVorgaengerCur['text'] = "-"
        
        self.imgLeftAktuell = ImageTk.PhotoImage(Image.open(self.pdfFileLeft.pageList[selection][2]))
        self.panelLeftAktuell['image'] = self.imgLeftAktuell
        self.labelLeftAktuellCur['text'] = self.pdfFileLeft.pageList[selection][4]
        
        if selection != self.pdfFileLeft.pageCount-1:
            self.imgLeftNachfolger = ImageTk.PhotoImage(Image.open(self.pdfFileLeft.pageList[selection+1][2]))
            self.panelLeftNachfolger['image'] = self.imgLeftNachfolger
            self.labelLeftNachfolgerCur['text'] = self.pdfFileLeft.pageList[selection+1][4]
        else:
            self.imgLeftNachfolger = []
            self.panelLeftNachfolger['image'] = self.imgLeftNachfolger
            self.labelLeftNachfolgerCur['text'] = "-"

    def SelectItem(self, evt):
        if evt == -1:
            selection = 0
        else:
            try:
                selection = evt.widget.curselection()[0]
            except:
                return
        self.SelectItemLeft(selection)

    def SelectItemRight(self, evt):
        if evt == -1:
            selection = 0
        else:
            try:
                # print("Right: ", evt.widget.curselection()[0])
                selection = evt.widget.curselection()[0]
            except:
                return
            
        if selection != 0:
            self.imgRightVorgaenger = ImageTk.PhotoImage(Image.open(self.pdfFileRightList[-1].pageList[selection-1][2]))
            self.panelRightVorgaenger['image'] = self.imgRightVorgaenger
            self.labelRightVorgaengerCur['text'] = self.pdfFileRightList[-1].pageList[selection-1][4]
        else:
            self.imgRightVorgaenger = []
            self.panelRightVorgaenger['image'] = self.imgRightVorgaenger
            self.labelRightVorgaengerCur['text'] = "-"
        
        self.imgRightAktuell = ImageTk.PhotoImage(Image.open(self.pdfFileRightList[-1].pageList[selection][2]))
        self.panelRightAktuell['image'] = self.imgRightAktuell
        self.labelRightAktuellCur['text'] = self.pdfFileRightList[-1].pageList[selection][4]
        
        if selection != self.pdfFileRightList[-1].pageCount-1:
            self.imgRightNachfolger = ImageTk.PhotoImage(Image.open(self.pdfFileRightList[-1].pageList[selection+1][2]))
            self.panelRightNachfolger['image'] = self.imgRightNachfolger
            self.labelRightNachfolgerCur['text'] = self.pdfFileRightList[-1].pageList[selection+1][4]
        else:
            self.imgRightNachfolger = []
            self.panelRightNachfolger['image'] = self.imgRightNachfolger
            self.labelRightNachfolgerCur['text'] = "-"


    def Insert(self):
        pageNumberLeft  = self.frameLeft.list.index(tk.ACTIVE)
        pageNumberRight = self.frameRight.list.index(tk.ACTIVE)
        new = []
        for i in range(0, len(self.pdfFileLeft.pageList)):
            new.append(self.pdfFileLeft.pageList[i])
            if i == pageNumberLeft:
                new.append(self.pdfFileRightList[-1].pageList[pageNumberRight])

        self.pdfFileLeft.pageList = new
        self.pdfFileLeft.pageCount = self.pdfFileLeft.pageCount+1
        self.frameLeft.fillList(self.pdfFileLeft.getPageListNames())
        self.frameLeft.list.selection_set(pageNumberLeft+1)
        self.frameLeft.list.activate(pageNumberLeft+1)
        self.SelectItemLeft(pageNumberLeft+1)
            


    def Save(self):
        if self.pdfFileLeft.fileName == "":
            messagebox.showerror(title="Fehler", message="Keine pdf-Datei zur Bearbeitung ausgewählt!")
            return

        fileTypes = (("pdf-Datei", "*.pdf"), ("Alle Dateien", "*"))
        fh = fd.asksaveasfile(mode='w', title="Speichere pdf-Datei", initialfile=self.pdfFileLeft.fileName, filetypes=fileTypes, defaultextension = fileTypes)
        filePath = fh.name
        fh.close()
        if filePath[-4:] != ".pdf":
            filePath = filePath + ".pdf"
        # print(filePath)

        # merger = PdfFileMerger()
        # for line in self.pdfFileLeft.pageList:
        #     print(line[1], line[0])
        #     merger.append(fileobj = line[1], page = line[0]-1)
        #     merger.append()
        # #     pass
        # print(self.pdfFileLeft.pageList)

        # merger.append(fileobj = input1, pages = (0,3))

        # output = open(filePath, "wb")
        # merger.write(output)

        output = PdfFileWriter()
        for line in self.pdfFileLeft.pageList:
            pdfFile = PdfFileReader(open(line[1], "rb"))
            output.addPage(pdfFile.getPage(line[0]-1))
        outputStream = open(filePath,"wb")
        output.write(outputStream)
        outputStream.close()


if __name__ == "__main__":
    app = App()
    app.mainloop()