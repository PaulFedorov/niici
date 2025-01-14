import wx
import wx.grid as gridlib
import csv


class CSVEditor(wx.Frame):
    def __init__(self, *args, **kwargs):
        super(CSVEditor, self).__init__(*args, **kwargs)

        self.current_file = None

        self.InitUI()

    def InitUI(self):
        # Меню
        menubar = wx.MenuBar()
        fileMenu = wx.Menu()

        openItem = fileMenu.Append(wx.ID_OPEN, 'Open', 'Open a CSV file')
        saveItem = fileMenu.Append(wx.ID_SAVE, 'Save', 'Save the file')
        saveAsItem = fileMenu.Append(wx.ID_SAVEAS, 'Save As...', 'Save to a new file')
        fileMenu.AppendSeparator()
        exitItem = fileMenu.Append(wx.ID_EXIT, 'Exit', 'Exit the application')

        aboutMenu = wx.Menu()
        aboutItem = aboutMenu.Append(wx.ID_ABOUT, 'About', 'About this application')

        menubar.Append(fileMenu, 'File')
        menubar.Append(aboutMenu, 'About')

        self.SetMenuBar(menubar)

        # Привязка событий
        self.Bind(wx.EVT_MENU, self.OnOpen, openItem)
        self.Bind(wx.EVT_MENU, self.OnSave, saveItem)
        self.Bind(wx.EVT_MENU, self.OnSaveAs, saveAsItem)
        self.Bind(wx.EVT_MENU, self.OnExit, exitItem)
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)

        # Таблица
        self.grid = gridlib.Grid(self)
        self.grid.CreateGrid(0, 0)

        # Кнопка Auto Size
        autoSizeBtn = wx.Button(self, label='Auto Size')
        autoSizeBtn.Bind(wx.EVT_BUTTON, self.OnAutoSize)

        # Макет
        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.Add(self.grid, 1, wx.EXPAND | wx.ALL, 5)
        vbox.Add(autoSizeBtn, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        self.SetSizer(vbox)

        self.SetTitle('CSV Reader')
        self.SetSize((800, 600))
        self.Centre()

    def OnOpen(self, event):
        with wx.FileDialog(self, "Open CSV file", wildcard="CSV files (*.csv)|*.csv",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return

            path = fileDialog.GetPath()
            self.LoadCSV(path)

    def LoadCSV(self, path):
        try:
            with open(path, newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                data = list(reader)

            self.grid.ClearGrid()
            if self.grid.GetNumberRows() > 0:
                self.grid.DeleteRows(0, self.grid.GetNumberRows())
            if self.grid.GetNumberCols() > 0:
                self.grid.DeleteCols(0, self.grid.GetNumberCols())

            self.grid.AppendRows(len(data))
            self.grid.AppendCols(len(data[0]))

            for row_idx, row in enumerate(data):
                for col_idx, value in enumerate(row):
                    self.grid.SetCellValue(row_idx, col_idx, value)

            self.current_file = path
        except Exception as e:
            wx.LogError(f"Cannot open file '{path}': {e}")

    def OnSave(self, event):
        if self.current_file:
            self.SaveCSV(self.current_file)
        else:
            self.OnSaveAs(event)

    def OnSaveAs(self, event):
        with wx.FileDialog(self, "Save CSV file", wildcard="CSV files (*.csv)|*.csv",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return

            path = fileDialog.GetPath()
            self.SaveCSV(path)

    def SaveCSV(self, path):
        try:
            # Указываем кодировку utf-8-sig для совместимости с Excel
            with open(path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)

                for row in range(self.grid.GetNumberRows()):
                    writer.writerow([self.grid.GetCellValue(row, col) for col in range(self.grid.GetNumberCols())])

            self.current_file = path
        except Exception as e:
            wx.LogError(f"Cannot save file '{path}': {e}")

    def OnExit(self, event):
        self.Close()

    def OnAutoSize(self, event):
        self.grid.AutoSize()

    def OnAbout(self, event):
        wx.MessageBox('CSV Reader\nVersion 1.0', 'About', wx.OK | wx.ICON_INFORMATION)


if __name__ == '__main__':
    app = wx.App()
    editor = CSVEditor(None)
    editor.Show()
    app.MainLoop()
