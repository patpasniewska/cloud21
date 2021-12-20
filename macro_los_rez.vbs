'Input Excel File's Full Path
  ExcelFilePath = "C:\Users\Patrycja_Project\Desktop\Projekt\obliczenia.xlsm"

'Input Module/Macro name within the Excel File
  MacroPath = "Module2.los_rez"

'Create an instance of Excel
  Set ExcelApp = CreateObject("Excel.Application")

'Do you want this Excel instance to be visible?
  ExcelApp.Visible = False

'Prevent any App Launch Alerts (ie Update External Links)
  ExcelApp.DisplayAlerts = False

'Open Excel File
  Set wb = ExcelApp.Workbooks.Open(ExcelFilePath)

'Execute Macro Code
  ExcelApp.Run MacroPath

'Save Excel File (if applicable)
  wb.Save

'Reset Display Alerts Before Closing
  ExcelApp.DisplayAlerts = True

'Close Excel File
  wb.Close

'End instance of Excel
  ExcelApp.Quit
  
