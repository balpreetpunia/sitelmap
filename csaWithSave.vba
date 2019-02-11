Sub Surfy()

Rows(1).EntireRow.Delete
Columns(1).EntireColumn.Delete
Columns(1).EntireColumn.Delete
Columns(1).EntireColumn.Delete
Columns(1).EntireColumn.Delete
Columns(3).EntireColumn.Delete
Columns(3).EntireColumn.Delete


Dim r As Integer

For r = Sheet1.UsedRange.Rows.Count To 1 Step -1
    If (Cells(r, "B") = "null") Then
        Sheet1.Rows(r).EntireRow.Delete
    End If
Next

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd1 = "Electrolux -"
fnd2 = "Electrolux Icon -"
fnd3 = "Bosch -"
fnd4 = "Bosch 300 Series -"
fnd5 = "Bosch 500 Series -"
fnd6 = "Bosch 800 Series -"
fnd7 = "Bosch Ascenta Series -"
fnd8 = "Bosch Benchmark Series -"
fnd9 = "Frigidaire -"
fnd10 = "Frigidaire Gallery -"
fnd11 = "Frigidaire Professional -"
fnd12 = "KitchenAid -"
fnd13 = "LG -"
fnd14 = "LG Studio -"
fnd15 = "Maytag -"
fnd16 = "Whirlpool -"
fnd17 = "Samsung -"
fnd18 = "Samsung Chef Collection -"


rplc = ""

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd1, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd2, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd3, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd4, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd5, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd6, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd7, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd8, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd9, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd10, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd11, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd12, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd13, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd14, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd15, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd16, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd17, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht
For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd18, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht


For r = Sheet1.UsedRange.Rows.Count To 1 Step -1
    Cells(r, "B").Value = Cells(r, "B") - 1
Next

For r = Sheet1.UsedRange.Rows.Count To 1 Step -1
    Cells(r, "A").Value = Replace(Cells(r, "A").Value, Chr(10), "")
Next

Columns(2).EntireColumn.Insert
Columns(2).EntireColumn.Insert
    
Active = ActiveWorkbook.Name
new_name = Replace(Active, ".csv", ".xlsx", 1, 1)
relativePath = ThisWorkbook.Path & "\" & new_name
    ActiveWorkbook.SaveAs Filename:=relativePath, FileFormat:=51
    ActiveWorkbook.Close SaveChanges:=True
End Sub

Sub LoopAllExcelFilesInFolder()
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com

Dim wb As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
    .InitialFileName = "D:\Teletime\scrape"
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.csv*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(Filename:=myPath & myFile)
    
    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    
    Call Surfy
      
    'Ensure Workbook has closed before moving on to next line of code
      DoEvents

    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
