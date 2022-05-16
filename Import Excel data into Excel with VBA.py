Sub Import_Data()

Dim oFSO as Object
Dim oFolder as Object
Dim oFile as Object

Dim Output_Path as String
Dim path as String 
Dim nwb as Workbook
Dim sh as Worksheet
Dim owb as Workbook
Dim ws as Worksheet

Dim sec as Integer
Dim model as Integer
Dim data as Integer

Application.DisplayAlerts = False
Application.ScreenUpdating = False

path = Left(Output_Path, Len(Output_Path) - 1)
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFso.GetFolder(path)
If Right(path, 1) <> "\" Then path = path & "\"

Set nwb = ThisWorkbook

Dim found as Boolean
found = False
For Each sh in ThisWorkbook.Sheets
    If sh.name = "Metrics" Then
        found = True
        Exit For
    End If
Next

If Not found Then
    Sheets.Add(After:=Sheets(Sheets.Count)).name = "Metrics"
End If

Set sh = nwb.Worksheets("Metrics")
sh.Cells.Clear
sh.Range("B2").Value = "Security_Name"
sh.Range("C2").Value = "Model_Name"
sh.Range("D2").Value = "Stats"

sec = 3
model = 3
data = 3

For each oFile in oFolder.Files

    'open each file
    Set owb = Workbooks.Open(path & oFile.name)
    
    'input security name
    sec_name = Split(Split(oFile.name, "_")(1), ".")(0)
    sh.Cells(sec, 2) = sec_name
    sec = sec + 3 * owb.Worksheets.Count
    
    For Each ws In owb.Worksheets
    
        'input model name
        sh.Cells(model, 3) = ws.name
        If Right(sh.Cells(model,3),1) = "0" Then
            sh.Cells(model, 3) = Replace(sh.Cells(model, 3), "0", "Full_Data")
        Else 
            sh.Cells(model, 3) = Replace(sh.Cells(model, 3), "1", "Test_Data")
        End If
        model = model + 3
        
        'input data
        ws.Range("A1:L3").Copy sh.Cells(data,4)
        data = data + 4
        
    Next ws
    
    owb.Close SaveChanges:=False #dont make changes to excel that is copying
        
Next oFile

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

