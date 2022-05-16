Sub Read_File()

Dim batch as String
Dim Input_Path as String
Dim Output_Path as String
Dim Script as String
Dim Securities as String 
Dim execute as String
Dim strquote as String

strquote = Chr$(34)

'get batch file to activate command prompt with py environments
batch = Range("batch")

'get filepath from excel cell
Input_Path = Range('Input_Path')
Output_Path = Range('Output_Path')

'get python script
Script = "/Users/linngyinn/Desktop/Py_Scripts/Script_to_Run.py"

'get variables to pass to py class files
Securities = Range("Securities")

execute = "call " & batch & " base & python " & strquote & Script & strquote _
        & " " & strquote & Input_Path & strquote & " " & strquote & Output_Path _
        & " " & strquote & Securities & strquote
Call Shell("cmd.exe /S /C" & execute, vbHide)

End Sub

