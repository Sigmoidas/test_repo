Attribute VB_Name = "Module1"
Sub ExportToTxt()
    'export activesheet as txt file
    Dim myPath As String, myFile As String, thisPath As String, thisFile As String
    
    'Getting this Excel file path
    thisPath = Application.ThisWorkbook.Path
    thisFile = Application.ThisWorkbook.FullName
    
    'Getting this Excel file name
    t = Split(thisFile, "\")
    fName = Mid(t(UBound(t)), 1, InStr(1, t(UBound(t)), ".") - 1)
    
    'Getting date for "03_beelden" folder
    fDate = Split(fName, "_")(1)
    
    'Get directory two level up
    ChDir thisPath
    ChDir ".."
    ChDir ".."
    TwoLvlUpDir = CurDir()
    
    '"03_beelden" folder directory
    myPath = TwoLvlUpDir & "\03_beelden" & "\" & fDate & "\"
    
    '"*.txt" file name
    myFile = fName & ".txt"
    
    'Check directory
    CheckDir = Dir(myPath, vbDirectory)

    If CheckDir <> "" Then
        'MsgBox CheckDir & " folder exists"
        Dim WB As Workbook, newWB As Workbook
        Set WB = ThisWorkbook
        Application.ScreenUpdating = False
        Set newWB = Workbooks.Add
        WB.ActiveSheet.UsedRange.Copy newWB.Sheets(1).Range("A1")
        With newWB
            Application.DisplayAlerts = False
            .SaveAs Filename:=myPath & myFile, FileFormat:=xlText
            .Close True
            Application.DisplayAlerts = True
        End With
        WB.Save
        Application.ScreenUpdating = True
    Else
        MkDir TwoLvlUpDir & "\03_beelden" & "\" & fDate
        MsgBox "A folder has been created with the name – " & fDate
    End If
  
'    myPath = TwoLvlUpDir & "\03_beelden" & "\" & fDate & "\"
'    myFile = fName & ".txt"
'    Dim WB As Workbook, newWB As Workbook
'    Set WB = ThisWorkbook
'    Application.ScreenUpdating = False
'    Set newWB = Workbooks.Add
'    WB.ActiveSheet.UsedRange.Copy newWB.Sheets(1).Range("A1")
'    With newWB
'        Application.DisplayAlerts = False
'        .SaveAs Filename:=myPath & myFile, FileFormat:=xlText
'        .Close True
'        Application.DisplayAlerts = True
'    End With
'    WB.Save
'    Application.ScreenUpdating = True
'

    
    
    'MsgBox thisPath & vbCrLf & TwoLvlUpDir & vbCrLf & myPath
    
End Sub
