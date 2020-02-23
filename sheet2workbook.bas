Option Explicit
 
Sub SaveSheetsAsBook()

  On Error GoTo Error_Handler

    Dim fileExtension, fileNamePrefix, OSpathSep
    Dim Sheet As Worksheet, SheetName$, MyFilePath$, N&
    
    ' For Mac OS
    OSpathSep = "/"
    ' For Windows
    'OSpathSep = "\"
    
    ' << Set a filename prefix to the new Workbook
    fileNamePrefix = InputBox("Please provide the prefix of the filename", "Filename prefix")
    
    If Len(fileNamePrefix) = 0 Then
        fileNamePrefix = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
    End If
    
    If vbYes = MsgBox(fileNamePrefix & vbCrLf & "Do you want to continue?", vbYesNo, "Continue?") Then
        MyFilePath$ = ActiveWorkbook.Path & OSpathSep & fileNamePrefix
        With Application
            .ScreenUpdating = False
            .DisplayAlerts = False
             '      End With
            On Error Resume Next '<< a folder exists
            MkDir MyFilePath '<< create a folder
            For N = 1 To Sheets.Count
                Sheets(N).Activate
                SheetName = ActiveSheet.Name
                Cells.Copy
                Workbooks.Add (xlWBATWorksheet)
                With ActiveWorkbook
                    With .ActiveSheet
                        .Paste
                        .Name = SheetName
                        [A1].Select
                    End With
                     'save book in this folder
                    .SaveAs FileName:=MyFilePath _
                    & OSpathSep & Left(ThisWorkbook.Name, 28) & SheetName & ".xlsx"
                    .Close SaveChanges:=True
                End With
                .CutCopyMode = False
            Next
        End With
        Sheet1.Activate
    Else
        Exit Sub
    End If

Error_Handler:
    MsgBox "The following error has occurred." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: SaveSheetsAsBook" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occurred!"
    Exit Sub
End Sub
