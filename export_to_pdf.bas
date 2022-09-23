'Made with the power of the internet by NotAllur (https://github.com/NotAllur).
Sub Export_to_PDF()
    Dim boxTitle As String
    boxTitle = "Export to PDF"
    Dim currentPath As String
        currentPath = ThisWorkbook.Path
    Dim ws As Worksheet
    Dim desired As Integer
        desired = MsgBox("Is this the desired export directory? The current directory path is" & currentPath & ".", vbYesNo + vbQuestion, boxTitle)
            If desired = vbNo Then
                Dim manInput As Integer
                manInput = MsgBox("Would you like to manually enter the path to the desired target directory?", vbYesNo + vbQuestion, boxTitle)
                    If manInput = vbYes Then
                        Dim srcPath As String
                        srcPath = InputBox("Enter the path to the desired target directory.", boxTitle, currentPath)
                            For Each ws In Worksheets
                                ws.Select
                                nm = ws.Name
                                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                Filename:=srcPath & "\" & nm & ".pdf", _
                                Quality:=xlQualityStandard, IncludeDocProperties:=False, _
                                IgnorePrintAreas:=True, OpenAfterPublish:=False
                            Next ws
                                MsgBox "Success! All worksheets have been exported to " & srcPath & ".", vbInformation, boxTitle
                            Exit Sub
                    Else
                        MsgBox "Export aborted. Transfer your Excel file to the desired export folder or enter the desired directory path manually.", vbCritical, boxTitle
                        Exit Sub
                End If
            End If
    For Each ws In Worksheets
        ws.Select
        nm = ws.Name

        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=currentPath & "\" & nm & ".pdf", _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=True, OpenAfterPublish:=False
    Next ws
        MsgBox "Success! All worksheets have been exported.", vbInformation, boxTitle
End Sub

