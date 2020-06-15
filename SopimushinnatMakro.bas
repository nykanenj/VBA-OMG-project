
Public thisWB As Workbook
Public resultSheet As Worksheet


Sub contractColumnsMacro()

    Set thisWB = ThisWorkbook
    Set resultSheet = createSheet()

End Sub

Private Function createSheet() As Worksheet

    Dim sheetName As String
    Dim ws As Worksheet

    Dim i As Integer

    Dim y As Integer
    Dim m As Integer
    Dim d As Integer

    y = Year(Now())
    m = Month(Now())
    d = Day(Now())

    sheetName = "Lopputulos_" & Day(Now()) & "_" & Month(Now()) & "_klo_" & Hour(Now()) & "_" & Minute(Now())

    i = 0
    Do While i < 10
        i = i + 1

        For Each ws In thisWB.Worksheets
            If ws.Name = sheetName Then GoTo NextCycle
        Next ws
        GoTo Continue

NextCycle:
        sheetName = sheetName & "(" & i & ")"
    Loop

Continue:

    Set createPDFsheet = thisWB.Sheets.Add(After:=thisWB.Sheets(thisWB.Sheets.Count))
    createPDFsheet.Name = sheetName
    DoEvents

End Function

Private Sub displayInstructions()

    MsgBox "Ohjeet:" & vbCrLf & vbCrLf & _
    "1. Täytä 'Sopimushinnat' -välilehti." & vbCrLf & _
    "2. Lisää ohjelmasta saatu tuntiraportti samaan kansioon tämän tiedoston kanssa." & vbCrLf & _
    "3. Paina nappia 'Lisää sopimushinnat.'" & vbCrLf & _
    "4. Yhdistetty lopputulos ilmestyy uudelle välilehdelle."
End Sub

Private Sub addError(i As Variant, error As String)

    Dim errorText As String

    If D_errors.exists(i) Then
        errorText = D_errors(i) & " " & error
        D_errors.Add i, errorText
    Else
        D_errors.Add i, error
    End If
    
End Sub

Private Sub addWarning(i As Variant, warning As String)

    Dim warningText As String

    If D_warnings.exists(i) Then
        warningText = D_warnings(i) & " " & warning
        D_warnings.Add i, warningText
    Else
        D_warnings.Add i, warning
    End If
    
End Sub

Sub saveByDateTime()

    Dim filenameAndPath As String

    filenameAndPath = ThisWorkbook.Path & "\SopimusHinnatPohja_" & Year(Now()) & "_" & Month(Now()) & "_" & Day(Now()) & "_klo_" & Hour(Now()) & "_" & Minute(Now()) & ".xlsm"

    ActiveWorkbook.SaveAs Filename:=filenameAndPath

End Sub


