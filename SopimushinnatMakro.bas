
Sub contractColumnsMacro()


End Sub

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
