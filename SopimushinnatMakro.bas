
Public thisWB As Workbook
Public sourceWB as Workbook
Public contractPricesSheet As Worksheet
Public sourceSheet As Worksheet
Public resultSheet As Worksheet
Public errorSheet As Worksheet

Public D_warnings As Scripting.Dictionary
Public D_errors As Scripting.Dictionary

Public Const CNST_errorSheetName As String = "Virheet Makroajossa"
Public Const CNST_contractPricesSheetName As String = "Sopimushinnat"

Sub contractColumnsMacro()

    Call setDictsWorkbooksAndSheets()
    Call gatherContractPrices()
    Call cleanup()

End Sub

Private Sub runChecks()

    Dim sheetName As String
    Dim ws As Worksheet

    Dim bool_contractPricesSheet As Boolean
    Dim bool_errorSheet As Boolean

    bool_contractPricesSheet = False
    bool_errorSheet = False

    For Each ws in thisWB.Worksheets
        Select Case ws.Name
            Case CNST_contractPricesSheetName
                bool_contractPricesSheet = True
            Case CNST_errorSheetName
                bool_errorSheet = True
                Set errorSheet = ws
        End Select
    Next ws

    If Not bool_contractPricesSheet Then 
        D_errors.Add "Sopimushinnat -välilehti puuttuu", True
    Else
        Set contractPricesSheet = thisWB.Sheets(CNST_contractPricesSheetName)
    End If
    If Not bool_errorSheet Then 
        Set errorSheet = thisWB.Sheets.Add(After:=thisWB.Sheets(thisWB.Sheets.Count))
        errorSheet.Name = CNST_errorSheetName
    Else
        Set errorSheet = thisWB.Sheets(CNST_errorSheetName)
    End If

End Sub

Private Sub setDictsWorkbooksAndSheets()

    Set D_warnings = New Scripting.Dictionary
    Set D_errors = New Scripting.Dictionary
    Set thisWB = ThisWorkbook
    Call runChecks()
    Set sourceSheet = setSourceSheet()

    If D_errors.Count > 0 Then
        Call warningsAndErrors
        Exit Sub
    End If

    Set resultSheet = createSheet()

End Sub

Private Function setSourceSheet() As WorkSheet

    Dim strPath As String
    Dim strFile As String
    Dim errorMsg As String
    Dim errorCode As String

    strPath = thisWB.Path & "\"
    strFile = Dir(strPath & "*.xlsx")

    If strFile = "" Then
        errorCode = "404"
        errorMsg = "Could not find a file in current folder with file ending .xlsx"
        GoTo handleError
    End If

    'Do While strFile <> ""
    Set sourceWB = Workbooks.Open(Filename:=strPath & strFile)
    DoEvents

    If sourceWB.Worksheets.count > 1 Then
        errorCode = "500"
        errorMsg = "Tried to read from workbook " & strFile & vbCrLf & _ 
                    "Error: Too many worksheets in workbook. Should have only one worksheet."
        GoTo handleError
    End If
    
    Set setSourceSheet = sourceWB.Worksheets(1)

    'strFile = Dir 'This moves the value of strFile to the next file.
    'Loop

    Exit Function
handleError:
    Call addError(errorCode, errorMsg)

End Function

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

    sourceSheet.Copy Before:=ThisWB.Sheets(1)
    Set createSheet = thisWB.Sheets(1)
    createSheet.Name = sheetName
    DoEvents

End Function

Sub gatherContractPrices()

    Dim contractPrices As Variant
    Dim i As Integer
    Dim j As Integer
    Dim key as Variant


    contractPrices = contractPricesSheet.Range("A1").CurrentRegion.Value

    For i = 3 to UBound(contractPrices, 1)

        key = contractPrices(i, 1)

        contractPricesObj.partnersSopimushinta = contractPrices(i, 2)
        contractPricesObj.servicesSopimushinta = contractPrices(i, 3)
        contractPricesObj.planningSopimushinta = contractPrices(i, 4)
        contractPricesObj.insightSopimushinta = contractPrices(i, 5)
        contractPricesObj.dashTech = contractPrices(i, 6)
        contractPricesObj.digAnalytic = contractPrices(i, 7)
        contractPricesObj.marketScien = contractPrices(i, 8)
        contractPricesObj.stratCons = contractPrices(i, 9)

    Next i


End Sub

Sub cleanup()

    sourceWB.Close SaveChanges:= False

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

Private Sub warningsAndErrors()

    Dim key As Variant
    Dim error As String

    warning = ""
    For Each key In D_warnings.Keys()
            warning = warning & vbCrLf & D_warnings(key)
    Next key
    If warning <> "" Then MsgBox warning, vbExclamation, "Varoitukset"

    error = ""

    For Each key In D_errors.Keys()
            error = error & vbCrLf & key
    Next key
    If error <> "" Then MsgBox error, vbCritical, "Virheet makron ajossa"

    errorSheet.Range("A1") = warning
    errorSheet.Range("A2") = error
    DoEvents

End Sub

Sub saveByDateTime()

    Dim filenameAndPath As String

    filenameAndPath = ThisWorkbook.Path & "\SopimusHinnatPohja_" & Year(Now()) & "_" & Month(Now()) & "_" & Day(Now()) & "_klo_" & Hour(Now()) & "_" & Minute(Now()) & ".xlsm"

    ActiveWorkbook.SaveAs Filename:=filenameAndPath

End Sub


Private Sub testFuncCopySheet()

'Not used, remove when done.
    Application.ScreenUpdating = False
 
    Set closedBook = Workbooks.Open("D:\Dropbox\excel\articles\example.xlsm")
    closedBook.Sheets("Sheet1").Copy Before:=ThisWorkbook.Sheets(1)
    closedBook.Close SaveChanges:=False
 
    Application.ScreenUpdating = True

End Sub


