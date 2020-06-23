Public partnersSopimushinta As Double
Public servicesSopimushinta As Double
Public planningSopimushinta As Double
Public insightSopimushinta As Double
Public dashTech As Double
Public digAnalytic As Double
Public marketScien As Double
Public stratCons As Double

Public Function fetchColumnHeading(columnHeading As String) As String

    Select Case columnHeading
    Case "ClientPartners KTH"
        fetchColumnHeading = "Client Partners Sopimushinta"
    Case "ClientServices KTH"
        fetchColumnHeading = "Client Services Sopimushinta"
    Case "CI Services Planning KTH"
        fetchColumnHeading = "CI Services Planning Sopimushinta"
    Case "CustomerInsight KTH"
        fetchColumnHeading = "Customer Insight Sopimushinta"
    Case "Dash&TechKTH"
        fetchColumnHeading = "Dash&Tech Sopimushinta"
    Case "DigAnalyticKTH"
        fetchColumnHeading = "Dig Analytic Sopimushinta"
    Case "Market ScienKTH"
        fetchColumnHeading = "Market Scien Sopimushinta"
    Case "Strat&ConsKTH"
        fetchColumnHeading = "Strat&Cons Sopimushinta"
    Case Else
        fetchColumnHeading = ""
    End Select

End Function

Public Function fetchCorrectValue(columnHeading As String) As Double

    Select Case columnHeading
    Case "ClientPartners KTH"
        fetchCorrectValue = partnersSopimushinta
    Case "ClientServices KTH"
        fetchCorrectValue = servicesSopimushinta
    Case "CI Services Planning KTH"
        fetchCorrectValue = planningSopimushinta
    Case "CustomerInsight KTH"
        fetchCorrectValue = insightSopimushinta
    Case "Dash&TechKTH"
        fetchCorrectValue = dashTech
    Case "DigAnalyticKTH"
        fetchCorrectValue = digAnalytic 
    Case "Market ScienKTH"
        fetchCorrectValue = marketScien
    Case "Strat&ConsKTH"
        fetchCorrectValue = stratCons
    Case Else
        fetchCorrectValue = -404
    End Select

End Function