Public partnersSopimushinta As Double
Public servicesSopimushinta As Double
Public planningSopimushinta As Double
Public insightSopimushinta As Double
Public dashTech As Double
Public digAnalytic As Double
Public marketScien As Double
Public stratCons As Double
Public clservDigplan As Double
Public some As Double
Public sem As Double
Public prog As Double
Public cxSeoCpoCont As Double
Public cxCustDev As Double
Public cxInsDmp As Double
Public proOffline As Double
Public proCS As Double
Public videoOffline As Double
Public videoCS As Double
Public ia As Double
Public bonusKord As Double

Public Function fetchColumnHeading(columnHeading As String) As String

    Select Case columnHeading
    Case "ClientPartners KTH"
        fetchColumnHeading = "Client Partners Sopimushinta"
    Case "ClientServices KTH"
        fetchColumnHeading = "Client Services Sopimushinta"
    Case "Cl ServicesPlanning KTH"
        fetchColumnHeading = "Cl Services Planning Sopimushinta"
    Case "CustomerInsight KTH"
        fetchColumnHeading = "Customer Insight Sopimushinta"
    Case "Dash&TechKTH"
        fetchColumnHeading = "Dash&Tech Sopimushinta"
    Case "Dig AnalyticKTH"
        fetchColumnHeading = "Dig Analytic Sopimushinta"
    Case "Market ScienKTH"
        fetchColumnHeading = "Market Scien Sopimushinta"
    Case "Strat&ConsKTH"
        fetchColumnHeading = "Strat&Cons Sopimushinta"
    Case "cl serv /digplan KTH"
        fetchColumnHeading = "cl serv /dig Sopimushinta"
    Case "SOMEKTH"
        fetchColumnHeading = "SOME Sopimushinta"
    Case "SEMKTH"
        fetchColumnHeading = "SEM Sopimushinta"
    Case "PROGKTH"
        fetchColumnHeading = "PROG Sopimushinta"
    Case "CX SEO,CPO ContKTH"
        fetchColumnHeading = "CX SEO,CPO Cont Sopimushinta"
    Case "CX Cust DevKTH"
        fetchColumnHeading = "CX Cust Dev Sopimushinta"
    Case "CX Ins.&DMPKTH"
        fetchColumnHeading = "CX Ins.&DMP Sopimushinta"
    Case "PRO OfflineKTH"
        fetchColumnHeading = "PRO Offline Sopimushinta"
    Case "PRO CSKTH"
        fetchColumnHeading = "PRO CS Sopimushinta"
    Case "Video OfflineKTH"
        fetchColumnHeading = "Video Offline Sopimushinta"
    Case "Video CSKTH"
        fetchColumnHeading = "Video CS Sopimushinta"
    Case "I&AKTH" 'Is it I Or l? It is I
        fetchColumnHeading = "I&A Sopimushinta"
    Case "Bonus & KordKTH"
        fetchColumnHeading = "Bonus & Kord Sopimushinta"
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
    Case "Cl ServicesPlanning KTH" 'In ih I or l? It is l.
        fetchCorrectValue = planningSopimushinta
    Case "CustomerInsight KTH"
        fetchCorrectValue = insightSopimushinta
    Case "Dash&TechKTH"
        fetchCorrectValue = dashTech
    Case "Dig AnalyticKTH"
        fetchCorrectValue = digAnalytic 
    Case "Market ScienKTH"
        fetchCorrectValue = marketScien
    Case "Strat&ConsKTH"
        fetchCorrectValue = stratCons
    Case "cl serv /digplan KTH"
        fetchCorrectValue = clservDigplan
    Case "SOMEKTH"
        fetchCorrectValue = some
    Case "SEMKTH"
        fetchCorrectValue = sem
    Case "PROGKTH"
        fetchCorrectValue = prog
    Case "CX SEO,CPO ContKTH"
        fetchCorrectValue = cxSeoCpoCont
    Case "CX Cust DevKTH"
        fetchCorrectValue = cxCustDev
    Case "CX Ins.&DMPKTH"
        fetchCorrectValue = cxInsDmp
    Case "PRO OfflineKTH"
        fetchCorrectValue = proOffline
    Case "PRO CSKTH"
        fetchCorrectValue = proCS
    Case "Video OfflineKTH"
        fetchCorrectValue = videoOffline
    Case "Video CSKTH"
        fetchCorrectValue = videoCS
    Case "I&AKTH" 'Is it I Or l? It is I
        fetchCorrectValue = ia
    Case "Bonus & KordKTH"
        fetchCorrectValue = bonusKord
    Case Else
        fetchCorrectValue = -404
    End Select

End Function