Attribute VB_Name = "main"
Option Explicit
Option Base 1
Sub main()

'procédure mettant en oeuvre le modèle core-satellites

Dim ws As Worksheet
Dim rg As Range
Dim fs As fSeries
Dim val() As Double 'vecteur de valeurs
Dim rm() As Double 'vecteur des rendements de l'indice de marché
Dim rf() As Double 'vecteur des rendements de l'indice monétaire
Dim r() As Double 'vecteur des rendements
Dim series() As Variant 'variant récupérant les séries
Dim cs As coreSatellites 'variable type core satellites

Dim observ As Integer 'nbre de périodes
Dim nSeries As Integer 'nbre de séries

Dim i As Integer, j As Integer, k As Integer

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% I. Récupération des données %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'affectation à ws de la feuille mkt
Set ws = ThisWorkbook.Worksheets("indice_mkt")

'calcul du nombre de périodes
observ = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1


'=====================================================================================================
'affectation à rg de la plage des valeurs de l'indice de marché
Set rg = ws.Cells(2, 2).Resize(observ, 1)

'récupération des valeurs de l'indice de marché (avec fnImpVal)
fs = fnImpVal(rg)


'récupération de l'unique série financière dans val
val = fs.series(1)

'transformation des valeurs en rendements (avec fnVal2Rend) et récupération dans rm
rm = fnVal2Rend(val)



'==================================================================================================
'affectation à ws de la feuille "indice_monetaire", à rg de la plage des valeurs
Set ws = ThisWorkbook.Worksheets("indice_monetaire")
Set rg = ws.Cells(2, 2).Resize(observ, 1)


'récupération des valeurs de l'indice monétaire (avec fnImpVal)
fs = fnImpVal(rg)

'récupération de l'unique série financière dans val
val = fs.series(1)


'transformation des valeurs en rendements (avec fnVal2Rend) et récupération dans rf
rf = fnVal2Rend(val)



'=====================================================================================================

'affectation à ws de la feuille "actions"
Set ws = ThisWorkbook.Worksheets("actions")

'calcul du nombre de séries
nSeries = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column - 1

'affectation à rg de la plage des valeurs
Set rg = ws.Cells(2, 2).Resize(observ, nSeries)

'récupération des valeurs (avec fnImpVal) dans fs
fs = fnImpVal(rg)

'redimensionnement de series (variant destiné à recevoir les séries de rendements)
ReDim series(1 To nSeries)

'boucle for next sur les séries (compteur i)
For i = 1 To nSeries

    'récupération dans val de la i-eme série (de fs)
    val = fs.series(i)
    
    'calcul des rendements (avec fnVal2Rend) (et récupération dans r)
    r = fnVal2Rend(val)
    
    'report dans series de son i-eme élément
    series(i) = r

'fin de la boucle sur les titres
Next i



'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% II. Mise en oeuvre du modèle core-satellites %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'calcul de l'allocation core satellites
cs = fnCoreSatellites(series, rm, rf, 5)


'affectation à ws de la feuille core satellites
Set ws = ThisWorkbook.Worksheets("core-satellites")

'effacement des données
ws.Cells.ClearContents


'mise en place intitulés
With ws
    .Cells(1, 1).Resize(1, 5).Value = Array("composante", "part", , , "beta")
    .Cells(2, 1).Resize(3, 1).Value = WorksheetFunction.Transpose(Array("satellites", "core", "cash"))
    .Cells(10, 1).Resize(1, 7).Value = Array("stats", "beta", "alpha", "risque actif", "IR", "AR", "part")
    .Cells(11, 1).Resize(7, 1).Value = WorksheetFunction.Transpose(Array("min", "5%", "25%", "50%", "75%", "95%", "Max"))
    .Cells(20, 1).Resize(1, 7).Value = Array("titre", "beta", "alpha", "risque actif", "IR", "AR", "part")
    .Cells.NumberFormat = "0.000"
End With


'affectation à rg de la plage des résultats (à partir de la ligne 21)
Set rg = ws.Cells(21, 1).Resize(nSeries, 7)

'report des noms puis des résultats (à partir des éléments de cs)
With WorksheetFunction
    rg.Columns(1).Value = .Transpose(fs.noms)
    rg.Columns(2).Value = .Transpose(cs.betas)
    rg.Columns(3).Value = .Transpose(cs.alphas)
    rg.Columns(4).Value = .Transpose(cs.se_alphas)
    rg.Columns(5).Value = .Transpose(cs.IR)
    rg.Columns(6).Value = .Transpose(cs.AR)
    rg.Columns(7).Value = .Transpose(cs.x)
End With
    
    
    
'report des résultats sur les parts globales, sur le beta
ws.Cells(2, 2).Value = WorksheetFunction.Sum(cs.x)
ws.Cells(3, 2).Value = cs.xmkt
ws.Cells(4, 2).Value = cs.xf
ws.Cells(2, 5).Value = cs.beta_cible

'calcul des stats (avec les fonctions Excel Min, Percentile, Max)
With WorksheetFunction
    For j = 2 To 7
        ws.Cells(11, j).Resize(7, 1).Value = _
        .Transpose(Array(.Min(rg.Columns(j)), .Percentile(rg.Columns(j), 0.05), _
        .Percentile(rg.Columns(j), 0.25), .Percentile(rg.Columns(j), 0.5), _
        .Percentile(rg.Columns(j), 0.75), .Percentile(rg.Columns(j), 0.95), _
        .Max(rg.Columns(j), 0.05)))
    Next j
End With

End Sub
