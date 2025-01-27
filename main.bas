Attribute VB_Name = "main"
Option Explicit
Option Base 1
Sub main()

'proc�dure mettant en oeuvre le mod�le core-satellites

Dim ws As Worksheet
Dim rg As Range
Dim fs As fSeries
Dim val() As Double 'vecteur de valeurs
Dim rm() As Double 'vecteur des rendements de l'indice de march�
Dim rf() As Double 'vecteur des rendements de l'indice mon�taire
Dim r() As Double 'vecteur des rendements
Dim series() As Variant 'variant r�cup�rant les s�ries
Dim cs As coreSatellites 'variable type core satellites

Dim observ As Integer 'nbre de p�riodes
Dim nSeries As Integer 'nbre de s�ries

Dim i As Integer, j As Integer, k As Integer

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% I. R�cup�ration des donn�es %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'affectation � ws de la feuille mkt
Set ws = ThisWorkbook.Worksheets("indice_mkt")

'calcul du nombre de p�riodes
observ = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1


'=====================================================================================================
'affectation � rg de la plage des valeurs de l'indice de march�
Set rg = ws.Cells(2, 2).Resize(observ, 1)

'r�cup�ration des valeurs de l'indice de march� (avec fnImpVal)
fs = fnImpVal(rg)


'r�cup�ration de l'unique s�rie financi�re dans val
val = fs.series(1)

'transformation des valeurs en rendements (avec fnVal2Rend) et r�cup�ration dans rm
rm = fnVal2Rend(val)



'==================================================================================================
'affectation � ws de la feuille "indice_monetaire", � rg de la plage des valeurs
Set ws = ThisWorkbook.Worksheets("indice_monetaire")
Set rg = ws.Cells(2, 2).Resize(observ, 1)


'r�cup�ration des valeurs de l'indice mon�taire (avec fnImpVal)
fs = fnImpVal(rg)

'r�cup�ration de l'unique s�rie financi�re dans val
val = fs.series(1)


'transformation des valeurs en rendements (avec fnVal2Rend) et r�cup�ration dans rf
rf = fnVal2Rend(val)



'=====================================================================================================

'affectation � ws de la feuille "actions"
Set ws = ThisWorkbook.Worksheets("actions")

'calcul du nombre de s�ries
nSeries = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column - 1

'affectation � rg de la plage des valeurs
Set rg = ws.Cells(2, 2).Resize(observ, nSeries)

'r�cup�ration des valeurs (avec fnImpVal) dans fs
fs = fnImpVal(rg)

'redimensionnement de series (variant destin� � recevoir les s�ries de rendements)
ReDim series(1 To nSeries)

'boucle for next sur les s�ries (compteur i)
For i = 1 To nSeries

    'r�cup�ration dans val de la i-eme s�rie (de fs)
    val = fs.series(i)
    
    'calcul des rendements (avec fnVal2Rend) (et r�cup�ration dans r)
    r = fnVal2Rend(val)
    
    'report dans series de son i-eme �l�ment
    series(i) = r

'fin de la boucle sur les titres
Next i



'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% II. Mise en oeuvre du mod�le core-satellites %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'calcul de l'allocation core satellites
cs = fnCoreSatellites(series, rm, rf, 5)


'affectation � ws de la feuille core satellites
Set ws = ThisWorkbook.Worksheets("core-satellites")

'effacement des donn�es
ws.Cells.ClearContents


'mise en place intitul�s
With ws
    .Cells(1, 1).Resize(1, 5).Value = Array("composante", "part", , , "beta")
    .Cells(2, 1).Resize(3, 1).Value = WorksheetFunction.Transpose(Array("satellites", "core", "cash"))
    .Cells(10, 1).Resize(1, 7).Value = Array("stats", "beta", "alpha", "risque actif", "IR", "AR", "part")
    .Cells(11, 1).Resize(7, 1).Value = WorksheetFunction.Transpose(Array("min", "5%", "25%", "50%", "75%", "95%", "Max"))
    .Cells(20, 1).Resize(1, 7).Value = Array("titre", "beta", "alpha", "risque actif", "IR", "AR", "part")
    .Cells.NumberFormat = "0.000"
End With


'affectation � rg de la plage des r�sultats (� partir de la ligne 21)
Set rg = ws.Cells(21, 1).Resize(nSeries, 7)

'report des noms puis des r�sultats (� partir des �l�ments de cs)
With WorksheetFunction
    rg.Columns(1).Value = .Transpose(fs.noms)
    rg.Columns(2).Value = .Transpose(cs.betas)
    rg.Columns(3).Value = .Transpose(cs.alphas)
    rg.Columns(4).Value = .Transpose(cs.se_alphas)
    rg.Columns(5).Value = .Transpose(cs.IR)
    rg.Columns(6).Value = .Transpose(cs.AR)
    rg.Columns(7).Value = .Transpose(cs.x)
End With
    
    
    
'report des r�sultats sur les parts globales, sur le beta
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
