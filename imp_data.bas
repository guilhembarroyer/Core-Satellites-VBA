Attribute VB_Name = "imp_data"
Option Explicit
Option Base 0
Type fSeries
    dates() As Long
    series() As Variant
    noms() As String
    descrip As String
End Type
Function fnImpVal(rg As Range, Optional descrip As Boolean = False) As fSeries

'fonction r�cup�rant les valeurs de rg dans une variable type fSeries.
'rg est suppos�e �tre la place des valeurs, sa ligne 0 est celle des noms des s�ries, sa colonne 0 est celle des dates.

Dim fs As fSeries 'variable type r�cup�rant les donn�es
Dim x() As Variant
Dim noms() As String 'vecteur (ligne) des noms des s�ries
Dim series() As Variant 'variant r�cup�rant les s�ries
Dim val() As Double 'vecteur des valeurs (d'une s�rie)
Dim dates() As Long 'vecteur des dates (en entier)
Dim nSeries As Integer 'nbre de s�ries
Dim nT As Integer 'nbre de p�riodes

Dim i As Integer, j As Integer, k As Integer 'compteurs

'r�cup�ration des donn�es dans le variant x
x = rg.Value
'calcul des dimensions
nSeries = rg.Columns.Count
nT = rg.Rows.Count
'redimensionnement du variant series
ReDim series(1 To nSeries)

'boucle sur les s�ries
For j = 1 To nSeries
    'redimensionnement du double val r�cup�rant la j-eme s�rie
    ReDim val(1 To nT)
    'boucle sur les p�riodes
    For i = 1 To nT
        val(i) = x(i, j)
    Next i
    'r�cup�ration comme j-eme �l�ment de series
    series(j) = val
Next j

'r�cup�ration des dates
With rg.Columns(0)
    .NumberFormat = "General"
    x = .Value
    .NumberFormat = "dd/mm/yy"
End With

'r�cup�ration dans dates
ReDim dates(1 To nT)
For i = 1 To nT
    dates(i) = x(i, 1)
Next i

'r�cup�ration des noms
If rg.Columns.Count > 1 Then
    x = rg.Rows(0).Value
    ReDim noms(1 To nSeries)
    For j = 1 To nSeries
        noms(j) = x(1, j)
    Next j
Else
    ReDim noms(1 To 1)
    noms(1) = rg.Cells(0, 1).Value
End If

'r�cup�ration dans la variable type
With fs
    .dates = dates
    .noms = noms
    .series = series
    If descrip = True Then
        .descrip = InputBox("Entrez la description de la s�rie.", "description des s�ries")
    End If
End With

'r�sultat
fnImpVal = fs

End Function
