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

'fonction récupérant les valeurs de rg dans une variable type fSeries.
'rg est supposée être la place des valeurs, sa ligne 0 est celle des noms des séries, sa colonne 0 est celle des dates.

Dim fs As fSeries 'variable type récupérant les données
Dim x() As Variant
Dim noms() As String 'vecteur (ligne) des noms des séries
Dim series() As Variant 'variant récupérant les séries
Dim val() As Double 'vecteur des valeurs (d'une série)
Dim dates() As Long 'vecteur des dates (en entier)
Dim nSeries As Integer 'nbre de séries
Dim nT As Integer 'nbre de périodes

Dim i As Integer, j As Integer, k As Integer 'compteurs

'récupération des données dans le variant x
x = rg.Value
'calcul des dimensions
nSeries = rg.Columns.Count
nT = rg.Rows.Count
'redimensionnement du variant series
ReDim series(1 To nSeries)

'boucle sur les séries
For j = 1 To nSeries
    'redimensionnement du double val récupérant la j-eme série
    ReDim val(1 To nT)
    'boucle sur les périodes
    For i = 1 To nT
        val(i) = x(i, j)
    Next i
    'récupération comme j-eme élément de series
    series(j) = val
Next j

'récupération des dates
With rg.Columns(0)
    .NumberFormat = "General"
    x = .Value
    .NumberFormat = "dd/mm/yy"
End With

'récupération dans dates
ReDim dates(1 To nT)
For i = 1 To nT
    dates(i) = x(i, 1)
Next i

'récupération des noms
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

'récupération dans la variable type
With fs
    .dates = dates
    .noms = noms
    .series = series
    If descrip = True Then
        .descrip = InputBox("Entrez la description de la série.", "description des séries")
    End If
End With

'résultat
fnImpVal = fs

End Function
