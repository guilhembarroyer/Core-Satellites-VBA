Attribute VB_Name = "calc_rendements"
Option Explicit
Option Base 1
Function fnVal2Rend(val() As Double) As Double()

'fonction calculant � partir du vecteur ligne les rendements g�om�triques

Dim r() As Double 'vecteur des rendements
Dim nT As Integer 'nombre de p�riodes
Dim i As Integer 'compteur de la boucle

'calcul du nombre de valeurs (nombre de p�riodes)
nT = UBound(val) - LBound(val) + 1

'redimensionnement du vecteur r
ReDim r(1 To nT - 1)

'boucle sur les rendements
For i = 1 To nT - 1
    r(i) = val(i + 1) / val(i) - 1
Next i

'r�sultat
fnVal2Rend = r

End Function
