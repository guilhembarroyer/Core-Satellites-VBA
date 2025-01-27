Attribute VB_Name = "calcul_alphas"
Option Explicit
Option Base 0
Type estim_alphas
    nom As String 'nom du titre
    modele As String 'mod�le utilis�
    dates() As Integer 'dates des s�ries
    r() As Double 'rendements du titre
    rm() As Double 'rendements de l'indice de march� utilis�
    rf() As Double 'rendements de l'indice mon�taire
    beta As Double 'beta du titre
    se_beta As Double 'erreur-type du beta
    t_beta As Double 't de Student du beta
    alpha As Double 'alpha du titre (au sens du CAPM)
    se_alpha As Double 'erreur-type de l'alpha
    t_alpha As Double 't de Student des rendements en exc�s du CAPM
    R2 As Double 'coefficient de d�termination
    se_eq As Double 'erreur-type du r�sidu
    F As Double 'stat de Fisher
    p_F As Double 'p-value de F
    r_exces() As Double 'vecteur des rendements en exc�s du CAPM
End Type
Function fnEstim_Alphas(r() As Double, rm() As Double, rf() As Double) As estim_alphas

'fonction estimant pour une s�rie son alpha et retournant les r�sultats sous la forme d'une variable type estim_alphas

Dim estim As estim_alphas 'variable type r�cup�rant les r�sultats
Dim nT As Integer  'nombre de p�riodes
Dim e() As Double 'vecteur des rendements en exc�s
Dim mat_linest() As Variant 'matrice renvoy�e par la fonction Linest (de r�gression)
Dim i As Integer

'calcul du nombre de p�riodes (en supposant r vecteur ligne)
nT = UBound(r) - LBound(r) + 1

'estimation du mod�le de march�
With WorksheetFunction
    mat_linest = WorksheetFunction.LinEst(r, rm, True, True)
End With

'r�cup�ration des premiers r�sultats
With estim
    .modele = "standard avec capm comme r�f�rence"
    .r = r
    .rm = rm
    .rf = rf
    .beta = mat_linest(1, 1)
    .se_beta = mat_linest(2, 1)
    .t_beta = mat_linest(1, 1) / mat_linest(2, 1)
    .R2 = mat_linest(3, 1)
    .se_eq = mat_linest(3, 2)
    .F = mat_linest(4, 1)
    .p_F = mat_linest(4, 2)
End With

'r�cup�ration des rendements en exc�s
ReDim e(1 To nT)
For i = 1 To nT
    e(i) = r(i) - rf(i) - mat_linest(1, 1) * (rm(i) - rf(i))
Next i

'calcul des alphas
With estim
    .r_exces = e
    .alpha = WorksheetFunction.Average(e)
    .se_alpha = WorksheetFunction.StDev(e)
    .t_alpha = .alpha / .se_alpha
End With

'r�sultat
fnEstim_Alphas = estim

End Function
