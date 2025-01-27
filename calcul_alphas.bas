Attribute VB_Name = "calcul_alphas"
Option Explicit
Option Base 0
Type estim_alphas
    nom As String 'nom du titre
    modele As String 'modèle utilisé
    dates() As Integer 'dates des séries
    r() As Double 'rendements du titre
    rm() As Double 'rendements de l'indice de marché utilisé
    rf() As Double 'rendements de l'indice monétaire
    beta As Double 'beta du titre
    se_beta As Double 'erreur-type du beta
    t_beta As Double 't de Student du beta
    alpha As Double 'alpha du titre (au sens du CAPM)
    se_alpha As Double 'erreur-type de l'alpha
    t_alpha As Double 't de Student des rendements en excès du CAPM
    R2 As Double 'coefficient de détermination
    se_eq As Double 'erreur-type du résidu
    F As Double 'stat de Fisher
    p_F As Double 'p-value de F
    r_exces() As Double 'vecteur des rendements en excès du CAPM
End Type
Function fnEstim_Alphas(r() As Double, rm() As Double, rf() As Double) As estim_alphas

'fonction estimant pour une série son alpha et retournant les résultats sous la forme d'une variable type estim_alphas

Dim estim As estim_alphas 'variable type récupérant les résultats
Dim nT As Integer  'nombre de périodes
Dim e() As Double 'vecteur des rendements en excès
Dim mat_linest() As Variant 'matrice renvoyée par la fonction Linest (de régression)
Dim i As Integer

'calcul du nombre de périodes (en supposant r vecteur ligne)
nT = UBound(r) - LBound(r) + 1

'estimation du modèle de marché
With WorksheetFunction
    mat_linest = WorksheetFunction.LinEst(r, rm, True, True)
End With

'récupération des premiers résultats
With estim
    .modele = "standard avec capm comme référence"
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

'récupération des rendements en excès
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

'résultat
fnEstim_Alphas = estim

End Function
