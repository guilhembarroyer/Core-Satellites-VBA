Attribute VB_Name = "core_satellites"
Option Explicit
Option Base 1
Type coreSatellites
    alphas() As Double 'alphas des titres
    se_alphas() As Double 'risques actifs (erreur-types)
    betas() As Double 'vecteur des betas
    AR() As Double 'appraisal ratios (alphas / risque actif^2)
    IR() As Double 'ratio d'information (alpha/risque actif)
    x() As Double 'vecteur des parts des titres
    xmkt As Double 'part indicielle
    xf As Double 'part en cash
    beta_cible As Double 'beta ciblé
End Type
Function fnCoreSatellites(series() As Variant, rm() As Double, rf() As Double, Optional aversion As Double = 3) As coreSatellites

'fonction calculant l'allocation core satellites en fonction des séries de rendements des titres (series), de la série de l'indice (rm), _
de la série du taux certain (rf)
'ATTENTION à la forme de series : series est un vecteur ligne dont chaque élément est le vecteur des rendements d'un titre!


Dim nSeries As Integer
Dim est As estim_alphas
Dim cs As coreSatellites
Dim r() As Double
Dim alphas() As Double
Dim se_alphas() As Double
Dim betas() As Double
Dim AR() As Double
Dim IR() As Double
Dim x() As Double


Dim i As Integer, j As Integer, k As Integer

'calcul du nombre de séries
nSeries = UBound(series) 'ubound(series)-Lbound(series)+1

'redimensionnement des vecteurs
ReDim alphas(1 To nSeries)
ReDim se_alphas(1 To nSeries)
ReDim betas(1 To nSeries)
ReDim AR(1 To nSeries)
ReDim IR(1 To nSeries)
ReDim x(1 To nSeries)

'boucle sur les séries
For i = 1 To nSeries

    'estimation des alphas et récupération
    r = series(i)
    est = fnEstim_Alphas(r, rm, rf)
    alphas(i) = est.alpha
    se_alphas(i) = est.se_alpha
    betas(i) = est.beta

    'calcul de IR et de AR
    IR(i) = est.alpha / est.se_alpha
    AR(i) = est.alpha / (est.se_alpha) ^ 2
    
    'calcul des parts
    x(i) = AR(i) / aversion

Next i

'récupération des résultats dans cs
cs.alphas = alphas
cs.se_alphas = se_alphas
cs.betas = betas
cs.IR = IR
cs.AR = AR
cs.x = x

'définition de la classe des fonctions Excel comme objet par défaut
With WorksheetFunction
    'calcul du beta ciblé
    cs.beta_cible = (.Average(rm) - .Average(rf)) / .Var(rm) / aversion

    'calcul de l'investissement indiciel
    cs.xmkt = cs.beta_cible - .SumProduct(cs.x, cs.betas)

    'calcul de l'investissement en cash
    cs.xf = 1 - .Sum(cs.x) - cs.xmkt

End With

'résultat
fnCoreSatellites = cs

End Function
