Attribute VB_Name = "ChangementTaux"
Option Explicit
Option Base 1

Function ChangeTaux(dDateDeCalcul As Date, dDateMaturite As Date, dblDonnee As Double, iTypeDonnee As Integer, iFrequence As Integer, iBase As Integer, iFrequenceCible As Integer, iTypeDonneeCible As Integer, iBaseCible As Integer) As Double

'Cette fonction permet de la conversion d'un taux à un autre en modifiant la base / type

'Déclaration des variables

Dim dblFA As Double
Dim dblValRet As Double
Dim dblF1 As Double
Dim dblF2 As Double

dblF1 = FractionAnnee(dDateDeCalcul, dDateMaturite, iBase)
dblF2 = FractionAnnee(dDateDeCalcul, dDateMaturite, iBaseCible)

'Angorithme
'Conversion de la donnée initiale en facteur d'actualisation

Select Case iTypeDonnee
    
    Case 0 'initialement taux simple
    dblFA = (1 + dblDonnee * dblF1)
    dblFA = 1 / dblFA
    
    Case 1 'initialement taux composé
    dblFA = (1 + dblDonnee / iFrequence) ^ (iFrequence * dblF1)
    dblFA = 1 / dblFA
    
    Case 2 'initialement facteur d'actualisation
    dblFA = dblDonnee
    
    Case 3 'initialement taux continu
    dblFA = Exp(-dblDonnee * dblF1)
    
    Case Else 'type inconnu
    dblFA = 0
    
End Select

Select Case iTypeDonneeCible

    Case 0 'on veut passer en taux simple
    dblValRet = ((1 / dblFA) - 1) * (1 / dblF2)
    
    Case 1 'on veut passer en taux composé
    dblValRet = (((1 / dblFA) ^ (1 / (iFrequenceCible * dblF2))) - 1) * iFrequenceCible
    
    Case 2 'on veut un taux d'actualisation
    dblValRet = dblFA
    
    Case 3 'on veut un taux continu
    dblValRet = (Log(1 / dblFA)) / dblF2
    
End Select

ChangeTaux = dblValRet

End Function
    


