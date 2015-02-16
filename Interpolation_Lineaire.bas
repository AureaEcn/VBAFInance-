Attribute VB_Name = "Interpolation_Lineaire"
Option Explicit

Function InterpolationLineaire(TableauMaturites, TableauDonnees, DatesCalculees, Optional EstFactActua As Boolean = False, Optional DateCalcul As Date)
'Déclaration des variables

Dim TabMat
Dim TabData
Dim TabDates
Dim i As Integer, j As Integer
Dim TabRetour
Dim DateCalc As Date
Dim DateInf As Date
Dim DateSup As Date
Dim dblTxInf As Double
Dim dblTxSup As Double
Dim dblTaux As Double

'Conversion des arguments en tableaux

TabMat = CTableau(TableauMaturites)
TabData = CTableau(TableauDonnees)
TabDates = CTableau(DatesCalculees)

'On redimensionne le retour en fonction des dates en entrée

ReDim TabRetour(LBound(TabDates) To UBound(TabDates))

'Algorithme

Select Case EstFactActua
    
    Case True 'Les données sont des facteurs d'actualisation
    MsgBox ("Il s'agit de facteurs d'actualisation")
    For i = LBound(TabDates) To i = UBound(TabDates) Step 1 'Boucle sur les dates
        
        If TabDates(i) <= TabMat(1) Then    'Borne inférieure
        dblTaux = ChangeTaux(DateCalcul, (TabMat(1)), (TabData(1)), 0, 1, 2, 2, 1, 2)
        TabRetour(i) = ChangeTaux(DateCalcul, (TabDates(i)), dblTaux, 0, 1, 2, 2, 1, 2)
        
        ElseIf TabDates(i) >= TabMat(UBound(TabMat)) Then 'Borne superieure
        dblTaux = ChangeTaux(DateCalcul, (TabMat(UBound(TabMat))), (TabData(UBound(TabData))), 2, 1, 2, 0, 1, 2)
        TabRetour(i) = ChangeTaux(DateCalcul, (TabDates(i)), dblTaux, 0, 1, 2, 2, 1, 2)
        
        Else 'Tous les autres cas
        
            For j = 1 To UBound(TabMat)
                If TabMat(j) >= TabDates(i) Then
                Exit For
        End If
        
    Next
    
'Donnees pour l'interpolation
dblTxInf = TabData(j - 1)
dblTxSup = TabData(j)
DateInf = TabMat(j - 1)
DateSup = TabMat(j)
DateCalc = TabDates(i)

'Calcul de l'interpolation
            TabRetour(i) = ((DateCalcul - DateInf) * dblTxSup + (DateSup - DateCalcul) * dblTxInf) / (DateSup - DateInf)

        End If

    Next
    
    Case Else 'Les données sont des taux
    
MsgBox ("Il s'agit de taux")

    For i = LBound(TabDates) To UBound(TabDates) Step 1 'Boucle sur les dates
    
        If TabDates(i) <= TabMat(1) Then 'Borne inférieure
        TabRetour(i) = TabData(1)
        
        ElseIf TabDates(i) >= TabMat(UBound(TabMat)) Then 'Borne supérieure
        TabRetour(i) = TabData(UBound(TabMat))
        
        Else 'Tous les autres cas
            For j = 1 To UBound(TabMat)
                If TabMat(j) > TabDates(i) Then
                    Exit For
                End If
            Next
            
'Donnees pour l'interpoaltion
dblTxInf = TabData(j - 1)
dblTxSup = TabData(j)
DateInf = TabMat(j - 1)
DateSup = TabMat(j)
DateCalc = TabDates(i)


'Calcul de l'interpolation
        TabRetour(i) = ((DateCalc - DateInf) * dblTxSup + (DateSup - DateCalc) * dblTxInf) / (DateSup - DateInf)
       
        
    End If
    
Next

End Select

InterpolationLineaire = TabRetour

End Function

