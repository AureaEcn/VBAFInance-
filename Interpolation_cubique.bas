Attribute VB_Name = "Interpolation_cubique"
Option Explicit
Option Base 1


Function InterpolationCubique(TableauMaturites, TableauDonnees, DatesCalculees, Optional EstFactActua As Boolean = False, Optional DateDeCalcul As Date)

'Déclaration des variables
Dim TabMat
Dim TabDates
Dim TabData
Dim i As Integer, j As Integer
Dim Tabretour
Dim MatriceDate(4, 4)
Dim MatDateInv
Dim VecTaux(4, 1)
Dim VecParam

'Conversion des arguments en tableaux
TabMat = CTableau(TableauMaturites)
TabData = CTableau(TableauDonnees)
TabDates = CTableau(DatesCalculees)

ReDim Tabretour(LBound(TabDates) To UBound(TabDates))

'Algorithme


For i = LBound(TabDates) To UBound(TabDates) Step 1 'Boucles sur les dates
 
   
    If TabDates(i) <= TabMat(2) Then 'Borne inférieure
    MsgBox ("Borne inferieure")
    Tabretour(i) = InterpolationLineaire(TableauMaturites, TableauDonnees, TabDates(i), EstFactActua, DateDeCalcul)(1)
  
   
    ElseIf TabDates(i) > TabMat(UBound(TabMat) - 1) Then ' Borne superieure
    MsgBox ("Borne superieure")
    Tabretour(i) = InterpolationLineaire(TableauMaturites, TableauDonnees, TabDates(i), EstFactActua, DateDeCalcul)(1)
   
   
   
   Else 'Les autres cas
   
    For j = 1 To UBound(TabMat) - 2 'premiere date supérieure a la date de calcul
        If TabMat(j) > TabDates(i) Then
          
            Exit For
            
        End If
        
    Next
    
    MsgBox ("matrice")
    'On renseigne la matrice de date
    MatriceDate(1, 1) = CDbl(TabMat(j - 2) ^ 3)
    MatriceDate(1, 2) = TabMat(j - 2) ^ 2
    MatriceDate(1, 3) = CDbl(TabMat(j - 2))
    MatriceDate(1, 4) = 1
    MatriceDate(2, 1) = CDbl(TabMat(j - 1) ^ 3)
    MatriceDate(2, 2) = TabMat(j - 1) ^ 2
    MatriceDate(2, 3) = CDbl(TabMat(j - 1))
    MatriceDate(2, 4) = 1
    MatriceDate(3, 1) = CDbl(TabMat(j) ^ 3)
    MatriceDate(3, 2) = TabMat(j) ^ 2
    MatriceDate(3, 3) = CDbl(TabMat(j))
    MatriceDate(3, 4) = 1
    MatriceDate(4, 1) = CDbl(TabMat(j + 1) ^ 3)
    MatriceDate(4, 2) = TabMat(j + 1) ^ 2
    MatriceDate(4, 3) = CDbl(TabMat(j + 1))
    MatriceDate(4, 4) = 1
    
    'On renseigne le vecteur de données
    VecTaux(1, 1) = TabData(j - 2)
    VecTaux(2, 1) = TabData(j - 1)
    VecTaux(3, 1) = TabData(j)
    VecTaux(4, 1) = TabData(j + 1)
    
    
    'Inversion de la matrice des dates
    MatDateInv = Application.WorksheetFunction.MInverse(MatriceDate)
    
    'Résolution du systeme
    VecParam = Application.WorksheetFunction.MMult(MatDateInv, VecTaux)
   
    
    'Calcul de la valeur interpolée
    Tabretour(i) = VecParam(1, 1) * TabDates(i) ^ 3 + VecParam(2, 1) * TabDates(i) ^ 2 + VecParam(3, 1) * TabDates(i) + VecParam(4, 1)
    
    
    End If
    
   Next
   
   InterpolationCubique = Tabretour
   

End Function

