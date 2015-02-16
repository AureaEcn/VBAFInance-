Attribute VB_Name = "DatesFlux"
Option Explicit

Function DatesDesFlux(dDateDeCalcul As Date, dDateMaturite As Date, iFrequence As Integer, dDateDeDepart As Date, Optional TypeCouponBrise = 0, Optional ModeAjustement = 1)

'La fonction renvoie un tableau comprenant les dates des coupons d'un instrument

'Décaration des variables

Dim TableauRetour() As Double
Dim TableauTemporaire() As Double
Dim i As Integer
Dim iTemp As Integer
Dim dblDerniere As Double
Dim bDonneesUtiles As Boolean
Dim iPeriodeCoupon As Integer

'Initialisation

If dDateDeDepart > dDateMaturite Or dDateDeCalcul > dDateMaturite Or dDateDeCalcul = 0 Or dDateMaturite = 0 Then ' Cas impossibles
    MsgBox ("Cas impossibles")
    ReDim TableauRetour(0)
    TableauRetour(0) = 1
    DatesDesFlux = TableauRetour
    Exit Function
End If

'Cas du zero coupon

If iFrequence = 0 Then
    MsgBox ("Zero coupon")
    MsgBox dDateDeDepart
    MsgBox dDateDeCalcul
    If dDateDeDepart > dDateDeCalcul Then 'Forward
    
        MsgBox ("Zero coupon forward")
        ReDim TableauRetour(0 To 1)
        TableauRetour(0) = AjusteDate(dDateDeDepart, ModeAjustement)
        TableauRetour(1) = AjusteDate(dDateMaturite, ModeAjustement)
        DatesDesFlux = TableauRetour
        Exit Function
        
        Else 'Ce n'est pas un forward
        MsgBox ("Zero coupon non forward")
        ReDim TableauRetour(0)
        TableauRetour(0) = AjusteDate(dDateMaturite, ModeAjustement)
        DatesDesFlux = TableauRetour
        Exit Function
        
End If
End If

    
   

'Autres cas
'iPeriodeCoupon est le nombre de mois entre deux coupons
iPeriodeCoupon = 12 / iFrequence
ReDim TableauTemporaire(0)

'======================================================================================================================================
'Calcul de l'echeancier
'======================================================================================================================================

Select Case TypeCouponBrise

    Case "Court Fin", 3, "Long Fin", 4
    'Coupon brisé à la fin, on part de la date de depart jusqu'à la date de maturité
    TableauTemporaire(0) = dDateDeDepart
    dblDerniere = TableauTemporaire(0)
    
    While TableauTemporaire(UBound(TableauTemporaire)) < dDateMaturite
        ReDim Preserve TableauTemporaire(UBound(TableauTemporaire) + 1)
        TableauTemporaire(UBound(TableauTemporaire)) = DateOffset(Year(dblDerniere), Month(dblDerniere), Day(dDateDeDepart), iPeriodeCoupon)
        dblDerniere = TableauTemporaire(UBound(TableauTemporaire))
    Wend

    Case Else
    'On part de la maturite et on recule suivant la frequence jusqu'a la date de calcul
    
    If dDateDeDepart = 0 Then
    dDateDeDepart = dDateDeCalcul
    End If
    TableauTemporaire(0) = dDateMaturite
    dblDerniere = TableauTemporaire(0)
    While TableauTemporaire(UBound(TableauTemporaire)) > dDateDeDepart
        ReDim Preserve TableauTemporaire(UBound(TableauTemporaire) + 1)
        TableauTemporaire(UBound(TableauTemporaire)) = DateOffset(Year(dblDerniere), Month(dblDerniere), Day(dblDerniere), -iPeriodeCoupon)
        dblDerniere = TableauTemporaire(UBound(TableauTemporaire))
    Wend
    
    End Select
    
    Select Case TypeCouponBrise
    
    Case "Court Debut", 1
    'Coupon court au début de la vie de l'instrument
    TableauTemporaire(UBound(TableauTemporaire)) = dDateDeDepart
    
    Case "Long Debut", 2
    'Coupon long à la fin de la vie de l'instrument, il faut retirer un flux
    ReDim Preserve TableauTemporaire(UBound(TableauTemporaire) - 1)
    'et retirer la premiere date
    TableauTemporaire(UBound(TableauTemporaire)) = dDateDeDepart
    
    Case "Court Fin", 3
    'Coupon court en fin de vie de l'instrument, il faut tronquer la dernière date
    TableauTemporaire(UBound(TableauTemporaire)) = dDateMaturite
    
    Case "Long Fin", 4
    'Coupon long en fin de vie de l'instrument, il faut retirer un flux
    ReDim Preserve TableauTemporaire(UBound(TableauTemporaire) - 1)
    TableauTemporaire(UBound(TableauTemporaire)) = dDateMaturite
    
End Select

'======================================================================================================================================
'Donnees
'======================================================================================================================================
    
If dDateDeDepart >= dDateDeCalcul Then 'Forward
    Select Case TypeCouponBrise
        Case "Court Fin", 3, "Long Fin", 4
        For i = 0 To UBound(TableauTemporaire)
            ReDim Preserve TableauRetour(i)
            TableauRetour(i) = TableauTemporaire(i)
            Next
        
        Case Else
        iTemp = UBound(TableauTemporaire)
        For i = iTemp To 0 Step -1
            ReDim Preserve TableauRetour(Abs(i - iTemp))
            TableauRetour(Abs(i - iTemp)) = TableauTemporaire(i)
            Next
            
        End Select
        
Else 'Pas un forward
Select Case TypeCouponBrise
    Case "Court Fin", 3, "Long Fin", 4
    bDonneesUtiles = False
    'On trie le tableau par ordre croissant
    For i = 0 To UBound(TableauTemporaire)
        If TableauTemporaire(i) >= dDateDeCalcul + 1 And Not bDonneesUtiles Then
            bDonneesUtiles = True
            iTemp = i
        End If
        
        If bDonneesUtiles Then
            ReDim Preserve TableauRetour(i - iTemp)
            TableauRetour(i - iTemp) = TableauTemporaire(i - 1)
        End If
        
    Next
    
    ReDim Preserve TableauRetour(i - iTemp)
    TableauRetour(i - iTemp) = TableauTemporaire(i - 1)
    
    Case Else 'Coupon court en début d'instrument ou pas de coupon brisé
    bDonneesUtiles = False
    
    For i = UBound(TableauTemporaire) - 1 To 0 Step -1
        If TableauTemporaire(i) >= dDateDeCalcul + 1 And Not bDonneesUtiles Then
        bDonneesUtiles = True
        iTemp = i
        End If
        
        If bDonneesUtiles Then
        ReDim Preserve TableauRetour(iTemp - i)
        TableauRetour(iTemp - i) = TableauTemporaire(i + 1)
        End If
    Next
    
    ReDim Preserve TableauRetour(iTemp - i)
    TableauRetour(iTemp - i) = TableauTemporaire(i + 1)
    
    End Select
    
End If

For i = 0 To UBound(TableauRetour)
    TableauRetour(i) = AjusteDate((TableauRetour(i)), ModeAjustement)
    Next
    
    DatesDesFlux = TableauRetour
   


    
End Function

