Attribute VB_Name = "DatesDesFlux"
Option Explicit

Function DatesDesFLux(dDateDeCalcul As Date, dDateDeMaturité As Date, iFrequence As Integer, Optional TypeCouponBrise = 0, Optional dDateDeDepart = 0, Optional ModeAjustement = 0)

'La fonction renvoie un tableau comprenant les dates des coupons d'un instrument

'Décaration des variables

Dim TableauRetour() As Double
Dim TableauTemporaire() As Double
Dim i As Integer
Dim dblDerniere As Double
Dim bDonneesUtiles As Boolean
Dim iPeriodeCoupon As Integer

'Initialisation

If DateDeDepart > DateMaturite Or DateDeCalcul > DateMaturite Or DateDeCalcul = 0 Or DateDeMaturite = 0 Then ' Cas impossibles
    ReDim TableauRetour(0)
    TableauRetour(0) = 0
    DatesDesFLux = TableauRetour
    Exit Function
End If

'Cas du zero coupon

If iFrequence = 0 Then
    If dDateDeDepart > dDateDeCalcul Then 'Forward
        ReDim TableauRetour(0 To 1)
        TableauRetour(0) = AjusteDate((dDateDeDepart), ModeAjustement)
        TableauRetour(1) = AjusteDate((dDateMaturite), ModeAjustement)
        Else 'Ce n'est pas un forward
        ReDim TableauRetour(1)
        TableauRetour(1) = AjusteDate((dDateMaturite), ModeAjustement)
    End If
    
DatesDesFLux = TableauRetour

'Autres cas
'iPeriodeCoupon est le nombre de mois entre deux coupons
iPeriodeCoupon = 12 / iFrequence
RedimTableauTemporaire (0)
'======================================================================================================================================
'Calcul de l'echeancier
'======================================================================================================================================

Select Case TypeCouponBrise

    Case "Court Fin", 3, "Long Fin", 4
    'Coupon brisé à la fin, on part de la date de depart jusqu'à la date de maturité
    TableauTemporaire(0) = dDateDeDepart
    dblDerniere = TableauTemporaire(0)
    
    While TableauTemporaire(UBound(TableauTemproaire)) < DateDeMaturite
        ReDim Preserve TableauTemporaire(UBound(TableauTemporaire) + 1)
        TableauTemporaire(UBound(TableauTemporaire)) = DateOffset(Year(dblDerniere, Month(dblDerniere), Day(DateDeDepart), iPeriodCoupon))
        dblDerniere = TableauTemporaire(UBound(TableauTemporaire))
    Wend
    End Select
    
Select Case TypeCouponBrise

    Case Else
    'On part de la maturite et on recule suivant la frequence jusqu'a la date de calcul
    
    If DateDeDepart = 0 Then
    DateDeDepart = DateDeCalcul
    End If
    TableauTemporaire(0) = DateDeMaturite
    dblDerniere = TableauTemporaire(0)
    While TableauTemporaire(UBound(TableauTemporaire)) > DateDeDepart
        ReDim Preserve TableauTemporaire(UBound(TableauTemporaire) + 1)
        TableauTemporaire(UBound(TableauTemporaire)) = DateOffset(Year(dblDerniere), Month(dblDerniere), Day(dblDerniere), -iPeriodeCoupon)
        dblDerniere = TableauTemporaire(UBound(TableauTemporaire))
    Wend
    
    
    Case "Court Debut", 1
    'Coupon court au début de la vie de l'instrument
    TableauTemporaire(UBound(TableauTemporaire)) = DateDeDepart
    
    Case "Long Debut", 2
    'Coupon long à la fin de la vie de l'instrument, il faut retirer un flux
    ReDim Preserve TableauTemporaire(UBound(TableauTemporaire) - 1)
    'et retirer la premiere date
    TableauTemporaire(UBound(TableauTemporaire)) = DateDeDepart
    
    Case "Court Fin", 3
    'Coupon court en fin de vie de l'instrument, il faut tronquer la dernière date
    TableauTemporaire(UBound(TableauTemporaire)) = DateDeMaturite
    
    Case "Long Fin", 4
    'Coupon long en fin de vie de l'instrument, il faut retirer un flux
    ReDim Preserve TableauTemporaire(UBound(TableauTemproaire) - 1)
    TableauTemporaire(UBound(TableauTemporaire)) = DateDeMaturite
    
End Select

'======================================================================================================================================
'Donnees
'======================================================================================================================================
    
If DateDeDepart >= DateDeCalcul Then 'Forward
    Select Case TypeCouponBrise
        Case "Court Fin", 3, "Long Fin", 4
        For i = 0 To UBound(TableauTemporaire)
            ReDim Preserve TableauRetour(i)
            TableauRetour(i) = TableauTemporaire(i)
            Next
        
        


End Function



