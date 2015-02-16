Attribute VB_Name = "Dernier_Flux"
Option Explicit
Function DernierFlux(dDateDeCalcul As Date, dDateMaturite As Date, Frequence As Integer, Optional ModeAjustement = 0, Optional TypeCouponBrise = 0, Optional DateDeDepart = 0) As Double

'Declaration des variables

Dim dblDateFlux As Double

'Algorithme

If DateDeDepart > dDateDeCalcul Then 'Forward
    dblDateFlux = 0
Else
    dblDateFlux = DatesDesFlux(dDateDeCalcul, dDateMaturite, Frequence, ModeAjustement, TypeCouponBrise, DateDeDepart)(0)
End If

DernierFlux = dblDateFlux


End Function

