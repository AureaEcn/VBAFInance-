Attribute VB_Name = "Prochain_FLux"
Option Explicit

Function ProchainFLux(dDateDeCalcul As Date, dDateMaturite As Date, Frequence As Integer, Optional ModeAjustement = 0, Optional TypeCouponBrise = 0, Optional DateDeDepart = 0) As Double

Dim dblDateFlux As Double

On Error Resume Next

dblDateFlux = DatesDesFlux(dDateDeCalcul, dDateMaturite, Frequence, ModeAjustement, TypeCouponBrise, DateDeDepart)(1)

On Error GoTo 0

ProchainFLux = dblDateFlux

End Function
