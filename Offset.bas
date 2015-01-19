Attribute VB_Name = "Offset"
Option Explicit

Function DateOffset(iAnnee As Integer, iMois As Integer, iJour As Integer, iMoisOffset As Integer) As Double

'Déclaration des variables

Dim dblDate As Double

'Algorithme

dblDate = DateSerial(iAnnee, iMois + iMoisOffset, iJour)
While Month(dblDate) <> Month(DateSerial(iAnnee, iMois + iMoisOffset, 1))
    dblDate = dblDate - 1
Wend

DateOffset = dblDate


End Function
