Attribute VB_Name = "Pentecote"
Option Explicit

Function ESTLundiDePentecote(dJour As Date) As Date

'Déclaration des variables

Dim iAnnee As Integer
Dim dDimanche As Date
Dim dLundi As Date

'Algorithme

iAnnee = Year(dJour)
dDimanche DimancheDePaques(iAnnee)
dLundi = dDimanche + 50

If Weekday(dLundi) = 2 Then
ESTLundiDePentecote = dLundi
End If


End Function
