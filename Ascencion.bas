Attribute VB_Name = "Ascencion"
Option Explicit

Function ESTJourAscencion(dJour As Date) As Date

'Déclaration des variables

Dim iAnnee As Integer
Dim dDimanche As Date
Dim dJeudi As Date

'Algorithme

iAnnee = Year(dJour)
dDimanche = DimancheDePaques(iAnnee)
dJeudi = dDimanche + 39

If Weekday(dJeudi) = 5 Then
ESTJourAscencion = dJeudi
End If


End Function
