Attribute VB_Name = "Travaille"
Option Explicit

Function ESTJourTravaille(dJour As Date) As Boolean

'Declaration des variables

Static TableauDimanchePaques() As Double 'Liste des dimanches de pâques
Dim iTest As Integer
Dim i As Integer

'Algorithme

On Error Resume Next
iTest = UBound(TableauDimanchePaques)
On Error GoTo 0

If iTest = 0 Then
ReDim TableauDimanchePaques(1970 To 2100)
    For i = 1970 To 2100 Step 1
    TableauDimanchePaques(i) = DimancheDePaques(i)
    Next i
End If

'Test des jours "communs"

If Weekday(dJour) = 7 Then  'Samedi
ESTJourTravaille = False

ElseIf Weekday(dJour) = 1 Then  'Dimanche
ESTJourTravaille = False

ElseIf Day(dJour) = 25 And Month(dJour) = 12 Then '25 decembre
ESTJourTravaille = False

ElseIf Day(dJour) = 1 And Month(dJour) = 1 Then '1 janvier
ESTJourTravaille = False

ElseIf Day(dJour) = 1 And Month(dJour) = 5 Then '1 mai
ESTJourTravaille = False

ElseIf Day(dJour) = 14 And Month(dJour) = 7 Then    '14 juillet
ESTJourTravaille = False

ElseIf Day(dJour) = 15 And Month(dJour) = 8 Then    '15 aout
ESTJourTravaille = False

ElseIf Day(dJour) = 1 And Month(dJour) = 11 Then    '1 novembre
ESTJourTravaille = False

ElseIf Day(dJour) = 11 And Month(dJour) = 11 Then   '11 novembre
ESTJourTravaille = False

ElseIf Day(dJour) = (DimancheDePaques(Year(dJour)) + 1) Then
ESTJourTravaille = False

ElseIf Day(dJour) = (ESTJourAscencion(dJour)) Then
ESTJourTravaille = False

ElseIf Day(dJour) = (ESTLundiDePentecote(dJour)) Then
ESTJourTravaille = False

Else
ESTJourTravaille = True

End If


End Function
