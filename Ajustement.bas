Attribute VB_Name = "Ajustement"
Option Explicit

Function AjusteDate(dJour As Date) As Date

'Déclaration des variables
 Dim Date_a As Double
 
 'Algorithme
 
 Select Case ModeAjustement
 
Case "Forward", "Following", 1
    While Not ESTJourTravaille(dJour)
    dJour = dJour + 1
    Wend
    
Case "Modified Forward", "Modified Following", 2
    Date_a = dJour
    While Not ESTJourTravaille(dJour) And (Month(Date_a) = Month(dJour))
    Wend
        If Not ESTJourTravaille(dJour) Or Not (Month(Date_a) = Month(dJour)) Then
            While Not ESTJourTravaille(dJour) Or Not (Month(Date_a) = Month(dJour))
            dJour = dJour - 1
            Wend
        End If
    
Case "Backward", "Preceding", 3
    While Not ESTJourTravaille(dJour)
    dJour = dJour - 1
    Wend
    
Case "Modified Preceding", "Modified Bakward", 4
    Date_a = dJour
    While Not ESTJourTravaille(dJour) And Month(Date_a) = Month(dJour)
    dJour = dJour - 1
    Wend
        If Not ESTJourTravaille(dJour) Or Not (Month(Date_a) = Month(dJour)) Then
            While Not ESTJourTravaille(dJour) Or Not (Month(Date_a) = Month(dJour))
            dJour = dJour + 1
            Wend
        End If
End Select
AjusteDate = dJour

End Function

        
    
    
End Function
