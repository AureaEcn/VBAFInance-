Attribute VB_Name = "Bissextile"
Option Explicit

Function ESTAnneeBissextile(iAnnee As Integer) As Boolean

'Declaration des variables

Dim bReponse As Boolean
bReponse = False

'Algorithme de la fonction disponible sur wikipedia: http://fr.wikipedia.org/wiki/Bissextile

If bReponse Mod 4 = 0 Then
bReponse = True
    
    If bReponse Mod 100 = 0 Then
    bReponse = False
    
        If bReponse Mod 400 = 0 Then
        bReponse = True
        
        End If
    End If
End If

ESTAnneeBissextile = bReponse
        
End Function

