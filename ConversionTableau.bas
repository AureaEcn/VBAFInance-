Attribute VB_Name = "ConversionTableau"
Option Explicit

Function CTableau(vDonnees)

'Renvoie un tableau de doubles sans savoir a partir de différents types de données

'Declaration des variables

Dim TableauRetour(), i As Integer

'Algorithme

Select Case TypeName(vDonnees)

    Case "Range"
    'C'est une plage de cellules
    ReDim TableauRetour(1 To vDonnees.Cells.Count)
    For i = 0 To vDonnees.Cells.Count Step 1
        TableauRetour(i) = vDonnees.Cells(i)
    Next
    
    Case "Integer()", "Double()", "Variant()", "Date()"
    'C'est un tableau
    ReDim TableauRetour(1 To UBound(vDonnees))
    For i = 0 To UBound(vDonnees) Step 1
        TableauRetour(i) = vDonnees(i)
    Next
    
    Case Else
    'C'est une date/nombre
    ReDim TableauRetour(1 To 1)
    TableauRetour(1) = vDonnees
    
End Select

CTableau = TableauRetour

    
End Function
