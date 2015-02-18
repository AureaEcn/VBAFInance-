Attribute VB_Name = "Choix_Interpolation"
Option Explicit

Function Interpolation(TableauMaturites, TableauDonnees, DatesCalculees, Optional TypeInterpolation As Boolean = False, Optional TypeDonnees As Boolean = False, Optional DateDeCalcul As Date)


Select Case TYpeInterpoaltion

    Case True
    Interpolation = InterpolationCubique(TableauMaturites, TableauDonnees, DatesCalculees, TypeDonnees, DateDeCalcul)
    
    Case False
    Interpolation = InterpoaltionLineaire(TableauMaturites, TableauDonnees, DatesCalculees, TypeDonnees, DateDeCalcul)
    
End Select

    
End Function
