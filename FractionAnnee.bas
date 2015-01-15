Attribute VB_Name = "FractionAnnee"
Option Explicit

Function FractionAnnee(dDebutPeriode As Date, dFinPeriode As Date, Convention) As Double

'Déclaration des variables

Dim Nbjd As Double
Dim Nbjf As Double
Dim Date_a As Double
Dim Date_b As Double

'Algorithme
Select Case Convention

    Case "Exact/Exact", 1
    If ESTAnneeBissextile(Year(dDebutPeriode)) Then
    Nbjd = 366
    Else: Nbjd = 365
    If ESTAnneeBissextile(Year(dFinPeriode)) Then
    Nbjf = 366
    Else: Nbjf = 365
    FractionAnnee = Year(dFinPeriode) - Year(dDebutPeriode)
    FractionAnnee = FractionAnnee - (dDebutPeriode - DateSerial(Year(dDebutPeriode), 1, 1)) / Nbjd
    FractionAnnee = FractionAnnee + (dFinPeriode - DateSerial(Year(dFinPeriode), 1, 1)) / Nbjf
    
    Case "Exact/365", 2
    FractionAnnee = (dFinPeriode - dDebutPeriode) / 365
    
    Case "Exact/360", 3
    FractionAnnee = (dFinPeriode - dDebutPeriode) / 360
    
    Case Else
    'Convention 30/360
    If Day(dDebutPeriode) = 31 Then Date_a = dDebutPeriode - 1
    Else: Date_a = dDebutPeriode
        If Day(dFinPeriode) = 31 And Day(dDebutPeriode) >= 30 Then
        Date_b = dFinPeriode - 1
        
        ElseIf Day(dFinPeriode) = 31 And Day(dDebutPeriode) < 30 Then
        Date_b = dFinPeriode + 1
        
        Else
        Date_b = dFinPeriode
        
        End If
        
    FractionAnnee = (Year(Date_b) - Year(Date_a)) * 360 + (Month(Date_b) - Month(Date_a)) * 30 + (Day(Date_b) - Day(Date_a))
    FractionAnnee = FractionAnnee / 360

    End Select
    
    End Function

 
