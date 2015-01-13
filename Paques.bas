Attribute VB_Name = "Paques"
Option Explicit

Function DimancheDePaques(iAnnee As Integer) As Date

' Déclaration des variables

Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer
Dim g As Integer, h As Integer, i As Integer, k As Integer, l As Integer, m As Integer
Dim iMois As Integer, iJour As Integer

'Algorithme de Meeus/Jones/butcher disponible sur wikipedia: http://en.wikipedia.org/wiki/Computus

a = iAnnee Mod 19
b = Int(iAnnee / 100)
c = iAnnee Mod 100
d = Int(b / 4)
e = b Mod 4
f = Int((b + 8) / 25)
g = Int((b - f + 1) / 3)
h = (19 * a + b - d - g + 15) Mod 30
i = Int(c / 4)
k = c Mod 4
l = (32 + 2 * e + 2 * i - h - k) Mod 7
m = Int((a + 11 * h + 22 * l) / 451)

iMois = Int((h + l - 7 * m + 114) / 31)
iJour = Int(((h + l - 7 * m + 114) Mod 31) + 1)

DimancheDePaques = DateSerial(iAnnee, iMois, iJour)

End Function



