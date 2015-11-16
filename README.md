sapflux-Excel
=============

Functions to generate predictions of whole tree sap flow based on sap flux density observations.

Usage
-----

qtot(v, treeRadius, woodType, probeDepth, probeStart, sapRadius)

#Arguments:	

| Variable | Definition |
| -------- | ---------- |
| v | Sap flux density observations (g / m^2 / time) |
| treeRadius | tree radius minus bark (in m) |
| woodType | xylem type for radial profile selection (choose from Tracheid, Diffuse-porous, or Ring-porous) |
| probeDepth | distance from cambium (in m) for inner end of sap flux probe |
| probeStart (optional) | distance from cambium (in m) for outer end of sap flux probe (default = 0) |
| sapRadius (optional) | sapwood depth (in m) (default is unknown) |

Functions
---------

These are the VBA functions that can work behind Excel (like in SapfluxScaling.xlsm)

```VB.net
Function incgma(a, x)
With WorksheetFunction
    incgma = Exp(.GammaLn(a)) * (.GammaDist(x, a, 1, True))
End With
End Function

Function Qc(R, S, a, b, alp, bet)
 If S = 0 Then
    Scoef = (R * incgma(alp + 1, bet * R) / bet - incgma(alp + 2, bet * R) / (bet ^ 2))
 Else
    Scoef = (R * incgma(alp + 1, bet * S) / bet - incgma(alp + 2, bet * S) / (bet ^ 2))
 End If
    Qc = Scoef / ((R * incgma(alp + 1, bet * b) / bet - incgma(alp + 2, bet * b) / (bet ^ 2)) - (R * incgma(alp + 1, bet * a) / bet - incgma(alp + 2, bet * a) / (bet ^ 2)))
End Function

Public Function qtot(v As Double, treeRadius As Double, woodType As String, probeDepth As Double, Optional probeStart As Double = 0, Optional sapRadius As Double = 0)
 If sapRadius = 0 Then
    sapRadius = treeRadius
 End If
 
 If woodType = "Tracheid" Then
    alp = 0.3965098
    bet = 32.92486
 Else
    If woodType = "Diffuse-porous" Then
        alp = 0.2639205
        bet = 39.07246
    Else
        alp = 0.1270407
        bet = 81.05592
    End If
 End If

cfac = Qc(treeRadius, sapRadius, probeStart, probeDepth, alp, bet)
Ameas = WorksheetFunction.Pi * treeRadius ^ 2 - WorksheetFunction.Pi * (treeRadius - probeDepth) ^ 2
qtot = v * Ameas * cfac
 
End Function
```

### About
 - Author: Aaron Berdanier
 - License: MIT