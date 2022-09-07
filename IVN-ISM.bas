Type Bound
U As Double
L As Double
End Type

Type IVN
Tr As Bound
In As Bound
Fa As Bound

End Type

Sub IVNISM()

Dim x, z As Integer





x = InputBox("Enter the number of factors:")
z = InputBox("Enter the number of experts:")





Dim IVNExpert() As IVN
ReDim IVNExpert(x, x, z) As IVN

Dim IVNDec() As IVN
ReDim IVNDec(x, x) As IVN

Dim CrispDec() As Double
ReDim CrispDec(x, x) As Double

Dim RM() As Double
ReDim RM(x, x) As Double

Dim BRM() As Double
ReDim BRM(1 To x, 1 To x) As Double

Dim FRM() As Double
ReDim FRM(1 To x, 1 To x) As Double

Dim YRM() As Double
ReDim YRM(1 To x, 1 To x) As Double

Dim MICRM() As Double
ReDim MICRM(1 To x, 1 To x) As Double

Dim DR() As Double
ReDim DR(1 To x) As Double

Dim DE() As Double
ReDim DE(1 To x) As Double

Dim Total As IVN




For i = 1 To x * z

    For j = 1 To x
    
    
    
    If i Mod x = 0 Then
    t = i / x
    
    
    Select Case Cells(i, j)
    Case 0
    IVNExpert(x, j, t).Tr.L = 0
    IVNExpert(x, j, t).Tr.U = 0
    IVNExpert(x, j, t).In.L = 0
    IVNExpert(x, j, t).In.U = 0
    IVNExpert(x, j, t).Fa.L = 1
    IVNExpert(x, j, t).Fa.U = 1
    Case 1
    IVNExpert(x, j, t).Tr.L = 0
    IVNExpert(x, j, t).Tr.U = 0.25
    IVNExpert(x, j, t).In.L = 0
    IVNExpert(x, j, t).In.U = 0.1
    IVNExpert(x, j, t).Fa.L = 0.7
    IVNExpert(x, j, t).Fa.U = 0.95
    Case 2
    IVNExpert(x, j, t).Tr.L = 0.2
    IVNExpert(x, j, t).Tr.U = 0.5
    IVNExpert(x, j, t).In.L = 0.1
    IVNExpert(x, j, t).In.U = 0.2
    IVNExpert(x, j, t).Fa.L = 0.5
    IVNExpert(x, j, t).Fa.U = 0.75
    Case 3
    IVNExpert(x, j, t).Tr.L = 0.5
    IVNExpert(x, j, t).Tr.U = 0.75
    IVNExpert(x, j, t).In.L = 0.1
    IVNExpert(x, j, t).In.U = 0.2
    IVNExpert(x, j, t).Fa.L = 0.2
    IVNExpert(x, j, t).Fa.U = 0.5
    Case 4
    IVNExpert(x, j, t).Tr.L = 0.7
    IVNExpert(x, j, t).Tr.U = 0.95
    IVNExpert(x, j, t).In.L = 0
    IVNExpert(x, j, t).In.U = 0.1
    IVNExpert(x, j, t).Fa.L = 0
    IVNExpert(x, j, t).Fa.U = 0.25
    
    End Select
    
    
    
    Else
    t = Int(i / x) + 1
    
    Select Case Cells(i, j)
    Case 0
    IVNExpert(i Mod x, j, t).Tr.L = 0
    IVNExpert(i Mod x, j, t).Tr.U = 0
    IVNExpert(i Mod x, j, t).In.L = 0
    IVNExpert(i Mod x, j, t).In.U = 0
    IVNExpert(i Mod x, j, t).Fa.L = 1
    IVNExpert(i Mod x, j, t).Fa.U = 1
    Case 1
    IVNExpert(i Mod x, j, t).Tr.L = 0
    IVNExpert(i Mod x, j, t).Tr.U = 0.25
    IVNExpert(i Mod x, j, t).In.L = 0
    IVNExpert(i Mod x, j, t).In.U = 0.1
    IVNExpert(i Mod x, j, t).Fa.L = 0.7
    IVNExpert(i Mod x, j, t).Fa.U = 0.95
    Case 2
    IVNExpert(i Mod x, j, t).Tr.L = 0.2
    IVNExpert(i Mod x, j, t).Tr.U = 0.5
    IVNExpert(i Mod x, j, t).In.L = 0.1
    IVNExpert(i Mod x, j, t).In.U = 0.2
    IVNExpert(i Mod x, j, t).Fa.L = 0.5
    IVNExpert(i Mod x, j, t).Fa.U = 0.75
    Case 3
    IVNExpert(i Mod x, j, t).Tr.L = 0.5
    IVNExpert(i Mod x, j, t).Tr.U = 0.75
    IVNExpert(i Mod x, j, t).In.L = 0.1
    IVNExpert(i Mod x, j, t).In.U = 0.2
    IVNExpert(i Mod x, j, t).Fa.L = 0.2
    IVNExpert(i Mod x, j, t).Fa.U = 0.5
    Case 4
    IVNExpert(i Mod x, j, t).Tr.L = 0.7
    IVNExpert(i Mod x, j, t).Tr.U = 0.95
    IVNExpert(i Mod x, j, t).In.L = 0
    IVNExpert(i Mod x, j, t).In.U = 0.1
    IVNExpert(i Mod x, j, t).Fa.L = 0
    IVNExpert(i Mod x, j, t).Fa.U = 0.25
    
    End Select
    
    
    End If
    
   Next j
    
Next i


For t = 1 To z
For i = 1 To x
For j = 1 To x

Cells(i + x * (t - 1), j + x + 3) = "[" & IVNExpert(i, j, t).Tr.L & ";" & IVNExpert(i, j, t).Tr.U & "]" & ";" & "[" & IVNExpert(i, j, t).In.L & ";" & IVNExpert(i, j, t).In.U & "]" & ";" & "[" & IVNExpert(i, j, t).Fa.L & ";" & IVNExpert(i, j, t).Fa.U & "]"

Next j
Next i
Next t
TotalRM = 0
For i = 1 To x
For j = 1 To x
cTrL = 1
cInL = 1
cFaL = 1
cTrU = 1
cInU = 1
cFaU = 1

For t = 1 To z
cTrL = cTrL * ((1 - IVNExpert(i, j, t).Tr.L) ^ (1 / z))
cTrU = cTrU * ((1 - IVNExpert(i, j, t).Tr.U) ^ (1 / z))
cInL = cInL * (IVNExpert(i, j, t).In.L ^ (1 / z))
cInU = cInU * (IVNExpert(i, j, t).In.U ^ (1 / z))
cFaL = cFaL * (IVNExpert(i, j, t).Fa.L ^ (1 / z))
cFaU = cFaU * (IVNExpert(i, j, t).Fa.U ^ (1 / z))

Next t
IVNDec(i, j).Tr.L = 1 - cTrL
IVNDec(i, j).Tr.U = 1 - cTrU
IVNDec(i, j).In.L = cInL
IVNDec(i, j).In.U = cInU
IVNDec(i, j).Fa.L = cFaL
IVNDec(i, j).Fa.U = cFaU
Cells(i + 1, j + 2 * x + 4) = "(" & "[" & Round(IVNDec(i, j).Tr.L, 2) & ";" & Round(IVNDec(i, j).Tr.U, 2) & "]" & ";" & "[" & Round(IVNDec(i, j).In.L, 2) & ";" & Round(IVNDec(i, j).In.U, 2) & "]" & ";" & "[" & Round(IVNDec(i, j).Fa.L, 2) & ";" & Round(IVNDec(i, j).Fa.U, 2) & "]" & ")"
CrispDec(i, j) = (IVNDec(i, j).Tr.L + IVNDec(i, j).Tr.U + (1 - IVNDec(i, j).Fa.L) + (1 - IVNDec(i, j).Fa.U) + (IVNDec(i, j).Tr.L * IVNDec(i, j).Tr.U) + ((1 - IVNDec(i, j).Fa.L) * (1 - IVNDec(i, j).Fa.U)) ^ (1 / 2)) * ((1 - (IVNDec(i, j).In.L + IVNDec(i, j).In.U) / 2) * ((1 - IVNDec(i, j).In.L) * (1 - IVNDec(i, j).In.U)) ^ (1 / 2)) / 12
Cells(i + 1, j + 3 * x + 6) = Round(CrispDec(i, j), 4)
RM(i, j) = CrispDec(i, j)
TotalRM = TotalRM + CrispDec(i, j)
Next j
Next i

TrH = TotalRM / (x * x)


Cells(1, x + 2) = "Experts' opinions in IVN numbers form"

Cells(1, 2 * x + Int(x / 2)) = "IVN Relationship Matrix"
Cells(1, 3 * x + Int(x / 2)) = "Relationship Matrix"


Cells(1, 4 * x + 8) = "The threshold value "
Cells(2, 4 * x + 8) = TrH



Cells(1, 6 * x + 16) = "The initial relationship matrix  "

For i = 1 To x
For j = 1 To x

If RM(i, j) > TrH Then
BRM(i, j) = 1
Else
BRM(i, j) = 0

End If
If i = j Then
BRM(i, j) = 1
End If

Cells(i + 1, j + 6 * x + 14) = BRM(i, j)
FRM(i, j) = BRM(i, j)
Next j
Next i

For i = 1 To x

For j = 1 To x

If FRM(i, j) = 1 Then
For k = 1 To x
If FRM(j, k) = 1 Then
FRM(i, k) = 1
End If
Next k
End If



Cells(i + 1, j + 7 * x + 16) = FRM(i, j)

Next j

Next i





For i = 1 To x


For j = 1 To x





Cells(i + 1, j + 7 * x + 16) = FRM(i, j)

If BRM(i, j) <> FRM(i, j) Then
Cells(i + 1, j + 8 * x + 18) = FRM(i, j) & "*"
Else
Cells(i + 1, j + 8 * x + 18) = FRM(i, j)
End If

YRM(i, j) = FRM(i, j)

MICRM(i, j) = FRM(i, j)

Next j

Next i




lev = 0

100: lev = lev + 1

For i = 1 To x
For j = 1 To x
YRM(i, j) = FRM(i, j)
Next j
Next i




For i = 1 To x
Cells((lev - 1) * x + i + 1, 9 * x + 20) = i
sayx = 0
sayy = 0
sayxy = 0

For j = 1 To x

If YRM(i, j) = 1 Then

sayx = sayx + 1

If sayx = 1 Then
Cells((lev - 1) * x + i + 1, 9 * x + 21) = j
ElseIf sayx > 1 Then

Cells((lev - 1) * x + i + 1, 9 * x + 21) = Cells((lev - 1) * x + i + 1, 9 * x + 21).Value & ";" & j

End If

End If

If YRM(j, i) = 1 Then
sayy = sayy + 1

If sayy = 1 Then
Cells((lev - 1) * x + i + 1, 9 * x + 22) = j
ElseIf sayy > 1 Then

Cells((lev - 1) * x + i + 1, 9 * x + 22) = Cells((lev - 1) * x + i + 1, 9 * x + 22).Value & ";" & j

End If

End If


If YRM(i, j) = 1 And YRM(j, i) = 1 Then
sayxy = sayxy + 1

If sayxy = 1 Then
Cells((lev - 1) * x + i + 1, 9 * x + 23) = j
ElseIf sayxy > 1 Then

Cells((lev - 1) * x + i + 1, 9 * x + 23) = Cells((lev - 1) * x + i + 1, 9 * x + 23).Value & ";" & j

End If

End If



Next j

If Cells((lev - 1) * x + i + 1, 9 * x + 21) <> 0 And Cells((lev - 1) * x + i + 1, 9 * x + 21).Value = Cells((lev - 1) * x + i + 1, 9 * x + 23).Value Then
Cells((lev - 1) * x + i + 1, 9 * x + 24) = lev

For s = 1 To x
FRM(i, s) = 0
FRM(s, i) = 0
Next s
End If



Next i

sayson = 0
For i = 1 To x


For j = 1 To x



If YRM(i, j) > 0 Then
sayson = sayson + 1
End If

Next j

Next i

If sayson > 0 Then
GoTo 100
End If




For i = 1 To x


For j = 1 To x
DR(i) = MICRM(i, j) + DR(i)
DE(i) = MICRM(j, i) + DE(i)




Next j
Cells(i + 1, 9 * x + 27) = DR(i)
Cells(i + 1, 9 * x + 26) = DE(i)
'Cells(i + 1, 9 * x + 26) = "(" & DE(i) & "," & DR(i) & ")"
Next i





Cells(1, 9 * x + 20) = "Element (Pi)"

Cells(1, 9 * x + 21) = "Reachability set: R (Pi)"
Cells(1, 9 * x + 22) = "Antecedent set: A (Pi)"
Cells(1, 9 * x + 23) = "Intersection R (Pi)n A (Pi)"
Cells(1, 9 * x + 24) = "Level"

Cells(1, 7 * x + 18) = "The final relationship matrix  "

Cells(1, 9 * x + 27) = "The driving power "
Cells(1, 9 * x + 26) = "The dependence power"
Cells(1, 9 * x + 29) = "Factor number"
Cells(1, 9 * x + 30) = "Factor type"
For i = 1 To x
Cells(i + 1, 9 * x + 29) = i
If DR(i) >= x / 2 And DE(i) < x / 2 Then
Cells(i + 1, 9 * x + 30) = "Driving"
Cells(i + 1, 9 * x + 30).Font.ColorIndex = 3

ElseIf DR(i) >= x / 2 And DE(i) >= x / 2 Then
Cells(i + 1, 9 * x + 30) = "Linkage"
Cells(i + 1, 9 * x + 30).Font.ColorIndex = 5

ElseIf DR(i) < x / 2 And DE(i) >= x / 2 Then
Cells(i + 1, 9 * x + 30) = "Dependent"
Cells(i + 1, 9 * x + 30).Font.ColorIndex = 4

ElseIf DR(i) < x / 2 And DE(i) < x / 2 Then
Cells(i + 1, 9 * x + 30) = "Autonomous"
Cells(i + 1, 9 * x + 30).Font.ColorIndex = 6
End If


Next i






End Sub
