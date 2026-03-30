VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Random"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'RandomLib by Arfur 2026
'Github: https://github.com/therealarfu
Option Explicit
Private initialized As Boolean

Private Sub Class_Initialize()
    If Not initialized Then
        Randomize
        initialized = True
    End If
End Sub

Public Sub SetSeed(ByVal Seed As Long)
    Rnd -1
    Randomize Seed
    initialized = True
End Sub

Public Function NextDouble(Optional ByVal Minimum As Double = 0, Optional ByVal Maximum As Double = 1, Optional ByVal DecimalPlaces As Integer = -1) As Double
    Dim value As Double
    If Maximum < Minimum Then Err.Raise 5, "Random", "Maximum must be >= Minimum"
    
    value = (Maximum - Minimum) * Rnd + Minimum

    If DecimalPlaces > -1 Then
        value = Round(value, DecimalPlaces)
    End If

    NextDouble = value
End Function

Public Function NextBoolean(Optional probabilityTrue As Double = 0.5) As Boolean
    If probabilityTrue < 0 Or probabilityTrue > 1 Then
        Err.Raise 5, "Random", "Probability must be between 0 and 1"
    End If
    NextBoolean = (Rnd < probabilityTrue)
End Function

Public Function NextInt(Optional ByVal Minimum As Long = 0, Optional ByVal Maximum As Long = 1) As Long
    If Maximum < Minimum Then Err.Raise 5, "Random", "Maximum must be >= Minimum"
    NextInt = Int((Maximum - Minimum + 1) * Rnd) + Minimum
End Function

Public Function Choice(arr As Variant) As Variant
    If Not IsArray(arr) Then Exit Function
    Choice = arr(Int((UBound(arr) - LBound(arr) + 1) * Rnd) + LBound(arr))
End Function

Public Function Shuffled(arr As Variant) As Variant
    If Not IsArray(arr) Then Exit Function
    Dim i As Long
    Dim j As Long
    Dim upper As Long
    Dim lower As Long
    Dim temp As Variant
    Dim temp_arr As Variant
    
    upper = UBound(arr)
    lower = LBound(arr)
    ReDim temp_arr(lower To upper)
    For i = lower To upper
        temp_arr(i) = arr(i)
    Next
    
    For i = upper To lower + 1 Step -1
        j = Int((i - lower + 1) * Rnd) + lower
        temp = temp_arr(i)
        temp_arr(i) = temp_arr(j)
        temp_arr(j) = temp
    Next
    
    Shuffled = temp_arr
End Function

Public Function NextString(ByVal Length As Long, Optional ByVal Alphabet As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz") As String
    Dim i As Long
    Dim lenalf As Long
    If Length < 0 Then Err.Raise 5, "Random", "Length must be >= 0"
    lenalf = Len(Alphabet)

    For i = 1 To Length
        NextString = NextString & Mid(Alphabet, Int(lenalf * Rnd) + 1, 1)
    Next
End Function

Public Function WeightedChoice(Weights As Variant, Values As Variant) As Variant
    WeightedChoice = False
    If Not IsArray(Weights) Or Not IsArray(Values) Then Exit Function
    
    Dim i As Long
    Dim upper As Long
    Dim lower As Long
    Dim prev As Double
    Dim rand As Double
    Dim total As Double
    Dim w As Double
    prev = 0
    
    If UBound(Weights) <> UBound(Values) Or LBound(Weights) <> LBound(Values) Then Exit Function
    upper = UBound(Weights)
    lower = LBound(Weights)
    
    For i = lower To upper
        w = Abs(Weights(i))
        total = total + w
    Next
    
    rand = Rnd * total
    
    For i = lower To upper
        w = Abs(Weights(i))
        If rand < w + prev Then
            WeightedChoice = Values(i)
            Exit Function
        End If
        
        prev = prev + w
    Next
End Function

