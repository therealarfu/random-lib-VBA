# 🎲 Random Utilities in VBA — ``Random.cls``
A simple class module with functions that can manipulate random values.
Inclunding: Shuffled(Fisher-Yates Algorithm), Weighted Choices and Seed control.

--

# 💡 How to use
```vb
Sub Main()
  Dim randomtest As New Random
  Dim randomint as Integer
  Dim arr as Variant
  arr = Array(1, 2, 3, 4)
  randomint = randomtest.NextInt(0, 10)

  MsgBox "Random Integer: " & randomint
  MsgBox "Shuffled Array: " & Join(randomtest.Shuffled(arr), ", ")
  Msgbox "Weighted Choice: " & randomtest.WeightedChoice(Array(50, 30, 20), Array("Banana", "Apple", "Chocolate"))
End Sub
```

--

# 🎰 Random numbers
They return a random number within a range.

- ``NextDouble(Optional ByVal Minimum As Double = 0, Optional ByVal Maximum As Double = 1, Optional ByVal DecimalPlaces As Integer = -1) As Double``
- ``NextInt(Optional Minimum As Long = 0, Optional Maximum As Long = 1) As Long``

--

# 🎯 Random boolean

- ``NextBoolean(Optional probabilityTrue As Double = 0.5)``

--

# 📦 Random choice

- ``Choice(arr As Variant)``

--

# 🔀 Shuffle (Fisher-Yates)

- ``Shuffled(arr As Variant)``

--

# 🔤 Random string

- ``NextString(Length As Long, Optional Alphabet As String)``

--

# ⚖️ Weighted choice

- ``WeightedChoice(Weights As Variant, Values As Variant)``

--

# 🎲 Deterministic seed

- ``SetSeed(Seed As Long)``


