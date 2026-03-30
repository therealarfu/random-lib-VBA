# 🎲 Random Utilities in VBA — ``Random.cls``
A simple class module with functions that can manipulate random values.
Inclunding: Shuffled(Fisher-Yates Algorithm), Weighted Choices and Seed control.

---

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

---

# 🎰 Random numbers
They return a random number within a range.

- ``NextDouble(Optional ByVal Minimum As Double = 0, Optional ByVal Maximum As Double = 1, Optional ByVal DecimalPlaces As Integer = -1)``
- ``NextInt(Optional Minimum As Long = 0, Optional Maximum As Long = 1)``

---

# 🎯 Random boolean
Returns a random boolean

- ``NextBoolean(Optional probabilityTrue As Double = 0.5)``

---

# 📦 Random choice
Returns a random value in an array

- ``Choice(arr As Variant)``

---

# 🔀 Shuffle (Fisher-Yates)
Returns a shuffled version of an array with the Fisher-Yates algorithm

- ``Shuffled(arr As Variant)``

---

# 🔤 Random string
Returns a random string based on the length and the alphabet

- ``NextString(Length As Long, Optional Alphabet As String)``

---

# ⚖️ Weighted choice
Returns a random choice based on the weights

- ``WeightedChoice(Weights As Variant, Values As Variant)``

---

# 🎲 Deterministic seed

- ``SetSeed(Seed As Long)``


