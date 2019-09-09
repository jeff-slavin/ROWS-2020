Option Explicit

Public Sub ConvertArrayToString(ByRef asArray() As String, ByRef sResult As String)

     'Variable Declarations
     Dim i As Integer

     'Set initial state for the ByRef sResult parameter (empty string)
     sResult = ""
     
     'make sure we have an array
     If Not IsArray(asArray) Then Exit Sub
     
     'Loop through the array and convert it to a string
     For i = LBound(asArray) To UBound(asArray)
          sResult = sResult & asArray(i) & "|"
     Next i

End Sub