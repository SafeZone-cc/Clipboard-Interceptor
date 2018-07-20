Attribute VB_Name = "HexToDec"
Option Explicit

Public Sub RegConvert()
    'arrHexValues = Split(Replace(strHexValues, "hex:", ""), ",")
    'arrDecValues = DecimalNumbers(arrHexValues)
End Sub

Function DecimalNumbers(arrHex)
   Dim i, strDecValues
   For i = 0 To UBound(arrHex)
     If IsEmpty(strDecValues) Then
       strDecValues = CLng("&H" & arrHex(i))
     Else
       strDecValues = strDecValues & "," & CLng("&H" & arrHex(i))
     End If
   Next
   DecimalNumbers = Split(strDecValues, ",")
End Function

