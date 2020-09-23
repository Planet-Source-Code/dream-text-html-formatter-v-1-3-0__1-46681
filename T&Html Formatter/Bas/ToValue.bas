Attribute VB_Name = "ToValue"

Option Explicit

'ToValue...
'05.2003 mrk Change/Add

Public Function ToValueIntMinIs0( _
              ByVal Value As Integer) As Integer
  Select Case Value
    Case Is < 0: ToValueIntMinIs0 = 0
    Case Else:   ToValueIntMinIs0 = Value
  End Select
End Function '05.2003 mrk Change/Add
Public Function ToValueIntMinMax( _
              ByVal Value As Integer, _
              ByVal Min As Integer, _
              ByVal Max As Integer) As Integer
  Select Case Value
    Case Is < Min: ToValueIntMinMax = Min
    Case Is > Max: ToValueIntMinMax = Max
    Case Else:     ToValueIntMinMax = Value
  End Select
End Function '05.2003 mrk Change/Add


Public Function ToValueLngMinIs0( _
              ByVal Value As Long) As Long
  Select Case Value
    Case Is < 0: ToValueLngMinIs0 = 0
    Case Else:   ToValueLngMinIs0 = Value
  End Select
End Function '05.2003 mrk Change/Add
Public Function ToValueLngMinMax( _
              ByVal Value As Long, _
              ByVal Min As Long, _
              ByVal Max As Long) As Long
  Select Case Value
    Case Is < Min: ToValueLngMinMax = Min
    Case Is > Max: ToValueLngMinMax = Max
    Case Else:     ToValueLngMinMax = Value
  End Select
End Function '05.2003 mrk Change/Add

Public Function ToValueSngMinIs0( _
              ByVal Value As Single) As Single
  Select Case Value
    Case Is < 0: ToValueSngMinIs0 = 0
    Case Else:   ToValueSngMinIs0 = Value
  End Select
End Function '05.2003 mrk Change/Add
Public Function ToValueSngMinMax( _
              ByVal Value As Single, _
              ByVal Min As Single, _
              ByVal Max As Single) As Single
  Select Case Value
    Case Is < Min: ToValueSngMinMax = Min
    Case Is > Max: ToValueSngMinMax = Max
    Case Else:     ToValueSngMinMax = Value
  End Select
End Function '05.2003 mrk Change/Add



