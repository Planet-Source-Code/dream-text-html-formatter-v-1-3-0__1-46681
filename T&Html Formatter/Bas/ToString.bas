Attribute VB_Name = "ToString"

Option Explicit

'ToStr...
'05.2003 mrk Change/Add

Public Function ToStrScrollToLeft( _
              ByVal Text As String) As String
  Select Case Len(Text)
    Case Is < 2: ToStrScrollToLeft = Text
    Case Else:   ToStrScrollToLeft = Mid$(Text, 2) & Left$(Text, 1)
  End Select
End Function '05.2003 mrk Change/Add
              
Public Function ToStrScrollToRight( _
              ByVal Text As String) As String
  Select Case Len(Text)
    Case Is < 2: ToStrScrollToRight = Text
    Case Else:   ToStrScrollToRight = Right$(Text, 1) & Mid$(Text, 1, Len(Text) - 1)
  End Select
End Function '05.2003 mrk Change/Add
              
Public Function ToStrRemoveExtra( _
              ByVal Text As String, _
     Optional ByVal CharToOne As String = " ") As String
'1.2.3.4.! = RemoveExtra("1.2..3...4....!", ".")
  
  Dim nPos As Long, nLen As Long
  Dim sCharNow As String, sReturn As String
  
  nLen = Len(Text)
  If nLen = 0 Then
     ToStrRemoveExtra = ""
Exit Function
  End If
  
  Select Case Len(CharToOne)
   Case Is > 1: CharToOne = Mid$(CharToOne, 1, 1)
   Case Is < 1: CharToOne = Chr$(32)
  End Select

  sReturn = Mid$(Text, 1, 1)
  
  For nPos = 2 To nLen
      sCharNow = Mid$(Text, nPos, 1)
      If sCharNow = CharToOne Then
         If Right$(sReturn, 1) <> CharToOne Then
            sReturn = sReturn & sCharNow
         End If
        
        Else
         sReturn = sReturn & sCharNow
      End If
  Next
  ToStrRemoveExtra = sReturn
End Function '05.2003 mrk Change/Add
              
Public Function ToStrCRC32( _
              ByVal Text As String) As Long
    Dim i As Long
    Dim J As Long
    Dim l As Long
    Dim nPowers(0 To 7) As Long
    Dim nCRC As Long
    Dim nByte As Integer
    Dim nBit As Boolean

    nPowers(0) = 1
    nPowers(1) = 2
    nPowers(2) = 4
    nPowers(3) = 8
    nPowers(4) = 16
    nPowers(5) = 32
    nPowers(6) = 64
    nPowers(7) = 128
    
    l = Len(Text)
    For i = 1 To l
        nByte = Asc(Mid$(Text, i, 1))
        For J = 7 To 0 Step -1
            nBit = CBool((nCRC And 32768) = 32768) Xor _
                        ((nByte And nPowers(J)) = nPowers(J))
            nCRC = (nCRC And 32767&) * 2&
            If nBit Then
               nCRC = nCRC Xor &H8005&
            End If
        Next 'j
    Next 'i
    Erase nPowers
    ToStrCRC32 = nCRC
End Function '05.2003 mrk Change/Add


