Attribute VB_Name = "HTMLCOMPILER"
Option Explicit
'credits Sherif Rofael
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function LockWindowUpdate& Lib "user32" (ByVal hwndLock As Long)

' Global Constants
'Global Const MainTitle = "HTML Messages Encoder"
'Global Const Title = "The VB Source Code After Being Converted Into HTML Format"
'Global Const AllFilter = "Hyper Text Markup Language (*.html)|*.html|Rich Text Format (*.rtf)|*.rtf|Plain-Text (*.txt)|*.txt"
Global Const Navy = &H800000
Global Const SRCCOPY = &HCC0020
'Global Const HWND_TOPMOST = -1
'Global Const HWND_NOTOPMOST = -2
'Global Const SWP_NOACTIVATE = &H10
'Global Const SWP_SHOWWINDOW = &H40
'
'' Message Box Constants
'Global Const MB_YESNO = 4
'Global Const MB_ICONQUESTION = 32
'Global Const MB_DEFBUTTON1 = &H0&
'Global Const MB_DEFBUTTON2 = 256
'Global Const IDYES = 6


' Public Variables
Public iRet As Long
Public currentCounter As Long
Public destFile As String
Public Msg As String
Public dgDef
Public Response
Public QT As String

' Global Variables
Global FontNameIndex%, FontSizeIndex%, i%
Public Function FileExists(Path$) As Integer
    Dim X
    X = FreeFile
    On Error Resume Next
    Open Path$ For Input As X
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close X
End Function
Public Function HTMLCompile(baseFont As String, baseSize As Long, baseColor As String) As String
    
    ' HTMLCompile is the actual function that converts the RichText to HTML.
        
    Dim arLine() As String, arWord() As String
    Dim curLine As Long, curWord As Long
    Dim curChar As Long, lastPos As Long
    Dim lastTag As String, curTag As String
    Dim lastFormat As String, lastCloseFormat As String
    Dim curFormat As String, curCloseFormat As String
    Dim InTag As Boolean, html As String
    Dim textLen As Long, Char As String
    
    With frmMDI.RTB
    textLen = Len(.Text)
    InTag = False
    html = ""
    lastTag = ""
    lastFormat = ""
    lastCloseFormat = ""
    lastPos = 0
    .SelStart = 0
    arLine = Split(.Text, vbCrLf)
    
    Call UpdateProgress(1, textLen, 1)
    
    For curLine = LBound(arLine) To UBound(arLine)
        arWord = Split(arLine(curLine), " ")
        If arLine(curLine) = "" Then lastPos = lastPos + 1
        For curWord = LBound(arWord) To UBound(arWord)
            For curChar = 1 To Len(arWord(curWord))
                
                .SelStart = lastPos
                .SelLength = 1
                curTag = ""
                curFormat = ""
                curCloseFormat = ""
                Char = HTMLChar(Mid$(arWord(curWord), curChar, 1))
                If .SelBold Then curFormat = curFormat & "<B>"
                If .SelItalic Then curFormat = curFormat & "<I>"
                If .SelUnderline Then curFormat = curFormat & "<U>"
                If .SelStrikeThru Then curFormat = curFormat & "<STRIKE>"
                
                If .SelStrikeThru Then curCloseFormat = curCloseFormat & "</STRIKE>"
                If .SelUnderline Then curCloseFormat = curCloseFormat & "</U>"
                If .SelItalic Then curCloseFormat = curCloseFormat & "</I>"
                If .SelBold Then curCloseFormat = curCloseFormat & "</B>"
                
                If Not baseFont = .SelFontName Then curTag = " face='" & QT & .SelFontName & QT & "'"
                If Not baseSize = .SelFontSize Then curTag = curTag & " size=" & QT & HTMLSize(.SelFontSize) & QT
                If Not baseColor = .SelColor Then curTag = curTag & " color=" & QT & HTMLColor(.SelColor) & QT
                If Not curTag = "" Then curTag = "FONT" & curTag
                
                If curTag = lastTag Then
                    If curFormat = lastFormat Then
                        html = html & Char
                    Else
                        html = html & lastCloseFormat & curFormat & Char
                    End If
                Else
                    html = html & lastCloseFormat
                    If Not lastTag = "" Then html = html & "</" & TagKeyword(lastTag) & ">": InTag = False
                    If Not curTag = "" Then html = html & "<" & curTag & ">": InTag = True
                    html = html & curFormat & Char
                End If
                lastTag = curTag
                lastFormat = curFormat
                lastCloseFormat = curCloseFormat
                lastPos = lastPos + 1
                Call UpdateProgress(0, textLen, 1)

            Next curChar
        Select Case InTag
        Case True
            html = html & " "
        Case False
            html = html & "&#32;"
        End Select
        lastPos = lastPos + 1
        Call UpdateProgress(0, textLen, 1)
        Next curWord
        html = html & vbCrLf
        'html = html & "<BR>" & vbCrLf
        lastPos = lastPos + 1
        Call UpdateProgress(0, textLen, 1)

    Next curLine
    If Not lastTag = "" Then html = html & "</" & TagKeyword(lastTag) & ">"
    End With
    
    HTMLCompile = html

End Function
Private Function HTMLSize(FontPoint As Long) As Long
    Select Case FontPoint
    Case 8 To 9
        HTMLSize = 1
    Case 10 To 11
        HTMLSize = 2
    Case 12 To 13
        HTMLSize = 3
    Case 14 To 17
        HTMLSize = 4
    Case 18 To 23
        HTMLSize = 5
    Case 24 To 35
        HTMLSize = 6
    Case 36 To 100
        HTMLSize = 7
    End Select
End Function
Private Function HTMLColor(FontColor As Long) As String
    
    Dim curColorIndex As Long
    Dim curColorString As String
    Dim resultColorString As String
    
    curColorString = Hex$(FontColor)
    If Len(curColorString) < 6 Then
        For curColorIndex = 1 To 6 - Len(curColorString)
            curColorString = "0" & curColorString
        Next
    End If
    resultColorString = ""
    
    ' Convert color to #RGB format
    For curColorIndex = 2 To 6 Step 2
        resultColorString = resultColorString & Left(Right$(curColorString, curColorIndex), 2)
    Next
    HTMLColor = "#" & resultColorString
    
End Function
Public Sub UpdateProgress(PositionReset As Long, Total As Long, ProgressBlock As Long)
    On Error Resume Next
    ' Updates the progress bar
    
    Static position
    Dim r As Long
    
    If PositionReset <> 0 Then
        position = 0
      
    End If
    position = position + CSng((ProgressBlock / Total) * 100)
    If position > 100 Then
        position = 100
    End If
        
End Sub


Public Function TagKeyword(sourceTag As String) As String
    
    ' This function takes out the tag's keyword
    ' by excluding any attribute that exists.
    ' For e.g: TagKeyword("FONT size=2 color=red") = "FONT"
    
    Dim curTagChar As Long
    For curTagChar = 1 To Len(sourceTag)
        If Mid$(sourceTag, curTagChar, 1) = " " Then GoTo TagCut
    Next curTagChar
    TagKeyword = sourceTag
    Exit Function
TagCut:
    TagKeyword = Left$(sourceTag, curTagChar - 1)
    Exit Function
End Function

Public Function HTMLChar(inChar As String) As String
    
    Dim chrAsc As Integer
    chrAsc = Asc(inChar)
        
    If chrAsc = 46 Or chrAsc = 44 Or chrAsc = 39 Or chrAsc = 40 Or chrAsc = 41 Or chrAsc = 45 Or chrAsc = 58 Then
        HTMLChar = inChar
    ElseIf chrAsc < 32 Then
        HTMLChar = "&#" & chrAsc & ";"
    ElseIf (chrAsc >= 33) And (chrAsc <= 47) Then
        HTMLChar = "&#" & chrAsc & ";"
    ElseIf (chrAsc >= 58) And (chrAsc <= 63) Then
        HTMLChar = "&#" & chrAsc & ";"
    ElseIf (chrAsc >= 91) And (chrAsc <= 96) Then
        HTMLChar = "&#" & chrAsc & ";"
    ElseIf (chrAsc >= 123) Then
        HTMLChar = "&#" & chrAsc & ";"
    Else
        HTMLChar = inChar
    End If
    
End Function
Public Function HTMLString(inString As String) As String
    
    Dim curChr As Long
    Dim curChrAsc As Integer
    
    HTMLString = ""
    For curChr = 1 To Len(inString)
        
        curChrAsc = Asc(Mid$(inString, curChr, 1))
        If curChrAsc = 46 Or curChrAsc = 44 Or curChrAsc = 39 Or curChrAsc = 40 Or curChrAsc = 41 Or curChrAsc = 45 Or curChrAsc = 58 Then
            HTMLString = HTMLString & Mid$(inString, curChr, 1)
        ElseIf curChrAsc < 32 Then
            HTMLString = HTMLString & "&#" & curChrAsc & ";"
        ElseIf (curChrAsc >= 33) And (curChrAsc <= 47) Then
            HTMLString = HTMLString & "&#" & curChrAsc & ";"
        ElseIf (curChrAsc >= 58) And (curChrAsc <= 63) Then
            HTMLString = HTMLString & "&#" & curChrAsc & ";"
        ElseIf (curChrAsc >= 91) And (curChrAsc <= 96) Then
            HTMLString = HTMLString & "&#" & curChrAsc & ";"
        ElseIf (curChrAsc >= 123) Then
            HTMLString = HTMLString & "&#" & curChrAsc & ";"
        Else
            HTMLString = HTMLString & Mid$(inString, curChr, 1)
        End If
    Next curChr
    
End Function

