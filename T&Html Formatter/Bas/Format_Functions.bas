Attribute VB_Name = "Format_Functions"

Option Explicit

Public Const SPC_VL = "&nbsp;"
Public strHTMLPreview As String
Public strHTMLStart As String

Public Wrapped          As Boolean   'Tells us if we have previewed and WRAPPED already
Public DoubledTxt       As Boolean   'Lets us know txt spaces doubled(either option)

Public Sub AddText(ByVal Text As String, _
                    Optional ByVal AddvbCrLf As Boolean = True)  ' Add text to textbox
  If AddvbCrLf Then
     frmMDI.RTB.SelText = Text & vbCrLf           'If last item in textbox tell it boolean false
    Else                                   'So it doesnt add the vbCrLf onto the end for
     frmMDI.RTB.SelText = Text                    'the next new line cause there isnt one
  End If
End Sub

Public Sub DoubleUpMargin(Optional existing As Boolean = False)
      Dim i As Integer
      Dim TextLines As Long         'No of lines in RTB
      Dim TextBuff As String
      Dim CharRet As Long
      Dim strLine As String         'Current line in RTB being processed
      Dim Count As Integer          'Current position within strLine
      Dim strNew As String          'String holding NEW Text:
                                      'Replace RTB with this when done
10    On Error GoTo LogError
20    Screen.MousePointer = 11
30    LockedWindow = True
40    frmMDI.SB.Panels(1).Text = "Status: Spacing"               '## Print
50    With frmMDI.RTB
60        TextLines = SendMessage(.hwnd, &HBA, 0, 0)    'Get number of lines in text box
  
70        For i = 0 To TextLines - 1                           'Extract & print each line in TextBox
80            TextBuff = Space(1000)
90            Mid(TextBuff, 1, 1) = Chr(79 And &HFF)             'Setup buffer for the line!
100           Mid(TextBuff, 2, 1) = Chr(79 \ &H100)
110           CharRet = SendMessage(.hwnd, &HC4, i, ByVal TextBuff) 'Get the data from the line
120           strLine = Left(TextBuff, CharRet)

130           For Count = 1 To Len(strLine)
140               Select Case existing
                         Case False
150                           If Mid(strLine, Count, 1) = Chr$(32) Then
160                              strNew = strNew & Chr$(32) & Chr$(32)
170                            Else
180                              If i = TextLines - 1 Then
190                                 strNew = strNew & Mid(strLine, Count)
200                               Else
210                                 strNew = strNew & Mid(strLine, Count) ' & vbCrLf
220                              End If                   ' If you want a spacey look in RTB Uncomment...
230                              GoTo Done                ' the & vbCrLf above!
240                           End If
250                      Case True
260                           strNew = strNew & Chr$(32) & Chr$(32)
270                           strNew = strNew & Mid(strLine, Count)
280                           GoTo Done
290               End Select
300            Next Count
Done:
310        Next i
320       .Text = strNew
330    End With
340    Screen.MousePointer = 1
350    With frmMDI.SB
360            .Panels(1).Text = "Status: Done"
370            .Panels(2).Text = "Margin Formatted           "
380            .Panels(3).Text = "Lines: " & TextLines - 1
390    End With
400    DoubledTxt = True
410    LockedWindow = False
     
420   On Error GoTo 0
430   Exit Sub
LogError:
440     LockedWindow = False
450     Screen.MousePointer = 1
460     LockWindowUpdate 0&
470     frmMDI.SB.Panels(1).Text = "Status: Margin Error"
480     frmMDI.SB.Panels(2).Text = Err.Number & "  " & Err.Description & "           "
       'Log the error
490     If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                               "{DoubleUpMargin}", _
                               "{Public Sub}", _
                               "{Format_Functions}", _
                               "{Format_Functions}") Then
500     End If
510     Resume Next ' Resume 'Exit Sub
End Sub

Public Sub FormatText()
    On Error GoTo LogError

20      Screen.MousePointer = 11                      'Format the Text to Html

        LockWindowUpdate frmMDI.RTB.hwnd
40      strHTMLPreview = HTMLCompile("System", 1, vbBlack)
        LockWindowUpdate 0&
            
        frmMDI.RTB.SelStart = 0
50      frmMDI.SB.Panels(1).Text = "Status: Formatted"
60      Screen.MousePointer = 1
        frmMDI.Text1.Visible = True
On Error GoTo 0
70  Exit Sub
LogError:
  If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                               "{FormatText}", _
                               "{Public Sub}", _
                               "{Format_Functions}", _
                               "{Format_Functions}") Then
80    End If
90    Resume Next ' Resume 'Exit Sub
End Sub

Public Sub Insert(sText As String)                           'RTB Text Insertion
   frmMDI.RTB.SelText = sText
End Sub

Public Sub PreviewHtml(Optional Wrap As Boolean = True)
     Dim Inputval As String
10   On Error GoTo LogError
20   If frmMDI.mnuFrameWrap.Checked = True Then        'If Frame Wrap On Preview is selected
   

30    If Wrap = True Then                      'If preview not example from frmmain_load
        Inputval = InputBox("What Is Your Article Title?", "Title", "My Title Here")
       Else
        Inputval = "Text & Html Formatter Example"
      End If
      
        If frmMDI.mnuHtmlPage.Checked = True Then
           strHTMLStart = "<HTML><HEAD><TITLE>" & Inputval & "</TITLE></HEAD><BODY>" & vbNewLine
        End If
        
        'Table start at beginning of RTB
80      strHTMLStart = strHTMLStart & "<center><table width= 100% bordercolor = '#1010BA' " & _
                      "cellspacing='0' cellpadding='20' border='2'><tr><td>"
 
        'and now add the following
90      strHTMLStart = strHTMLStart & vbNewLine & "<center><font size = 4 color = #1010BA>" & _
                 vbNewLine & "<b>"
        
        If Inputval <> "" Then strHTMLStart = strHTMLStart & Inputval      ' Text & Html Formatter
                 
        strHTMLStart = strHTMLStart & "</b></font><hr color=#1010BA width = 90%><br></center><Pre>" & vbNewLine
        
        'go to end of RTB Text
100     strHTMLPreview = strHTMLStart & strHTMLPreview
        
        'and now add the following
110     strHTMLPreview = strHTMLPreview & "</Pre><center><hr color=#1010BA width = 90%><font size = 2>" & _
                 "Formatted with 'Text & Html Formatter'.<br>" & vbNewLine & "<a href = " & _
                 "http://www.dream-domain.net>www.dream-domain.net</a href></td></tr></table></center>"

        If frmMDI.mnuHtmlPage.Checked = True Then
           strHTMLPreview = strHTMLPreview & "</BODY></HTML>"
        End If

130       frmMDI.RTB.SelStart = 0                     'Set RTB Cursor back to beginning of RTB
140       Wrapped = True                       'Set the Wrapped boolean to True

160        End If
170

        frmMDI.Text1 = strHTMLPreview
        strHTMLStart = vbNullString
        
180     ToFileSave ToAppPath & "preview.htm", strHTMLPreview           'Save to file
'190     frmHTML.WB1.Navigate ToAppPath & "preview.htm"                   'Preview HTML
        frmMDI.Caption = " " & Inputval & " - Html Preview"
        frmMDI.WB1.Navigate ToAppPath & "preview.htm"
200     frmMDI.SB.Panels(1).Text = "Status: Html Previewed"           'if enabled text unformatted

220     frmMDI.SB.Panels(2).Text = "        "
230     Screen.MousePointer = 1
        frmMDI.Tab1.TabIndex = 2
        frmMDI.WB1.Visible = True
250   On Error GoTo 0
260   Exit Sub
LogError:
270     If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                               "{PreviewHtml}", _
                               "{Public Sub}", _
                               "{Format_Functions}", _
                               "{Format_Functions}") Then
280     End If
290     Resume Next ' Resume 'Exit Sub
End Sub

Public Function StripTags(ByVal strHTML As String, ByVal strIgnoreTags As String)
   Dim i As Long, J As Long, ii As Long, jj As Long
   Dim arr_strIgnoreTags()
   Dim strIgnoreTag As String
   Dim css_Tags() As String
   Dim strCssTag As String
   Dim bIgnoreTags As Boolean, bIgnoreTag As Boolean, iIndex As Long
10      On Error GoTo LogError
  
20      strCssTag = "font,bold,body,px,position,class,href,link,visibility,color,background,form,action"
 
       'Put CSS Tags Into Array
30      css_Tags() = Split(strCssTag, ",")
       'Put Ignore Tags Into Array
40      If Len(strIgnoreTags) > 0 Then
50         arr_strIgnoreTags = Split(strIgnoreTags, ",")
60         bIgnoreTags = True
70      End If
  
80      strHTML = Trim(strHTML)
90      If IsNull(strHTML) Then strHTML = vbNullString
  
       'STRIP text breakpoints
100     strHTML = Replace(strHTML, vbCrLf, vbNullString)
       'STRIP html, replace space tag with " "
110     strHTML = Replace(strHTML, LCase(SPC_VL), Chr$(32))
       'STRIP <br>, replace with vbCrLf
120     strHTML = Replace(strHTML, LCase("<br>"), vbCrLf)  'replace <br> with vbCrLf
  
       'STRIP <> Tags
130     i = InStr(1, strHTML, "<")
140     Do While i <> 0
150        bIgnoreTag = False
160        If bIgnoreTags Then
170           For iIndex = 0 To UBound(arr_strIgnoreTags)
180             strIgnoreTag = Trim(arr_strIgnoreTags(iIndex))
              ' Debug.Print UCase(Mid(strHtml, i, Len(strIgnoreTag)))
190             If UCase(Mid(strHTML, i, Len(strIgnoreTag))) = UCase(strIgnoreTag) Then
200                bIgnoreTag = True
210                Exit For
220             End If
230           Next
240        End If

250        If Not bIgnoreTag Then
260           J = InStr(i + 1, strHTML, ">")
270           If J <> 0 Then
    
             '###################################
             'STRIP Java Scripts
            ' Debug.Print "START" & LCase(Mid(strHtml, i, J - i + 1))
280           If InStr(1, LCase(Mid(strHTML, i, J - i + 1)), "script") And InStr(1, LCase(Mid(strHTML, i, J - i + 1)), "a href") = False Then
             'If LCase(Mid(strHtml, i, J - i + 1)) = "<script>" Then
290             ii = InStr(J + 1, strHTML, "<")
300                 Do While ii <> 0
310                       ii = InStr(J + 1, strHTML, "<")
320                       If ii = 0 Then Exit Do
330                       jj = InStr(ii + 1, strHTML, ">")
340                       If jj <> 0 Then
                           ' Debug.Print "FINISH" & LCase(Mid(strHtml, ii, jj - ii + 1))
350                          If InStr(1, LCase(Mid(strHTML, ii, jj - ii + 1)), "/script") Then
                            'If LCase(Mid(strHtml, ii, jj - ii + 1)) = "</script>" Then
360                             strHTML = Left(strHTML, i - 1) & Mid(strHTML, jj + 1)
370                             GoTo Done
380                           Else
390                             strHTML = Left(strHTML, ii - 1) & Mid(strHTML, jj + 1)
400                          End If
410                         Else
420                          GoTo Done
430                       End If
440                 Loop
450           End If
             '##################################
  
             '###################################
             'CSS STRIPPER
            ' Debug.Print "CSSTART" & LCase(Mid(strHtml, i, J - i + 1))
460           If InStr(1, LCase(Mid(strHTML, i, J - i + 1)), "style") Then
470              For iIndex = 0 To UBound(css_Tags)
480                  strCssTag = Trim(css_Tags(iIndex))
490                  If InStr(1, LCase(Mid(strHTML, i, J - i + 1)), strCssTag) Then GoTo FalseCSS
500              Next
  
               'InStr(1, LCase(Mid(strHtml, i, J - i + 1)), "input") = False Then
               'If LCase(Mid(strHtml, i, J - i + 1)) = "<script>" Then
510             ii = InStr(J + 1, strHTML, "<")
520             Do While ii <> 0
530                ii = InStr(J + 1, strHTML, "<")
540                If ii = 0 Then Exit Do
550                jj = InStr(ii + 1, strHTML, ">")
560                If jj <> 0 Then
                    ' Debug.Print "CSSFINISH" & LCase(Mid(strHtml, ii, jj - ii + 1))
570                   If InStr(1, LCase(Mid(strHTML, ii, jj - ii + 1)), "/style") Then
                        'If LCase(Mid(strHtml, ii, jj - ii + 1)) = "</script>" Then
580                      strHTML = Left(strHTML, i - 1) & Mid(strHTML, jj + 1)
590                      GoTo Done
600                     Else
610                      strHTML = Left(strHTML, ii - 1) & Mid(strHTML, jj + 1)
620                   End If
630                  Else
640                   GoTo Done
650                 End If
660              Loop
670            End If
              '##################################
  
  
FalseCSS:
              'can replace other items here !!
680            strHTML = Left(strHTML, i - 1) & Mid(strHTML, J + 1)
690           Else
              'skip rest off strHtml
700            strHTML = Left(strHTML, i - 1)
710          End If
720        Else
730          i = i + 1 'So Next tag is searched
Done:
740       End If
750       i = InStr(i, strHTML, "<")
760     Loop
  
       'Strip Symbols etc ie: &#149;
770     i = InStr(1, strHTML, "&")
780     Do While i <> 0
790        If Mid(strHTML, i + 5, 1) = ";" Then
          'Debug.Print i
          'Debug.Print Mid(strHtml, i, 6)
800        strHTML = Left(strHTML, i - 1) & Mid(strHTML, i + 6)
810        i = i - 1
820        End If
830        i = i + 1
840        i = InStr(i, strHTML, "&")
850      DoEvents
860      Loop

870      If Len(strHTML) = 0 Then strHTML = Chr$(32)
880      StripTags = strHTML
890    On Error GoTo 0
900   Exit Function
LogError:
910   If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                         "{StripTags}", _
                         "{Public Function}", _
                         "{Format_Functions}", _
                         "{Format_Functions}") Then
920     End If
930   Resume Next ' Resume 'Exit Sub
End Function

Public Sub TableStart()
     Dim Inputval As String
10    On Error GoTo LogError
20    Inputval = 0 'InputBox("What width do you want the frame to be?", "Frame Width", "90")
30    Do Until IsNumeric(Inputval) = True          ' Check to see numeric values entered!
40        Inputval = InputBox("Please Enter A Numeric Value!", "User Input", "")
        ' Check the user entered a valid value
50        If IsNumeric(Inputval) = True Then Exit Do
60    Loop
70    strHTMLStart = strHTMLStart & "<center><table width= " & Inputval & "% bordercolor = '#1010BA' " & _
                      "cellspacing='0' cellpadding='20' border='2'><tr><td>"
80    On Error GoTo 0
90    Exit Sub
LogError:
100   If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                         "{TableStart}", _
                         "{Public Sub}", _
                         "{Format_Functions}", _
                         "{Format_Functions}") Then
110     End If
120   Resume Next ' Resume 'Exit Sub
End Sub

