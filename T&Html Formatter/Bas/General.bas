Attribute VB_Name = "General"

Option Explicit
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI)
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
'Printer Shit
Public Declare Function SetTextAlign Lib "gdi32.dll" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
                                   (ByVal hwnd As Long, ByVal wMsg As Long, _
                                    ByVal wParam As Long, lParam As Any) As Long
Public Const TA_CENTER = 4
 
'' Types used for fonts & images
Public Type PointAPI
    X As Long
    Y As Long
End Type
 
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Enum StartWindowState
    START_HIDDEN = 0
    START_NORMAL = 4
    START_MINIMIZED = 2
    START_MAXIMIZED = 3
End Enum

 Public Const SW_SHOWNORMAL = 1
'For speed we exit the loop when a match is found so for best
'results we put the most common filetypes first!
'Const FILE_TYPE = "Txt,Html,Exe,Zip,Bas,Ctl,Frm,Bmp,Gif,Jpg,Log,Dll,Ocx,Ini,Dat" 'For imagelist index
 Const FILE_TYPE = "Exe,Txt,Zip,Html,Log,Ini,Dat,Jpg,Gif,Bmp,Dll,Ocx,Frm,Bas,Ctl"
 Const FILTER_HTML = "Html,Txt,Vtml"
 Const FILTER_TXT = "Txt,Ini,Log"
 Const FILTER_VB = "Bas,Cls,Frm,Vbp,Ctl"

#Const conDebug = "stop on Error"

#If conDebug = "stop on Error" Then
    Dim bStopErrorMsg As Boolean
#End If
Public bOpen As Integer
 
Public TTT As String
Public OpenGate As Boolean

Public Function ErrorLogAndStop(ErrNum As Long, _
    ErrDescription As String, _
    ErrLine As Long, _
    ErrSource As String, _
    sProcName As String, _
    sProcType As String, _
    sModuleName As String, _
    sModuleFileName As String) As Boolean
 '======================================
 ' ERROR LOG RooterProcedure'
 '=======================================
        Dim s As String
        Dim iMsg As Integer ' Dim ifileNum As Integer
        Dim S1 As String
    
10      On Error GoTo ErrorH
    
        ' 1) Rooting Actions
        ' Select Case Err.Number
        ' better I think is to set here the Set
        '     objcls for User Defined Error Objects in
        '     the App
    
        '  Handler Of Unhandled
        ' (A) Create Error Message String
20      s = s & "Message Of Unhandled" & vbCrLf & _
        "***************" & vbCrLf & _
        CStr(Now) & vbCrLf & _
        "***************" & vbCrLf & _
        "ErrNumber = " & CStr(ErrNum) & vbCrLf & vbCrLf & _
        "ErrDescription = " & ErrDescription & vbCrLf & vbCrLf & _
        "At Error Line No: " & CStr(ErrLine) & vbCrLf & vbCrLf & _
        "In: " & sProcName & vbCrLf & vbCrLf & _
        "Procedure Type: " & sProcType & vbCrLf & vbCrLf & _
        "ModuleName: " & sModuleName & vbCrLf & vbCrLf & _
        "ErrorSource =" & ErrSource & vbCrLf & _
        "ModuleFileName =" & sModuleFileName & vbCrLf & vbCrLf & _
        "ExePath = " & App.Path & "\" & App.EXEName & " (AppRevision = " & App.Revision & ")" & vbCrLf & _
        vbCrLf & vbCrLf & vbCrLf
    
        ' (B) MsgBox Only in Debug Mode


  #If conDebug = "stop on Error" Then
  
30        If Not bStopErrorMsg Then 'true and Nothing Happens (only in DebugMode)
      
40          S1 = "------------------" & vbCrLf & vbCrLf & _
            "Do you want To Break ?" & vbCrLf & vbCrLf & _
            "Press YES To BREAK and immediately access the raisingErrorProcedure or " & vbCrLf & _
            "NO To Continue execution and Log to ErrorFile" & vbCrLf & _
            "CANCEL To stop ErrorMessaging"
      
50          iMsg = MsgBox(s & S1, vbCritical + vbMsgBoxSetForeground + vbYesNoCancel, "Oh No Another Error In " & App.Title)
            '+ vbSystemModal
      
60          Select Case iMsg
              Case vbYes
70              ErrorLogAndStop = True ' For Break
80              Exit Function ' without fileLog, comment the line To log also to the file
90            Case vbCancel
100             iMsg = MsgBox("Turn Off Error Checking in the current AppRunning Period ?", _
                vbYesNo + vbQuestion, "No More Error Messages In " & App.Title & " For the time being ?")
110             If iMsg = vbYes Then bStopErrorMsg = True
               'the error goes to the file
               'Case vbNo
               'and the code continues & log to error
               'file
120         End Select
130       End If
     #End If

      ' If user answers No Then log the Error
      '     to File
140   S1 = App.Path & "\" & "ErrorLog.log" ' If Not Exists Then Create
150   iMsg = FreeFile(1)

160   Open S1 For Append As iMsg

170   Print #iMsg, s

180   Close iMsg 'ifileNum = 0
190   Exit Function
ErrorH:
200   MsgBox "Error LogFile Process Not Successful, cause of " & CStr(Err.Number) & " " & Err.Description
210   Stop 'Debug.Assert Err.Number = 0
      'Resume Next ' Resume 'Exit Sub
End Function

Public Sub ExampleLoad()
10       On Error GoTo LogError
   Dim ansa As String
  
20       If frmMDI.RTB.Text <> "" Then
30    ansa = MsgBox("This will clear the Text box of current Text, Proceed?", vbYesNo, "Show Example")
40    If ansa = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
50       End If
        
60       Screen.MousePointer = vbHourglass
70       frmMDI.mnuFormat.Enabled = True
90       LockedWindow = True
100      frmMDI.RTB.Text = ToFileLoad(ToAppPath & "Example.txt")
110      ColorIn frmMDI.RTB
120      LockedWindow = False

'130      Call PreviewHtml(False) 'False is No Frame Wrap
140      frmMDI.SB.Panels(2).Text = ToAppPath & "Example.txt           "
150      frmMDI.Timer1.Enabled = True
160      Screen.MousePointer = vbDefault
         frmMDI.Tab1.TabIndex = 0
         frmMDI.Tab1.Visible = True
         frmMDI.RTB.Visible = True
170   On Error GoTo 0
180   Exit Sub
LogError:
190      If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                         "{ExampleLoad}", _
                         "{Public Sub}", _
                         "{General}", _
                         "{General}") Then
200     End If
210     Resume Next ' Resume 'Exit Sub
End Sub

Public Function GetSysDir() As String
  Dim strSyspath As String
  strSyspath = String(145, Chr(0))
  strSyspath = Left(strSyspath, GetSystemDirectory(strSyspath, 145))
  GetSysDir = strSyspath
End Function

Public Sub OpenLink(Frm As Form, URL As String)
  ShellExecute Frm.hwnd, "Open", URL, "", "", 1
End Sub

Public Sub PrintOut()              'Print Out Text or Html
   Dim i As Integer
   Dim ta As Long
   Dim TextLines As Long
   Dim TextBuff As String
   Dim CharRet As Long
10       On Error GoTo LogError
20    Screen.MousePointer = 11
30    frmMDI.SB.Panels(1).Text = "Status: Printing"               '## Print
40    Printer.Print " "
50    Printer.Print , , "Text To Html Formatter"
60    Printer.Print " "
70    Printer.Print , , Now
  
80    With frmMDI.RTB
90        ta = SetTextAlign(Printer.hDC, TA_CENTER)            'Center text on printer object
100       Printer.CurrentY = (Printer.ScaleHeight / .Parent.ScaleHeight) * .Top
110       TextLines = SendMessage(.hwnd, &HBA, 0, 0)    'Get number of lines in text box
120       For i = 0 To TextLines - 1                           'Extract & print each line in TextBox
130           TextBuff = Space(1000)
140           Printer.CurrentX = (Printer.ScaleWidth / .Parent.ScaleWidth) * (.Left + (.Width / 2))
150           Mid(TextBuff, 1, 1) = Chr(79 And &HFF)             'Setup buffer for the line!
160           Mid(TextBuff, 2, 1) = Chr(79 \ &H100)
170           CharRet = SendMessage(.hwnd, &HC4, i, ByVal TextBuff)
180           Printer.Print Left(TextBuff, CharRet)
190       Next i
200   End With

210   ta = SetTextAlign(Printer.hDC, ta)        'Reset alignment back to original setting
220   Printer.EndDoc
230   Screen.MousePointer = 1
240   frmMDI.SB.Panels(1).Text = "Status: Idle"
250   frmMDI.SB.Panels(2).Text = "Finished Sending Data To Printer"
260   On Error GoTo 0
270   Exit Sub
LogError:
280      Screen.MousePointer = 1
290      frmMDI.SB.Panels(1).Text = "Status: Error Printing"
300      frmMDI.SB.Panels(2).Text = Err.Number & "   " & Err.Description & "        "
310      If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                         "{Printout}", _
                         "{Public Sub}", _
                         "{General}", _
                         "{General}") Then
320     End If
330     Resume Next ' Resume 'Exit Sub
End Sub

Public Sub ReList_Files()
     Dim Char As String
     Dim i As Integer
     Dim bb As Integer
     Dim arr2() As String
     Dim pic As String
     Dim Filter_IN As Boolean
     Dim FCount As Integer
10    On Error GoTo ErrControl
20    Screen.MousePointer = 11 'vbHourGlass
30    With frmMDI
40        If .File1.ListCount > 0 Then
50            Select Case .Combo1.ListIndex
                    Case 0: arr2() = Split(FILE_TYPE, ","): Filter_IN = True
60                  Case 1: arr2() = Split(FILTER_HTML, ",")
70                  Case 2: arr2() = Split(FILTER_TXT, ",")
80                  Case 3: arr2() = Split(FILTER_VB, ",")
90            End Select
    
100          .lvCode.ListItems.Clear
110           LockWindowUpdate .lvCode.hwnd
120           For i = 0 To .File1.ListCount - 1
130
                'Get file extension
140              Char = Right(.File1.List(i), 3)
    
                'Filter out unwanted file types &
                'Set the image for the file
150              For bb = 0 To UBound(arr2)
160                  If InStr(1, UCase(arr2(bb)), UCase(Char)) Then
170                     pic = arr2(bb)
180                     Filter_IN = True
190                     Exit For
200                  End If
210              Next

220              Select Case .Combo1.ListIndex
                        Case 0: Filter_IN = True
230               Case Else:  'Blah
240              End Select

250              If Filter_IN = True Then
                   'Else we give it a default image
260                 Select Case pic
                           Case "Ocx", vbNullString, "Cls", "Vbp": pic = "Item"
270                        Case "Dat": pic = "Log"
280                        Case "Frm": pic = "Exe"
290                        Case "Vtml": pic = "Html"
300                 End Select
310                 FCount = FCount + 1
320                .lvCode.ListItems.Add FCount, m_CurrentDirectory & "\" & .File1.List(i), .File1.List(i), pic, pic
330                 pic = vbNullString
340              End If
350              Filter_IN = False
360              DoEvents
370           Next
380       End If
390      .SB.Panels(3).Text = FCount & " Files"
400      .SB.Panels(2).Text = .File1.Path & "           "
410       LockWindowUpdate 0&
420       Screen.MousePointer = 1 'vbDefault
430      .Timer1 = True
440   End With
450        On Error GoTo 0
460   Exit Sub
ErrControl:
470      If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                         "{ReList_Files}", _
                         "{Public Sub}", _
                         "{General}", _
                         "{General}") Then
480     End If
490     Resume Next ' Resume 'Exit Sub
End Sub

Public Function Reset() As Boolean
10      On Error GoTo LogError
20      If frmMDI.RTB.Text <> "" Then
    Dim ansa As String                  'If text present then ask to save?
30        ansa = MsgBox("You will lose Unsaved Information, Continue?", vbYesNo, "Save Before Continuing")
40        If ansa = vbNo Then
50        GoTo EndMe
60        End If
70      End If
 'Now its safe to clear the form
80      Wrapped = False  'After wrapping once shut down function(this resets it)
'  cmdFormat.Enabled = True
'  cmdUndo.Enabled = False
90      frmMDI.mnuFormat.Enabled = True

110     frmMDI.RTB.Text = ""
120     frmMDI.SB.Panels(1).Text = "Status: Reset"
130     frmMDI.SB.Panels(2).Text = "       "
150     LockedWindow = False
160   Reset = True
170   Exit Function
EndMe:
180   Reset = False
190   On Error GoTo 0
200   Exit Function
LogError:
210      If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                         "{Reset}", _
                         "{Public Function}", _
                         "{General}", _
                         "{General}") Then
220     End If
230     Resume Next ' Resume 'Exit Sub
End Function

Public Function ToAppPath() As String                   'Check "\" on end of filename
  Dim Text As String
  Text = App.Path
  If Right$(Text, 1) <> "\" Then  'if doesnt end with "\"
     ToAppPath = Text & "\"       'Add one onto the end, why?...
    Else                          'so our path looks like C:\Temp\MyApp.exe
     ToAppPath = Text             'and not look like      C:\TempMyApp.exe    when you...
  End If                          'Add a filename onto the end of it
End Function '05.2003 mrk Change/Add


