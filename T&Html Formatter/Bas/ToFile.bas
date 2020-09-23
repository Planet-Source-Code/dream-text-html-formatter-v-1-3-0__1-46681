Attribute VB_Name = "ToFile"

Option Explicit

Public Enum eToFileLoadTextType
 Default = 0 'Load all
 VBDoc = 1     'Load Code
End Enum
'
'ToFile...
'05.2003 mrk Change/Add

Public Sub FileLoader()
    Dim Char As String
    Dim strPath As String
10    On Error GoTo LogError
20      strPath = ShowOpen(frmMDI.hwnd, _
    "All Files (*.*)" & Chr(0) & "*.*", "T&H Open File", False)
30      If strPath = vbNullString Then Exit Sub
    
40      If Reset = True Then
  
50         LockedWindow = True          'Tell RTB change to cease while loading file
60         Screen.MousePointer = 11
70         frmMDI.SB.Panels(1).Text = "Status: Loading File"
80         Char = Mid(strPath, Len(strPath) - 2)
     
    'Check filetype for .frm .bas & .cls
    'If VBDoc then tell ToFileLoad what kinda form!
90         Select Case Char
           'You can customize the filetypes and the load procedure in the ToFile.bas
           'to load just about any filetype you wish (text type files)
                  Case "frm", "bas", _
                      "cls": frmMDI.RTB.Text = ToFileLoad(strPath, _
                      VBDoc) 'if a VBDoc then filter i
     
100               Case Else: frmMDI.RTB.Text = ToFileLoad(strPath)   'or ToFileLoad(.FileName, Default)
110        End Select
     
120        ColorIn frmMDI.RTB
           frmMDI.RTB.Visible = True
           frmMDI.Tab1.Visible = True
130        frmMDI.Timer1.Enabled = True
140        LockedWindow = False
'150        Call PreviewHtml(False)
160        Screen.MousePointer = 1
170        frmMDI.SB.Panels(1).Text = "Status: File Loaded: "
180        frmMDI.SB.Panels(2).Text = strPath & "    "
190    End If
200   On Error GoTo 0
210   Exit Sub
LogError:
220      If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                                        "{FileLoader}", _
                                        "{Public Sub}", _
                                        "{ToFile}", _
                                        "{ToFile}") Then
230     End If
240     Resume Next ' Resume 'Exit Sub
End Sub

Public Sub FileSaver()
10      On Error GoTo LogError
  Dim strPath As String
  
20      strPath = ShowSave(frmMDI.hwnd, "All Files (*.*)" & Chr(0) & "*.*", "T&H Save Text/Html", False)
  
30      If strPath = vbNullString Then Exit Sub  'If this call returns a filename then..
    
40        Screen.MousePointer = 11
          frmMDI.RTB.SaveFile strPath, rtfText
50       'ToFileSave strPath, frmText.RTB.Text        'Call the save file function

60        frmMDI.SB.Panels(1).Text = "Status: File Saved: "
70        frmMDI.SB.Panels(2).Text = strPath & "          "
80        Screen.MousePointer = 1
90    On Error GoTo 0
100   Exit Sub
LogError:
110      If ErrorLogAndStop(Err.Number, Err.Description, Erl, Err.Source, _
                                       "{FileSaver}", _
                                       "{Public Sub}", _
                                       "{ToFile}", _
                                       "{ToFile}") Then
120     End If
130     Resume Next ' Resume 'Exit Sub
End Sub

Public Function ToFileLoad( _
              ByVal FileName As String, _
     Optional ByVal TextType As eToFileLoadTextType = Default) As String
  On Error Resume Next
  
  Dim FF As Integer
  Dim sText As String
  Dim fText As String
  Dim TextToAdd As Long
    
  TextToAdd = 0
  
  FF = ToFileFree
  
  Select Case TextType
    
    Case eToFileLoadTextType.Default
         Open FileName For Binary Access Read As #FF
         sText = Space$(LOF(FF))
         Get FF, , sText
    
    Case eToFileLoadTextType.VBDoc
         Open FileName For Input As #FF

         Do Until EOF(FF)
            Line Input #FF, fText
            
            Select Case TextToAdd
              Case 0
                   If InStr(1, Left$(fText, 10), "Attribute") = 1 Then
                      TextToAdd = 1
                   End If
              
              Case 1
                   If InStr(1, Left$(fText, 10), "Attribute") = 0 Then
                      TextToAdd = 2
                      sText = sText & fText & vbCrLf   'load each line into txtdata
                   End If
              
              Case 2
                   'read read read
                   sText = sText & fText & vbCrLf   'load each line into txtdata
            End Select
            
            DoEvents
         Loop
  End Select
  
  If FF > 0 Then
     Close #FF
  End If

  ToFileLoad = sText
  sText = ""
  fText = ""
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Function

Public Sub ToFileSave( _
              ByVal FileName As String, _
              ByVal Text As String)
  On Error Resume Next
  
  Dim FF As Integer
  
  FF = ToFileFree
  Open FileName For Output As #FF
  Print #FF, Text
  Close #FF
  
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Sub '05.2003 mrk Change/Add

Public Sub ToFileKill( _
              ByVal FileName As String)
  On Error Resume Next
  
  If ToFileIsExist(FileName) Then
     ToFileAttrSet FileName, vbNormal
     Kill FileName
  End If
  
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Sub '05.2003 mrk Change/Add

Public Function ToFileIsExist( _
              ByVal FileName As String) As Boolean
  On Error Resume Next
  
  Dim Value As Boolean
  Value = CBool(Len(Dir$(FileName, vbArchive Or vbHidden Or vbNormal Or vbReadOnly)) > 0)
  
  If Err.Number <> 0 Then
     Err.Clear
     Value = False
  End If
  
  ToFileIsExist = Value
  
  On Error GoTo 0
End Function '05.2003 mrk Change/Add

Private Sub ToFileAttrSet( _
              ByVal FileName As String, _
              ByVal Attri As VbFileAttribute)
  On Error Resume Next
  
  If ToFileIsExist(FileName) Then
     SetAttr FileName, Attri
  End If
  
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Sub '05.2003 mrk Change/Add

Private Function ToFileAttrGet( _
              ByVal FileName As String) As VbFileAttribute
  On Error Resume Next
  
  If ToFileIsExist(FileName) Then
     ToFileAttrGet = GetAttr(FileName)
  End If
  
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Function '05.2003 mrk Change/Add


Private Function ToFileFree() As Integer
  On Error Resume Next
  
  ToFileFree = FreeFile
  
  If Err.Number <> 0 Then
     '--> ToFileFree = 0
     Err.Clear
  End If
  
  On Error GoTo 0
End Function '05.2003 mrk Change/Add
