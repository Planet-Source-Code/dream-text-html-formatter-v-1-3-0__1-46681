Attribute VB_Name = "ToDlg"

Option Explicit

Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As TCHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
                                   (ByVal hwnd As Long, ByVal szApp As String, _
                                    ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type TCHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Property Get GetColor(Optional AppHwnd As Long = 0, Optional Multiple As Boolean) As Long
Dim Color As TCHOOSECOLOR
Color.flags = 2 'Show full box includes custom
If Multiple = False Then 'If they only want one opened at a time give it an hwnd
    Color.hwndOwner = AppHwnd 'when skipped program will keep focus
End If
Color.lStructSize = Len(Color)
Color.lpCustColors = 100
If CHOOSECOLOR(Color) <> 0 Then
    GetColor = Color.rgbResult 'Set GetColor to the long value despite it saying rgb
Else
    GetColor = vbNullString 'Not 0 because then it would be though black not an error
End If
End Property



'ToDlg...
'05.2003 mrk Change/Add

Public Function ShowOpen( _
                AppHwnd As Long, _
                Filter As String, _
                Title As String, _
                Optional Multiple As Boolean) As String
On Error GoTo ErrorLoc
Dim OpenF As OPENFILENAME
OpenF.flags = &H4 ' no open as readonly box
If Multiple = True Then
    OpenF.hwndOwner = AppHwnd 'set the window handle
End If
OpenF.lpstrFile = String(500, Chr(0))
OpenF.lpstrFileTitle = String(500, Chr(0))
OpenF.lpstrFilter = Filter
OpenF.lpstrTitle = Title
OpenF.lStructSize = Len(OpenF)
OpenF.nMaxFile = 501
OpenF.nMaxFileTitle = 501
If GetOpenFileName(OpenF) Then
    ShowOpen = Replace(OpenF.lpstrFile, Chr(0), vbNullString)
Else
ErrorLoc:
    ShowOpen = vbNullString 'No file error
End If
End Function

Public Function ShowSave(AppHwnd As Long, _
                Filter As String, _
                Title As String, _
                Optional Multiple As Boolean) As String
                
On Error GoTo ErrorLoc
Dim SaveF As OPENFILENAME

SaveF.flags = &H2 Or &H4 'Prompt on overwrite, no read only box

If Multiple = False Then
    SaveF.hwndOwner = AppHwnd
End If

SaveF.lpstrFile = String(500, Chr(0))
SaveF.lpstrFileTitle = String(500, Chr(0))
SaveF.lpstrFilter = Filter
SaveF.lpstrTitle = Title
SaveF.lStructSize = Len(SaveF)
SaveF.nMaxFile = 501
SaveF.nMaxFileTitle = 501

If GetSaveFileName(SaveF) Then
    ShowSave = Replace(SaveF.lpstrFile, Chr(0), vbNullString)
Else

ErrorLoc:
    ShowSave = vbNullString 'vbNullString 'Error no file
End If
End Function

Public Sub ToDlgAbout( _
              ByVal Form As Form, _
              ByVal App As String, _
              ByVal OtherStuff As String, _
              ByVal Icon)
  On Error Resume Next
  
  ShellAbout Form.hwnd, _
             App, _
             OtherStuff, _
             CLng(Icon)
  
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Sub '05.2003 mrk Change/Add
