VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Compare Text

'Credits To M.R.K. Thanks Martin!

Private propColl As Collection

Private Sub Class_Initialize()
  Set propColl = New Collection
End Sub
Private Sub Class_Terminate()
  Clear
  Set propColl = Nothing
End Sub

Public Function Add(ParamArray Items() As Variant) As Boolean
  On Error Resume Next
  
  Dim l As Long
  Dim u As Long
  Dim i As Long
  
  Add = True
  
  l = LBound(Items)
  u = UBound(Items)
  
  With propColl
    For i = l To u
        .Add Items(i)
    Next
  End With
  
  If Err.Number <> 0 Then
     Err.Clear
     Add = False
  End If
  
  On Error GoTo 0
End Function

Public Function Count() As Long
  Count = propColl.Count
End Function

Public Function Item( _
              ByVal Index As Long) As Variant
  If IndexOK(Index) Then
     Item = propColl.Item(Index)
  End If
End Function

Public Function Remove( _
              ByVal Index As Long) As Boolean
  If IndexOK(Index) Then
     propColl.Remove Index
     Remove = False
  End If
End Function

Public Sub Clear()
  Dim i As Long
  
  With propColl
    For i = 1 To .Count
       .Remove 1
    Next i
  End With
End Sub

Public Function Find( _
              ByVal Value As Variant) As Long
'Item    .Item   =
'abc123  abc123  n
'abc123          -1
  
  Dim u As Long
  Dim i As Long
  
  Find = -1
  
  With propColl
    u = .Count
    For i = 1 To u
        If .Item(i) = Value Then
           Find = i
    Exit For
        End If
    Next
  End With
End Function

Public Function FindPartStr( _
              ByVal Text As String, _
     Optional ByVal AtStart As Boolean = True) As Long
'CStr comparison

'Item AtStart .Item  =
'abc  True    abc123 n
'd    True           -1
'abc  False   123abc n
'd    False          -1

  Dim u As Long
  Dim i As Long
  Dim l As Long
  
  FindPartStr = -1
  l = Len(Text)
  
  With propColl
    u = .Count
  
    If AtStart Then
       For i = 1 To u
           If Left$(CStr(.Item(i)), l) = Text Then
              FindPartStr = i
       Exit For
           End If
       Next
      
      Else
       For i = 1 To u
           If Right$(CStr(.Item(i)), l) = Text Then
              FindPartStr = i
       Exit For
           End If
       Next
    End If
  End With
End Function

Public Sub Sort( _
              Optional ByVal OrderA2Z As Boolean = True)
   Dim u As Long, _
       i As Long, _
       J As Long, _
       n As Long, _
       v As Variant
   
   With propColl
     n = propColl.Count
  
     u = CLng(VBA.Log(n) / VBA.Log(2))
     u = 2 ^ u - 1
     
     If OrderA2Z Then
        Do While u > 0
           For i = 1 To n - u
               For J = i To 1 Step -u
                   If .Item(J) < .Item(J + u) Then
               Exit For
                   End If
                   v = .Item(J)
                   .Add .Item(J + u), , , J
                   .Remove J
                   .Add v, , , J + u
                   .Remove J + u
               Next J
           Next i
           u = (u / 2)
        Loop
       
       Else
        Do While u > 0
           For i = 1 To n - u
               For J = i To 1 Step -u
                   If .Item(J) > .Item(J + u) Then
               Exit For
                   End If
                   v = .Item(J)
                   .Add .Item(J + u), , , J
                   .Remove J
                   .Add v, , , J + u
                   .Remove J + u
               Next J
           Next i
           u = (u / 2)
        Loop
     End If
  End With
End Sub

'------------------------------------------------------------------------------

Private Function IndexOK( _
              ByVal Index As Long) As Boolean
  On Error Resume Next
  
  IndexOK = CBool(Index > 0 And Index <= propColl.Count)
  
  If Err.Number <> 0 Then
     Err.Clear
  End If
  
  On Error GoTo 0
End Function

