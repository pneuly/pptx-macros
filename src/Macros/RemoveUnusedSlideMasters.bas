Option Explicit
Sub RemoveUnusedSlideMasters()
  Dim dsn As Design
  Dim i As Long
  On Error GoTo ERR_HNDL

  For Each dsn In ActivePresentation.Designs
    With dsn.SlideMaster.CustomLayouts
      'delete unused layout
      For i = .Count To 1 Step -1
        .Item(i).Delete
      Next i
      'Delete unused slide master
      If .Count = 0 Then
        dsn.SlideMaster.Delete
      End If
    End With
  Next dsn
  Exit Sub

  ERR_HNDL:
  Select Case Err.Number
    Case -2147188160
      'Debug.Print Err.Description
      Err.Clear
      Resume Next
    Case Else
      MsgBox Err.Number
      Exit Sub
    End Select
End Sub

