Option Explicit
Sub SaveAsPPAM()
    Dim presentationPath As String
    Dim presentationName As String
    Dim savePath As String

    ' Get the full path of the current presentation
    presentationPath = ActivePresentation.FullName
    ' Get the current presentation name
    presentationName = ActivePresentation.Name
    ' Get the current path
    savePath = Left(presentationPath, InStrRev(presentationPath, "\"))

    ' Save as .ppam
    ActivePresentation.SaveAs savePath & Left(presentationName, InStrRev(presentationName, ".") - 1) & ".ppam", ppSaveAsAddIn
End Sub

