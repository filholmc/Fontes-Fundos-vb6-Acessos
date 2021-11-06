Attribute VB_Name = "moduErrors"
Option Explicit
Public Function rgfMsgBox(ByVal vstrTxtMsg As String, ByVal vmsgTipMsg As NumMsg, _
                                             Optional ByVal vbytNumAju As Byte = 0) As Integer
       Select Case vmsgTipMsg
              Case MsgErr
               If (vbytNumAju = 0) Or (Not gbooAjuHab) Then
                   MsgBox vstrTxtMsg, vbCritical, App.Title
               Else
                   MsgBox vstrTxtMsg, vbCritical + vbMsgBoxHelpButton, App.Title, gstrPthAju, vbytNumAju
               End If
              Case MsgInf
                   MsgBox vstrTxtMsg, vbInformation, App.Title
              Case MsgNao
                rgfMsgBox = _
                   MsgBox(vstrTxtMsg, vbQuestion + vbYesNo + vbDefaultButton2, App.Title)
       End Select
End Function
Public Sub rgsSetarFoco(vforFormes As Form, ByVal vstrSource As String)
       Dim lintQtdFor As Integer

       For lintQtdFor = 0 To vforFormes.Controls.Count - 1
       If (InStr(1, vforFormes.Controls.Item(lintQtdFor).Name, vstrSource, vbTextCompare) <> 0) Then
           vforFormes.Controls.Item(lintQtdFor).SetFocus
           Exit For
       End If
       Next
End Sub
Public Sub rgsTratarErro(ByVal vobjErrors As Object, ByVal verrErrors As Errors, vforFormes As Form)
       Dim lerrErrors As Error

       If (vobjErrors.Number <> 0) Then
           rgfMsgBox vobjErrors.Description, MsgErr
           rgsSetarFoco vforFormes, vobjErrors.Source
       End If

       If (Not (verrErrors Is Nothing)) Then
       For Each lerrErrors In verrErrors
           rgfMsgBox lerrErrors.Description, MsgErr
       Next
       End If
End Sub
