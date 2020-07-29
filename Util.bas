Attribute VB_Name = "Util"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub log(txt As String)
    frmMain.txtLog.Text = frmMain.txtLog.Text + txt + vbNewLine
    frmMain.txtLog.SelStart = Len(frmMain.txtLog.Text)
End Sub
