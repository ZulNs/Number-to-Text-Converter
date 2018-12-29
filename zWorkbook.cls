VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const cPass As Double = 7895123.19701105

Private Sub Workbook_Activate()
    Dim Wbk As Workbook
    If Workbooks.Count = 0 Then Workbooks.Add
    For Each Wbk In Workbooks
        If Wbk.Name <> ThisWorkbook.Name Then
            Workbooks(Wbk.Name).Activate
            Exit For
        End If
    Next
    ThisWorkbook.IsAddin = True
    Application.Run "zmMain.RestoreOrgProperty", cPass
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.Run "zmMain.RestoreOrgProperty", cPass
    Application.OnKey "^%+R"
    Application.OnKey "^%+V"
    Application.OnKey "^%+K"
    Application.OnKey "^%+C"
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Workbook_Open()
    Application.Run "zmMain.WbkOpen", cPass
End Sub
