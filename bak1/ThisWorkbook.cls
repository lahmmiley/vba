VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub Workbook_AddinInstall()
    With Application.CommandBars.Add(name:="myCmdbar")
        .Position = msoBarTop
        With .Controls.Add
        .Caption = "…˙≥…≈‰÷√±Ì"
        .OnAction = "Main.ShowUserForm1"
        End With
        .Visible = True
    End With
End Sub

Sub Workbook_AddinUninstall()
    Application.CommandBars("myCmdbar").Delete
End Sub

