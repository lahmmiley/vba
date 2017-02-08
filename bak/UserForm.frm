VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   ClientHeight    =   3000
   ClientLeft      =   -6400
   ClientTop       =   -4400
   ClientWidth     =   6000
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FResultLable As Label

Private Sub command_Click()
    ResultLabel.Caption = "开始生成"
    Call Main.Main(client.value, server.value)
End Sub
