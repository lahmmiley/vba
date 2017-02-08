VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public ExportPath As String
Public ClientExport As Boolean
Public ServerExport As Boolean
Public fileContent As String

Public Sub Init(pExportPath As String, pClientExport As Boolean, pServerExport As Boolean)
    ExportPath = pExportPath
    ClientExport = pClientExport
    ServerExport = pServerExport
End Sub
