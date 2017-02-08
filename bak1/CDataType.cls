VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_clientField As Boolean
Private m_serverField As Boolean
Private m_typeWritor As String
Private m_name As String

Property Get GetClientField()
    GetClientField = m_clientField
End Property

Property Get GetServerField()
    GetServerField = m_serverField
End Property

Property Get GetTypeWritor()
    GetTypeWritor = m_typeWritor
End Property

Property Get GetName()
    GetName = m_name
End Property

Property Let LetExportClient(exportClient As Boolean)
    m_clientField = exportClient
End Property

Property Let LetExportServer(exportServer As Boolean)
    m_serverField = exportServer
End Property

Property Let LetTypeWritor(typeWritor As String)
    m_typeWritor = typeWritor
End Property

Property Let LetName(name As String)
    m_name = name
End Property
