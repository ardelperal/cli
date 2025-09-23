VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EMapeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_idMapeo As Long
Private m_nombrePlantilla As String
Private m_nombreCampoTabla As String
Private m_valorAsociado As String
Private m_nombreCampoWord As String

' Propiedades p?blicas
Public Property Get IdMapeo() As Long
    IdMapeo = m_idMapeo
End Property

Public Property Let IdMapeo(ByVal value As Long)
    m_idMapeo = value
End Property

Public Property Get NombrePlantilla() As String
    NombrePlantilla = m_nombrePlantilla
End Property

Public Property Let NombrePlantilla(ByVal value As String)
    m_nombrePlantilla = value
End Property

Public Property Get NombreCampoTabla() As String
    NombreCampoTabla = m_nombreCampoTabla
End Property

Public Property Let NombreCampoTabla(ByVal value As String)
    m_nombreCampoTabla = value
End Property

Public Property Get ValorAsociado() As String
    ValorAsociado = m_valorAsociado
End Property

Public Property Let ValorAsociado(ByVal value As String)
    m_valorAsociado = value
End Property

Public Property Get NombreCampoWord() As String
    NombreCampoWord = m_nombreCampoWord
End Property

Public Property Let NombreCampoWord(ByVal value As String)
    m_nombreCampoWord = value
End Property
