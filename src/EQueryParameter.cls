VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EQueryParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_name As String
Private m_value As Variant

' Constructor
Public Sub Initialize(ByVal paramName As String, ByVal paramValue As Variant)
    m_name = paramName
    m_value = paramValue
End Sub

' Propiedades
Public Property Get ParameterName() As String
    ParameterName = m_name
End Property

Public Property Let ParameterName(ByVal value As String)
    m_name = value
End Property

Public Property Get ParameterValue() As Variant
    If IsObject(m_value) Then
        Set ParameterValue = m_value
    Else
        ParameterValue = m_value
    End If
End Property

Public Property Let ParameterValue(ByVal value As Variant)
    m_value = value
End Property

Public Property Set ParameterValue(ByVal value As Variant)
    Set m_value = value
End Property
