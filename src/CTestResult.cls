VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Name As String
Public Passed As Boolean
Public ErrorMessage As String

Public Sub Initialize(ByVal testName As String)
    Name = testName
    Passed = False
    ErrorMessage = ""
End Sub

Public Sub Pass()
    Passed = True
    ErrorMessage = ""
End Sub

Public Sub Fail(Optional ByVal message As String = "Test failed without a specific message.")
    Passed = False
    ErrorMessage = message
End Sub
