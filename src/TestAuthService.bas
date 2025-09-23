Attribute VB_Name = "TestAuthService"
Option Compare Database
Option Explicit


Public Function TestAuthServiceRunAll() As CTestSuiteResult
    Set TestAuthServiceRunAll = New CTestSuiteResult
    TestAuthServiceRunAll.Initialize "TestAuthService"
    TestAuthServiceRunAll.AddResult TestGetUserRole_Admin_ReturnsAdmin()
End Function

Private Function TestGetUserRole_Admin_ReturnsAdmin() As CTestResult
    Set TestGetUserRole_Admin_ReturnsAdmin = New CTestResult
    TestGetUserRole_Admin_ReturnsAdmin.Initialize "GetUserRole con datos de Admin devuelve RolAdmin"
    
    Dim serviceImpl As CAuthService
    Dim mockRepo As CMockAuthRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockError As CMockErrorHandlerService
    Dim mockConfig As CMockConfig
    Dim service As IAuthService
    
    On Error GoTo TestFail
    
    ' Arrange
    Set mockRepo = New CMockAuthRepository
    Set mockLogger = New CMockOperationLogger
    Set mockError = New CMockErrorHandlerService
    Set mockConfig = New CMockConfig
    
    Dim authData As New EAuthData
    authData.UserExists = True
    authData.IsGlobalAdmin = True
    mockRepo.ConfigureGetUserAuthData authData
    
    Set serviceImpl = New CAuthService
    serviceImpl.Initialize mockConfig, mockLogger, mockRepo, mockError
    Set service = serviceImpl
    
    ' Act
    Dim resultRole As UserRole
    resultRole = service.GetUserRole("admin@test.com")
    
    ' Assert
    AssertEquals RolAdmin, resultRole, "El rol devuelto debería ser Admin."
    
    TestGetUserRole_Admin_ReturnsAdmin.Pass
Cleanup:
    Exit Function
TestFail:
    TestGetUserRole_Admin_ReturnsAdmin.Fail "Error: " & Err.Description
    Resume Cleanup
End Function
