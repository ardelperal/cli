Attribute VB_Name = "TestAppManager"
Option Compare Database
Option Explicit


Public Function TestAppManagerRunAll() As CTestSuiteResult
    Set TestAppManagerRunAll = New CTestSuiteResult
    TestAppManagerRunAll.Initialize "TestAppManager"
    TestAppManagerRunAll.AddResult TestStartApplication_AdminUser_Success()
End Function

Private Function TestStartApplication_AdminUser_Success() As CTestResult
    Set TestStartApplication_AdminUser_Success = New CTestResult
    TestStartApplication_AdminUser_Success.Initialize "StartApplication con Admin tiene éxito"
    
    Dim appManagerImpl As CAppManager
    Dim mockAuth As CMockAuthService
    Dim mockConfig As CMockConfig
    Dim mockError As CMockErrorHandlerService
    Dim appManager As IAppManager
    
    On Error GoTo TestFail
    
    ' Arrange
    Set mockAuth = New CMockAuthService
    Set mockConfig = New CMockConfig
    Set mockError = New CMockErrorHandlerService
    
    mockConfig.SetSetting "IsInitialized", True
    mockAuth.ConfigureGetUserRole RolAdmin
    
    Set appManagerImpl = New CAppManager
    appManagerImpl.Initialize mockAuth, mockConfig, mockError
    Set appManager = appManagerImpl
    
    ' Act
    Dim success As Boolean
    success = appManager.StartApplication("admin@test.com")
    
    ' Assert
    modAssert.AssertTrue success, "La aplicación debería iniciarse."
    AssertEquals RolAdmin, appManager.GetCurrentUserRole(), "El rol debería ser Admin."
    
    TestStartApplication_AdminUser_Success.Pass
Cleanup:
    Exit Function
TestFail:
    TestStartApplication_AdminUser_Success.Fail "Error: " & Err.Description
    Resume Cleanup
End Function
