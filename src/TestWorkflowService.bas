Attribute VB_Name = "TestWorkflowService"
Option Compare Database
Option Explicit



Public Function TestWorkflowServiceRunAll() As CTestSuiteResult
    Set TestWorkflowServiceRunAll = New CTestSuiteResult
    TestWorkflowServiceRunAll.Initialize "TestWorkflowService"
    TestWorkflowServiceRunAll.AddResult TestValidateTransition_ValidCase()
    TestWorkflowServiceRunAll.AddResult TestValidateTransition_InvalidFromFinalState()
    TestWorkflowServiceRunAll.AddResult TestGetNextStates_ValidState()
    TestWorkflowServiceRunAll.AddResult TestIsEstadoFinal_AprobadaState()
    TestWorkflowServiceRunAll.AddResult TestValidateTransition_CalidadCanOverrideRepository()
    TestWorkflowServiceRunAll.AddResult TestValidateTransition_TecnicoIsRestrictedByRepository()
End Function

Private Function TestValidateTransition_ValidCase() As CTestResult
    Set TestValidateTransition_ValidCase = New CTestResult
    TestValidateTransition_ValidCase.Initialize "ValidateTransition con mock válido debe pasar"
    
    Dim serviceImpl As CWorkflowService
    Dim mockRepo As CMockWorkflowRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim service As IWorkflowService
    
    On Error GoTo TestFail
    
    ' Arrange
    Set mockRepo = New CMockWorkflowRepository
    mockRepo.ConfigureIsValidTransition True
    
    Set mockLogger = New CMockOperationLogger
    Set mockErrorHandler = New CMockErrorHandlerService
    
    Set serviceImpl = New CWorkflowService
    serviceImpl.Initialize mockRepo, mockLogger, mockErrorHandler
    Set service = serviceImpl
    
    ' Act
    Dim result As Boolean
    result = service.ValidateTransition(1, "Registrado", "Desarrollo", "", "Calidad")
    
    ' Assert
    modAssert.AssertTrue result, "La transición Registrado -> Desarrollo debería ser válida."
    
    TestValidateTransition_ValidCase.Pass
Cleanup:
    Exit Function
TestFail:
    TestValidateTransition_ValidCase.Fail "Error: " & Err.Description
    Resume Cleanup
End Function

Private Function TestValidateTransition_InvalidFromFinalState() As CTestResult
    Set TestValidateTransition_InvalidFromFinalState = New CTestResult
    TestValidateTransition_InvalidFromFinalState.Initialize "ValidateTransition desde estado final debe fallar para CUALQUIER rol"
    
    Dim serviceImpl As CWorkflowService
    Dim mockRepo As CMockWorkflowRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim service As IWorkflowService
    
    On Error GoTo TestFail
    
    ' Arrange
    ' No es necesario configurar el mockRepo, ya que la lógica del servicio
    ' debe cortar la ejecución antes de llegar a consultarlo.
    Set mockRepo = New CMockWorkflowRepository
    Set mockLogger = New CMockOperationLogger
    Set mockErrorHandler = New CMockErrorHandlerService
    
    Set serviceImpl = New CWorkflowService
    serviceImpl.Initialize mockRepo, mockLogger, mockErrorHandler
    Set service = serviceImpl
    
    ' Act
    Dim result As Boolean
    ' Se utiliza el rol "Calidad" a propósito para verificar que la regla del estado final
    ' tiene prioridad sobre la regla del rol privilegiado.
    result = service.ValidateTransition(1, "Aprobada", "Desarrollo", "", "Calidad")
    
    ' Assert
    modAssert.AssertFalse result, "La transición desde 'Aprobada' no debería ser válida, incluso para un rol privilegiado."
    
    TestValidateTransition_InvalidFromFinalState.Pass
Cleanup:
    Exit Function
TestFail:
    TestValidateTransition_InvalidFromFinalState.Fail "Error: " & Err.Description
    Resume Cleanup
End Function

Private Function TestGetNextStates_ValidState() As CTestResult
    Set TestGetNextStates_ValidState = New CTestResult
    TestGetNextStates_ValidState.Initialize "GetNextStates debe retornar estados válidos"
    
    Dim serviceImpl As CWorkflowService
    Dim mockRepo As CMockWorkflowRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim service As IWorkflowService
    Dim mockDict As Object
    
    On Error GoTo TestFail
    
    ' Arrange
    Set mockDict = New Scripting.Dictionary
    mockDict.Add 2, "Desarrollo"
    
    Set mockRepo = New CMockWorkflowRepository
    mockRepo.ConfigureGetNextStates mockDict
    
    Set mockLogger = New CMockOperationLogger
    Set mockErrorHandler = New CMockErrorHandlerService
    
    Set serviceImpl = New CWorkflowService
    serviceImpl.Initialize mockRepo, mockLogger, mockErrorHandler
    Set service = serviceImpl
    
    ' Act
    Dim result As Object
    Set result = service.GetNextStates("Registrado", "", "Calidad")
    
    ' Assert
    modAssert.AssertEquals 1, result.Count, "Debe haber exactamente un estado siguiente."
    modAssert.AssertTrue result.Exists(2), "Debe existir el estado ID 2 (Desarrollo)."
    
    TestGetNextStates_ValidState.Pass
Cleanup:
    Exit Function
TestFail:
    TestGetNextStates_ValidState.Fail "Error: " & Err.Description
    Resume Cleanup
End Function

Private Function TestIsEstadoFinal_AprobadaState() As CTestResult
    Set TestIsEstadoFinal_AprobadaState = New CTestResult
    TestIsEstadoFinal_AprobadaState.Initialize "IsEstadoFinal debe identificar Aprobada como final"
    
    Dim serviceImpl As CWorkflowService
    Dim mockRepo As CMockWorkflowRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim service As IWorkflowService
    
    On Error GoTo TestFail
    
    ' Arrange
    Set mockRepo = New CMockWorkflowRepository
    Set mockLogger = New CMockOperationLogger
    Set mockErrorHandler = New CMockErrorHandlerService
    
    Set serviceImpl = New CWorkflowService
    serviceImpl.Initialize mockRepo, mockLogger, mockErrorHandler
    Set service = serviceImpl
    
    ' Act
    Dim result As Boolean
    result = service.IsEstadoFinal("Aprobada")
    
    ' Assert
    modAssert.AssertTrue result, "Aprobada debería ser identificado como estado final."
    
    TestIsEstadoFinal_AprobadaState.Pass
Cleanup:
    Exit Function
TestFail:
    TestIsEstadoFinal_AprobadaState.Fail "Error: " & Err.Description
    Resume Cleanup
End Function

Private Function TestValidateTransition_CalidadCanOverrideRepository() As CTestResult
    Set TestValidateTransition_CalidadCanOverrideRepository = New CTestResult
    TestValidateTransition_CalidadCanOverrideRepository.Initialize "ValidateTransition debe devolver True para Calidad incluso si el repositorio dice False"
    
    Dim service As IWorkflowService
    Dim mockRepo As New CMockWorkflowRepository
    
    On Error GoTo TestFail
    
    ' Arrange: Configurar el mock repo para que devuelva False
    mockRepo.ConfigureIsValidTransition False
    
    ' Crear el servicio inyectando el mock
    Set service = modWorkflowServiceFactory.CreateWorkflowServiceWithMocks(mockRepo)
    
    ' Act: Validar una transición como rol "Calidad"
    Dim result As Boolean
    result = service.ValidateTransition(1, "PC", "EstadoA", "EstadoB", "Calidad")

    
    ' Assert: El servicio debe anular el False del repositorio y devolver True
    modAssert.AssertTrue result, "El rol Calidad debería haber anulado la respuesta del repositorio."
    
    TestValidateTransition_CalidadCanOverrideRepository.Pass
    GoTo Cleanup
TestFail:
    TestValidateTransition_CalidadCanOverrideRepository.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set service = Nothing
    Set mockRepo = Nothing
End Function

Private Function TestValidateTransition_TecnicoIsRestrictedByRepository() As CTestResult
    Set TestValidateTransition_TecnicoIsRestrictedByRepository = New CTestResult
    TestValidateTransition_TecnicoIsRestrictedByRepository.Initialize "ValidateTransition debe devolver False para Tecnico si el repositorio dice False"
    
    Dim service As IWorkflowService
    Dim mockRepo As New CMockWorkflowRepository
    
    On Error GoTo TestFail
    
    ' Arrange: Configurar el mock repo para que devuelva False
    mockRepo.ConfigureIsValidTransition False
    
    ' Crear el servicio inyectando el mock
    Set service = modWorkflowServiceFactory.CreateWorkflowServiceWithMocks(mockRepo)
    
    ' Act: Validar una transición como rol "Tecnico"
    Dim result As Boolean
    result = service.ValidateTransition(1, "PC", "EstadoA", "EstadoB", "Tecnico")
    
    
    ' Assert: El servicio debe respetar el False del repositorio para roles no privilegiados
    modAssert.AssertFalse result, "El rol Tecnico NO debería haber anulado la respuesta del repositorio."
    
    TestValidateTransition_TecnicoIsRestrictedByRepository.Pass
    GoTo Cleanup
TestFail:
    TestValidateTransition_TecnicoIsRestrictedByRepository.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set service = Nothing
    Set mockRepo = Nothing
End Function
