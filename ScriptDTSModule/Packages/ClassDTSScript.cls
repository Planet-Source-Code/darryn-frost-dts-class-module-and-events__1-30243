VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDTSScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************************************
'* This is the module created when you retrieve a package from the server and script it
'* To run your own, simply remove this module from the project and add one you have scripted

Option Explicit

Private WithEvents goPackage As DTS.Package
Attribute goPackage.VB_VarHelpID = -1

Private m_Provider As String
Private m_DataSource As String
Private m_Catalog As String
Private m_WorkStation As String
Private m_UserID As String
Private m_Pwd As String
Private m_ConnectImmediate As Boolean
Private m_ConnectionTimeout As Long
Private m_UseNTSecurity As Boolean
Private m_Cancel As Boolean
Private m_Totaltasks As Integer
Private m_TasksCompleted As Integer

Public Event ErrorOccurred(ByVal pErr As Long, ByVal pSource As String, ByVal pDescription As String)
Public Event PercentDone(ByVal percent As Integer)
Public Event RowsCopied(ByVal RowsCopied As String)
Public Event Currenttask(ByVal pCurrenttask As String)
Public Event CurrentStep(ByVal pCurrentStep As String)
       
Public Property Let CancelTask(pCancel As Boolean)
  m_Cancel = pCancel
End Property


Private Sub ParseConnectionString(ByVal strConnect As String)

  Dim i As Long
  Dim strProps() As String
  Dim strpropName As String
  Dim strPropVal As String
  Dim lPos As Long

  m_Provider = ""
  m_DataSource = ""
  m_UserID = "admin"
  m_Pwd = ""
  m_WorkStation = ""
  m_Catalog = ""
  m_UseNTSecurity = False

  strProps = Split(strConnect, ";", , vbTextCompare)

  For i = 0 To UBound(strProps)
    lPos = InStr(1, strProps(i), "=")
    If Not lPos = 0 Then
      strpropName = Left(strProps(i), lPos - 1)
      strPropVal = Mid(strProps(i), lPos + 1)

      Select Case strpropName
        Case "Provider"
          m_Provider = strPropVal
          If m_Provider = "SQLOLEDB.1" Then m_UserID = "sa"
        Case "User ID"
          If Not Len(strPropVal) = 0 Then m_UserID = strPropVal
        Case "Data Source"
          m_DataSource = strPropVal
        Case "Initial Catalog"
          m_Catalog = strPropVal
        Case "WorkStation ID"
          m_WorkStation = strPropVal
        Case "Persist Security Info"
          If strPropVal = "SSP1" And m_Provider = "SQLOLEDB.1" Then
            m_UseNTSecurity = True
          Else
            m_UseNTSecurity = strPropVal
          End If
        Case "Password"
          m_Pwd = strPropVal
        Case "ConnectionTimeout"
          m_ConnectionTimeout = strPropVal
      End Select

    End If
    
  Next i

End Sub

Private Sub AddConnection(ByVal pProvider As String, ByVal strConnectionName As String, _
                          ByVal lonConnectionID As Long, _
                          ByVal pDataSource As String, _
                          ByVal pCatalog As String, ByVal pUserID As String, _
                          Optional ByVal pPwd As String = "", _
                          Optional ByVal pNTSecurity As Boolean = False, _
                          Optional ByVal pConnectImmediate As Boolean = False, _
                          Optional ByVal pConnectionTimeout As Long = 0, _
                          Optional ByVal pReusable As Boolean = False)


  Dim oConnection As DTS.Connection

  Set oConnection = goPackage.Connections.New(pProvider)

  With oConnection
    .Name = strConnectionName
    .ID = lonConnectionID
    .Reusable = pReusable
    .ConnectImmediate = pConnectImmediate
    .DataSource = pDataSource
    .UserID = pUserID
    .Password = pPwd
    .ConnectionTimeout = pConnectionTimeout
    .Catalog = pCatalog
    .UseTrustedConnection = pNTSecurity
  End With

  goPackage.Connections.Add oConnection
  Set oConnection = Nothing

End Sub

Private Sub CreateStep(ByVal pStepName As String, ByVal pStepDescription As String, _
                       ByVal pTaskName As String, Optional ByVal pCloseConnection As Boolean = False, _
                       Optional ByVal pCommitOnSuccess = False, Optional ByVal pRollbackOnFailure = False)

  Dim oStep As DTS.Step

  Set oStep = goPackage.Steps.New

  With oStep
    .Name = pStepName
    .Description = pStepDescription
    .ExecutionStatus = 1
    .TaskName = pTaskName
    .CommitSuccess = pCommitOnSuccess
    .RollbackFailure = pRollbackOnFailure
    .ScriptLanguage = "VBScript"
    .AddGlobalVariables = True
    .RelativePriority = 3
    .CloseConnection = pCloseConnection
    .ExecuteInMainThread = True
    .IsPackageDSORowset = False
    .JoinTransactionIfPresent = False
    .DisableStep = False
  End With

  goPackage.Steps.Add oStep

  Set oStep = Nothing

End Sub


Private Sub AddExecuteSQLTask(ByVal pTaskName As String, ByVal pTaskDescription As String, _
                              ByVal pSQL As String, ByVal pConnectionID As Long, Optional ByVal pConnectionTimeout = 0)

  Dim oCustomTask As DTS.ExecuteSQLTask
  Dim oTask As DTS.Task

  Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
      
  Set oCustomTask = oTask.CustomTask
  
    oCustomTask.Name = pTaskName
    oCustomTask.Description = pTaskDescription
    oCustomTask.SQLStatement = pSQL
    oCustomTask.ConnectionID = pConnectionID
    oCustomTask.CommandTimeout = pConnectionTimeout
  
  goPackage.Tasks.Add oTask
  
  Set oCustomTask = Nothing
  Set oTask = Nothing

End Sub


Private Sub AddColumnTransformation(ByRef oCustomTask As DTS.DataPumpTask, _
                                    ByVal pTransformName As String, _
                                    ByVal pSourceColName As String, ByVal pDestColName As String, _
                                    ByVal pSourceDatatype As Long, _
                                    ByVal pDestDataType As Long, ByVal pSourceFieldSize As Long, _
                                    ByVal pDestFieldSize As Long, ByVal pSourceFlags As Long, _
                                    ByVal pDestFlags As Long, _
                                    Optional ByVal pSourceNullable As Boolean = False, _
                                    Optional ByVal pDestNullable As Boolean = False, _
                                    Optional ByVal pSourceOrdinal As Long = 1, _
                                    Optional ByVal pDestOrdinal As Long = 1)

  Dim oColumn As DTS.Column
  Dim oTransformation As DTS.Transformation
  
  Set oTransformation = oCustomTask.Transformations.New("DTS.DataPumpTransformCopy.1")
    oTransformation.Name = pTransformName
    oTransformation.TransformFlags = 63
    oTransformation.ForceSourceBlobsBuffered = 0
    oTransformation.ForceBlobsInMemory = False
    oTransformation.InMemoryBlobSize = 1048576
      
  Set oColumn = oTransformation.SourceColumns.New(pSourceColName, 1)
          
    oColumn.Name = pSourceColName
    oColumn.Ordinal = pSourceOrdinal
    oColumn.Flags = pSourceFlags
    oColumn.Size = pSourceFieldSize
    oColumn.DataType = pSourceDatatype
    oColumn.Precision = 0
    oColumn.NumericScale = 0
    oColumn.Nullable = pSourceNullable
        
  oTransformation.SourceColumns.Add oColumn
        
  Set oColumn = Nothing

  Set oColumn = oTransformation.DestinationColumns.New(pDestColName, 1)
          
    oColumn.Name = pDestColName
    oColumn.Ordinal = pDestOrdinal
    oColumn.Flags = pDestFlags
    oColumn.Size = pDestFieldSize
    oColumn.DataType = pDestDataType
    oColumn.Precision = 0
    oColumn.NumericScale = 0
    oColumn.Nullable = pDestNullable
        
  oTransformation.DestinationColumns.Add oColumn
        
  Set oColumn = Nothing
  
  oCustomTask.Transformations.Add oTransformation
  Set oTransformation = Nothing
      
End Sub


Private Sub AddPrecedenceConstraint(ByVal pStep As String, ByVal priorStep As String, _
                                    ByVal pConstraintBasis As Long, _
                                    ByVal pConstraintResult As Long)

  Dim oStep As DTS.Step
  Dim oPrecConstraint As DTS.PrecedenceConstraint

  Set oStep = goPackage.Steps(pStep)
  Set oPrecConstraint = oStep.PrecedenceConstraints.New(priorStep)
    oPrecConstraint.StepName = priorStep
    oPrecConstraint.PrecedenceBasis = pConstraintBasis
    oPrecConstraint.Value = pConstraintResult
  oStep.PrecedenceConstraints.Add oPrecConstraint
  Set oPrecConstraint = Nothing

End Sub


Private Sub CreateCustomTask(ByRef pTask As DTS.Task, ByVal pTaskName As String, _
                             ByVal pTaskDescription As String, ByVal pSourceConnectionID As Long, _
                             ByVal pSourceObjectName As String, ByVal pDestConnectionID As Long, _
                             ByVal pDestObjectName As String, _
                             Optional ByVal pSourceSQL As String = "", _
                             Optional ByVal pDestSQL As String = "", _
                             Optional ByVal pFetchBuffer As Long = 1, _
                             Optional ByVal pProgressCount As Long = 1000, _
                             Optional ByVal pFastLoad As Boolean = True, _
                             Optional ByVal pFastLoadOptions As Long = 2, _
                             Optional ByVal pMaxErrors As Integer = 0, _
                             Optional ByVal pAllowIdentityInserts As Boolean = False, _
                             Optional ByVal pFirstRow As Long = 0, _
                             Optional ByVal pLastRow As Long = 0)

  pTask.CustomTask.Properties("Name") = pTaskName
  pTask.CustomTask.Properties("Description") = pTaskDescription
  pTask.CustomTask.Properties("SourceConnectionID") = pSourceConnectionID
  pTask.CustomTask.Properties("DestinationConnectionID") = pDestConnectionID
  
  If pSourceSQL = "" Then
    pTask.CustomTask.Properties("SourceObjectName") = pSourceObjectName
  Else
    pTask.CustomTask.Properties("SourceSQLStatement") = pSourceSQL
  End If
  
  If pDestSQL = "" Then
    pTask.CustomTask.Properties("DestinationObjectName") = pDestObjectName
  Else
    pTask.CustomTask.Properties("DestinationSQLStatement") = pDestSQL
  End If
  
  pTask.CustomTask.Properties("ProgressRowCount") = pProgressCount
  pTask.CustomTask.Properties("MaximumErrorCount") = pMaxErrors
  pTask.CustomTask.Properties("FetchBufferSize") = pFetchBuffer
  pTask.CustomTask.Properties("UseFastLoad") = pFastLoad
  pTask.CustomTask.Properties("InsertCommitSize") = 0
  pTask.CustomTask.Properties("ExceptionFileColumnDelimiter") = "|"
  pTask.CustomTask.Properties("ExceptionFileRowDelimiter") = vbCrLf
  pTask.CustomTask.Properties("AllowIdentityInserts") = pAllowIdentityInserts
  pTask.CustomTask.Properties("FirstRow") = pFirstRow
  pTask.CustomTask.Properties("LastRow") = pLastRow
  pTask.CustomTask.Properties("FastLoadOptions") = pFastLoadOptions

End Sub

Public Sub CreatePackage(Optional ByVal pSourceConnectionString As String, _
                         Optional ByVal pSourceConnectionID As Long = 1, _
                         Optional ByVal pDestConnectionString As String, _
                         Optional ByVal pDestConnectionID As Long = 2)

Dim oConnProperty As DTS.OleDBProperty
Dim oTask As DTS.Task

Set goPackage = New DTS.Package

  goPackage.Name = "TestPkg"
  'Readonly property: goPackage.PackageID = {CB64F4FC-3CD3-48E3-9A75-EED7CE90625A}
  goPackage.Description = "WinSort Access Volume Insert"
  goPackage.WriteCompletionStatusToNTEventLog = True
  goPackage.LogFileName = "C:\Documents and Settings\jester\Desktop\DTSTest\Database\DTSErrlog.txt"
  goPackage.FailOnError = False
  'Readonly property: goPackage.VersionID = {B95CE19E-6389-4A70-A824-E20F695EE581}
  'Readonly property: goPackage.CreatorName = ICE\jester
  'Readonly property: goPackage.CreatorComputerName = ICE
  'Readonly property: goPackage.CreationDate = 12/31/2001 8:57:57 AM
  goPackage.PackagePriorityClass = 3
  goPackage.MaxConcurrentSteps = 4
  goPackage.LineageOptions = 0
  goPackage.UseTransaction = True
  goPackage.TransactionIsolationLevel = 4096
  goPackage.AutoCommitTransaction = True
  goPackage.RepositoryMetadataOptions = 0
  goPackage.UseOLEDBServiceComponents = True

'***********************************************************************************************
'TO USE ADO CONNECTION STRINGS THAT YOU PASS IN AS YOUR SOURCE AND DESTINATION, SIMPLY
'UNCOMMENT THE FOLLOWING LINES AND COMMENT OUT THE NEXT SECTION OF ADD CONNECTION LINES
'If you pass in connectionStrings, you need to make sure you passed in the correct
'ConnectionID for each, or it will set the Source to 1, and the Destination to 2
  If Not Len(pSourceConnectionString) = 0 Then
    Call ParseConnectionString(pSourceConnectionString)
    Call AddConnection(m_Provider, "SourceConnection", pSourceConnectionID, _
                       m_DataSource, m_Catalog, m_UserID, m_Pwd, m_UseNTSecurity, m_ConnectImmediate, _
                       m_ConnectionTimeout, True)
  End If
  If Not Len(pDestConnectionString) = 0 Then
    Call ParseConnectionString(pDestConnectionString)
    Call AddConnection(m_Provider, "DestinationConnection", pDestConnectionID, _
                       m_DataSource, m_Catalog, m_UserID, m_Pwd, m_UseNTSecurity, m_ConnectImmediate, _
                       m_ConnectionTimeout, True)
  End If

'TO USE ADO CONNECTION STRINGS THAT YOU PASS IN AS YOUR SOURCE AND DESTINATION, SIMPLY
'COMMENT OUT THE FOLLOWING LINES IN THIS SECTION IF YOU UNCOMMENTED THE ONES ABOVE
'  Call AddConnection("Microsoft.Jet.OLEDB.4.0", "Microsoft Access FROM", 1, "C:\Documents and Settings\jester\Desktop\ScriptDTSModule\Database\From.mdb", "", "Admin", "", False, False, 60, True)

'  Call AddConnection("Microsoft.Jet.OLEDB.4.0", "Microsoft Access To", 2, "C:\Documents and Settings\jester\Desktop\ScriptDTSModule\Database\To.mdb", "", "Admin", "", False, False, 60, True)

'***********************************************************************************************
  Call CreateStep("STEP01_INSERT_NEW_RECORDS", "INSERT_NEW_RECORDS", "TASK01_INSERT_NEW_RECORDS", False, False, False)
  Call CreateStep("STEP02_DELETE_TEMP_RECORDS", "DELETE_TEMP_RECORDS", "TASK02_DELETE_TEMP_RECORDS", False, False, False)
  Call CreateStep("STEP03_INSERT_UPDATE_RECORDS", "INSERT_UPDATE_RECORDS", "TASK03_INSERT_UPDATE_RECORDS", False, False, False)
  Call CreateStep("STEP04_UPDATE_DEST_TABLE", "UPDATE_DEST_TABLE", "TASK04_UPDATE_DEST_TABLE", True, False, False)
  Call CreateStep("DTSStep_DTSDataPumpTask_1", "Transfer_Records_Amount1", "DTSTask_DTSDataPumpTask_1", False, False, False)
  Call CreateStep("DTSStep_DTSExecuteSQLTask_1", "SP_DELETE_AMOUNT1", "DTSTask_DTSExecuteSQLTask_1", False, False, False)

'Scripting Precedence Constraints for Package
  Call AddPrecedenceConstraint("STEP02_DELETE_TEMP_RECORDS", "STEP01_INSERT_NEW_RECORDS", 1, 0)
  Call AddPrecedenceConstraint("STEP03_INSERT_UPDATE_RECORDS", "STEP02_DELETE_TEMP_RECORDS", 1, 0)
  Call AddPrecedenceConstraint("STEP04_UPDATE_DEST_TABLE", "STEP03_INSERT_UPDATE_RECORDS", 1, 0)
  Call AddPrecedenceConstraint("DTSStep_DTSDataPumpTask_1", "STEP04_UPDATE_DEST_TABLE", 0, 4)
  Call AddPrecedenceConstraint("DTSStep_DTSExecuteSQLTask_1", "DTSStep_DTSDataPumpTask_1", 1, 0)

'Scripting Tasks For Package
  Set oTask = goPackage.Tasks.New("DTSDataPumpTask")

  Call CreateCustomTask(oTask, "TASK01_INSERT_NEW_RECORDS", "Task01_Insert", 1, "", 2, "TestTable", "SELECT Pass, OldID, NewID, Pocket, Amount FROM TestTable WHERE FldUpdate = 'New'", , 1, 1000, True, 2, 0, False, 0, 0)

    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_1", "Pass", "Pass", 130, 130, 255, 2, 102, 120, True, True, 1, 1)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_2", "OldID", "OldID", 130, 130, 255, 2, 102, 120, True, True, 1, 1)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_3", "NewID", "NewID", 130, 130, 255, 2, 102, 102, True, True, 1, 1)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_4", "Pocket", "Pocket", 130, 130, 255, 2, 102, 102, True, True, 1, 1)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_5", "Amount", "Amount", 130, 130, 255, 2, 32, 102, True, True, 1, 1)

  goPackage.Tasks.Add oTask
  Set oTask = Nothing

  Call AddExecuteSQLTask("TASK02_DELETE_TEMP_RECORDS", "Task02_Delete", "DELETE * FROM TempTestTable", 2, 0)

  Set oTask = goPackage.Tasks.New("DTSDataPumpTask")

  Call CreateCustomTask(oTask, "TASK03_INSERT_UPDATE_RECORDS", "Task03_Insert", 1, "", 2, "TempTestTable", "SELECT Pass, OldID, NewID, Pocket, Amount FROM TestTable WHERE FldUpdate = 'Update'", , 1, 1000, True, 2, 0, False, 0, 0)

    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_1", "Pass", "Pass", 130, 130, 255, 2, 102, 120, True, True, 1, 1)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_2", "OldID", "OldID", 130, 130, 255, 2, 102, 120, True, True, 1, 1)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_3", "NewID", "NewID", 130, 130, 255, 2, 102, 102, True, True, 1, 1)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_4", "Pocket", "Pocket", 130, 130, 255, 2, 102, 102, True, True, 1, 1)
    Call AddColumnTransformation(oTask.CustomTask, "TestTableCol_5", "Amount", "Amount", 130, 130, 255, 2, 32, 102, True, True, 1, 1)

  goPackage.Tasks.Add oTask
  Set oTask = Nothing

  Call AddExecuteSQLTask("TASK04_UPDATE_DEST_TABLE", "Task04_Delete", "UPDATE TestTable INNER JOIN TempTestTable ON ((TestTable.Pass = TempTestTable.Pass) AND (TestTable.OldID = TempTestTable.OldID)) SET TestTable.NewID = TempTestTable.OldID, TestTable.Amount = TempTesttable.Amount", 2, 0)

  Set oTask = goPackage.Tasks.New("DTSDataPumpTask")

  Call CreateCustomTask(oTask, "DTSTask_DTSDataPumpTask_1", "Transfer_Records_Amount1", 1, "", 2, "TempTesttable", "SELECT * " _
    & "FROM Testtable " _
    & "WHERE Amount = 1", , 1, 1000, True, 2, 0, False, 0, 0)

    Call AddColumnTransformation(oTask.CustomTask, "DTSTransformation__1", "Pass", "Pass", 130, 130, 8, 8, 102, 102, True, True, 1, 1)
    Call AddColumnTransformation(oTask.CustomTask, "DTSTransformation__2", "OldID", "OldID", 130, 130, 4, 4, 102, 102, True, True, 1, 1)
    Call AddColumnTransformation(oTask.CustomTask, "DTSTransformation__3", "NewID", "NewID", 130, 130, 4, 4, 102, 102, True, True, 1, 1)
    Call AddColumnTransformation(oTask.CustomTask, "DTSTransformation__4", "Pocket", "Pocket", 130, 130, 2, 2, 102, 102, True, True, 1, 1)
    Call AddColumnTransformation(oTask.CustomTask, "DTSTransformation__5", "Amount", "Amount", 3, 3, 0, 0, 118, 118, True, True, 1, 1)

  goPackage.Tasks.Add oTask
  Set oTask = Nothing

  Call AddExecuteSQLTask("DTSTask_DTSExecuteSQLTask_1", "SP_DELETE_AMOUNT1", "EXEC DELETE_TEST_AMOUNT 1", 2, 0)
    
End Sub

Public Sub SavePackage()
goPackage.SaveToSQLServer "(local)", "sa", ""
End Sub
Private Sub Class_Terminate()
  Set goPackage = Nothing
End Sub


Public Function ExecutePackage() As Boolean
  Dim errStep As Long

On Error Resume Next

'Establish the count for progress events
  m_Totaltasks = goPackage.Tasks.Count
  m_TasksCompleted = 0

  RaiseEvent PercentDone(0)
  
  ExecutePackage = True

  goPackage.Execute

  For errStep = 1 To goPackage.Steps.Count
    If goPackage.Steps(errStep).ExecutionResult = DTSStepExecResult_Failure Then
      ExecutePackage = False
      m_Cancel = True
      Debug.Print "Step "; goPackage.Steps(errStep).Name; " has failed " _
      & vbCrLf; "Affected tables have been rolled back to previous state."
    End If
  Next errStep

  RaiseEvent PercentDone(100)

End Function

Private Sub goPackage_OnStart(ByVal EventSource As String)
  RaiseEvent Currenttask(Right(EventSource, Len(EventSource) - 7))
  DoEvents
End Sub

Private Sub goPackage_OnFinish(ByVal EventSource As String)
  m_TasksCompleted = m_TasksCompleted + 1
  RaiseEvent PercentDone((m_TasksCompleted / m_Totaltasks) * 100)
  RaiseEvent Currenttask("")
  DoEvents
End Sub

Private Sub goPackage_OnProgress(ByVal EventSource As String, ByVal ProgressDescription As String, ByVal PercentComplete As Long, ByVal ProgressCountLow As Long, ByVal ProgressCountHigh As Long)
  RaiseEvent RowsCopied(ProgressDescription)
  DoEvents
End Sub

Private Sub goPackage_OnError(ByVal EventSource As String, ByVal ErrorCode As Long, ByVal source As String, ByVal Description As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal IDofInterfaceWithError As String, pbCancel As Boolean)

  RaiseEvent ErrorOccurred(ErrorCode, source, Description)

End Sub

Private Sub goPackage_OnQueryCancel(ByVal EventSource As String, pbCancel As Boolean)
  If m_Cancel = True Then
    pbCancel = True
  Else
    pbCancel = False
  End If
End Sub
