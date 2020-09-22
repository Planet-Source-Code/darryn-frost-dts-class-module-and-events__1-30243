Attribute VB_Name = "ScriptPkg"
Option Explicit
Public goPackage As New DTS.Package
Public gnScriptFile As Long         'The name and path of the file to place the VB script in
Public QUOTE As String              'The Quote character (used as a constant)
Const TABSIZE = 2                   'Set tab size = 2 spaces in the script file

'********************************************************************
'* ScriptPkg - a sample DTS program to create VB code from packages
'*
'* This program script a DTS Package saved on your local server as VB code.
'* It attempts to login to the server as 'sa' with no password. To change that, edit the LoadServerPkg routine at the end of this file
'* The resulting script is put in your temp directory with the same filename as the package and a .txt extention.
'*
'* Known issues:
'*
'*   1) if the package contains too many columns or tables, the script that is generated will be bigger then VB can hold in one function.
'* To get around this, manually breakup the script into multiple functions.
'*
'*   2) Sometimes connection providers will identify readonly properties. If the scripted package fails when setting a
'* connection property, comment out the line that sets that property - it's not needed.
'*
'********************************************************************
Public Sub getPackage(ByVal strPkgName As String, ByVal strPath As String)
Dim i%, j%
Dim strTaskName$
QUOTE = Chr(&H22)
    
    Call LoadServerPkg(strPkgName)  'load the package from the server
    
    'open the script file
    gnScriptFile = FreeFile
    If gnScriptFile = 0 Then FAIL "Unable to get a filenumber"
    
    Open strPath For Output As gnScriptFile
    
    Call ScriptClassModule
    
    Scriptout "Option Explicit"
    Scriptout ""
    Scriptout "Private WithEvents goPackage As DTS.Package"
    Scriptout ""
    
    Scriptout "Private m_Provider As String"
    Scriptout "Private m_DataSource As String"
    Scriptout "Private m_Catalog As String"
    Scriptout "Private m_WorkStation As String"
    Scriptout "Private m_UserID As String"
    Scriptout "Private m_Pwd As String"
    Scriptout "Private m_ConnectImmediate As Boolean"
    Scriptout "Private m_ConnectionTimeout As Long"
    Scriptout "Private m_UseNTSecurity As Boolean"
    
    Scriptout "Private m_Cancel As Boolean"
    Scriptout "Private m_Totaltasks As Integer"
    Scriptout "Private m_TasksCompleted As Integer"
    Scriptout ""
    Scriptout "Public Event ErrorOccurred(ByVal pErr As Long, ByVal pSource As String, ByVal pDescription As String)"
    Scriptout "Public Event PercentDone(ByVal percent As Integer)"
    Scriptout "Public Event RowsCopied(ByVal RowsCopied As String)"
    Scriptout "Public Event Currenttask(ByVal pCurrenttask As String)"
    Scriptout "Public Event CurrentStep(ByVal pCurrentStep As String)"
    Scriptout "       "
    Scriptout "Public Property Let CancelTask(pCancel As Boolean)"
    Scriptout "m_Cancel = pCancel", 1
    Scriptout "End Property"
    Scriptout ""
    
    Call ScriptParseConnection
    Call ScriptAddConnection
    Call ScriptAddStep
    Call ScriptAddSQLTask
    Call ScriptAddColumnTransformation
    Call ScriptAddConstraint
    Call ScriptCreateCustomTask
    
    Scriptout "Public Sub CreatePackage(Optional Byval pSourceConnectionString as string, _ "
    Scriptout "                         Optional Byval pSourceConnectionID as Long = 1, _"
    Scriptout "                         Optional Byval pDestConnectionString as string, _"
    Scriptout "                         Optional Byval pDestConnectionID as Long = 2)"
    Scriptout ""
    Scriptout "Dim oConnProperty As DTS.OleDBProperty"
    Scriptout "Dim oTask AS DTS.Task"
    Scriptout ""
    Scriptout "Set goPackage = new DTS.Package"
    Scriptout ""
    'set the package properties
    ScriptObject goPackage, "goPackage", 1
    Scriptout ""
    
    Dim oConnProperty As DTS.OleDBProperty
    
    
    'Script Globals
    If goPackage.GlobalVariables.Count > 0 Then
        Dim oGlobal As DTS.GlobalVariable
        Scriptout "dim oGlobal as DTS.GlobalVariable"
        For Each oGlobal In goPackage.GlobalVariables
            Scriptout "Set oGlobal = goPackage.GlobalVariables.New(" + QUOTE + oGlobal.Name + QUOTE + ")", 1
            ScriptObject oGlobal, "oGlobal", 1
            Scriptout "goPackage.GlobalVariables.Add oGlobal"
            Scriptout "set oGlobal = Nothing"
            Scriptout ""
        Next oGlobal
    End If
    
    Scriptout "'***********************************************************************************************"
    Scriptout "'TO USE ADO CONNECTION STRINGS THAT YOU PASS IN AS YOUR SOURCE AND DESTINATION, SIMPLY "
    Scriptout "'UNCOMMENT THE FOLLOWING LINES AND COMMENT OUT THE NEXT SECTION OF ADD CONNECTION LINES"
    Scriptout "'If you pass in connectionStrings, you need to make sure you passed in the correct "
    Scriptout "'ConnectionID for each, or it will set the Source to 1, and the Destination to 2 "
    Scriptout "'If Not Len(pSourceConnectionString) = 0 then ", 1
    Scriptout "'Call ParseConnectionString(pSourceConnectionString)", 2
    Scriptout "'Call AddConnection(m_Provider," & QUOTE & "SourceConnection" & QUOTE & ", pSourceConnectionID, _", 2
    Scriptout "'                   m_DataSource, m_Catalog, m_UserID, m_Pwd, m_UseNTSecurity, m_ConnectImmediate, _", 2
    Scriptout "'                   m_ConnectionTimeout, True)", 2
    Scriptout "'End If", 1
    Scriptout "'If Not Len(pDestConnectionString) = 0 then ", 1
    Scriptout "'Call ParseConnectionString(pDestConnectionString)", 2
    Scriptout "'Call AddConnection(m_Provider," & QUOTE & "DestinationConnection" & QUOTE & ", pDestConnectionID, _", 2
    Scriptout "'                   m_DataSource, m_Catalog, m_UserID, m_Pwd, m_UseNTSecurity, m_ConnectImmediate, _", 2
    Scriptout "'                   m_ConnectionTimeout, True)", 2
    Scriptout "'End If", 1
    Scriptout ""
     
    Scriptout "'TO USE ADO CONNECTION STRINGS THAT YOU PASS IN AS YOUR SOURCE AND DESTINATION, SIMPLY "
    Scriptout "'COMMENT OUT THE FOLLOWING LINES IN THIS SECTION IF YOU UNCOMMENTED THE ONES ABOVE"
    
    'Script Connections
    If goPackage.Connections.Count > 0 Then
        Dim oConnection As DTS.Connection
        Dim strCon As String
        
        For Each oConnection In goPackage.Connections
            strCon = QUOTE & oConnection.ProviderID & QUOTE & ", "
            strCon = strCon & QUOTE & oConnection.Name & QUOTE & ", "
            strCon = strCon & oConnection.ID & ", "
            strCon = strCon & QUOTE & oConnection.DataSource & QUOTE & ", "
            strCon = strCon & QUOTE & oConnection.Catalog & QUOTE & ", "
            strCon = strCon & QUOTE & oConnection.UserID & QUOTE & ", "
            strCon = strCon & QUOTE & oConnection.Password & QUOTE & ", "
            strCon = strCon & oConnection.UseTrustedConnection & ", "
            strCon = strCon & oConnection.ConnectImmediate & ", "
            strCon = strCon & oConnection.ConnectionTimeout & ", "
            strCon = strCon & oConnection.Reusable
            Scriptout "Call AddConnection(" & strCon & ")", 1
            Scriptout ""
        Next oConnection
    
    End If
    
    Scriptout "'***********************************************************************************************"
    'Script Steps
    If goPackage.Steps.Count > 0 Then
        Dim oStep As DTS.Step
        Dim oPrecConstraint As DTS.PrecedenceConstraint
                
        'make two passes. First script the steps, then add the precedence constraints
        'in this way, the steps will all exist when the precedence constraints are added
        Dim strStep As String
        For Each oStep In goPackage.Steps
           strStep = ""
           strStep = QUOTE & oStep.Name & QUOTE & ", "
           strStep = strStep & QUOTE & oStep.Description & QUOTE & ", "
           strStep = strStep & QUOTE & oStep.TaskName & QUOTE & ", "
           strStep = strStep & oStep.CloseConnection & ", "
           strStep = strStep & oStep.CommitSuccess & ", "
           strStep = strStep & oStep.RollbackFailure
           Scriptout "Call CreateStep(" & strStep & ")", 1
        Next oStep
        Set oStep = Nothing
        Scriptout ""
        Scriptout "'Scripting Precedence Constraints for Package"
        Dim strPrec As String
        For Each oStep In goPackage.Steps
            'Script Precedence Constrainsts
            If oStep.PrecedenceConstraints.Count > 0 Then
                
                For Each oPrecConstraint In oStep.PrecedenceConstraints
                  strPrec = ""
                  strPrec = QUOTE & oStep.Name & QUOTE & ", "
                  strPrec = strPrec & QUOTE & oPrecConstraint.StepName & QUOTE & ", "
                  strPrec = strPrec & oPrecConstraint.PrecedenceBasis & ", "
                  strPrec = strPrec & oPrecConstraint.Value
                  Scriptout "Call AddPrecedenceConstraint(" & strPrec & ")", 1
                Next oPrecConstraint
                
            End If
        Next oStep
        Set oStep = Nothing
    End If
    Scriptout ""
    Scriptout "'Scripting Tasks For Package"
   'Script Tasks
   If goPackage.Tasks.Count > 0 Then
        
        Dim oTask As DTS.Task
        Dim oPumpTask As DTS.DataPumpTask
        Dim oDDQTask As DTS.DataDrivenQueryTask
        Dim strTask As String
        
        j = 0 'j = number of tasks
        For Each oTask In goPackage.Tasks
            'create each task and script the general task properties
            j = j + 1
            strTask = ""
            If InStr(1, oTask.CustomTaskID, "DTSExecuteSQLTask") <> 0 Then
              strTask = QUOTE & oTask.CustomTask.Properties("Name") & QUOTE & ", "
              strTask = strTask & QUOTE & oTask.CustomTask.Properties("Description") & QUOTE & ", "
              strTask = strTask & strFixSQL(oTask.CustomTask.Properties("SQLStatement")) & ", "
              strTask = strTask & oTask.CustomTask.Properties("ConnectionID") & ", "
              strTask = strTask & oTask.CustomTask.Properties("CommandTimeout")
              Scriptout "Call AddExecuteSQLTask(" & strTask & ")", 1
              Scriptout ""
            Else
              Scriptout "Set oTask = goPackage.Tasks.New(" & QUOTE & oTask.CustomTaskID & QUOTE & ")", 1
              Scriptout ""
              'if it's a datapump task
              If InStr(LCase(oTask.CustomTaskID), "datapump") > 0 Then
                Set oPumpTask = oTask.CustomTask
                strTask = "oTask, "
                strTask = strTask & QUOTE & oPumpTask.Name & QUOTE & ", "
                strTask = strTask & QUOTE & oPumpTask.Description & QUOTE & ", "
                strTask = strTask & oPumpTask.SourceConnectionID & ", "
                strTask = strTask & QUOTE & oPumpTask.SourceObjectName & QUOTE & ", "
                strTask = strTask & oPumpTask.DestinationConnectionID & ", "
                strTask = strTask & QUOTE & oPumpTask.DestinationObjectName & QUOTE & ", "
                strTask = strTask & strFixSQL(oPumpTask.SourceSQLStatement) & ", "
                strTask = strTask & strFixSQL(oPumpTask.DestinationSQLStatement) & ", "
                strTask = strTask & oPumpTask.FetchBufferSize & ", "
                strTask = strTask & oPumpTask.ProgressRowCount & ", "
                strTask = strTask & oPumpTask.UseFastLoad & ", "
                strTask = strTask & oPumpTask.FastLoadOptions & ", "
                strTask = strTask & oPumpTask.MaximumErrorCount & ", "
                strTask = strTask & oPumpTask.AllowIdentityInserts & ", "
                strTask = strTask & oPumpTask.FirstRow & ", "
                strTask = strTask & oPumpTask.LastRow
                Scriptout "Call CreateCustomtask(" & strTask & ")", 1
                Scriptout ""
                ScriptPump oPumpTask, strTaskName, False
              ElseIf InStr(LCase(oTask.CustomTaskID), "datadrivenquery") > 0 Then
                ScriptObject oTask, strTaskName, 2
                Set oDDQTask = oTask.CustomTask
                ScriptPump oDDQTask, strTaskName, True
              End If

              Scriptout "goPackage.Tasks.Add oTask", 1
              Scriptout "Set oTask = Nothing", 1
              Scriptout ""
            End If
        Next oTask
        Set oTask = Nothing
    End If
        
    Scriptout "End Sub"
    
    Call ScriptoutClassTerminate
    Call ScriptoutExecutePkg
    Call ScriptoutPkgStart
    Call ScriptoutPkgFinish
    Call ScriptoutPkgProgress
    Call ScriptOutPkgErr
    Call ScriptoutPkgCancel
    Call ScriptSavePkg
    
    MsgBox "Package " & goPackage.Name & " scripted as " & strPath & "."
        
    'cleanup
    Close #gnScriptFile
    goPackage.UnInitialize
    Set goPackage = Nothing
    Exit Sub
End Sub

'********************************************************************
'* Subroutine:  ScriptPump
'*
'* Description: for datapump and DDQ tasks, get the custom task and script it's properties
'*
'* Parameters:
'*      oPumpTask - the object to script (either a Pump or DDQ custom task)
'*      strTaskName$ - name of the task being scripted
'********************************************************************
Sub ScriptPump(oPumpTask As Object, strTaskName$, fDDQ%)
Dim oTransformation As DTS.Transformation
Dim oTransProps As DTS.Properties
Dim i%, fTransObject%
Dim oProperty As DTS.Property
Dim oConnProperty As DTS.OleDBProperty
    
    'Script the Transformations
    If oPumpTask.Transformations.Count > 0 Then
        Dim oLookup As DTS.Lookup
        For i = 1 To oPumpTask.Transformations.Count
            Set oTransformation = oPumpTask.Transformations(i)
                        
            'Script Columns
            ScriptTransColumns oTransformation.SourceColumns(1), oTransformation.DestinationColumns(1), oTransformation.Name
                                    
            If fDDQ Then
                'Script Delete Query Columns
                ScriptColumns oPumpTask.DeleteQueryColumns, strTaskName & ".DeleteQueryColumns"
                            
                'Script Insert Query Columns
                ScriptColumns oPumpTask.InsertQueryColumns, strTaskName & ".InsertQueryColumns"
                
                'Script Update Query Columns
                ScriptColumns oPumpTask.UpdateQueryColumns, strTaskName & ".UpdateQueryColumns"
                
                'Script User Query Columns
                ScriptColumns oPumpTask.UserQueryColumns, strTaskName & ".UserQueryColumns"
            End If
            
            'Script Lookups
            If oPumpTask.Lookups.Count > 0 Then
                For Each oLookup In oPumpTask.Lookups
                    Scriptout "set oLookup = " & strTaskName & ".Lookups.New (" & QUOTE & oLookup.Name & QUOTE & ")", 3
                    ScriptObject oLookup, "oLookup", 3
                    Scriptout strTaskName & ".Lookups.Add oLookup", 3
                    Scriptout "Set oLookup = Nothing", 3
                    Scriptout ""
                Next oLookup
                Set oLookup = Nothing
            End If
            
            'Script Transform Server if it exists and isn't the copy transform (copy has no properties)
            fTransObject = True
            If InStr(oTransformation.TransformServerID, "DataPumpTransformCopy") = 0 Then
                On Error GoTo NoTransformObject
                Set oTransProps = oTransformation.TransformServerProperties
                If fTransObject Then
                    Scriptout "Set oTransProps = oTransformation.TransformServerProperties", 4
                            
                    For Each oProperty In oTransProps
                        ScriptProperty oProperty, "oTransProps(" & QUOTE & oProperty.Name & QUOTE & ")", 4
                    Next oProperty
                
                    Set oProperty = Nothing
                    Set oTransProps = Nothing
                    Scriptout "Set oTransProps = Nothing", 4
                    Scriptout ""
                End If
            End If
                
        Set oTransformation = Nothing
        
        Next i 'next transformation
                
        'script SourceCommandProperties
        If oPumpTask.SourceCommandProperties.Count > 0 Then
            For Each oConnProperty In oPumpTask.SourceCommandProperties
                If oConnProperty.Value <> "" Then
                    If VarType(oConnProperty) = vbString Then
                        Scriptout strTaskName & ".SourceCommandProperties(" & oConnProperty.PropertyID & ") = " & QUOTE & oConnProperty.Value & QUOTE, 2
                    Else
                        Scriptout strTaskName & ".SourceCommandProperties(" & oConnProperty.PropertyID & ") = " & oConnProperty.Value, 2
                    End If
                End If
            Next oConnProperty
            Set oConnProperty = Nothing
            Scriptout ""
        End If
        
        'Script DestinationCommandProperties
        'If oPumpTask.DestinationCommandProperties.Count > 0 Then
        '    For Each oConnProperty In oPumpTask.DestinationCommandProperties
        '        If oConnProperty.Value <> "" Then
        '            If VarType(oConnProperty) = vbString Then
        '                Scriptout strTaskName & ".DestinationCommandProperties(" & oConnProperty.PropertyID & ") = " & QUOTE & oConnProperty.Value & QUOTE, 2
        '            Else
        '                Scriptout strTaskName & ".DestinationCommandProperties(" & oConnProperty.PropertyID & ") = " & oConnProperty.Value, 2
        '            End If
        '        End If
        '    Next oConnProperty
        '    Set oConnProperty = Nothing
        '    Scriptout ""
        'End If
    End If
    Set oPumpTask = Nothing

Exit Sub

NoTransformObject:
    fTransObject = False
    Resume Next
End Sub

Sub ScriptTransColumns(oSourceColumn As DTS.Column, oDestColumn As DTS.Column, pTransName As String)
    
  Dim strColTrans As String
   
  strColTrans = "oTask.Customtask, "
  strColTrans = strColTrans & QUOTE & pTransName & QUOTE & ", "
  strColTrans = strColTrans & QUOTE & oSourceColumn.Name & QUOTE & ", "
  strColTrans = strColTrans & QUOTE & oDestColumn.Name & QUOTE & ", "
  strColTrans = strColTrans & oSourceColumn.DataType & ", "
  strColTrans = strColTrans & oDestColumn.DataType & ", "
  strColTrans = strColTrans & oSourceColumn.Size & ", "
  strColTrans = strColTrans & oDestColumn.Size & ", "
  strColTrans = strColTrans & oSourceColumn.Flags & ", "
  strColTrans = strColTrans & oDestColumn.Flags & ", "
  strColTrans = strColTrans & oSourceColumn.Nullable & ", "
  strColTrans = strColTrans & oDestColumn.Nullable & ", "
  strColTrans = strColTrans & oSourceColumn.Ordinal & ", "
  strColTrans = strColTrans & oDestColumn.Ordinal
  
  Scriptout "Call AddColumnTransformation(" & strColTrans & ")", 2
    
End Sub

Sub ScriptColumns(oColumns As DTS.Columns, strColumnsObjName As String)
Dim oColumn As DTS.Column
    If oColumns.Count > 0 Then
        For Each oColumn In oColumns
            Scriptout "Set oColumn = " & strColumnsObjName & ".New(" + QUOTE & oColumn.Name & QUOTE + "," & oColumn.Ordinal & ")", 3
            ScriptObject oColumn, "oColumn", 3
            Scriptout strColumnsObjName & ".Add oColumn", 3
            Scriptout "Set oColumn = Nothing", 3
            Scriptout ""
        Next oColumn
        Set oColumn = Nothing
    End If

End Sub
'********************************************************************
'* Subroutine:  ScriptObject
'*
'* Description: script all the properties of a DTS object
'*
'* Parameters:
'*      oObject - the object to script
'*      strObjectName$ - the sub string to find and replace
'*      nIndent% - number of tabs to indent the line
'********************************************************************
Sub ScriptObject(oObject, strObjectName$, Optional nIndent%)
Dim oProperty As DTS.Property
    For Each oProperty In oObject.Properties
        ScriptProperty oProperty, strObjectName, nIndent
    Next oProperty
End Sub

'********************************************************************
'* Subroutine:  ScriptProperty
'*
'* Description: script out a property of an object
'*
'* Parameters:
'*      oProperty - the property to script
'*      strObjectName$ - the sub string to find and replace
'*      nIndent% - number of tabs to indent the line
'********************************************************************
Sub ScriptProperty(oProperty As DTS.Property, strObjectName$, Optional nIndent%)
Dim strObjPart$
    strObjPart = strObjectName + "." + oProperty.Name
    If InStr(strObjectName, "(") > 0 Then
        strObjPart = strObjectName
    End If
    If oProperty.Set = True Then ' if writable
        Select Case oProperty.Type
            Case vbString  'text
                ScriptStringValue strObjPart, oProperty.Value, nIndent
            Case vbDate  'date
                Scriptout strObjPart & " = #" & oProperty.Value & "#", nIndent
            Case vbVariant
                If Not (IsEmpty(oProperty.Value) Or IsNull(oProperty.Value)) Then
                    Scriptout strObjPart & " = " & oProperty.Value, nIndent
                End If
            Case Else
                Scriptout strObjPart & " = " & oProperty.Value, nIndent
        End Select
    Else
        Scriptout "'Readonly property: " + strObjectName + "." + oProperty.Name + " = " & oProperty.Value, nIndent
    End If
End Sub

'********************************************************************
'* Subroutine:  ScriptStringValue
'*
'* Description: a special function to script a string property. With string properties
'* we have to deal with quotes and carriage returns.
'*
'* Parameters:
'*      strObjectName - the name of the object being scripted
'*      strValue - the string value to script
'*      nIndent - number of tabs to indent the line
'*      strComment - Optional. A comment to be added on the same line.
'********************************************************************
Private Sub ScriptStringValue(ByVal strObjectName As String, ByVal strValue As String, _
                                   Optional ByVal nIndent As Integer, Optional ByVal strComment As String)

  Dim strtemp As String

    If strComment <> "" And InStr(strComment, "'") = 0 Then
        strComment = " '" & strComment
    End If
    If strValue <> "" Then 'skip blank properties - that's the default anyway
        strtemp = strReplaceSubString(strValue, QUOTE, QUOTE & QUOTE) 'double the quotes
        If Right$(strValue, Len(vbCrLf)) = vbCrLf Then 'handle strings that end in CRLF
            If Len(strValue) = Len(vbCrLf) Then
                Scriptout strObjectName & " = vbCrLf" & strComment, nIndent
            Else
                strtemp = Left(strtemp, Len(strtemp) - Len(vbCrLf)) & QUOTE & " & vbCRLF"
                'replace all internal CRLF's with vbCrLf
                'strTemp = strReplaceSubString(strTemp, vbCrLf, QUOTE & " & vbCRLF _ " & vbCrLf & Space((nIndent + 1) * TABSIZE) & "& " & QUOTE)
                strtemp = strReplaceSubString(strtemp, vbCrLf, QUOTE & " & vbCRLF" & vbCrLf & Space(nIndent * TABSIZE) & strObjectName & " = " & strObjectName & " & " & QUOTE)
                Scriptout strObjectName & " = " & QUOTE & strtemp & strComment, nIndent
            End If
        Else
            If InStr(strtemp, vbCrLf) > 0 Then  'replace all internal CRLF's with vbCrLf
                strtemp = strReplaceSubString(strtemp, vbCrLf, QUOTE & " & vbCRLF" & vbCrLf & Space(nIndent * TABSIZE) & strObjectName & " = " & strObjectName & " & " & QUOTE)
            End If
            Scriptout strObjectName & " = " & QUOTE & strtemp & QUOTE & strComment, nIndent
        End If
    End If
End Sub

'*****************************************************************************
'Sub to fix SQL Statements that have CRLF in them

Private Function strFixSQL(ByVal strIn As String) As String
Dim istrFindlen As Integer
Dim istrInLen As Integer
Dim iPos As Integer
Dim ilastPos As Integer
Dim strNewChunk As String
Dim strOut As String
Dim strtemp As String

    'Check String Length
    istrInLen = Len(strIn)
    
    'check args
    If (istrInLen = 0) Then
        strFixSQL = strIn
        Exit Function
    End If
    
    'strFind = vbCrLf
    
    If Right$(strIn, Len(vbCrLf)) = vbCrLf Then 'handle strings that end in CRLF
            strtemp = Left(strtemp, Len(strtemp) - Len(vbCrLf)) & QUOTE & " & vbCRLF"
    End If
            'replace all internal CRLF's with vbCrLf
    strtemp = strReplaceSubString(strIn, vbCrLf, QUOTE & "  _ " & vbCrLf & Space((2) * TABSIZE) & "& " & QUOTE)
    strOut = QUOTE & strtemp & QUOTE
           
    strFixSQL = strOut
  Debug.Print strFixSQL
End Function

'********************************************************************
'* Subroutine:  strReplaceSubString$
'*
'* Description: a simple string manipulation function to replace one substring
'* with another in a string
'*
'* Parameters:
'*      strIn - the main string
'*      strFind - the sub string to find and replace
'*      strReplace - the substring to replace it with
'********************************************************************
Private Function strReplaceSubString(ByVal strIn As String, ByVal strFind As String, ByVal strReplace As String) As String
Dim istrFindlen As Integer
Dim istrInLen As Integer
Dim iPos As Integer
Dim ilastPos As Integer
Dim strNewChunk As String
Dim strOut As String

    'get the lengths of the strings
    istrFindlen = Len(strFind)
    istrInLen = Len(strIn)
    
    'check args
    If (istrFindlen = 0) Or (istrFindlen = 0) Then
        strReplaceSubString = strIn
        Exit Function
    End If
    
    ilastPos = 1
    strOut = strIn  'don't modify the input arg
    iPos = InStr(strOut, strFind) 'find position of first substring
    While iPos > 0
        strNewChunk = Mid(strOut, ilastPos, iPos - 1)  'string from last pos up to next substring
        strOut = strNewChunk + strReplace + Mid(strOut, iPos + istrFindlen)
        iPos = InStr(iPos + Len(strReplace), strOut, strFind)
    Wend
    strReplaceSubString = strOut
End Function

Private Sub ScriptAddConnection()
  
    Scriptout "Private Sub AddConnection(Byval pProvider as string, ByVal strConnectionName As String, _"
    Scriptout "                          ByVal lonConnectionID As Long, _"
    Scriptout "                          ByVal pDataSource As String, _"
    Scriptout "                          Byval pCatalog as string, Byval pUserID as string, _"
    Scriptout "                          Optional byval pPwd as string = " & QUOTE & QUOTE & ", _"
    Scriptout "                          Optional byval pNTSecurity as Boolean = false, _"
    Scriptout "                          Optional byval pConnectImmediate as Boolean = false, _"
    Scriptout "                          Optional byval pConnectionTimeout as Long = 0, _"
    Scriptout "                          Optional ByVal pReusable As Boolean = False)"
    Scriptout ""
    Scriptout ""
    Scriptout "Dim oConnection As DTS.Connection", 1
    Scriptout ""
    Scriptout "Set oConnection = goPackage.Connections.New(pProvider)", 1
    Scriptout ""
    Scriptout "With oConnection", 1
    Scriptout ".Name = strConnectionName", 2
    Scriptout ".ID = lonConnectionID", 2
    Scriptout ".Reusable = pReusable", 2
    Scriptout ".ConnectImmediate = pConnectImmediate", 2
    Scriptout ".DataSource = pDataSource", 2
    Scriptout ".UserID = pUserID", 2
    Scriptout ".Password = pPwd", 2
    Scriptout ".ConnectionTimeout = pConnectionTimeout", 2
    Scriptout ".Catalog = pCatalog", 2
    Scriptout ".UseTrustedConnection = pNTSecurity", 2
    Scriptout "End With", 1
    Scriptout ""
    Scriptout "goPackage.Connections.Add oConnection", 1
    Scriptout "Set oConnection = Nothing", 1
    Scriptout ""
    Scriptout "End Sub"
    
End Sub

Private Sub ScriptParseConnection()
  Scriptout ""
  Scriptout "Private Sub ParseConnectionString(ByVal strConnect As String)"
  Scriptout ""
  Scriptout "Dim i As Long", 1
  Scriptout "Dim strProps() As String", 1
  Scriptout "Dim strpropName As String", 1
  Scriptout "Dim strPropVal As String", 1
  Scriptout "Dim lPos As Long", 1
  Scriptout ""
  Scriptout "m_Provider = " & QUOTE & QUOTE, 1
  Scriptout "m_DataSource = " & QUOTE & QUOTE, 1
  Scriptout "m_UserID = " & QUOTE & "admin" & QUOTE, 1
  Scriptout "m_Pwd = " & QUOTE & QUOTE, 1
  Scriptout "m_WorkStation = " & QUOTE & QUOTE, 1
  Scriptout "m_Catalog = " & QUOTE & QUOTE, 1
  Scriptout "m_UseNTSecurity = False", 1
  Scriptout ""
  Scriptout "strProps = Split(strConnect, " & QUOTE & ";" & QUOTE & ", , vbTextCompare)", 1
  Scriptout ""
  Scriptout "For i = 0 To UBound(strProps)", 1
  Scriptout "lPos = InStr(1, strProps(i), " & QUOTE & "=" & QUOTE & ")", 2
  Scriptout "If Not lPos = 0 Then", 2
  Scriptout "strpropName = Left(strProps(i), lPos - 1)", 3
  Scriptout "strPropVal = Mid(strProps(i), lPos + 1)", 3
  Scriptout ""
  Scriptout "Select Case strpropName", 3
  Scriptout "Case " & QUOTE & "Provider" & QUOTE, 4
  Scriptout "m_Provider = strPropVal", 5
  Scriptout "If m_Provider = " & QUOTE & "SQLOLEDB.1" & QUOTE & " Then m_UserID = " & QUOTE & "sa" & QUOTE, 5
  Scriptout "Case " & QUOTE & "User ID" & QUOTE, 4
  Scriptout "If Not Len(strPropVal) = 0 Then m_UserID = strPropVal", 5
  Scriptout "Case " & QUOTE & "Data Source" & QUOTE, 4
  Scriptout "m_DataSource = strPropVal", 5
  Scriptout "Case " & QUOTE & "Initial Catalog" & QUOTE, 4
  Scriptout "m_Catalog = strPropVal", 5
  Scriptout "Case " & QUOTE & "WorkStation ID" & QUOTE, 4
  Scriptout "m_WorkStation = strPropVal", 5
  Scriptout "Case " & QUOTE & "Persist Security Info" & QUOTE, 4
  Scriptout "If strpropval = " & QUOTE & "SSP1" & QUOTE & " And m_Provider = " & QUOTE & "SQLOLEDB.1" & QUOTE & " Then", 5
  Scriptout "m_UseNTSecurity = True", 6
  Scriptout "Else", 5
  Scriptout "m_UseNTSecurity = strpropval", 6
  Scriptout "End If", 5
  Scriptout "Case " & QUOTE & "Password" & QUOTE, 4
  Scriptout "m_Pwd = strPropVal", 5
  Scriptout "Case " & QUOTE & "ConnectionTimeout" & QUOTE, 4
  Scriptout "m_ConnectionTimeout = strPropVal", 5
  Scriptout "End Select", 3
  Scriptout ""
  Scriptout "End If", 2
  Scriptout "    "
  Scriptout "Next i", 1
  Scriptout ""
  Scriptout "End Sub"
  Scriptout ""
End Sub

Private Sub ScriptAddStep()
  Scriptout ""
  Scriptout "Private Sub CreateStep(ByVal pStepName As String, ByVal pStepDescription As String, _"
  Scriptout "                       ByVal pTaskName As String, Optional ByVal pCloseConnection As Boolean = False, _"
  Scriptout "                       Optional ByVal pCommitOnSuccess = False, Optional ByVal pRollbackOnFailure = False)"
  Scriptout ""
  Scriptout "Dim oStep As DTS.Step", 1
  Scriptout ""
  Scriptout "Set oStep = goPackage.Steps.New", 1
  Scriptout ""
  Scriptout "With oStep", 1
  Scriptout ".Name = pStepName", 2
  Scriptout ".Description = pStepDescription", 2
  Scriptout ".ExecutionStatus = 1", 2
  Scriptout ".TaskName = pTaskName", 2
  Scriptout ".CommitSuccess = pCommitOnSuccess", 2
  Scriptout ".RollbackFailure = pRollbackOnFailure", 2
  Scriptout ".ScriptLanguage = ""VBScript""", 2
  Scriptout ".AddGlobalVariables = True", 2
  Scriptout ".RelativePriority = 3", 2
  Scriptout ".CloseConnection = pCloseConnection", 2
  Scriptout ".ExecuteInMainThread = True", 2
  Scriptout ".IsPackageDSORowset = False", 2
  Scriptout ".JoinTransactionIfPresent = False", 2
  Scriptout ".DisableStep = False", 2
  Scriptout "End With", 1
  Scriptout ""
  Scriptout "goPackage.Steps.Add oStep", 1
  Scriptout ""
  Scriptout "Set oStep = Nothing", 1
  Scriptout ""
  Scriptout "End Sub"
  Scriptout ""
End Sub

Private Sub ScriptAddSQLTask()

  Scriptout ""
  Scriptout "Private Sub AddExecuteSQLTask(ByVal pTaskName As String, ByVal pTaskDescription As String, _"
  Scriptout "                              ByVal pSQL As String, ByVal pConnectionID As Long, Optional Byval pConnectionTimeout = 0)"
  Scriptout ""
  Scriptout "Dim oCustomTask As DTS.ExecuteSQLTask", 1
  Scriptout "Dim oTask As DTS.Task", 1
  Scriptout ""
  Scriptout "Set oTask = goPackage.Tasks.New(""DTSExecuteSQLTask"")", 1
  Scriptout "      "
  Scriptout "Set oCustomTask = oTask.CustomTask", 1
  Scriptout "  "
  Scriptout "oCustomTask.Name = pTaskName", 2
  Scriptout "oCustomTask.Description = pTaskDescription", 2
  Scriptout "oCustomTask.SQLStatement = pSQL", 2
  Scriptout "oCustomTask.ConnectionID = pConnectionID", 2
  Scriptout "oCustomTask.CommandTimeout = pConnectiontimeout", 2
  Scriptout "  "
  Scriptout "goPackage.Tasks.Add oTask", 1
  Scriptout "  "
  Scriptout "Set oCustomTask = Nothing", 1
  Scriptout "Set oTask = Nothing", 1
  Scriptout ""
  Scriptout "End Sub"
  Scriptout ""
  
End Sub

Private Sub ScriptAddColumnTransformation()

  Scriptout ""
  Scriptout "Private Sub AddColumnTransformation(ByRef oCustomTask As DTS.DataPumpTask, _"
  Scriptout "                                    ByVal pTransformName As String, _"
  Scriptout "                                    ByVal pSourceColName As String, ByVal pDestColName As String, _"
  Scriptout "                                    ByVal pSourceDatatype As Long, _"
  Scriptout "                                    ByVal pDestDataType As Long, ByVal pSourceFieldSize As Long, _"
  Scriptout "                                    ByVal pDestFieldSize As Long, ByVal pSourceFlags As Long, _"
  Scriptout "                                    ByVal pDestFlags As Long, _"
  Scriptout "                                    Optional ByVal pSourceNullable As Boolean = false, _"
  Scriptout "                                    Optional ByVal pDestNullable As Boolean = false, _"
  Scriptout "                                    Optional Byval pSourceOrdinal as long = 1, _"
  Scriptout "                                    Optional byval pDestOrdinal as long = 1)"
  Scriptout ""
  Scriptout "Dim oColumn As DTS.Column", 1
  Scriptout "Dim oTransformation As DTS.Transformation", 1
  Scriptout "  "
  Scriptout "Set oTransformation = oCustomTask.Transformations.New(" & QUOTE & "DTS.DataPumpTransformCopy.1" & QUOTE & ")", 1
  Scriptout "oTransformation.Name = pTransformName", 2
  Scriptout "oTransformation.TransformFlags = 63", 2
  Scriptout "oTransformation.ForceSourceBlobsBuffered = 0", 2
  Scriptout "oTransformation.ForceBlobsInMemory = False", 2
  Scriptout "oTransformation.InMemoryBlobSize = 1048576", 2
  Scriptout "      "
  Scriptout "Set oColumn = oTransformation.SourceColumns.New(pSourceColName, 1)", 1
  Scriptout "          "
  Scriptout "oColumn.Name = pSourceColName", 2
  Scriptout "oColumn.Ordinal = pSourceOrdinal", 2
  Scriptout "oColumn.Flags = pSourceFlags", 2
  Scriptout "oColumn.Size = pSourceFieldSize", 2
  Scriptout "oColumn.DataType = pSourceDatatype", 2
  Scriptout "oColumn.Precision = 0", 2
  Scriptout "oColumn.NumericScale = 0", 2
  Scriptout "oColumn.Nullable = pSourceNullable", 2
  Scriptout "        "
  Scriptout "oTransformation.SourceColumns.Add oColumn", 1
  Scriptout "        "
  Scriptout "Set oColumn = Nothing", 1
  Scriptout ""
  Scriptout "Set oColumn = oTransformation.DestinationColumns.New(pDestColName, 1)", 1
  Scriptout "          "
  Scriptout "oColumn.Name = pDestColName", 2
  Scriptout "oColumn.Ordinal = pDestOrdinal", 2
  Scriptout "oColumn.Flags = pDestFlags", 2
  Scriptout "oColumn.Size = pDestFieldSize", 2
  Scriptout "oColumn.DataType = pDestDataType", 2
  Scriptout "oColumn.Precision = 0", 2
  Scriptout "oColumn.NumericScale = 0", 2
  Scriptout "oColumn.Nullable = pDestNullable", 2
  Scriptout "        "
  Scriptout "oTransformation.DestinationColumns.Add oColumn", 1
  Scriptout "        "
  Scriptout "Set oColumn = Nothing", 1
  Scriptout "  "
  Scriptout "oCustomTask.Transformations.Add oTransformation", 1
  Scriptout "Set oTransformation = Nothing", 1
  Scriptout "      "
  Scriptout "End Sub"
  Scriptout ""
End Sub

Private Sub ScriptAddConstraint()
  Scriptout ""
  Scriptout "Private Sub AddPrecedenceConstraint(ByVal pStep As String, ByVal priorStep As String, _"
  Scriptout "                                    Byval pConstraintBasis as long, _"
  Scriptout "                                    ByVal pConstraintResult As Long)"
  Scriptout ""
  Scriptout "Dim oStep As DTS.Step", 1
  Scriptout "Dim oPrecConstraint As DTS.PrecedenceConstraint", 1
  Scriptout ""
  Scriptout "Set oStep = goPackage.Steps(pStep)", 1
  Scriptout "Set oPrecConstraint = oStep.PrecedenceConstraints.New(priorStep)", 1
  Scriptout "oPrecConstraint.StepName = priorStep", 2
  Scriptout "oPrecConstraint.PrecedenceBasis = pConstraintBasis", 2
  Scriptout "oPrecConstraint.Value = pConstraintResult", 2
  Scriptout "oStep.PrecedenceConstraints.Add oPrecConstraint", 1
  Scriptout "Set oPrecConstraint = Nothing", 1
  Scriptout ""
  Scriptout "End Sub"
  Scriptout ""
End Sub

Private Sub ScriptCreateCustomTask()
  Scriptout ""
  Scriptout "Private Sub CreateCustomTask(ByRef pTask As DTS.Task, ByVal pTaskName As String, _"
  Scriptout "                             ByVal pTaskDescription As String, ByVal pSourceConnectionID As Long, _"
  Scriptout "                             ByVal pSourceObjectName As String, ByVal pDestConnectionID As Long, _"
  Scriptout "                             ByVal pDestObjectName As String, _"
  Scriptout "                             Optional ByVal pSourceSQL As String = """", _"
  Scriptout "                             Optional ByVal pDestSQL As String = """", _"
  Scriptout "                             Optional ByVal pFetchBuffer As Long = 1, _"
  Scriptout "                             Optional ByVal pProgressCount As Long = 1000, _"
  Scriptout "                             Optional ByVal pFastLoad As Boolean = True, _"
  Scriptout "                             Optional ByVal pFastLoadOptions AS Long = 2, _"
  Scriptout "                             Optional ByVal pMaxErrors As Integer = 0, _"
  Scriptout "                             Optional ByVal pAllowIdentityInserts As Boolean = False, _"
  Scriptout "                             Optional ByVal pFirstRow AS Long = 0, _"
  Scriptout "                             Optional ByVal pLastRow AS Long = 0)"
  Scriptout ""
  Scriptout "pTask.CustomTask.Properties(""Name"") = pTaskName", 1
  Scriptout "pTask.CustomTask.Properties(""Description"") = pTaskDescription", 1
  Scriptout "pTask.CustomTask.Properties(""SourceConnectionID"") = pSourceConnectionID", 1
  Scriptout "pTask.CustomTask.Properties(""DestinationConnectionID"") = pDestConnectionID", 1
  Scriptout "  "
  Scriptout "If pSourceSQL = " & QUOTE & QUOTE & " Then", 1
  Scriptout "pTask.CustomTask.Properties(""SourceObjectName"") = pSourceObjectName", 2
  Scriptout "Else", 1
  Scriptout "pTask.CustomTask.Properties(""SourceSQLStatement"") = pSourceSQL", 2
  Scriptout "End If", 1
  Scriptout "  "
  Scriptout "If pDestSQL = " & QUOTE & QUOTE & " Then", 1
  Scriptout "pTask.CustomTask.Properties(""DestinationObjectName"") = pDestObjectName", 2
  Scriptout "Else", 1
  Scriptout "pTask.CustomTask.Properties(""DestinationSQLStatement"") = pDestSQL", 2
  Scriptout "End If", 1
  Scriptout "  "
  Scriptout "pTask.CustomTask.Properties(""ProgressRowCount"") = pProgressCount", 1
  Scriptout "pTask.CustomTask.Properties(""MaximumErrorCount"") = pMaxErrors", 1
  Scriptout "pTask.CustomTask.Properties(""FetchBufferSize"") = pFetchBuffer", 1
  Scriptout "pTask.CustomTask.Properties(""UseFastLoad"") = pFastLoad", 1
  Scriptout "pTask.CustomTask.Properties(""InsertCommitSize"") = 0", 1
  Scriptout "pTask.CustomTask.Properties(""ExceptionFileColumnDelimiter"") = ""|""", 1
  Scriptout "pTask.CustomTask.Properties(""ExceptionFileRowDelimiter"") = vbCrLf", 1
  Scriptout "pTask.CustomTask.Properties(""AllowIdentityInserts"") = pAllowIdentityInserts", 1
  Scriptout "pTask.CustomTask.Properties(""FirstRow"") = pFirstRow", 1
  Scriptout "pTask.CustomTask.Properties(""LastRow"") = pLastRow", 1
  Scriptout "pTask.CustomTask.Properties(""FastLoadOptions"") = pFastLoadOptions", 1
  Scriptout ""
  Scriptout "End Sub"
  Scriptout ""
  
End Sub
Private Sub ScriptoutExecutePkg()
  Scriptout ""
  Scriptout "Public Function ExecutePackage() As Boolean"
  Scriptout "Dim errStep As Long", 1
  Scriptout ""
  Scriptout "On Error Resume Next"
  Scriptout ""
  Scriptout "'Establish the count for progress events"
  Scriptout "m_Totaltasks = goPackage.Tasks.Count", 1
  Scriptout "m_TasksCompleted = 0", 1
  Scriptout ""
  Scriptout "RaiseEvent PercentDone(0)", 1
  Scriptout "  "
  Scriptout "ExecutePackage = True", 1
  Scriptout ""
  Scriptout "goPackage.Execute", 1
  Scriptout ""
  Scriptout "For errStep = 1 To goPackage.Steps.Count", 1
  Scriptout "If goPackage.Steps(errStep).ExecutionResult = DTSStepExecResult_Failure Then", 2
  Scriptout "ExecutePackage = False", 3
  Scriptout "m_Cancel = True", 3
  Scriptout "Debug.Print " & QUOTE & "Step " & QUOTE & "goPackage.Steps(errStep).Name" & QUOTE & " has failed " & QUOTE & " _", 3
  Scriptout "& vbCrLf " & QUOTE & "Affected tables have been rolled back to previous state." & QUOTE, 3
  Scriptout "End If", 2
  Scriptout "Next errStep", 1
  Scriptout ""
  Scriptout "RaiseEvent PercentDone(100)", 1
  Scriptout ""
  Scriptout "End Function"
End Sub

Private Sub ScriptoutClassTerminate()
  Scriptout ""
  Scriptout "Private Sub Class_Terminate()"
  Scriptout "Set goPackage = Nothing", 1
  Scriptout "End Sub"
  Scriptout ""
End Sub

Private Sub ScriptOutPkgErr()
  Scriptout ""
  Scriptout "Private Sub goPackage_OnError(ByVal EventSource As String, ByVal ErrorCode As Long, ByVal source As String, ByVal Description As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal IDofInterfaceWithError As String, pbCancel As Boolean)"
  Scriptout ""
  Scriptout "RaiseEvent ErrorOccurred(ErrorCode, source, Description)", 1
  Scriptout ""
  Scriptout "End Sub"
End Sub

Private Sub ScriptoutPkgFinish()
  Scriptout ""
  Scriptout "Private Sub goPackage_OnFinish(ByVal EventSource As String)"
  Scriptout "m_TasksCompleted = m_TasksCompleted + 1", 1
  Scriptout "RaiseEvent PercentDone((m_TasksCompleted / m_Totaltasks) * 100)", 1
  Scriptout "RaiseEvent Currenttask(" & QUOTE & QUOTE & ")", 1
  Scriptout "DoEvents", 1
  Scriptout "End Sub"
End Sub

Private Sub ScriptoutPkgProgress()
  Scriptout ""
  Scriptout "Private Sub goPackage_OnProgress(ByVal EventSource As String, ByVal ProgressDescription As String, ByVal PercentComplete As Long, ByVal ProgressCountLow As Long, ByVal ProgressCountHigh As Long)"
  Scriptout "RaiseEvent RowsCopied(ProgressDescription)", 1
  Scriptout "DoEvents", 1
  Scriptout "End Sub"
End Sub

Private Sub ScriptoutPkgCancel()
  Scriptout ""
  Scriptout "Private Sub goPackage_OnQueryCancel(ByVal EventSource As String, pbCancel As Boolean)"
  Scriptout "If m_Cancel = True Then", 1
  Scriptout "pbCancel = True", 2
  Scriptout "Else", 1
  Scriptout "pbCancel = False", 2
  Scriptout "End If", 1
  Scriptout "End Sub"
End Sub

Private Sub ScriptoutPkgStart()
  Scriptout ""
  Scriptout "Private Sub goPackage_OnStart(ByVal EventSource As String)"
  Scriptout "RaiseEvent Currenttask(EventSource)", 1
  Scriptout "DoEvents", 1
  Scriptout "End Sub"
End Sub

Private Sub ScriptClassModule(Optional ByVal strname As String = "ClassDTSScript")
  Scriptout "VERSION 1.0 CLASS"
  Scriptout "BEGIN"
  Scriptout "MultiUse = -1  'True"
  Scriptout "Persistable = 0  'NotPersistable"
  Scriptout "DataBindingBehavior = 0  'vbNone"
  Scriptout "DataSourceBehavior = 0   'vbNone"
  Scriptout "MTSTransactionMode = 0   'NotAnMTSObject"
  Scriptout "End"
  Scriptout "Attribute VB_Name = " & QUOTE & strname & QUOTE
  Scriptout "Attribute VB_GlobalNameSpace = False"
  Scriptout "Attribute VB_Creatable = True"
  Scriptout "Attribute VB_PredeclaredId = False"
  Scriptout "Attribute VB_Exposed = False"
  Scriptout ""
End Sub

Private Sub ScriptSavePkg()
  Scriptout "Public Sub SavePackage()"
  Scriptout "goPackage.SaveToSQLServer " & QUOTE & "(local)" & QUOTE & ", " & QUOTE & "sa" & QUOTE & ", " & QUOTE & QUOTE
  Scriptout "End Sub"
  Scriptout ""
End Sub
'********************************************************************
'* Subroutine:  Scriptout
'*
'* Description: writes a line of text to the script file at the specified indent level
'*
'* Parameters:
'*      strText - the text to write
'*      nIndent - the number of tabs to indent (tabs are replaced with spaces)
'********************************************************************
Sub Scriptout(ByVal strText As String, Optional nIndent As Integer)
Dim nSpace%
    nSpace = nIndent * TABSIZE
    Print #gnScriptFile, Space(nSpace) & strText
End Sub

'********************************************************************
'* Subroutine:  FAIL
'*
'* Description: Pops up a message box and stops the program
'*
'* Parameters: strMessage - the text in the messagebox
'********************************************************************
Private Sub FAIL(ByVal strMessage As String)
    MsgBox strMessage
    End
End Sub

'********************************************************************
'* Subroutine:  LoadServerPkg
'*
'* Description: Loads the DTS Package from the local server into the
'* global variable goPackage
'*
'* Parameters: none
'********************************************************************
Private Sub LoadServerPkg(ByVal strPkgName As String)
  
  Dim strServerName As String
  Dim strLogin As String
  Dim strPwd As String
  Dim bTrusted As Boolean
  Dim strPkgGuid As String
  Dim strPkgVersion As String

     
    Set goPackage = Nothing
    
    'set server info
    strServerName = "(local)" 'assumes local server, change this as you like or add a UI to set it.
    strLogin = "sa"
    strPwd = ""
    bTrusted = False
    If strLogin = "" Then bTrusted = True
           
  On Error GoTo LoadErr
    goPackage.LoadFromSQLServer strServerName, strLogin, strPwd, bTrusted, , , , strPkgName
        
     'Here's how to change this to load packages from the repository instead:
    'dim strPkgID
    'strPkgID = InputBox("Enter the package guid from the package properties dialog of the DTS Designer")
    'goPackage.LoadFromRepository strServerName, "msdb", strLogin, strPwd, strPkgID
  
          
Exit Sub
LoadErr:
        FAIL "Unable to load package " & strPkgName & ". Error: " & Error$
End Sub


