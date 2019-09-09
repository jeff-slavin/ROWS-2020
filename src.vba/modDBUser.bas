Option Explicit

Public Function DBUser_GetDefaultRolePermissions(ByVal sRole As String, ByRef asDefaultPermissions() As String) As Messages
'Returns the default permissions for the given sRole parameter

     'Variable Declarations
     Dim cROWSDB As New clsROWSDB
     Dim sSQL As String
     Dim i As Integer
     
     'Set initial states
     DBUser_GetDefaultRolePermissions = Messages.msgFalse
     If IsArray(asDefaultPermissions) Then Erase asDefaultPermissions
     
     'Build the SQL statement to get the default permissions for the given role
     sSQL = ""
     sSQL = sSQL & "SELECT [tblPermissionList].sPermissionName "
     sSQL = sSQL & "FROM [tblPermissionList], [tblUserRoles], [tblDefaultRolePermissions] "
     sSQL = sSQL & "WHERE [tblUserRoles].sRoleName = '" & sRole & "' "
     sSQL = sSQL & "AND [tblDefaultRolePermissions].iUserRoleID = [tblUserRoles].ID "
     sSQL = sSQL & "AND [tblDefaultRolePermissions].iPermissionID = [tblPermissionList].ID "
     sSQL = sSQL & "AND [tblDefaultRolePermissions].bIsActive = TRUE "
     sSQL = sSQL & "AND [tblPermissionList].bIsActive = TRUE;"
     
     DBUser_GetDefaultRolePermissions = cROWSDB.Query(sSQL, True)
     
     'Check for Error
     If DBUser_GetDefaultRolePermissions <> Messages.msgTrue Then GoTo DBUser_GetDefaultRolePermissions_Error
     
     'See if we have a record returned
     If cROWSDB.RecordCount < 1 Then
          DBUser_GetDefaultRolePermissions = Messages.msgFalse
          GoTo DBUser_GetDefaultRolePermissions_Error
     End If
     
     'We now have a valid response with a RecordCount > 0
     i = 1
     cROWSDB.MoveFirst
     ReDim asDefaultPermissions(1 To cROWSDB.RecordCount)
     
     While Not cROWSDB.EOF
     
          Call cROWSDB.Fields("sPermissionName", asDefaultPermissions(i))
     
          i = i + 1
          cROWSDB.MoveNext
     Wend

DBUser_GetDefaultRolePermissions_Error:
     'Clear memory
     Set cROWSDB = Nothing

End Function

Public Function DBUser_GetAllActivePermissions(ByRef asPermissions() As String) As Messages
'Returns information around all active permissions in the 'asPermissions' ByRef parameter

     'Variable Declarations
     Dim cROWSDB As New clsROWSDB
     Dim sSQL As String
     Dim i As Integer
     
     'Set initial states
     DBUser_GetAllActivePermissions = Messages.msgFalse
     If IsArray(asPermissions) Then Erase asPermissions
     
     'Build the SQL statement to get all active permissions
     sSQL = ""
     sSQL = sSQL & "SELECT [tblPermissionCategories].sPermissionCategory, [tblPermissionList].sPermissionName "
     sSQL = sSQL & "FROM [tblPermissionList], [tblPermissionCategories] "
     sSQL = sSQL & "WHERE [tblPermissionList].bIsActive = TRUE "
     sSQL = sSQL & "AND [tblPermissionCategories].bIsActive = TRUE "
     sSQL = sSQL & "AND [tblPermissionList].iPermissionCategoryID = [tblPermissionCategories].ID;"
     
     'Run the query
     DBUser_GetAllActivePermissions = cROWSDB.Query(sSQL, True)
     
     'Check for error
     If DBUser_GetAllActivePermissions <> Messages.msgTrue Then GoTo DBUser_GetAllActivePermissions_Error
     
     'See if we have a record returned
     If cROWSDB.RecordCount < 1 Then
          DBUser_GetAllActivePermissions = Messages.msgFalse
          GoTo DBUser_GetAllActivePermissions_Error
     End If
     
     'We have a response with values returned (RecordCount > 0)
     i = 1
     cROWSDB.MoveFirst
     ReDim asPermissions(1 To cROWSDB.RecordCount, 1 To 2)
     
     While Not cROWSDB.EOF
     
          Call cROWSDB.Fields("sPermissionCategory", asPermissions(i, 1))
          Call cROWSDB.Fields("sPermissionName", asPermissions(i, 2))
     
          i = i + 1
          cROWSDB.MoveNext
     Wend
     
DBUser_GetAllActivePermissions_Error:
     'Clear memory
     Set cROWSDB = Nothing

End Function

Public Function DBUser_GetRoleRank(ByVal sUsername As String, ByRef iRoleRank As Integer) As Messages

     'Variable Declarations
     Dim cROWSDB As New clsROWSDB
     Dim sSQL As String
     
     'Build the SQL statement to grab the role rank for the sUsername parameter
     sSQL = ""
     sSQL = sSQL & "SELECT [tblUserRoles].iRank "
     sSQL = sSQL & "FROM [tblUserRoles], [tblUsers] "
     sSQL = sSQL & "WHERE [tblUsers].sUsername = '" & sUsername & "' "
     sSQL = sSQL & "AND [tblUsers].iRoleID = [tblUserRoles].ID;"
     
     'Run the query
     DBUser_GetRoleRank = cROWSDB.Query(sSQL, True)
     
     'Check for error
     If DBUser_GetRoleRank <> Messages.msgTrue Then GoTo DBUser_GetRoleRank_Error
     
     'Check if a record was returned
     If cROWSDB.RecordCount < 1 Then
          DBUser_GetRoleRank = Messages.msgFalse
          GoTo DBUser_GetRoleRank_Error
     End If
     
     'Now set the role rank returned
     If cROWSDB.Fields("iRank", iRoleRank) = False Then DBUser_GetRoleRank = Messages.msgFalse
     
DBUser_GetRoleRank_Error:
     
     Set cROWSDB = Nothing

End Function

Public Function DBUser_GetLesserRoles(ByVal sUsername As String, ByRef asLesserRoles() As String) As Messages

     'Variable Declaration
     Dim cROWSDB As New clsROWSDB
     Dim sSQL As String
     Dim iRoleRank As Integer
     Dim i As Integer
     
     'Build the SQL statement to grab the lower ranked roles
     'Note: Roles are ranked where the higher ones have a higher Rank (e.g. Administrator = 1)
     'So while we say lower ranked roles, we're really grabbing roles with a higher rank number
     
     'clear out ByRef parameter
     If IsArray(asLesserRoles) Then Erase asLesserRoles
     
     'Get current Role Rank First
     DBUser_GetLesserRoles = DBUser_GetRoleRank(sUsername, iRoleRank)
     
     'check for error & let error bubble up
     If DBUser_GetLesserRoles <> msgTrue Then GoTo DBUser_GetLesserRoles_Error
     
'If the user is not an administrator find the lesser roles
     If g_cUser.Role <> "Administrator" Then
          'Now set SQL statement to find the roles that are lesser than our current iRoleRank variable
          'Using greater than (>) sign as lesser roles actually have a higher iRank number (e.g. Administrator is 1, everything else is higher)
          sSQL = ""
          sSQL = sSQL & "SELECT [tblUserRoles].sRoleName "
          sSQL = sSQL & "FROM [tblUserRoles] "
          sSQL = sSQL & "WHERE [tblUserRoles].iRank > " & iRoleRank & " "
          sSQL = sSQL & "AND [tblUserRoles].bIsActive = TRUE;"
     Else
          'User is an administrator, so get all active roles
          sSQL = ""
          sSQL = sSQL & "SELECT [tblUserRoles].sRoleName "
          sSQL = sSQL & "FROM [tblUserRoles] "
          sSQL = sSQL & "WHERE [tblUserRoles].bIsActive = TRUE;"
     End If
     
'if the user is an administrator then find all roles

     
     'Run the query
     DBUser_GetLesserRoles = cROWSDB.Query(sSQL, True)
     
     'check for error
     If DBUser_GetLesserRoles <> Messages.msgTrue Then GoTo DBUser_GetLesserRoles_Error
     
     'check that we got a record in return
     If cROWSDB.RecordCount < 1 Then
          DBUser_GetLesserRoles = Messages.msgFalse
          GoTo DBUser_GetLesserRoles_Error
     End If
     
     'We have a valid return, so let's capture the roles
     cROWSDB.MoveFirst
     i = 1
     ReDim asLesserRoles(1 To cROWSDB.RecordCount)
     
     While Not cROWSDB.EOF
          
          Call cROWSDB.Fields("sRoleName", asLesserRoles(i))
          
          i = i + 1
          cROWSDB.MoveNext
     Wend
          
DBUser_GetLesserRoles_Error:
     'Clear memory
     Set cROWSDB = Nothing

End Function

Public Function DBUser_SetLastLoginDateTime(ByVal sUsername As String) As Messages

     'Variable Declaration
     Dim cROWSDB As New clsROWSDB
     Dim sSQL As String
     
     'Build the SQL statement to update just the last login time for the user
     sSQL = ""
     sSQL = sSQL & "UPDATE [tblUsers] "
     sSQL = sSQL & "SET [tblUsers].dLastLogin = Now() "
     sSQL = sSQL & "WHERE [tblUsers].sUsername = '" & sUsername & "';"
     
     DBUser_SetLastLoginDateTime = cROWSDB.Query(sSQL, False)
     
     Set cROWSDB = Nothing

End Function

Public Function DBUser_ChangePassword(ByVal sUsername As String, ByVal sNewPassword As String, ByVal bIsTempPass As Boolean) As Messages
'Update the password for sUsername
'Also set bIsTempPass to variable that was passed

     Dim cROWSDB As New clsROWSDB
     Dim sSQL As String
     
     'Build the SQL statement to update the password for the given user
     sSQL = ""
     sSQL = sSQL & "UPDATE [tblUsers] "
     sSQL = sSQL & "SET [tblUsers].sPassword = '" & sNewPassword & "', "
     sSQL = sSQL & "[tblUsers].bIsTempPass = " & bIsTempPass & " "
     sSQL = sSQL & "WHERE [tblUsers].sUsername = '" & sUsername & "';"
     
     'Since we are not requesting any information from the database
     'just return the result of the query
     DBUser_ChangePassword = cROWSDB.Query(sSQL, False)
     
     'Free memory
     Set cROWSDB = Nothing

End Function

Public Function DBUser_CheckCredentials(ByVal sUsername As String, ByVal sPassword As String, ByRef bIsTempPass As Boolean, _
                                        ByRef sRoleName As String, ByRef iRoleRank As Integer) As Messages
'Purpose of this function is to be called if you want to check that a user has entered valid credentials (username/password)
'This will validate the username/password combination as well as grab a few extra parameters about the user:
'bTempPass, sRoleName, iRoleRank - through the ByRef parameters
'While these extra parameters that get returned don't need to be used, they are useful to grab now on this one call
'in case you are loggin in a new user (instead of doing a 2nd call to get them)

     'Variable Declaration
     Dim cROWSDB As New clsROWSDB
     Dim sSQL As String
     
     'Build SQL statement
     'Find a user with matching username/password and is set to active
     'Also grab role name, role rank and if password is temp
     sSQL = ""
     sSQL = sSQL & "SELECT [tblUsers].iRoleID, [tblUsers].bIsTempPass, [tblUserRoles].sRoleName, [tblUserRoles].iRank "
     sSQL = sSQL & "FROM [tblUsers], [tblUserRoles] "
     sSQL = sSQL & "WHERE [tblUsers].sUsername = '" & sUsername & "' "
     sSQL = sSQL & "AND [tblUsers].sPassword = '" & sPassword & "' "
     sSQL = sSQL & "AND [tblUsers].bIsActive = TRUE "
     sSQL = sSQL & "AND [tblUsers].iRoleID = [tblUserRoles].ID;"
     
     'Run the query and save the resulting Message
     DBUser_CheckCredentials = cROWSDB.Query(sSQL, True)
     
     'If query failed, exit the function (returning the failed error Message)
     If DBUser_CheckCredentials <> Messages.msgTrue Then GoTo DBUser_CheckCredentials_Error
     
     'Query was successful but no matching user
     If cROWSDB.RecordCount < 1 Then
          DBUser_CheckCredentials = Messages.msgFalse
          GoTo DBUser_CheckCredentials_Error
     End If
     
     'Successful query and matching user was found
     DBUser_CheckCredentials = Messages.msgTrue
     
     'Save the ByRef parameters
     Call cROWSDB.MoveFirst
     Call cROWSDB.Fields("bIsTempPass", bIsTempPass)
     Call cROWSDB.Fields("sRoleName", sRoleName)
     Call cROWSDB.Fields("iRank", iRoleRank)
     
DBUser_CheckCredentials_Error:
     'Free up memory
     Set cROWSDB = Nothing

End Function

Public Function DBUser_GetPermissions(ByVal sUsername As String, ByRef asPermissions() As String) As Messages

     'Variable Declarations
     Dim cROWSDB As New clsROWSDB
     Dim sSQL As String
     Dim i As Integer
     
     'Clear out ByRef array parameter (so if this fails that returns as empty)
     If IsArray(asPermissions) Then Erase asPermissions

     'Create SQL statement
     sSQL = ""
     sSQL = sSQL & "SELECT [tblPermissionList].sPermissionName "
     sSQL = sSQL & "FROM [tblPermissionList], [tblUsers], [tblUserPermissions] "
     sSQL = sSQL & "WHERE [tblPermissionList].ID = [tblUserPermissions].iPermissionID "
     sSQL = sSQL & "AND [tblUserPermissions].iUserID = [tblUsers].ID "
     sSQL = sSQL & "AND [tblUsers].sUsername = '" & sUsername & "' "
     sSQL = sSQL & "AND [tblUserPermissions].bIsActive = TRUE "
     sSQL = sSQL & "AND [tblPermissionList].bIsActive = TRUE;"
     
     'Run the query
     DBUser_GetPermissions = cROWSDB.Query(sSQL, True)
     
     'Check for an error
     If DBUser_GetPermissions <> Messages.msgTrue Then GoTo DBUser_GetPermissions_Error
     
     
     'Query was successful, let's see if we got any permissions returned
     'If no permissions, let's exit the function now
     'The user can be logged in, but will not have any permissions
     If cROWSDB.RecordCount < 1 Then GoTo DBUser_GetPermissions_Error
     
     'Save the permissions in the array ByRef parameter
     cROWSDB.MoveFirst
     i = 1
     ReDim asPermissions(1 To cROWSDB.RecordCount)
     
     While Not cROWSDB.EOF
          
          Call cROWSDB.Fields("sPermissionName", asPermissions(i))
          
          i = i + 1
          cROWSDB.MoveNext
     Wend
     
DBUser_GetPermissions_Error:
     'Free up memory
     Set cROWSDB = Nothing

End Function