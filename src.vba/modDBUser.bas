Option Explicit

Public Function DBUser_SetLastLoginDateTime(ByVal sUsername As String) As Messages

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


End Function