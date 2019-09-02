Option Explicit

'***************************************************************************************************
'Class Name:   clsUser
'---------------------------------------------------------------------------------------------------
'Purpose:
'***************************************************************************************************

'Private Member Variables
Private m_sUsername As String
Private m_sRole As String
Private m_iRoleRank As Integer
Private m_asPermissions() As String
Private m_sPermissions As String

Private Sub Class_Initialize()
'Constructor
     Call Me.Reset
End Sub

Private Sub Class_Terminate()
'Destructor
     Call Me.Reset
End Sub

Public Function GetLesserRoles() As Messages
'Returns a list of roles that are lesser than the current user's role

     Dim asLesserRoles() As String

     'Call the modDBUser Function
     GetLesserRoles = DBUser_GetLesserRoles(m_sUsername, asLesserRoles)

End Function

Public Sub Logout()
'Log the user out
'Resets all of the member variables
'Sends the user back to the login form (hides all other worksheets)
     
     Call Me.Reset

     Call wsWelcome_ShowSection("Login")
     Call Worksheets_ShowWelcomeAndOne("Welcome")

End Sub

Public Sub SavePermissionsToString()
'Saves the permissions in m_asPermissions to the string member variable m_sPermissions
     
     'Variable Declaration
     Dim i As Integer
     
     'Set initial state of member variable, no permissions
     m_sPermissions = ""
     
     'Check that we have permissions in our member array (otherwise exit)
     If Not IsArray(m_asPermissions) Then Exit Sub
     
     'Is the array valid?
     If (Not Not m_asPermissions) <> 0 Then
          'Valid Array - so let's convert this to the string
          
          For i = LBound(m_asPermissions) To UBound(m_asPermissions)
               m_sPermissions = m_sPermissions & m_asPermissions(i) & "|"
          Next i
          
     Else
          'Invalid array - exit
          
          Exit Sub
          
     End If

End Sub

Public Function Login(ByVal sUsername As String, ByVal sPassword As String) As Messages
'Attempt to login the user

     Dim bIsTempPass As Boolean
     Dim sRoleName As String
     Dim iRoleRank As Integer

     Login = DBUser_CheckCredentials(sUsername, sPassword, bIsTempPass, sRoleName, iRoleRank)

     'Check for failure, letting error message bubble up
     If Login <> Messages.msgTrue Then Exit Function
     
     'Check for temporary password for user.
     'If true, then return this and exit (without setting anything)
     'This allows the calling function to show the temp password reset
     'Until that is done, we don't want to actually fully login the user
     'Once temp password is resolved, this function can be called again
     If bIsTempPass = True Then
          Login = Messages.msgTemporaryPassword
          Exit Function
     End If
     
     'Set member variables (successful login - no temp password)
     m_sUsername = sUsername
     m_sRole = sRoleName
     m_iRoleRank = iRoleRank
     
     'Get User Permissions
     'This could return zero permissions (and so m_asPermissions would be an empty array)
     Login = DBUser_GetPermissions(m_sUsername, m_asPermissions)

     'Check for failure and set appropriate error response code
     If Login <> Messages.msgTrue Then Exit Function
     
     'Save permissions to the member string
     Call SavePermissionsToString

     'update the last login date/time in the database
     Call DBUser_SetLastLoginDateTime(sUsername)

End Function

Public Sub Reset()
'Reset the user data back to an empty state

     m_sUsername = ""
     m_sRole = ""
     m_iRoleRank = 0
     If IsArray(m_asPermissions) Then Erase m_asPermissions
     m_sPermissions = ""

End Sub

Property Get Username() As String

     Username = m_sUsername

End Property

Property Let Username(ByVal sUsername As String)

     m_sUsername = sUsername

End Property

Property Get Role() As String

     Role = m_sRole

End Property

Property Let Role(ByVal sRole As String)

     m_sRole = sRole

End Property

Property Get RoleRank() As Integer

     RoleRank = m_iRoleRank

End Property

Property Let RoleRank(ByVal iRoleRank As Integer)

     m_iRoleRank = iRoleRank

End Property

Property Get PermissionArray() As String()

     PermissionArray = m_asPermissions

End Property

Property Get Permissions() As String

     Permissions = m_sPermissions

End Property