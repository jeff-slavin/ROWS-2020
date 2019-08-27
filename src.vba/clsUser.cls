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

Property Let PermissionArray(ByRef asPermissions() As String)

     m_asPermissions = asPermissions

End Property
