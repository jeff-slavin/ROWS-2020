Option Explicit

Public Sub Test()

     Dim bTempPass As Boolean
     Dim sRoleName As String
     Dim iRoleRank As Integer
     
     Dim msgResponse As Messages
     
     msgResponse = DBUser_CheckCredentials("JFMS", "j3ffj3nn", bTempPass, sRoleName, iRoleRank)
     
     MsgBox "Response : " & msgResponse

End Sub