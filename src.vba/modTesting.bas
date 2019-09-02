Option Explicit

Public Sub Test()

     Dim iRoleRank As Integer
     Dim msgResponse As Messages
     
     msgResponse = DBUser_GetRoleRank("JFMS", iRoleRank)
     
     MsgBox "Response: " & msgResponse
     MsgBox "Rank : " & iRoleRank

End Sub

Public Sub PrintUserPermissions()

     Worksheets("Testing").Range("B6").Value = g_cUser.Permissions

End Sub