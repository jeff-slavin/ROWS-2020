Option Explicit

Public Sub Test()

     Dim asPermissions() As String
     Dim msgResponse As Messages
     Dim i As Integer
     
     msgResponse = DBUser_GetPermissions("JFMS", asPermissions)
     
     MsgBox "Response : " & msgResponse
     MsgBox "Records Found : " & UBound(asPermissions) - LBound(asPermissions)
     
     For i = LBound(asPermissions) To UBound(asPermissions)
          MsgBox "Permission Found #" & Str(i) & ": " & asPermissions(i)
     Next i
     
     If IsArray(asPermissions) Then Erase asPermissions

End Sub

Public Sub PrintUserPermissions()

     Worksheets("Testing").Range("B6").Value = g_cUser.Permissions

End Sub