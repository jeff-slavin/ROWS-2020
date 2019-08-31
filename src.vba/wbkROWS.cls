Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
'On closing the workbook, free the memory for any global instances of classes (e.g. cUser)

     Set g_cUser = Nothing

End Sub

Private Sub Workbook_Open()
'When the workbook opens...
'    1) Set the global instance of clsUser
'    2) Show the 'Welcome' sheet and only the login section
     
     Set g_cUser = New clsUser

     Call Worksheets_ShowWelcomeAndOne("Welcome")
     Call wsWelcome_ShowSection("Login")
     Worksheets("Welcome").Range(gc_wsWelcomeLoginError).Value = ""
     Worksheets("Welcome").Range(gc_wsWelcomeUsername).Value = ""
     Worksheets("Welcome").Range(gc_wsWelcomePassword).Value = ""
     Worksheets("Welcome").Range(gc_wsWelcomeUsername).Select

End Sub