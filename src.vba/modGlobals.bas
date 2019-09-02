Option Explicit

'Database Path (Relative)
Global Const gc_sROWSDBFilename As String = "/database/ROWS-2020.accdb"

'Global clsUser Instance
Global g_cUser As clsUser

'Function Return Statuses / Custom Error Codes
Enum Messages
     msgFalse = 0
     msgTrue = 1
     msgFailedDBConnection = 2
     msgFailedQuery = 3
     msgTemporaryPassword = 4
End Enum

'Worksheet Range Locations

     'Welcome Sheet Sections
     Global Const gc_wsLoginRange As String = "5:14"
     Global Const gc_wsLogoutRange As String = "15:23"
     Global Const gc_wsTempPasswordRange As String = "24:35"
     
     'Welcome Sheet Button Positions
     Global Const gc_wscmdLoginTop As Integer = 151
     Global Const gc_wscmdLogoutTop As Integer = 136
     Global Const gc_wscmdTempPasswordTop As Integer = 178
     
     'Welcome Error Cells
     Global Const gc_wsWelcomeLoginError As String = "C6"
     Global Const gc_wsWelcomeTempPasswordError As String = "C25"
     
     'Welcome Cell Inputs
     Global Const gc_wsWelcomeUsername As String = "D7"
     Global Const gc_wsWelcomePassword As String = "D9"
     Global Const gc_wsWelcomeLoggedInUsername As String = "D17"
     Global Const gc_wsWelcomeLoggedInRole As String = "D18"
     Global Const gc_wsWelcomeTempPassword As String = "E26"
     Global Const gc_wsWelcomeNewPassword As String = "E28"
     Global Const gc_wsWelcomeRetypePassword As String = "E30"