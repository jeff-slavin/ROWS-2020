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

'******************************************************************************
'Worksheet Range Locations
'******************************************************************************

'GLOBAL ERROR CELL
     Global Const gc_rErrorCell As String = "$C$1"

' WELCOME Worksheet

     'Welcome Sheet Sections
     Global Const gc_wsLoginRange As String = "5:14"
     Global Const gc_wsLogoutRange As String = "15:23"
     Global Const gc_wsTempPasswordRange As String = "24:35"
     
     'Welcome Sheet Button Positions
     Global Const gc_wscmdLoginTop As Integer = 151
     Global Const gc_wscmdLogoutTop As Integer = 136
     Global Const gc_wscmdTempPasswordTop As Integer = 178
     
     'Welcome Error Cells
     Global Const gc_wsWelcomeLoginError As String = "$C$6"
     Global Const gc_wsWelcomeTempPasswordError As String = "$C$25"
     
     'Welcome Cell Inputs
     Global Const gc_wsWelcomeUsername As String = "$D$7"
     Global Const gc_wsWelcomePassword As String = "$D$9"
     Global Const gc_wsWelcomeLoggedInUsername As String = "$D$17"
     Global Const gc_wsWelcomeLoggedInRole As String = "$D$18"
     Global Const gc_wsWelcomeTempPassword As String = "$E$26"
     Global Const gc_wsWelcomeNewPassword As String = "$E$28"
     Global Const gc_wsWelcomeRetypePassword As String = "$E$30"

' USER CREATE Worksheet

     Global Const gc_wsUserCreateUsername As String = "$E$7"
     Global Const gc_wsUserCreateTempPassword As String = "$E$9"
     Global Const gc_wsUserCreateRole As String = "$E$11"
     Global Const gc_wsUserCreateRoleRange As String = "$E$11:$F$11"
     Global Const gc_wsUserCreatePermissionStart As String = "$J$8"
     Global Const gc_wsUserCreatePermissionTableRange As String = "$J$8:$L$25"