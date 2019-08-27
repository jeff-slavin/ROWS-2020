Option Explicit

'Database Path (Relative)
Global Const gc_sROWSDBFilename As String = "/database/ROWS-2020.accdb"

'Global clsUser Instance
Global g_cUser As clsUser

'Function Return Statuses
Enum Messages
     msgFalse = 0
     msgTrue = 1
     msgFailedDBConnection = 2
     msgFailedQuery = 3
End Enum