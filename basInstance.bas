Attribute VB_Name = "basInstance"
Option Explicit

Public Const GW_HWNDPREV = 3
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'Keeps the user from starting two applications
Sub ActivatePrevInstance()

   Dim OldTitle As String
   Dim PrevHndl As Long
   Dim result As Long

   'Save the title of the application.
   OldTitle = App.Title

   'Rename the title of this application so FindWindow
   'will not find this application instance.
   App.Title = "unwanted instance"

   'Attempt to get window handle using VB4 class name.
   PrevHndl = FindWindow("ThunderRTMain", OldTitle)

    'Check if found
   If PrevHndl = 0 Then
        'Attempt to get window handle using VB6 class name
        PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
   End If

   'Check if found
   If PrevHndl = 0 Then
      'No previous instance found.
      Exit Sub
   End If

   'Get handle to previous window.
   PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)

   'Restore the program.
   result = OpenIcon(PrevHndl)

   'Activate the application.
   result = SetForegroundWindow(PrevHndl)

   'End the application.
   MsgBox "Application is already running.", vbSystemModal, "Shutting Down Application"
   End

End Sub
