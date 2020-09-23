Attribute VB_Name = "basUserInformation"
Option Explicit
'*************************************************************************
'This routine is for changing the password through ADO
'
''This program is documented well with comments.
'**************************************************************************
'
'UserInformation Module
'
'Copyright: (c) 2000 Nicholas K. Kruse
'PLEASE:  If you use my code that is fine, but please remember that I spent
'a lot of time developing this in the newest technology ADO.  I would
'appreaicate you leaving my comments in here and giving me credit.  Futhermore
'If you use my code please read the .txt file you downloaded with
'these modules.  Further you can reach me at steelcoil@bellsouth.net
'This is by far the best set of modules for this procedure on the web, remember!
'
'   Contains all the globals these can be changed but caution make sure you
'   understand what you are changing
'
'
'   History.
'       26 November 2000
'       Nicholas K. Kruse
'
'**************************************************************************
'Global Variables not neccesarily written in this order
Global UserID As String
Global UserPassword As String
Global UserLName As String
Global UserFName As String
Global UserMInitial As String
Global UserExpireDate As String
Global UserActivationDate As String
Global UserTaskLevel As String
'User specific constants
Global Const EXPIRE_TERM = 180 'password expiration interval in days
Global Const MINIMUM_PASSWORD_LENGTH = 8 'Minimum password length
Global Const DEFAULT_PASSWORD = "password"
'Password entry specifics
Global Const APP_PASSWORD_REQUIRED = True 'Enables password protection disable for development
Global Const NUM_TRIES = 3
Global Const PROVIDER = "Microsoft.Jet.OLEDB.4.0"
Global Const GBL_USER_CONNECT = "C:\UserInfo.mdb"
'The following fields need to be created in a Access Database and set a database password "nkk'
'and place it in the C:\drive for temp You can also use access 2000 change JET to 4.0
'UserID
'UserPassword
'UserTaskLevel
'UserActivationDate
'UserExpireDate
'UserLastName
'UserFirstName
'UserMiddleInitial
'UserNotes
Global Const DB_PASSWORD = "1Ov45FD56g"
Global Const TASK_LEVEL_5 = "5 - Administrator" 'You can name the levels what ever you want
Global Const TASK_LEVEL_4 = "4 - Claims"
Global Const TASK_LEVEL_3 = "3 - Material Control"
Global Const TASK_LEVEL_2 = "2 - Process Responsibility"
Global Const TASK_LEVEL_1 = "1 - Guest"
'This will by-pass logon when APP_PASSWORD_REQUIRED = False
'Makes development easy
Sub Main()
If APP_PASSWORD_REQUIRED = True Then
frmLogin.Show
Else
Form1.Show
End If

   If App.PrevInstance Then
      ActivatePrevInstance
   End If

End Sub
'Check the Password and Default/Expire
Sub PassCheck()
Dim Password As String
Dim UserTask As String
Dim UserExpire As String
Dim ID As String
Dim NewDate As Date
'Checking to make user password is correct.
Password = frmLogin.txtPasswordCheck.Text
UserTask = frmLogin.txtUserTaskLevel.Text
UserExpire = frmLogin.txtUserExpireDate.Text
If frmLogin.txtUserPassword = Password Then
'setting all globals, look at the Task level this how the text is taken care of
    UserPassword = (Password)
    UserTaskLevel = Left(UserTask, 1)
    UserExpireDate = UserExpire
    'If default password we need to change
    If UserPassword = DEFAULT_PASSWORD Then
        frmChangePass.Show
        MsgBox "Your Password is the Default and must be changed.", vbApplicationModal, "Password needs changed"
    Unload frmLogin
    
    Exit Sub
    End If
    'If password has expired we need to change
    NewDate = DateAdd("d", EXPIRE_TERM, UserExpireDate)
    If NewDate - Date <= 1 Then
        frmChangePass.Show
        MsgBox "Your Password has expire you must change it.", vbApplicationModal, "Password needs changed"
    Unload frmLogin
    Exit Sub
    End If
    'If all okay let's run on, this is where you would put your start up form, make sure
    'that you issue access to users form somewhere else I have it on this temp form
Form1.Show
Unload frmLogin
Else: MsgBox "Password is not correct for this user ID. Passwords are case-sensitive.", vbApplicationModal, "Password Incorrect."
        frmLogin.txtPasswordCheck = ""
        frmLogin.txtPasswordCheck.SetFocus
End If

End Sub


