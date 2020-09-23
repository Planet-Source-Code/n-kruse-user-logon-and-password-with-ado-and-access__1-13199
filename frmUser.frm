VERSION 5.00
Begin VB.Form frmUser 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Information Screen"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer cmdTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6075
      Top             =   4080
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      Caption         =   "&Close && Cancel will take five seconds"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   4125
      Width           =   3810
   End
   Begin VB.CommandButton cmdPwdReset 
      Caption         =   "&Reset Password "
      Height          =   375
      Left            =   5100
      TabIndex        =   23
      Top             =   3345
      Width           =   1560
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add User"
      Height          =   375
      Left            =   5100
      TabIndex        =   22
      Top             =   1605
      Width           =   1560
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Info"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5100
      TabIndex        =   21
      Top             =   2040
      Width           =   1560
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete User"
      Height          =   375
      Left            =   5100
      TabIndex        =   20
      Top             =   2475
      Width           =   1560
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Info Change"
      Height          =   375
      Left            =   5100
      TabIndex        =   19
      Top             =   2910
      Width           =   1560
   End
   Begin VB.CommandButton cmdMoveNextBack 
      Caption         =   "&<"
      Height          =   495
      Left            =   4365
      TabIndex        =   18
      Top             =   4050
      Width           =   420
   End
   Begin VB.CommandButton cmdMoveNextFor 
      Caption         =   "&>"
      Height          =   495
      Left            =   4785
      TabIndex        =   17
      Top             =   4050
      Width           =   420
   End
   Begin VB.ComboBox DBComboUserTaskLevel 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3825
      TabIndex        =   15
      Top             =   315
      Width           =   2880
   End
   Begin VB.TextBox txtPwdExpire 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM dd, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2490
      TabIndex        =   14
      Top             =   900
      Width           =   1815
   End
   Begin VB.TextBox txtUserNotes 
      DataField       =   "UserNotes"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2175
      Width           =   4095
   End
   Begin VB.TextBox txtUserID 
      DataField       =   "UserID"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM d, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2505
      TabIndex        =   5
      Top             =   1545
      Width           =   1815
   End
   Begin VB.TextBox txtUserPassword 
      DataField       =   "UserPassword"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   60
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1545
      Width           =   1815
   End
   Begin VB.TextBox txtUserActivationDate 
      DataField       =   "UserActivationDate"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM d, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   75
      TabIndex        =   3
      Top             =   900
      Width           =   1815
   End
   Begin VB.TextBox txtUserLastName 
      DataField       =   "USerLastName"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Top             =   315
      Width           =   1530
   End
   Begin VB.TextBox txtUserMiddleInitial 
      DataField       =   "UserMiddleInitial"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1635
      MaxLength       =   1
      TabIndex        =   1
      Top             =   315
      Width           =   360
   End
   Begin VB.TextBox txtUserFirstName 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   75
      TabIndex        =   0
      ToolTipText     =   "User's First Name"
      Top             =   315
      Width           =   1530
   End
   Begin VB.TextBox txtUserExpireDate 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   180
      TabIndex        =   16
      Top             =   4695
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblUserStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   26
      Top             =   4590
      Width           =   6840
   End
   Begin VB.Label lblRecordMove 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Move through Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4290
      TabIndex        =   25
      Top             =   3780
      Width           =   2610
   End
   Begin VB.Label lblUserNotes 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "User Notes --Enter anything Regarding this User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   105
      TabIndex        =   13
      Top             =   1935
      Width           =   4485
   End
   Begin VB.Label lblUserID 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID same as LOGON ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2490
      TabIndex        =   12
      Top             =   1260
      Width           =   2745
   End
   Begin VB.Label lblUserPassword 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "User Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   1275
      Width           =   2070
   End
   Begin VB.Label lblUserExpireDate 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "User Password Expire Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2100
      TabIndex        =   10
      Top             =   660
      Width           =   2340
   End
   Begin VB.Label lblUserActivationDate 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "User Activation Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   45
      TabIndex        =   9
      Top             =   660
      Width           =   1845
   End
   Begin VB.Label lblUserTaskLevel 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "User Task Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3855
      TabIndex        =   8
      Top             =   60
      Width           =   2025
   End
   Begin VB.Label lblUserFullName 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Full User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   75
      TabIndex        =   7
      Top             =   45
      Width           =   2610
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim adoRecordset As ADODB.Recordset
'*************************************************************************
'This routine is for adding users the password through ADO
'
'I won't detail everything here, I think you will figure most of the functions
'**************************************************************************
'
'frmUser
'
'Copyright: (c) 2000 Nicholas K. Kruse
'PLEASE:  If you use my code that is fine, but please remember that I spent
'a lot of time developing this in the newest technology ADO.  I would
'appreaicate you leaving my comments in here and giving me credit.  Futhermore
'If you use my code please read the .txt file you downloaded with
'these modules.  Further you can reach me at steelcoil@bellsouth.net
'This is by far the best set of modules for this procedure on the web, remember!
'
'
'       This code allows a user to be added, deleted, changes made, password re-sets
'       Also all the functions for re-setting dates on password re-set.
'       As designed only administrators can get in here.
'
'
'   History.
'       26 November 2000
'       Nicholas K. Kruse
'
'**************************************************************************
Sub Update()
On Error GoTo Errlbl:
adoRecordset.Update
Exit Sub
Errlbl:
MsgBox "Make sure all information is entered properly", vbApplicationModal, "Information not Correct"
adoRecordset.CancelUpdate
Resume Next
End Sub
Sub Requery()
adoRecordset.Resync
End Sub
Private Sub cmdAdd_Click()
On Error GoTo ErrLabel
    Update
adoRecordset.AddNew
cmdMoveNextFor.Enabled = False
cmdMoveNextBack.Enabled = False
txtUserPassword.Enabled = False
cmdPwdReset.Enabled = False
cmdDelete.Enabled = False
txtUserActivationDate.Enabled = False
txtUserID.Enabled = False
txtUserExpireDate.Enabled = False
txtPwdExpire.Text = ""
txtUserFirstName.Enabled = True
txtUserMiddleInitial.Enabled = True
txtUserLastName.Enabled = True
cmdAdd.Enabled = False
DBComboUserTaskLevel.Enabled = True
txtUserNotes.Enabled = True
Exit Sub
ErrLabel:
    MsgBox Err.Description
End Sub
Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUserStatus.Caption = "Add a User"
lblUserStatus.FontSize = 10
End Sub

Private Sub cmdChange_Click()
txtUserNotes.Enabled = True
DBComboUserTaskLevel.Enabled = True
End Sub
Private Sub cmdChange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUserStatus.Caption = "Change information about a user"
End Sub
Private Sub cmdClose_Click()
cmdTimer.Enabled = True
End Sub
Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUserStatus.Caption = "Delayed Close, will close form"
lblUserStatus.FontSize = 10
End Sub
Private Sub cmdDelete_Click()
Dim DeleteUser As String
Dim CurrentUser As String
'Setting up details to make sure the current user will not be deleted.
DeleteUser = txtUserID.Text
CurrentUser = UserID
DeleteUser = Format(DeleteUser, "<")
CurrentUser = Format(CurrentUser, "<")
If DeleteUser = CurrentUser Then
MsgBox "You cannot delete the current user.", vbCritical, "User Security"
    Exit Sub
    End If
adoRecordset.Delete
adoRecordset.MoveNext
If adoRecordset.EOF Then
    adoRecordset.MoveLast
    End If
End Sub
Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUserStatus.Caption = "Delete this user"
End Sub
Private Sub cmdMoveNextBack_Click()
adoRecordset.CancelUpdate
If Not adoRecordset.BOF Then
    adoRecordset.MovePrevious
End If
If adoRecordset.BOF And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
End If
End Sub
Private Sub cmdMoveNextBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUserStatus.Caption = "Move to the previous record"
End Sub
Private Sub cmdMoveNextFor_Click()
If Not adoRecordset.EOF Then
    adoRecordset.MoveNext
End If
If adoRecordset.EOF And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveLast
End If
End Sub
Private Sub cmdMoveNextFor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUserStatus.Caption = "Move to the next record"
End Sub
Private Sub cmdPwdReset_Click()
txtUserPassword.Text = "password"
adoRecordset.MoveNext
adoRecordset.MovePrevious
End Sub
Private Sub cmdPwdReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUserStatus.Caption = "Re-set this users password" & "" & "Password will = 'password' & case-sensitive."
lblUserStatus.FontSize = 8
End Sub
Private Sub cmdTimer_Timer()
Unload Me
End Sub
Private Sub cmdUpdate_Click()
On Error GoTo Errlbl:
adoRecordset.Update
cmdAdd.Enabled = True
cmdMoveNextFor.Enabled = True
cmdMoveNextBack.Enabled = True
DBComboUserTaskLevel.Enabled = False
txtUserNotes.Enabled = False
cmdUpdate.Enabled = False
cmdDelete.Enabled = True
cmdPwdReset.Enabled = True
Exit Sub
Errlbl:
MsgBox "Make sure all information is entered properly", vbApplicationModal, "Information not Correct"
adoRecordset.CancelUpdate
Resume Next
End Sub
Private Sub cmdUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUserStatus.Caption = "Update information about this user"
End Sub
Private Sub DBComboUserTaskLevel_Click()
cmdUpdate.Enabled = True
End Sub
Private Sub Form_Load()
Dim db As Connection
Set db = New Connection
Dim strSQL As String
'I know this connection must look bad, but if you are going to user on a single machine
'and do lots of record writing to a local batch for later updates, then you can also use
'connections of two kinds just demonstrating.  You could create another command in the
'dsrUserInfo
'Also note that I use a password protected database, so it can not be opened.  once this is
'compilied and installed it would take a slick individual to get that password.  It is set
'very light now make it something like, A3bn4r8PO451z1s2f35, that would make the programs
'out there how a rough time.  Yes there are cracking program available to break into Access,
'don't expect Microsoft (R) to tell you that.
With db
    .CursorLocation = adUseClient
    .PROVIDER = PROVIDER
    .Properties("Jet OLEDB:Database Password") = DB_PASSWORD
    .Open GBL_USER_CONNECT
End With
Set adoRecordset = New Recordset
strSQL = "SELECT UserID, UserPassword, UserTaskLevel," & _
        "UserActivationDate, UserExpireDate, UserLastName," & _
        "UserMiddleInitial, UserNotes, UserFirstName " & _
        "FROM tblUserInfo ORDER BY UserID"
adoRecordset.Open strSQL, db, adOpenDynamic, adLockOptimistic
Set txtUserExpireDate.DataSource = adoRecordset
    txtUserExpireDate.DataField = "UserExpireDate"
Set txtUserID.DataSource = adoRecordset
    txtUserID.DataField = "UserID"
Set txtUserFirstName.DataSource = adoRecordset
    txtUserFirstName.DataField = "UserFirstName"
Set txtUserMiddleInitial.DataSource = adoRecordset
    txtUserMiddleInitial.DataField = "UserMiddleInitial"
Set txtUserLastName.DataSource = adoRecordset
    txtUserLastName.DataField = "UserLastName"
Set txtUserPassword.DataSource = adoRecordset
    txtUserPassword.DataField = "UserPassword"
Set txtUserNotes.DataSource = adoRecordset
    txtUserNotes.DataField = "UserNotes"
Set txtUserActivationDate.DataSource = adoRecordset
    txtUserActivationDate.DataField = "UserActivationDate"
Set DBComboUserTaskLevel.DataSource = adoRecordset
    DBComboUserTaskLevel.DataField = "UserTaskLevel"
'This is where the task level come in, set by globals
'Remember the task level is trimmed so don't worry about all that text
DBComboUserTaskLevel.AddItem TASK_LEVEL_5
DBComboUserTaskLevel.AddItem TASK_LEVEL_4
DBComboUserTaskLevel.AddItem TASK_LEVEL_3
DBComboUserTaskLevel.AddItem TASK_LEVEL_2
DBComboUserTaskLevel.AddItem TASK_LEVEL_1
'This displays the actual expire date, not the expire date
NewDate = DateAdd("d", EXPIRE_TERM, txtUserExpireDate)
txtPwdExpire = Format(NewDate, "mmmm d, yyyy")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblUserStatus.Caption = "Material Tracking User Information" & "" & Now()
lblUserStatus.FontSize = 10
End Sub
'makes the user ID, I user 1 from first name, middle initial, and 6 (if there) from Last
Private Sub txtUserLastName_Change()
txtUserID = Left([txtUserFirstName], 1) & Left([txtUserMiddleInitial], 1) & Left([txtUserLastName], 6)
End Sub
'You can add this to keep user from putting spaces in the fields
'I will activate in all fields when tied to an application
'Remove comments from next seven lines
'Private Sub txtUserFirstName_KeyPress(KeyAscii As Integer)
'Select Case KeyCode
'    Case SpaceBar
'    Beep
'    KeyAscii = 0
'    End Select
'End Sub
'Entering all the fields that are needed for a new user activation
Private Sub txtUserLastName_KeyPress(KeyAscii As Integer)
txtUserPassword.Text = "password"
txtUserActivationDate = Date
txtUserExpireDate = Date
NewDate = DateAdd("d", EXPIRE_TERM, txtUserExpireDate)
txtPwdExpire = Format(NewDate, "mmmm d, yyyy")
'You can add this to keep user from putting spaces in the fields
'I will activate in all fields when tied to an application
'Remove comments from next five lines
'Select Case KeyCode
'    Case SpaceBar
'    Beep
'    KeyAscii = 0
'    End Select
End Sub
Private Sub txtUserNotes_Click()
cmdUpdate.Enabled = True
End Sub

