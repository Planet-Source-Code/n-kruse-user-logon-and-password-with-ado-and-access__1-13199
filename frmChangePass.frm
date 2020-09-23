VERSION 5.00
Begin VB.Form frmChangePass 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6045
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tempUserID 
      Height          =   330
      Left            =   1665
      TabIndex        =   13
      Top             =   2610
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton cmdChangePass 
      Caption         =   "&Change Password"
      Enabled         =   0   'False
      Height          =   420
      Left            =   2430
      TabIndex        =   4
      Top             =   1530
      Width           =   2040
   End
   Begin VB.TextBox txtNewPass1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2295
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   2310
   End
   Begin VB.TextBox txtNewPass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2295
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   585
      Width           =   2310
   End
   Begin VB.TextBox txtOldPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2295
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   135
      Width           =   2310
   End
   Begin VB.TextBox txtUserPassword 
      DataField       =   "UserPassword"
      DataMember      =   "ChangePass"
      DataSource      =   "dsrUserInfo"
      Height          =   285
      Left            =   2130
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2845
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.TextBox txtUserExpireDate 
      DataField       =   "UserExpireDate"
      DataMember      =   "ChangePass"
      DataSource      =   "dsrUserInfo"
      Height          =   285
      Left            =   2130
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2475
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtUserID 
      DataField       =   "UserID"
      DataMember      =   "ChangePass"
      DataSource      =   "dsrUserInfo"
      Height          =   285
      Left            =   2130
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2085
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password Verify"
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
      Left            =   270
      TabIndex        =   12
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New password"
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
      Left            =   270
      TabIndex        =   11
      Top             =   585
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
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
      Left            =   270
      TabIndex        =   10
      Top             =   135
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UserPassword:"
      Height          =   255
      Index           =   2
      Left            =   285
      TabIndex        =   8
      Top             =   2890
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UserExpireDate:"
      Height          =   255
      Index           =   1
      Left            =   285
      TabIndex        =   6
      Top             =   2510
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UserID:"
      Height          =   255
      Index           =   0
      Left            =   285
      TabIndex        =   3
      Top             =   2130
      Width           =   1815
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
'This routine is for changing the password through ADO
'
''This program is documented well with comments.
'**************************************************************************
'
'frmChangePass
'
'Copyright: (c) 2000 Nicholas K. Kruse
'PLEASE:  If you use my code that is fine, but please remember that I spent
'a lot of time developing this in the newest technology ADO.  I would
'appreaicate you leaving my comments in here and giving me credit.  Futhermore
'If you use my code please read the .txt file you downloaded with
'these modules.  Further you can reach me at steelcoil@bellsouth.net
'This is by far the best set of modules for this procedure on the web, remember!
'
'   This code will allow you to change password.  When changing passwords a
'   user can not select the DEFAULT_PASSWORD or any string that contains
'   "password", this is very common among users so we will stop it here.
'
'
'   History.
'       26 November 2000
'       Nicholas K. Kruse
'
'**************************************************************************

Sub Change_Pass()
Dim Lastrow
Dim findStr As String

    If Not dsrUserInfo.rsChangePass.Supports(adFind) Then
        MsgBox "Recordset doesn't support the FIND method"
    Else
        If dsrUserInfo.rsChangePass.Supports(adBookmark) Then
            Lastrow = dsrUserInfo.rsChangePass.Bookmark
        End If
        findStr = tempUserID.Text
        dsrUserInfo.rsChangePass.MoveFirst
        'Set current user infomation
        dsrUserInfo.rsChangePass.Find "UserID LIKE '*" & findStr & "*'"
            If dsrUserInfo.rsChangePass.EOF Then
            MsgBox "This user, " & findStr & " , does not have Access, a log will be generated.", vbSystemModal, "Security System"
            End
            End If
    End If
    txtUserPassword.Text = txtNewPass.Text 'Set info to update
    txtUserExpireDate.Text = Date 'Set info to update
    UserPassword = txtUserPassword.Text 'Update Globals
    UserExpireDate = txtUserExpireDate.Text
    Call Update
End Sub
Sub Update()
'update just the fields we changed
dsrUserInfo.rsChangePass.Update "UserPassword", txtNewPass.Text
dsrUserInfo.rsChangePass.Update "UserExpireDate", Date
Unload Me
End Sub
Private Sub cmdChangePass_Click()
'Call sub function
Call Change_Pass
End Sub
Private Sub Form_Load()
'need to set this at the beginning we will reference this when finding record
tempUserID.Text = UserID
End Sub
Private Sub Form_Unload(Cancel As Integer)
'You will need to set this to what form will be your next or maybe where you were
'Another idea here that I use is to send them back to where their task level takes them
Form1.Show
End Sub
'must make sure the correct old password is entered
Private Sub txtNewPass_Validate(KeepFocus As Boolean)
Dim count As String
Dim SearchStr, PassStr, MyStr As String
count = Len(txtNewPass.Text)
'Make sure password does not equal old
If txtNewPass.Text = txtOldPassword.Text Then
KeepFocus = True
MsgBox "Your password can be the same as the old.", vbApplicationModal, "New Password"

Exit Sub
End If
'Now make sure it is long enough, other they will use ABC then 123
If count < MINIMUM_PASSWORD_LENGTH Then
KeepFocus = True
MsgBox "Your password must be at least 8 characters.", vbApplicationModal, "New Password"
txtNewPass.Text = ""
Exit Sub
End If
'Make sure password does not equal password
If txtNewPass.Text = DEFAULT_PASSWORD Then
KeepFocus = True
MsgBox "The default 'password' can not be used.", vbCritical, "New Password."
txtNewPass.Text = ""
Exit Sub
End If
'Make sure password is not in the string any where, they will user password1, 2,3...
PassStr = "password"
SearchStr = txtNewPass.Text
MyStr = InStr(1, SearchStr, PassStr, vbTextCompare)
If MyStr <> 0 Then
KeepFocus = True
MsgBox "The default 'password' can not be in the password.", vbCritical, "New Password."
txtNewPass.Text = ""
Exit Sub
End If
'Make sure the password is not the user id as well
If txtNewPass.Text = UserID Then
KeepFocus = True
MsgBox "The 'password' can not be the UserID.", vbCritical, "New Password."
txtNewPass.Text = ""
Exit Sub
End If
'Last for this version, yes really there is more to come, but this checks the new password
'against the old to make sure they are not using similiar words, really have to mix it up
Dim midstr, midpassstr, newmidstr As String
midstr = txtOldPassword.Text
midpassstr = midstr
newmidstr = Mid(midstr, 3, 4)
SearchStr = txtNewPass.Text
MyStr = InStr(1, SearchStr, newmidstr, vbTextCompare)
If MyStr <> 0 Then
KeepFocus = True
MsgBox "The password can not contain similiar words from old password.", vbCritical, "New Password."
txtNewPass.Text = ""
Exit Sub
End If
End Sub
'When the new and verify are at least the same in length lets turn on the button
Private Sub txtNewPass1_Change()
Dim count As String
Dim count1 As String
count = Len(txtNewPass.Text)
count1 = Len(txtNewPass1.Text)
If count1 >= count Then
cmdChangePass.Enabled = True
End If
End Sub
'Let's verify that the two new passwords are equal
Private Sub txtNewPass1_Validate(KeepFocus As Boolean)
If txtNewPass1.Text <> txtNewPass.Text Then
txtNewPass.SetFocus
MsgBox "Your password and verify password do not match.", vbApplicationModal, "Password"
txtNewPass.Text = ""
txtNewPass1.Text = ""
Exit Sub
End If
cmdChangePass.Enabled = True
cmdChangePass.SetFocus
End Sub
'Old pass must equal old global, I do this so if someone sits down at a computer
'they can't make changes and go access it somewhere else.
Private Sub txtOldPassword_Validate(KeepFocus As Boolean)
Static Tries As Integer
Tries = Tries + 1
If Tries >= NUM_TRIES Then     'nope, otta hear, you have gotta go.
        MsgBox "YOUR ACCESS HAS BEEN DENIED PLEASE CALL MIS ADMINISTRATOR.", vbApplicationModal, "Password"
        End
        End If
If txtOldPassword.Text <> UserPassword Then
KeepFocus = True
MsgBox "Your old password does not match", vbSystemModal, "Security System"
txtOldPassword.Text = ""
Exit Sub
End If
End Sub
