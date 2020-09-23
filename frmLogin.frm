VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter User ID and Password"
   ClientHeight    =   1440
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleMode       =   0  'User
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserTaskLevel 
      DataField       =   "UserTaskLevel"
      DataMember      =   "Logon"
      DataSource      =   "dsrUserInfo"
      Height          =   285
      Left            =   2385
      TabIndex        =   13
      Top             =   3615
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.TextBox txtUserPassword 
      DataField       =   "UserPassword"
      DataMember      =   "Logon"
      DataSource      =   "dsrUserInfo"
      Height          =   285
      Left            =   2385
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.TextBox txtUserExpireDate 
      DataField       =   "UserExpireDate"
      DataMember      =   "Logon"
      DataSource      =   "dsrUserInfo"
      Height          =   285
      Left            =   2385
      TabIndex        =   9
      Top             =   2475
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtUserID 
      DataField       =   "UserID"
      DataMember      =   "Logon"
      DataSource      =   "dsrUserInfo"
      Height          =   285
      Index           =   2
      Left            =   2385
      TabIndex        =   7
      Top             =   2100
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtUserIDCheck 
      Height          =   345
      Left            =   1305
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1305
      TabIndex        =   4
      Top             =   900
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2475
      TabIndex        =   5
      Top             =   900
      Width           =   1140
   End
   Begin VB.TextBox txtPasswordCheck 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UserTaskLevel:"
      Height          =   255
      Index           =   5
      Left            =   540
      TabIndex        =   12
      Top             =   3660
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UserPassword:"
      Height          =   255
      Index           =   4
      Left            =   540
      TabIndex        =   10
      Top             =   3285
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UserExpireDate:"
      Height          =   255
      Index           =   2
      Left            =   540
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UserID:"
      Height          =   255
      Index           =   1
      Left            =   540
      TabIndex        =   6
      Top             =   2145
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
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
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
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
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************
'This routine is for logging on and checking password through ADO
'
''This program is documented well with comments.
'*************************************************************************
'
'frmLogin
'
'Copyright: (c) 2000 Nicholas K. Kruse
'PLEASE:  If you use my code that is fine, but please remember that I spent
'a lot of time developing this in the newest technology ADO.  I would
'appreaicate you leaving my comments in here and giving me credit.  Futhermore
'If you use my code please read the .txt file you downloaded with
'these modules.  Further you can reach me at steelcoil@bellsouth.net
'This is by far the best set of modules for this procedure on the web, remember!
'
'   This form is used for the secured login, please remember that the login
'   is the only security there is.
'
'
'   History.
'       26 November 2000
'       Nicholas K. Kruse
'
'**************************************************************************
Private Sub cmdCancel_Click()
   End
End Sub
Private Sub cmdOK_Click()
Static Tries As Integer
Tries = Tries + 1
'If to many tries we want to get rid of this user.
     If Tries >= NUM_TRIES Then     'nope, otta hear, you have gotta go.
        MsgBox "YOUR ACCESS HAS BEEN DENIED PLEASE CALL MIS ADMINISTRATOR.", vbApplicationModal, "Password"
        End
        End If
    'check to see that UserID field is not blank, I don't user validation here if want to
    'know why email me.
    If txtUserIDCheck = "" Then
        MsgBox "You must enter a UserID.", vbCritical, "UserID Missing."
        txtUserIDCheck.SetFocus
        Exit Sub
    Resume
    End If
    'same thing checking for values, same reason email me.
        If txtPasswordCheck = "" Then
            MsgBox "You must enter a Password.", vbCritical, "User Password Missing."
            txtPasswordCheck.SetFocus
        Exit Sub
        Resume
        End If
        Call Read_Pass
End Sub
'This is the read password module have to find the user first, read the dataenvirnoment
'connections carefully before changing.
Sub Read_Pass()
Dim Lastrow
Dim findStr As String
    If Not dsrUserInfo.rsLogon.Supports(adFind) Then
        MsgBox "Recordset doesn't support the FIND method"
    Else
        If dsrUserInfo.rsLogon.Supports(adBookmark) Then
            'Set this to keep system from locking up database if user doesn't exist, I think.
            Lastrow = dsrUserInfo.rsLogon.Bookmark
        End If
        findStr = txtUserIDCheck.Text
        'Only need to find first one, and must start in front
        dsrUserInfo.rsLogon.MoveFirst
        dsrUserInfo.rsLogon.Find "UserID LIKE '*" & findStr & "*'"
        UserID = findStr
                If dsrUserInfo.rsLogon.EOF Then
            MsgBox "This user, " & findStr & " , does not have Access, a log will be generated.", vbSystemModal, "Security System"
            End 'If no user programs goes I will update this in future versions,
                'will give the user more than one chance not for now. Not sure how it affects database
            End If
            End If
    Call PassCheck
End Sub
