VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Change Password"
      Height          =   420
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   990
      Width           =   1950
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Shows the globals that you should be concerned with."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   645
      Left            =   3735
      TabIndex        =   11
      Top             =   3735
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      Height          =   330
      Left            =   1080
      TabIndex        =   10
      Top             =   3720
      Width           =   2265
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      Caption         =   "User ID"
      Height          =   330
      Left            =   90
      TabIndex        =   9
      Top             =   3720
      Width           =   915
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FFFF&
      Caption         =   "Task Level"
      Height          =   330
      Left            =   90
      TabIndex        =   8
      Top             =   4590
      Width           =   915
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FF00&
      Caption         =   "Expire Date"
      Height          =   330
      Left            =   90
      TabIndex        =   7
      Top             =   4155
      Width           =   915
   End
   Begin VB.Label Label7 
      Caption         =   $"Form1.frx":0000
      Height          =   645
      Left            =   45
      TabIndex        =   6
      Top             =   90
      Width           =   4830
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   "Password"
      Height          =   330
      Left            =   90
      TabIndex        =   5
      Top             =   3285
      Width           =   915
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Height          =   330
      Left            =   1080
      TabIndex        =   3
      Top             =   4155
      Width           =   2265
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Top             =   4590
      Width           =   2265
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   3285
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":00D1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   45
      TabIndex        =   0
      Top             =   1620
      Width           =   4605
   End
   Begin VB.Menu mnu_Users 
      Caption         =   "Users"
   End
   Begin VB.Menu test 
      Caption         =   "Enable Disable Test"
   End
End
Attribute VB_Name = "Form1"
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
'Form1
'
'Copyright: (c) 2000 Nicholas K. Kruse
'PLEASE:  If you use my code that is fine, but please remember that I spent
'a lot of time developing this in the newest technology ADO.  I would
'appreaicate you leaving my comments in here and giving me credit.  Futhermore
'If you use my code please read the .txt file you downloaded with
'these modules.  Further you can reach me at steelcoil@bellsouth.net
'This is by far the best set of modules for this procedure on the web, remember!
'
'   Junk just to show possibilities
'
'
'   History.
'       26 November 2000
'       Nicholas K. Kruse
'
'**************************************************************************
Private Sub Command1_Click()
frmChangePass.Show
Unload Form1
End Sub
'This update the variables so you see them as you make your changes
Private Sub Form_GotFocus()
Me.Refresh
End Sub
Private Sub Form_Load()
Label2.Caption = UserPassword
Label3.Caption = UserTaskLevel
Label4.Caption = UserID
Label5.Caption = UserExpireDate
If UserTaskLevel = 5 Then
mnu_Users.Enabled = True
End If
End Sub
Private Sub mnu_Users_Click()
frmUser.Show
End Sub
Private Sub test_Click()
If UserTaskLevel = 5 Then
test.Enabled = False
End If
End Sub
