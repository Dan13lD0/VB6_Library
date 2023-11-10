VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LogIn"
   ClientHeight    =   2625
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      Height          =   405
      Left            =   3525
      TabIndex        =   5
      Top             =   2025
      Width           =   990
   End
   Begin VB.CommandButton btnSignIn 
      Caption         =   "Sign In"
      Height          =   405
      Left            =   2385
      TabIndex        =   4
      Top             =   2025
      Width           =   990
   End
   Begin VB.TextBox txtPassword 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   180
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "admin"
      Top             =   1575
      Width           =   2910
   End
   Begin VB.TextBox txtLogin 
      Height          =   360
      Left            =   195
      TabIndex        =   2
      Text            =   "admin"
      Top             =   705
      Width           =   2910
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   3360
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   645
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   180
      Left            =   180
      TabIndex        =   1
      Top             =   1215
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Login:"
      Height          =   180
      Left            =   195
      TabIndex        =   0
      Top             =   315
      Width           =   1080
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private utils As New utils
Dim bll As New UserBll
Private Sub btnClose_Click()
 If utils.MessageBoxConfirmation("Would you like to leave the system?", "Leave") = True Then
    Unload Me
 End If
End Sub

Private Sub btnSignIn_Click()
If bll.ValidateUser(txtLogin.text, txtPassword.text) = True Then
Dim main  As New FrmMain
main.Show
Unload Me
Else
MsgBox "User not found!", vbOK, "Login"
End If

End Sub
