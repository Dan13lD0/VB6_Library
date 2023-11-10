VERSION 5.00
Begin VB.Form FrmCreateUpdateUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "FrmCreateUpdateUser.frx":0000
      Left            =   3780
      List            =   "FrmCreateUpdateUser.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   330
      Width           =   1455
   End
   Begin VB.TextBox txtEmail 
      Height          =   315
      Left            =   105
      MaxLength       =   150
      TabIndex        =   12
      Top             =   2625
      Width           =   5085
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      Left            =   2640
      MaxLength       =   15
      TabIndex        =   10
      Top             =   1875
      Width           =   2535
   End
   Begin VB.TextBox txtLogin 
      Height          =   315
      Left            =   105
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1890
      Width           =   2475
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   3915
      TabIndex        =   6
      Top             =   3000
      Width           =   1260
   End
   Begin VB.CommandButton btnClean 
      Caption         =   "Clean"
      Height          =   315
      Left            =   2595
      TabIndex        =   5
      Top             =   3000
      Width           =   1260
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   1275
      TabIndex        =   4
      Top             =   3015
      Width           =   1260
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   105
      MaxLength       =   150
      TabIndex        =   3
      Top             =   1155
      Width           =   5085
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   390
      Width           =   1110
   End
   Begin VB.Label Label6 
      Caption         =   "Status"
      Height          =   195
      Left            =   3810
      TabIndex        =   13
      Top             =   75
      Width           =   765
   End
   Begin VB.Label Label5 
      Caption         =   "E-mail:"
      Height          =   210
      Left            =   90
      TabIndex        =   11
      Top             =   2325
      Width           =   660
   End
   Begin VB.Label Label4 
      Caption         =   "Password:"
      Height          =   210
      Left            =   2625
      TabIndex        =   9
      Top             =   1575
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Login:"
      Height          =   210
      Left            =   90
      TabIndex        =   7
      Top             =   1590
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   210
      Left            =   90
      TabIndex        =   2
      Top             =   855
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "Code:"
      Height          =   210
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   660
   End
End
Attribute VB_Name = "FrmCreateUpdateUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bll As New UserBll
Dim utils As New utils
Private Sub btnClean_Click()
txtName.text = ""
txtLogin.text = ""
txtPassword.text = ""
txtEmail.text = ""
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
Dim msg As String
If txtCode.text = "" Then
msg = "Record created with successfull!"
Else
msg = "Record updated with successfull!"
End If

If txtName.text = "" Or txtLogin.text = "" Or txtPassword.text = "" Or txtEmail.text = "" Then
    MsgBox "Insert the fields requerid!", vbOK, "Notice"
Else
    bll.CreateUpdateUser txtCode.text, txtName.text, txtLogin.text, txtPassword.text, txtEmail.text, cboStatus.text
    If txtCode.text = "" Then
    MsgBox "Record created with successfull!", vbOK, "Notice"
    Else
    MsgBox "Record update with successfull!", vbOK, "Notice"
    End If
    Unload Me
End If
End Sub

Private Sub Form_Activate()
If txtCode.text = "" Then
Me.Caption = "Create record"
cboStatus.text = "Active"
Else
Me.Caption = "Update record"
End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
KeyAscii = utils.NumberAndLetters(KeyAscii)
End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
KeyAscii = utils.NumberAndLetters(KeyAscii)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
KeyAscii = utils.NumberAndLetters(KeyAscii)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
KeyAscii = utils.NumberAndLetters(KeyAscii)
End Sub
