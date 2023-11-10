VERSION 5.00
Begin VB.Form FrmCreateUpdateAuthor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   3930
      TabIndex        =   5
      Top             =   1590
      Width           =   1260
   End
   Begin VB.CommandButton btnClean 
      Caption         =   "Clean"
      Height          =   315
      Left            =   2610
      TabIndex        =   4
      Top             =   1590
      Width           =   1260
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   1290
      TabIndex        =   3
      Top             =   1605
      Width           =   1260
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   120
      MaxLength       =   150
      TabIndex        =   2
      Top             =   1185
      Width           =   5085
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   1110
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "FrmCreateUpdateAuthor.frx":0000
      Left            =   3795
      List            =   "FrmCreateUpdateAuthor.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   210
      Left            =   105
      TabIndex        =   8
      Top             =   885
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "Code:"
      Height          =   210
      Left            =   90
      TabIndex        =   7
      Top             =   120
      Width           =   660
   End
   Begin VB.Label Label3 
      Caption         =   "Status"
      Height          =   195
      Left            =   3825
      TabIndex        =   6
      Top             =   105
      Width           =   765
   End
End
Attribute VB_Name = "FrmCreateUpdateAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bll As New AuthorBll
Dim utils As New utils
Private Sub btnClean_Click()
txtName.text = ""

End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
If txtName.text <> "" Then
    bll.CreateUpdateAuthor txtCode.text, txtName.text, cboStatus.text
    If txtCode.text = "" Then
        MsgBox "Author created with successful!", vbOKOnly, "Notice"
    Else
        MsgBox "Author updated with successful!", vbOKOnly, "Notice"
    End If
    Unload Me
Else
    MsgBox "Field name is required!", vbOKOnly, "Notice"
End If
End Sub

Private Sub Form_Activate()
If txtCode.text = "" Then
    Me.Caption = "Create Record"
    cboStatus.text = "Active"
Else
    Me.Caption = "Update Record"
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
KeyAscii = utils.NumberAndLetters(KeyAscii)
End Sub
