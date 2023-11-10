VERSION 5.00
Begin VB.Form FrmCreateUpdateCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "FrmCreateUpdateCategory.frx":0000
      Left            =   3810
      List            =   "FrmCreateUpdateCategory.frx":000A
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtCode 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   420
      Width           =   1110
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   135
      TabIndex        =   3
      Top             =   1185
      Width           =   5085
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   1305
      TabIndex        =   2
      Top             =   1605
      Width           =   1260
   End
   Begin VB.CommandButton btnClean 
      Caption         =   "Clean"
      Height          =   315
      Left            =   2625
      TabIndex        =   1
      Top             =   1590
      Width           =   1260
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   3945
      TabIndex        =   0
      Top             =   1590
      Width           =   1260
   End
   Begin VB.Label Label3 
      Caption         =   "Status"
      Height          =   195
      Left            =   3840
      TabIndex        =   7
      Top             =   105
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Code:"
      Height          =   210
      Left            =   105
      TabIndex        =   6
      Top             =   120
      Width           =   660
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   885
      Width           =   660
   End
End
Attribute VB_Name = "FrmCreateUpdateCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bll As New CategoryBll

Private Sub btnClean_Click()
txtName.text = ""

End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
If txtName.text <> "" Then
    bll.CreateUpdateCategory txtCode.text, txtName.text, cboStatus.text
    If txtCode.text = "" Then
        MsgBox "Category created with successful!", vbOKOnly, "Notice"
    Else
        MsgBox "Category updated with successful!", vbOKOnly, "Notice"
    End If
    Unload Me
Else
    MsgBox "Field name is required!", vbOKOnly, "Notice"
End If
End Sub

Private Sub Form_Load()
If txtCode.text = "" Then
    Me.Caption = "Create Record"
Else
    Me.Caption = "Update Record"
End If

End Sub