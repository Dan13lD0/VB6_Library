VERSION 5.00
Begin VB.Form FrmChoseCreateReturn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   2325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   75
      TabIndex        =   2
      Top             =   1455
      Width           =   2130
   End
   Begin VB.CommandButton btnAction 
      Caption         =   "Command1"
      Height          =   345
      Left            =   75
      TabIndex        =   1
      Top             =   1050
      Width           =   2130
   End
   Begin VB.Label lblMessage 
      Caption         =   "Label1"
      Height          =   390
      Left            =   180
      TabIndex        =   0
      Top             =   390
      Width           =   1950
   End
End
Attribute VB_Name = "FrmChoseCreateReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id As String
Dim bll As New BorrowBll
Private Sub btnAction_Click()
If btnAction.Caption = "Return" Then
bll.DeleteBorrow (id)
Unload Me
Else
Dim borrow As New FrmBorrow
borrow.bookId = id
borrow.Show vbModal
Unload Me
End If
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub
