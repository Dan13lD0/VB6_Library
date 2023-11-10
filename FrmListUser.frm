VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmListUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   14835
   Begin VB.Frame Frame1 
      Caption         =   "Filter:"
      Height          =   885
      Left            =   915
      TabIndex        =   4
      Top             =   0
      Width           =   11310
      Begin VB.TextBox txtLogin 
         Height          =   285
         Left            =   3915
         TabIndex        =   15
         Top             =   495
         Width           =   2175
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   6135
         TabIndex        =   13
         Top             =   480
         Width           =   2670
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   495
         Width           =   1185
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1365
         TabIndex        =   6
         Top             =   495
         Width           =   2505
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "FrmListUser.frx":0000
         Left            =   8880
         List            =   "FrmListUser.frx":000A
         TabIndex        =   5
         Text            =   "All"
         Top             =   480
         Width           =   2235
      End
      Begin VB.Label Label5 
         Caption         =   "Email:"
         Height          =   180
         Left            =   6135
         TabIndex        =   14
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Login:"
         Height          =   180
         Left            =   3915
         TabIndex        =   12
         Top             =   255
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Code:"
         Height          =   180
         Left            =   165
         TabIndex        =   10
         Top             =   255
         Width           =   810
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   180
         Left            =   1365
         TabIndex        =   9
         Top             =   255
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "Status:"
         Height          =   180
         Left            =   8850
         TabIndex        =   8
         Top             =   255
         Width           =   810
      End
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   800
      Left            =   13950
      TabIndex        =   3
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton btnEnabled 
      Caption         =   "Enable"
      Height          =   800
      Left            =   13155
      TabIndex        =   2
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      Height          =   800
      Left            =   12345
      TabIndex        =   1
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "New"
      Height          =   800
      Left            =   0
      TabIndex        =   0
      Top             =   105
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid dgv 
      Height          =   4965
      Left            =   0
      TabIndex        =   11
      Top             =   1005
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   8758
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmListUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bll As New UserBll
Dim utils As New utils

Private Sub btnClose_Click()
Unload Me
End Sub



Private Sub btnEnabled_Click()
If dgv.VisibleRows > 0 Then
    dgv.Col = 0
    bll.DeleteUser (dgv.text)
    LoadGrid
End If
End Sub

Private Sub btnNew_Click()
Dim createUpdate As New FrmCreateUpdateUser
createUpdate.Show vbModal
LoadGrid
End Sub

Private Sub Command2_Click()
LoadGrid
End Sub

Private Sub btnSearch_Click()
LoadGrid
End Sub

Private Sub dgv_DblClick()
    If dgv.VisibleRows > 0 Then
        Dim update As New FrmCreateUpdateUser
        dgv.Col = 0
        update.txtCode.text = dgv.text
        
        dgv.Col = 1
        update.txtName.text = dgv.text
        
        dgv.Col = 2
        update.txtLogin.text = dgv.text
        
        dgv.Col = 3
        update.txtEmail.text = dgv.text
        
        dgv.Col = 4
        update.cboStatus.text = dgv.text
                        
        update.Show vbModal
        LoadGrid
    End If
End Sub

Private Sub Form_Load()

LoadGrid
LoadComboBoxStatus

End Sub

Private Sub LoadGrid()
Set dgv.DataSource = bll.GetUsers(txtCode.text, txtName.text, txtLogin.text, txtEmail.text, cboStatus.text)

dgv.Columns(0).Caption = "Code"
dgv.Columns(0).Width = 1000
dgv.Columns(0).Alignment = dbgCenter

dgv.Columns(1).Caption = "Name"
dgv.Columns(1).Width = 4000

dgv.Columns(2).Caption = "Login"
dgv.Columns(2).Width = 4000


dgv.Columns(3).Caption = "email"
dgv.Columns(3).Width = 4000


dgv.Columns(4).Caption = "Status"
dgv.Columns(4).Width = 1300
dgv.Columns(4).Alignment = dbgCenter
End Sub

Private Sub LoadComboBoxStatus()
cboStatus.AddItem ("All")
cboStatus.AddItem ("Active")
cboStatus.AddItem ("Inactive")
cboStatus.text = "All"
End Sub


