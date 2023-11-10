VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmListClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Clients"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   14865
   Begin VB.Frame Frame1 
      Caption         =   "Filter:"
      Height          =   885
      Left            =   960
      TabIndex        =   4
      Top             =   15
      Width           =   11370
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "FrmListClient.frx":0000
         Left            =   8880
         List            =   "FrmListClient.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   2370
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1365
         MaxLength       =   100
         TabIndex        =   8
         Top             =   495
         Width           =   2505
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   150
         MaxLength       =   10
         TabIndex        =   7
         Top             =   495
         Width           =   1185
      End
      Begin VB.TextBox txtRg 
         Height          =   285
         Left            =   3915
         MaxLength       =   50
         TabIndex        =   6
         Top             =   495
         Width           =   2175
      End
      Begin VB.TextBox txtCpf 
         Height          =   285
         Left            =   6135
         MaxLength       =   50
         TabIndex        =   5
         Top             =   480
         Width           =   2670
      End
      Begin VB.Label Label3 
         Caption         =   "Status:"
         Height          =   180
         Left            =   8850
         TabIndex        =   14
         Top             =   255
         Width           =   810
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   180
         Left            =   1365
         TabIndex        =   13
         Top             =   255
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Code:"
         Height          =   180
         Left            =   165
         TabIndex        =   12
         Top             =   255
         Width           =   810
      End
      Begin VB.Label Label4 
         Caption         =   "RG:"
         Height          =   180
         Left            =   3915
         TabIndex        =   11
         Top             =   255
         Width           =   810
      End
      Begin VB.Label Label5 
         Caption         =   "CPF:"
         Height          =   180
         Left            =   6135
         TabIndex        =   10
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   800
      Left            =   13995
      TabIndex        =   3
      Top             =   105
      Width           =   800
   End
   Begin VB.CommandButton btnEnabled 
      Caption         =   "Enable"
      Height          =   800
      Left            =   13185
      TabIndex        =   2
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      Height          =   800
      Left            =   12375
      TabIndex        =   1
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "New"
      Height          =   800
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid dgv 
      Height          =   4965
      Left            =   45
      TabIndex        =   15
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
Attribute VB_Name = "FrmListClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bll As New ClientBll
Dim utils As New utils

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnEnabled_Click()
If dgv.VisibleRows > 0 Then
    dgv.Col = 0
    bll.DeleteClient (dgv.text)
    LoadGrid
End If
End Sub

Private Sub btnNew_Click()
Dim createUpdate As New FrmCreateUpdateClient
createUpdate.Show vbModal
LoadGrid
End Sub

Private Sub btnSearch_Click()
LoadGrid
End Sub

Private Sub dgv_DblClick()
    If dgv.VisibleRows > 0 Then
        Dim update As New FrmCreateUpdateClient
        dgv.Col = 0
        update.txtCode.text = dgv.text
                            
        update.Show vbModal
        LoadGrid
    End If
End Sub

Private Sub Form_Load()
LoadGrid
LoadComboBoxStatus
End Sub

Private Sub LoadComboBoxStatus()
cboStatus.AddItem ("All")
cboStatus.AddItem ("Active")
cboStatus.AddItem ("Inactive")
cboStatus.text = "All"
End Sub

Private Sub LoadGrid()
Set dgv.DataSource = bll.GetClients(txtCode.text, txtName.text, txtRg.text, txtCpf.text, cboStatus.text)

dgv.Columns(0).Caption = "Code"
dgv.Columns(0).Width = 900
dgv.Columns(0).Alignment = dbgCenter

dgv.Columns(1).Caption = "Name"
dgv.Columns(1).Width = 8000

dgv.Columns(2).Caption = "RG"
dgv.Columns(2).Width = 2000
dgv.Columns(2).Alignment = dbgCenter

dgv.Columns(3).Caption = "CPF"
dgv.Columns(3).Width = 2000
dgv.Columns(3).Alignment = dbgCenter


dgv.Columns(4).Caption = "Status"
dgv.Columns(4).Width = 1500
dgv.Columns(4).Alignment = dbgCenter
End Sub


Private Sub txtCode_KeyPress(KeyAscii As Integer)
KeyAscii = utils.OnlyNumbers(KeyAscii)
End Sub

Private Sub txtCpf_KeyPress(KeyAscii As Integer)
KeyAscii = utils.OnlyNumbers(KeyAscii)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
KeyAscii = utils.NumberAndLetters(KeyAscii)
End Sub

Private Sub txtRg_KeyPress(KeyAscii As Integer)
KeyAscii = utils.NumberAndLetters(KeyAscii)
End Sub
