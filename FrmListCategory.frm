VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmListCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   11760
   Begin MSDataGridLib.DataGrid dgv 
      Height          =   4830
      Left            =   105
      TabIndex        =   11
      Top             =   1110
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   8520
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Enabled         =   -1  'True
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
   Begin VB.Frame Frame1 
      Caption         =   "Filter:"
      Height          =   885
      Left            =   1005
      TabIndex        =   4
      Top             =   15
      Width           =   8145
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
         Width           =   5085
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   6510
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   1575
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
         Left            =   6480
         TabIndex        =   8
         Top             =   255
         Width           =   810
      End
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   800
      Left            =   10890
      TabIndex        =   3
      Top             =   105
      Width           =   800
   End
   Begin VB.CommandButton btnEnableDisable 
      Caption         =   "Enable"
      Height          =   800
      Left            =   10065
      TabIndex        =   2
      Top             =   105
      Width           =   800
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      Height          =   800
      Left            =   9255
      TabIndex        =   1
      Top             =   105
      Width           =   800
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "New"
      Height          =   800
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   800
   End
End
Attribute VB_Name = "FrmListCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bll As New CategoryBll
Dim utils As New utils

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnEnableDisable_Click()

If dgv.VisibleRows > 0 Then
    dgv.Col = 0
    bll.DeleteCategory (dgv.text)
    LoadGrid
End If
End Sub

Private Sub btnNew_Click()
Dim createUpdate As New FrmCreateUpdateCategory
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
        Dim update As New FrmCreateUpdateCategory
        dgv.Col = 0
        update.txtCode.text = dgv.text
        
        dgv.Col = 1
        update.txtName.text = dgv.text
        
        dgv.Col = 2
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
Set dgv.DataSource = bll.GetCategories(txtCode.text, txtName.text, cboStatus.text)

dgv.Columns(0).Caption = "Code"
dgv.Columns(0).Width = 1000
dgv.Columns(0).Alignment = dbgCenter

dgv.Columns(1).Caption = "Name"
dgv.Columns(1).Width = 9200


dgv.Columns(2).Caption = "Status"
dgv.Columns(2).Width = 1000
dgv.Columns(2).Alignment = dbgCenter
End Sub

Private Sub LoadComboBoxStatus()
cboStatus.AddItem ("All")
cboStatus.AddItem ("Active")
cboStatus.AddItem ("Inactive")
cboStatus.text = "All"
End Sub

