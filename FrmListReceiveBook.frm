VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmListReceiveBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      Height          =   800
      Left            =   9975
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   800
      Left            =   10830
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter:"
      Height          =   885
      Left            =   210
      TabIndex        =   0
      Top             =   30
      Width           =   9660
      Begin VB.TextBox txtClient 
         Height          =   285
         Left            =   4545
         TabIndex        =   6
         Top             =   495
         Width           =   4275
      End
      Begin VB.TextBox txtBook 
         Height          =   285
         Left            =   165
         TabIndex        =   1
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Client Name:"
         Height          =   180
         Left            =   4545
         TabIndex        =   7
         Top             =   255
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Book Name:"
         Height          =   180
         Left            =   165
         TabIndex        =   2
         Top             =   255
         Width           =   1080
      End
   End
   Begin MSDataGridLib.DataGrid dgv 
      Height          =   4965
      Left            =   60
      TabIndex        =   5
      Top             =   1035
      Width           =   11565
      _ExtentX        =   20399
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
Attribute VB_Name = "FrmListReceiveBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bll As New BorrowBll
Dim utils As New utils
Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnSearch_Click()
LoadGrid
End Sub

Private Sub LoadGrid()
Set dgv.DataSource = bll.GetBorrows2(txtClient.text, txtBook.text)

dgv.Columns(0).Width = 1500
dgv.Columns(0).Caption = "Code"

dgv.Columns(1).Width = 2700
dgv.Columns(1).Caption = "Book Name"

dgv.Columns(2).Width = 2500
dgv.Columns(2).Caption = "Client Name"

dgv.Columns(3).Width = 1500
dgv.Columns(3).Caption = "Borrow"

dgv.Columns(4).Width = 1500
dgv.Columns(4).Caption = "Return"

dgv.Columns(5).Width = 1500
dgv.Columns(5).Caption = "Status"
End Sub

Private Sub dgv_DblClick()
If dgv.VisibleRows > 0 Then
  If utils.MessageBoxConfirmation("Would you like to return this book?", "Return") = True Then
  dgv.Col = 0
  bll.BorrowChangeStatus (dgv.text)
  LoadGrid
  End If
End If
End Sub

Private Sub Form_Load()
LoadGrid
End Sub

