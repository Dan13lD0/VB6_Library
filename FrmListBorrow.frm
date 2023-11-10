VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmListBorrow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   14850
   Begin VB.Frame Frame1 
      Caption         =   "Filter:"
      Height          =   885
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   12990
      Begin VB.TextBox txtBook 
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   495
         Width           =   3015
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         Left            =   3225
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   2925
      End
      Begin VB.ComboBox cboAuthor 
         Height          =   315
         Left            =   6285
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   2400
      End
      Begin VB.ComboBox cboPublisher 
         Height          =   315
         Left            =   8700
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   3345
      End
      Begin VB.Label Label5 
         Caption         =   "Author:"
         Height          =   180
         Left            =   6300
         TabIndex        =   10
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Category:"
         Height          =   180
         Left            =   3240
         TabIndex        =   9
         Top             =   255
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Book Name:"
         Height          =   180
         Left            =   165
         TabIndex        =   8
         Top             =   255
         Width           =   810
      End
      Begin VB.Label Label6 
         Caption         =   "Publisher"
         Height          =   180
         Left            =   8715
         TabIndex        =   7
         Top             =   255
         Width           =   1305
      End
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   800
      Left            =   13950
      TabIndex        =   1
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      Height          =   800
      Left            =   13125
      TabIndex        =   0
      Top             =   90
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid dgv 
      Height          =   4965
      Left            =   15
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
Attribute VB_Name = "FrmListBorrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bll As New BorrowBll
Dim bllCategory As New CategoryBll
Dim bllAuthor As New AuthorBll
Dim bllPublisher As New PublisherBll
Dim utils As New utils

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnNew_Click()
Dim createUpdate As New FrmChoseCreateReturn
createUpdate.Show vbModal
End Sub

Private Sub btnSearch_Click()
LoadGrid
End Sub

Private Sub dgv_DblClick()
If dgv.VisibleRows > 0 Then

Dim action As New FrmChoseCreateReturn
dgv.Col = 5
If dgv.text <> "Unvaliable" Then
action.lblMessage.Caption = "Would you like to borrow this book?"
action.btnAction.Caption = "Borrow"
dgv.Col = 0
action.id = dgv.text
Else
action.lblMessage.Caption = "Would you like to return this book?"
action.btnAction.Caption = "Return"
dgv.Col = 0
action.id = dgv.text
End If
action.Show vbModal
LoadGrid
End If
End Sub

Private Sub Form_Load()
LoadComboBox
LoadGrid
End Sub


Private Sub LoadGrid()
Set dgv.DataSource = bll.GetBorrows(txtBook.text, cboCategory.text, cboAuthor.text, cboPublisher.text)

dgv.Columns(0).Width = 1000
dgv.Columns(1).Width = 3200
dgv.Columns(2).Width = 3000
dgv.Columns(3).Width = 3000
dgv.Columns(4).Width = 3000
End Sub




Private Sub LoadComboBox()

Dim recordCategory As New ADODB.Recordset

Set recordCategory = bllCategory.GetCategories("0", "", "Active")

cboCategory.AddItem ("All")
While Not recordCategory.EOF
cboCategory.AddItem (recordCategory.Fields.Item(1).Value)
recordCategory.MoveNext
Wend
cboCategory.text = "All"

Dim recordAuthor As New ADODB.Recordset

Set recordAuthor = bllAuthor.GetAuthors("0", "", "Active")

cboAuthor.AddItem ("All")
While Not recordAuthor.EOF
cboAuthor.AddItem (recordAuthor.Fields.Item(1).Value)
recordAuthor.MoveNext
Wend
cboAuthor.text = "All"

Dim recordPublisher As New ADODB.Recordset
cboPublisher.AddItem ("All")
Set recordPublisher = bllPublisher.GetPublishers("0", "", "Active")
While Not recordPublisher.EOF
cboPublisher.AddItem (recordPublisher.Fields.Item(1).Value)
recordPublisher.MoveNext
Wend
cboPublisher.text = "All"
End Sub
