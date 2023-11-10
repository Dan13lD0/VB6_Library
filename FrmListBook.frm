VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmListBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Books"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   14835
   Begin VB.Frame Frame1 
      Caption         =   "Filter:"
      Height          =   885
      Left            =   915
      TabIndex        =   4
      Top             =   0
      Width           =   11400
      Begin VB.ComboBox cboPublisher 
         Height          =   315
         Left            =   7740
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   480
         Width           =   1965
      End
      Begin VB.ComboBox cboAuthor 
         Height          =   315
         Left            =   6030
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   480
         Width           =   1680
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   2115
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   9735
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1365
         MaxLength       =   100
         TabIndex        =   6
         Top             =   495
         Width           =   2505
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   150
         MaxLength       =   10
         TabIndex        =   5
         Top             =   495
         Width           =   1185
      End
      Begin VB.Label Label6 
         Caption         =   "Publisher:"
         Height          =   180
         Left            =   7755
         TabIndex        =   17
         Top             =   255
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Status:"
         Height          =   180
         Left            =   9705
         TabIndex        =   12
         Top             =   255
         Width           =   810
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   180
         Left            =   1365
         TabIndex        =   11
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
      Begin VB.Label Label4 
         Caption         =   "Category:"
         Height          =   180
         Left            =   3915
         TabIndex        =   9
         Top             =   255
         Width           =   810
      End
      Begin VB.Label Label5 
         Caption         =   "Author:"
         Height          =   180
         Left            =   6045
         TabIndex        =   8
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   800
      Left            =   13965
      TabIndex        =   3
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enable"
      Height          =   800
      Left            =   13140
      TabIndex        =   2
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   800
      Left            =   12330
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
      Left            =   15
      TabIndex        =   13
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
Attribute VB_Name = "FrmListBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bllCategory As New CategoryBll
Dim bllAuthor As New AuthorBll
Dim bllPublisher As New PublisherBll
Dim bll As New BookBll
Dim utils As New utils
Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnNew_Click()
Dim createUpdate As New FrmCreateUpdateBook
createUpdate.Show vbModal
LoadGrid
End Sub

Private Sub Command2_Click()
LoadGrid
End Sub

Private Sub Command3_Click()
If dgv.VisibleRows > 0 Then
    dgv.Col = 0
    bll.DeleteBook (dgv.text)
    LoadGrid
End If
End Sub

Private Sub dgv_DblClick()
    If dgv.VisibleRows > 0 Then
        Dim update As New FrmCreateUpdateBook
        dgv.Col = 0
        update.txtCode.text = dgv.text
        
        dgv.Col = 1
        update.txtName.text = dgv.text
        
        dgv.Col = 2
        update.cboCategory.text = dgv.text
        
        dgv.Col = 3
        update.cboAuthor.text = dgv.text
        
        dgv.Col = 4
        update.cboPublisher.text = dgv.text
        
        dgv.Col = 5
        update.cboStatus.text = dgv.text
                        
        update.Show vbModal
        LoadGrid
    End If
End Sub

Private Sub Form_Load()
LoadComboBox
LoadGrid
End Sub



Private Sub LoadGrid()
Set dgv.DataSource = bll.GetBooks(txtName.text, cboCategory.text, cboAuthor.text, cboPublisher.text, cboStatus.text)

dgv.Columns(0).Caption = "Code"
dgv.Columns(0).Width = 1000
dgv.Columns(0).Alignment = dbgCenter

dgv.Columns(1).Caption = "Name"
dgv.Columns(1).Width = 4800

dgv.Columns(2).Caption = "Category"
dgv.Columns(2).Width = 2500

dgv.Columns(3).Caption = "Author"
dgv.Columns(3).Width = 2500

dgv.Columns(4).Caption = "Publisher"
dgv.Columns(4).Width = 2500


dgv.Columns(5).Caption = "Status"
dgv.Columns(5).Width = 1000
dgv.Columns(5).Alignment = dbgCenter
End Sub

Private Sub LoadComboBox()
cboStatus.AddItem ("All")
cboStatus.AddItem ("Active")
cboStatus.AddItem ("Inactive")
cboStatus.text = "All"


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

Private Sub txtCode_KeyPress(KeyAscii As Integer)
KeyAscii = utils.OnlyNumbers(KeyAscii)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
KeyAscii = utils.NumberAndLetters(KeyAscii)
End Sub
