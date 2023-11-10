VERSION 5.00
Begin VB.Form FrmCreateUpdateBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   1245
      TabIndex        =   17
      Top             =   5355
      Width           =   1260
   End
   Begin VB.CommandButton btnClean 
      Caption         =   "Clean"
      Height          =   315
      Left            =   2565
      TabIndex        =   16
      Top             =   5340
      Width           =   1260
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   3885
      TabIndex        =   15
      Top             =   5340
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Height          =   4500
      Left            =   135
      TabIndex        =   4
      Top             =   765
      Width           =   5040
      Begin VB.TextBox txtDescription 
         Height          =   1320
         Left            =   45
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "FrmCreateUpdateBook.frx":0000
         Top             =   3030
         Width           =   4770
      End
      Begin VB.ComboBox cboPublisher 
         Height          =   315
         Left            =   30
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2385
         Width           =   4890
      End
      Begin VB.ComboBox cboAuthor 
         Height          =   315
         Left            =   45
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1725
         Width           =   4890
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1095
         Width           =   4890
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   60
         MaxLength       =   150
         TabIndex        =   6
         Top             =   405
         Width           =   4860
      End
      Begin VB.Label Label7 
         Caption         =   "Description:"
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   2820
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "Publisher:"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   2130
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "Author:"
         Height          =   195
         Left            =   75
         TabIndex        =   9
         Top             =   1470
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "Category:"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   195
         Left            =   75
         TabIndex        =   5
         Top             =   150
         Width           =   765
      End
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Left            =   3690
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   420
      Width           =   1455
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   315
      Left            =   165
      TabIndex        =   2
      Top             =   435
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "Status"
      Height          =   195
      Left            =   3720
      TabIndex        =   1
      Top             =   165
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Code:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   765
   End
End
Attribute VB_Name = "FrmCreateUpdateBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bllCategory As New CategoryBll
Dim bllAuthor As New AuthorBll
Dim bllPublisher As New PublisherBll
Dim bll As New BookBll
Dim utils As New utils
Private Sub btnClean_Click()
txtName.text = ""
txtDescription.text = ""
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnSave_Click()

If txtName.text = "" Then
MsgBox "Field name is requerid!", vbOKOnly, "Notive"
Exit Sub
End If

If cboCategory.text = "" Then
MsgBox "Field category is requerid!", vbOKOnly, "Notive"
Exit Sub
End If

If cboAuthor.text = "" Then
MsgBox "Field author is requerid!", vbOKOnly, "Notive"
Exit Sub
End If

If cboPublisher.text = "" Then
MsgBox "Field publisher is requerid!", vbOKOnly, "Notive"
Exit Sub
End If


bll.CreateUpdateBook txtCode.text, txtName.text, cboCategory.text, cboAuthor.text, cboPublisher.text, txtDescription.text, cboStatus.text

If txtCode.text = "" Then
MsgBox "Record created with successful!", vbOKOnly, "Create"
Else
MsgBox "Record updated with successful!", vbOKOnly, "Update"
End If

Unload Me

End Sub

Private Sub Form_Activate()
If txtCode.text = "" Then
    Me.Caption = "Create Record"
    cboStatus.text = "Active"
Else
    Me.Caption = "Update Record"
End If
End Sub

Private Sub Form_Load()
LoadComboBox
End Sub

Private Sub LoadComboBox()

cboStatus.AddItem ("Active")
cboStatus.AddItem ("Inactive")
cboStatus.text = "Active"


Dim recordCategory As New ADODB.Recordset

Set recordCategory = bllCategory.GetCategories("0", "", "Active")

cboCategory.AddItem ("Select")
While Not recordCategory.EOF
cboCategory.AddItem (recordCategory.Fields.Item(1).Value)
recordCategory.MoveNext
Wend
cboCategory.text = "Select"

Dim recordAuthor As New ADODB.Recordset

Set recordAuthor = bllAuthor.GetAuthors("0", "", "Active")

cboAuthor.AddItem ("Select")
While Not recordAuthor.EOF
cboAuthor.AddItem (recordAuthor.Fields.Item(1).Value)
recordAuthor.MoveNext
Wend
cboAuthor.text = "Select"

Dim recordPublisher As New ADODB.Recordset
cboPublisher.AddItem ("Select")
Set recordPublisher = bllPublisher.GetPublishers("0", "", "Active")
While Not recordPublisher.EOF
cboPublisher.AddItem (recordPublisher.Fields.Item(1).Value)
recordPublisher.MoveNext
Wend
cboPublisher.text = "Select"
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
KeyAscii = utils.NumberAndLetters(KeyAscii)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

KeyAscii = utils.NumberAndLetters(KeyAscii)
End Sub
