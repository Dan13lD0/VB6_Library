VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBorrow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtpReturn 
      Height          =   330
      Left            =   1635
      TabIndex        =   37
      Top             =   9915
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      Format          =   203096065
      CurrentDate     =   45235
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   345
      Left            =   7815
      TabIndex        =   36
      Top             =   9795
      Width           =   1080
   End
   Begin VB.CommandButton btnClean 
      Caption         =   "Clean"
      Height          =   345
      Left            =   6705
      TabIndex        =   35
      Top             =   9795
      Width           =   1080
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   5580
      TabIndex        =   34
      Top             =   9795
      Width           =   1080
   End
   Begin VB.Frame Frame3 
      Caption         =   "Client"
      Height          =   5550
      Left            =   75
      TabIndex        =   2
      Top             =   3960
      Width           =   8835
      Begin VB.TextBox txtPerson 
         Height          =   315
         Left            =   6900
         TabIndex        =   33
         Top             =   5070
         Width           =   1800
      End
      Begin VB.TextBox txtContact 
         Height          =   315
         Left            =   5055
         TabIndex        =   31
         Top             =   5070
         Width           =   1800
      End
      Begin VB.TextBox txtType 
         Height          =   315
         Left            =   165
         TabIndex        =   29
         Top             =   5070
         Width           =   4815
      End
      Begin VB.TextBox txtState 
         Height          =   315
         Left            =   6900
         TabIndex        =   27
         Top             =   4125
         Width           =   1800
      End
      Begin VB.TextBox txtCity 
         Height          =   315
         Left            =   5055
         TabIndex        =   25
         Top             =   4125
         Width           =   1800
      End
      Begin VB.TextBox txtComplement 
         Height          =   315
         Left            =   165
         TabIndex        =   23
         Top             =   4125
         Width           =   4815
      End
      Begin VB.TextBox txtStreet 
         Height          =   315
         Left            =   165
         TabIndex        =   21
         Top             =   3300
         Width           =   8535
      End
      Begin VB.TextBox txtCpf 
         Height          =   315
         Left            =   1920
         TabIndex        =   19
         Top             =   1620
         Width           =   1680
      End
      Begin VB.TextBox txtRg 
         Height          =   315
         Left            =   165
         TabIndex        =   17
         Top             =   1620
         Width           =   1680
      End
      Begin VB.TextBox txtCep 
         Height          =   315
         Left            =   165
         TabIndex        =   15
         Top             =   2445
         Width           =   1155
      End
      Begin VB.TextBox txtClientName 
         Height          =   315
         Left            =   165
         TabIndex        =   13
         Top             =   720
         Width           =   8355
      End
      Begin VB.Label Label15 
         Caption         =   "For:"
         Height          =   210
         Left            =   6870
         TabIndex        =   32
         Top             =   4680
         Width           =   1200
      End
      Begin VB.Label Label14 
         Caption         =   "Contact:"
         Height          =   210
         Left            =   5025
         TabIndex        =   30
         Top             =   4680
         Width           =   1200
      End
      Begin VB.Label Label13 
         Caption         =   "Type:"
         Height          =   210
         Left            =   135
         TabIndex        =   28
         Top             =   4680
         Width           =   555
      End
      Begin VB.Label Label12 
         Caption         =   "State:"
         Height          =   210
         Left            =   6870
         TabIndex        =   26
         Top             =   3735
         Width           =   1200
      End
      Begin VB.Label Label11 
         Caption         =   "City:"
         Height          =   210
         Left            =   5025
         TabIndex        =   24
         Top             =   3735
         Width           =   1200
      End
      Begin VB.Label Label10 
         Caption         =   "Complement:"
         Height          =   210
         Left            =   135
         TabIndex        =   22
         Top             =   3735
         Width           =   1230
      End
      Begin VB.Label Label9 
         Caption         =   "Street:"
         Height          =   210
         Left            =   135
         TabIndex        =   20
         Top             =   2910
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "CPF:"
         Height          =   210
         Left            =   1890
         TabIndex        =   18
         Top             =   1230
         Width           =   555
      End
      Begin VB.Label Label7 
         Caption         =   "RG:"
         Height          =   210
         Left            =   135
         TabIndex        =   16
         Top             =   1230
         Width           =   555
      End
      Begin VB.Label Label6 
         Caption         =   "Cep:"
         Height          =   210
         Left            =   135
         TabIndex        =   14
         Top             =   2055
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "Name:"
         Height          =   210
         Left            =   135
         TabIndex        =   12
         Top             =   330
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search Client"
      Height          =   750
      Left            =   75
      TabIndex        =   1
      Top             =   3135
      Width           =   8835
      Begin VB.ComboBox cboClient 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   315
         Width           =   8490
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Book"
      Height          =   2985
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   8835
      Begin VB.TextBox txtAuthor 
         Height          =   315
         Left            =   150
         TabIndex        =   11
         Top             =   1605
         Width           =   8520
      End
      Begin VB.TextBox txtPublisher 
         Height          =   315
         Left            =   4440
         TabIndex        =   9
         Top             =   2430
         Width           =   4260
      End
      Begin VB.TextBox txtCategory 
         Height          =   315
         Left            =   150
         TabIndex        =   7
         Top             =   2430
         Width           =   4260
      End
      Begin VB.TextBox txtBookName 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   705
         Width           =   8520
      End
      Begin VB.Label Label4 
         Caption         =   "Author:"
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   1215
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Publisher:"
         Height          =   210
         Left            =   4410
         TabIndex        =   8
         Top             =   2040
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Category:"
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   555
      End
   End
   Begin MSComCtl2.DTPicker dtpBorrow 
      Height          =   330
      Left            =   105
      TabIndex        =   38
      Top             =   9930
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      Format          =   138674177
      CurrentDate     =   45235
   End
   Begin VB.Label Label17 
      Caption         =   "Date Return:"
      Height          =   180
      Left            =   1650
      TabIndex        =   40
      Top             =   9630
      Width           =   1140
   End
   Begin VB.Label Label16 
      Caption         =   "Date Borrow:"
      Height          =   270
      Left            =   135
      TabIndex        =   39
      Top             =   9630
      Width           =   1065
   End
End
Attribute VB_Name = "FrmBorrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bookId As String
Dim clientId As String
Dim bllBook As New BookBll
Dim bllClient As New ClientBll
Dim bllAddress As New AddressBll
Dim bllContact As New ContactBll
Dim bll As New BorrowBll
Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
If cboClient.text <> "" Then
bll.CreateUpdateBorrow bookId, clientId, dtpBorrow.Value, dtpReturn.Value
MsgBox "Borrow created with successful!", vbOK, "Borrow"
Unload Me
Else
MsgBox "Field client is requerid!", vbOK, "Notice"
End If
End Sub

Private Sub cboClient_LostFocus()
If cboClient.text <> "" Then
    Dim clientRecord As New ADODB.Recordset
    Dim addressRecord As New ADODB.Recordset
    Dim contactRecord As New ADODB.Recordset
    
    Set clientRecord = bllClient.GetClientByName(cboClient.text)
     With clientRecord.Fields
        clientId = .Item(0).Value
        txtClientName.text = .Item(1).Value
        txtRg.text = .Item(4).Value
        txtCpf.text = .Item(5).Value
        
        Set addressRecord = bllAddress.GetAddress(.Item(7).Value)
        
        With addressRecord.Fields
            txtCep.text = .Item(1).Value
            txtStreet.text = .Item(2).Value
            txtComplement.text = .Item(3).Value
            txtCity.text = .Item(4).Value
            txtState.text = .Item(5).Value
        End With
        
        
        Set contactRecord = bllContact.GetContact(.Item(6).Value)
                
        With contactRecord.Fields
            txtType.text = .Item(1).Value
            txtContact.text = .Item(2).Value
            txtPerson.text = .Item(3).Value
        End With
     
     End With
    
End If
End Sub

Private Sub Form_Load()
Dim bookRecord As New ADODB.Recordset
Dim clientRecord As New ADODB.Recordset

Set bookRecord = bllBook.GetBook(bookId)

If bookRecord.RecordCount > 0 Then
    With bookRecord.Fields
    txtBookName.text = .Item(1).Value
    txtCategory.text = .Item(2).Value
    txtAuthor.text = .Item(3).Value
    txtPublisher.text = .Item(4).Value
    End With
End If

Set clientRecord = bllClient.GetClients("0", "", "", "", "Active")
Dim clientId As String
If clientRecord.RecordCount > 0 Then
    While Not clientRecord.EOF
        With clientRecord.Fields
        cboClient.AddItem (.Item(1).Value)
        End With
    clientRecord.MoveNext
    Wend
End If




End Sub
