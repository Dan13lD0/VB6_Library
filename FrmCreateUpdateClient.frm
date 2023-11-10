VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCreateUpdateClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Caption:"
      Height          =   945
      Left            =   75
      TabIndex        =   9
      Top             =   5190
      Width           =   6420
      Begin VB.TextBox txtPerson 
         Height          =   315
         Left            =   3165
         TabIndex        =   25
         Top             =   510
         Width           =   3180
      End
      Begin VB.TextBox txtContact 
         Height          =   315
         Left            =   1365
         TabIndex        =   23
         Top             =   525
         Width           =   1710
      End
      Begin VB.TextBox txtType 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   525
         Width           =   1170
      End
      Begin VB.Label Label10 
         Caption         =   "For:"
         Height          =   195
         Left            =   3165
         TabIndex        =   24
         Top             =   255
         Width           =   765
      End
      Begin VB.Label Label9 
         Caption         =   "Contact:"
         Height          =   195
         Left            =   1380
         TabIndex        =   22
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label8 
         Caption         =   "Type:"
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   270
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Address:"
      Height          =   2805
      Left            =   90
      TabIndex        =   8
      Top             =   2310
      Width           =   6405
      Begin VB.TextBox txtCep 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   615
         Width           =   1155
      End
      Begin VB.TextBox txtStreet 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   1470
         Width           =   6180
      End
      Begin VB.TextBox txtComplement 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   2295
         Width           =   2550
      End
      Begin VB.TextBox txtCity 
         Height          =   315
         Left            =   2670
         TabIndex        =   27
         Top             =   2295
         Width           =   1800
      End
      Begin VB.TextBox txtState 
         Height          =   315
         Left            =   4515
         TabIndex        =   26
         Top             =   2295
         Width           =   1800
      End
      Begin VB.Label Label15 
         Caption         =   "Cep:"
         Height          =   210
         Left            =   90
         TabIndex        =   35
         Top             =   225
         Width           =   555
      End
      Begin VB.Label Label14 
         Caption         =   "Street:"
         Height          =   210
         Left            =   90
         TabIndex        =   34
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label13 
         Caption         =   "Complement:"
         Height          =   210
         Left            =   90
         TabIndex        =   33
         Top             =   1905
         Width           =   1230
      End
      Begin VB.Label Label11 
         Caption         =   "City:"
         Height          =   210
         Left            =   2640
         TabIndex        =   32
         Top             =   1905
         Width           =   1200
      End
      Begin VB.Label Label12 
         Caption         =   "State:"
         Height          =   210
         Left            =   4485
         TabIndex        =   31
         Top             =   1905
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   90
      TabIndex        =   7
      Top             =   690
      Width           =   6420
      Begin MSComCtl2.DTPicker dtpBirth 
         Height          =   330
         Left            =   105
         TabIndex        =   18
         Top             =   1020
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         Format          =   84017153
         CurrentDate     =   45235
      End
      Begin VB.TextBox txtCpf 
         Height          =   315
         Left            =   4560
         TabIndex        =   17
         Top             =   1005
         Width           =   1710
      End
      Begin VB.TextBox txtRg 
         Height          =   315
         Left            =   2775
         TabIndex        =   15
         Top             =   1020
         Width           =   1710
      End
      Begin VB.TextBox txtAge 
         Height          =   315
         Left            =   1530
         TabIndex        =   13
         Top             =   1020
         Width           =   1170
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   90
         TabIndex        =   11
         Top             =   420
         Width           =   6210
      End
      Begin VB.Label Label7 
         Caption         =   "Birthday:"
         Height          =   195
         Left            =   75
         TabIndex        =   19
         Top             =   750
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "CPF:"
         Height          =   195
         Left            =   4575
         TabIndex        =   16
         Top             =   750
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "RG:"
         Height          =   195
         Left            =   2790
         TabIndex        =   14
         Top             =   765
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Age:"
         Height          =   195
         Left            =   1545
         TabIndex        =   12
         Top             =   765
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   180
         Width           =   765
      End
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   2610
      TabIndex        =   6
      Top             =   6240
      Width           =   1260
   End
   Begin VB.CommandButton btnClean 
      Caption         =   "Clean"
      Height          =   315
      Left            =   3930
      TabIndex        =   5
      Top             =   6225
      Width           =   1260
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   5250
      TabIndex        =   4
      Top             =   6225
      Width           =   1260
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "FrmCreateUpdateClient.frx":0000
      Left            =   5055
      List            =   "FrmCreateUpdateClient.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   315
      Width           =   1455
   End
   Begin VB.TextBox txtCode 
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Top             =   330
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "Status"
      Height          =   195
      Left            =   5085
      TabIndex        =   1
      Top             =   60
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Code:"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   765
   End
End
Attribute VB_Name = "FrmCreateUpdateClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bll As New ClientBll
Dim bllAddress As New AddressBll
Dim bllContact As New ContactBll
Dim contactId As String
Dim addressId As String

Private Sub btnClean_Click()
    txtName.text = ""
    dtpBirth.Value = ""
    txtAge.text = ""
    txtRg.text = ""
    txtCpf.text = ""
    cboStatus.text = "Active"
    txtCep.text = ""
    txtStreet.text = ""
    txtComplement.text = ""
    txtCity.text = ""
    txtState.text = ""
    txtType.text = ""
    txtContact.text = ""
    txtPerson.text = ""
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
If txtName.text = "" Then
    MsgBox "Fild name is requerid", vbOK, "Notice"
    Exit Sub
End If

If txtRg.text = "" Then
    MsgBox "Fild rg is requerid", vbOK, "Notice"
    Exit Sub
End If

If txtCpf.text = "" Then
    MsgBox "Fild CPF is requerid", vbOK, "Notice"
    Exit Sub
End If

If txtCep.text = "" Then
    MsgBox "Fild cep is requerid", vbOK, "Notice"
    Exit Sub
End If

If txtStreet.text = "" Then
    MsgBox "Fild street is requerid", vbOK, "Notice"
    Exit Sub
End If

If txtComplement.text = "" Then
    MsgBox "Fild complement is requerid", vbOK, "Notice"
    Exit Sub
End If

If txtCity.text = "" Then
    MsgBox "Fild city is requerid", vbOK, "Notice"
    Exit Sub
End If

If txtState.text = "" Then
    MsgBox "Fild state is requerid", vbOK, "Notice"
    Exit Sub
End If

If txtType.text = "" Then
    MsgBox "Fild type is requerid", vbOK, "Notice"
    Exit Sub
End If

If txtContact.text = "" Then
    MsgBox "Fild name is requerid", vbOK, "Notice"
    Exit Sub
End If

If txtPerson.text = "" Then
    MsgBox "Fild name is requerid", vbOK, "Notice"
    Exit Sub
End If

bllContact.CreateUpdateContact contactId, txtType.text, txtContact.text, txtPerson.text

If contactId = "0" Then
    contactId = bllContact.GetLastId()
End If

bllAddress.CreateUpdateAddress addressId, txtCep.text, txtStreet.text, txtComplement.text, txtCity.text, txtState.text

If addressId = "0" Then
    addressId = bllAddress.GetLastId()
End If
bll.CreateUpdateClient txtCode.text, txtName.text, dtpBirth.Value, txtRg.text, txtCpf.text, contactId, addressId, cboStatus.text

If txtCode.text = "" Then
    MsgBox "Created Record", vbOK, "Notice"
Else
    MsgBox "Updated Record", vbOK, "Notice"
End If
Unload Me
End Sub



Private Sub Form_Activate()
contactId = "0"
addressId = "0"
If txtCode.text <> "" Then
Dim record As New ADODB.Recordset
Dim recordAddress As New ADODB.Recordset
Dim recordContact As New ADODB.Recordset

        Me.Caption = "Update record"

    
   Set record = bll.GetClient(txtCode.text)
   With record.Fields
    txtName.text = .Item(1).Value
    dtpBirth.Value = .Item(2).Value
    txtAge.text = .Item(3).Value
    txtRg.text = .Item(4).Value
    txtCpf.text = .Item(5).Value
    contactId = .Item(6).Value
    addressId = .Item(7).Value
    If .Item(8).Value = True Then
        cboStatus.text = "Active"
    Else
        cboStatus.text = "Inactive"
    End If
    
     Set recordAddress = bllAddress.GetAddress(.Item(7).Value)
     
     With recordAddress.Fields
        txtCep.text = .Item(1).Value
        txtStreet.text = .Item(2).Value
        txtComplement.text = .Item(3).Value
        txtCity.text = .Item(4).Value
        txtState.text = .Item(5).Value
     End With
    
     Set recordContact = bllContact.GetContact(.Item(6).Value)
    
      With recordContact.Fields
     txtType.text = .Item(1).Value
     txtContact.text = .Item(2).Value
     txtPerson.text = .Item(3).Value
     End With
   End With
Else
  Me.Caption = "Create new record"
    cboStatus.text = "Active"
End If
End Sub

