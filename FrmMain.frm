VERSION 5.00
Begin VB.MDIForm FrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu btnAction 
      Caption         =   "Action"
      Index           =   1
      Begin VB.Menu btnBook 
         Caption         =   "Book"
         Index           =   1
         Begin VB.Menu btnAuthor 
            Caption         =   "Author"
            Index           =   1
         End
         Begin VB.Menu btnBookList 
            Caption         =   "Book List"
            Index           =   3
         End
         Begin VB.Menu btnCategory 
            Caption         =   "Category"
            Index           =   2
         End
         Begin VB.Menu btnPublisher 
            Caption         =   "Publisher"
            Index           =   4
         End
      End
      Begin VB.Menu btnClient 
         Caption         =   "Client"
         Index           =   2
      End
      Begin VB.Menu btnUser 
         Caption         =   "User"
         Index           =   3
      End
   End
   Begin VB.Menu btnMovimentation 
      Caption         =   "Movimentation"
      Index           =   2
      Begin VB.Menu btnBorrow 
         Caption         =   "Borrow Book"
         Index           =   1
      End
      Begin VB.Menu btnReceive 
         Caption         =   "Receive Book"
         Index           =   2
      End
   End
   Begin VB.Menu btnAbout 
      Caption         =   "About"
      Index           =   2
   End
   Begin VB.Menu btnLogOut 
      Caption         =   "LogOut"
      Index           =   4
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private utils As New utils


Private Sub btnAbout_Click(Index As Integer)
    Dim about As New FrmAbout
    about.Show
End Sub

Private Sub btnAuthor_Click(Index As Integer)
    Dim author As New FrmListAuthor
    author.Show
End Sub

Private Sub btnBookList_Click(Index As Integer)
    Dim book As New FrmListBook
    book.Show
End Sub

Private Sub btnBorrow_Click(Index As Integer)
    Dim borrow As New FrmListBorrow
    borrow.Show
End Sub

Private Sub btnCategory_Click(Index As Integer)
    Dim category As New FrmListCategory
    category.Show
End Sub

Private Sub btnClient_Click(Index As Integer)
    Dim client As New FrmListClient
    client.Show
End Sub

Private Sub btnLogOut_Click(Index As Integer)
    If utils.MessageBoxConfirmation("Would you like to logout?", "LogOut") = True Then
        Unload Me
    End If
End Sub

Private Sub btnPublisher_Click(Index As Integer)
    Dim publisher As New FrmListPublisher
    publisher.Show
End Sub

Private Sub btnReceive_Click(Index As Integer)
    Dim receive As New FrmListReceiveBook
    receive.Show
End Sub

Private Sub btnUser_Click(Index As Integer)
    Dim user As New FrmListUser
    user.Show
End Sub
