VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Tom Pydeski's Wallpaper Printer"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4545
   Icon            =   "BlankForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1125
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Text            =   "Display our Update Message "
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub txtMessage_GotFocus()
txtMessage.SelStart = 0
txtMessage.SelLength = Len(txtMessage)
End Sub

Private Sub cmdUpdate_Click()
UpdateWall txtMessage.Text
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdUpdate_Click
End If
End Sub
