VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Validasi TextBox"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dalam contoh ini, kursor tidak akan dapat keluar dari 'textbox, sampai user mengetik: "abc". Untuk memeriksa 'coding ini, coba klik pada tombol atau enter textbox 'yang kedua.


Private Sub Text3_Validate(Cancel As Boolean)
Dim X As String
    Cancel = Text3.Text <> X & "." & X & "@" & X & "." & X

End Sub
