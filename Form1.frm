VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mendeteksi Lama waktu Windows Dijalankan"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   MsgBox "Windows telah berjalan selama " _
           & GetTickCount \ 1000 & " detik", _
           vbInformation, _
           "Informasi Lama Windows Berjalan"
End Sub


