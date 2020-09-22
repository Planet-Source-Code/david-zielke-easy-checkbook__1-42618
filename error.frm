VERSION 5.00
Begin VB.Form ErrorForm2 
   Caption         =   "ERROR"
   ClientHeight    =   1356
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   3744
   LinkTopic       =   "Form2"
   ScaleHeight     =   1356
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   252
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   972
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "You must enter the date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   1956
   End
End
Attribute VB_Name = "ErrorForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ErrorForm2.Hide
End Sub
