VERSION 5.00
Begin VB.Form Welcome 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton user 
      Caption         =   "USER"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   12480
      Picture         =   "Welcome.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton admin 
      Caption         =   "ADMIN"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   6240
      Picture         =   "Welcome.frx":E38C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO DMS"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   3600
      TabIndex        =   0
      Top             =   1440
      Width           =   15765
   End
   Begin VB.Image Image1 
      Height          =   12285
      Left            =   0
      Picture         =   "Welcome.frx":1C718
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22755
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rw As New ADODB.Recordset

Private Sub admin_Click()
admin_login.Show
WindowState = 2
End Sub

Private Sub user_Click()
user_login.Show
WindowState = 2
End Sub
