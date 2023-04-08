VERSION 5.00
Begin VB.Form user_page 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton feedback 
      Caption         =   "Feedback"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   7080
      Picture         =   "user_page.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton view 
      Caption         =   "View Bill"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   5520
      Picture         =   "user_page.frx":E38C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   3960
      Picture         =   "user_page.frx":1C718
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton profile 
      Caption         =   "My Profile"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   2400
      Picture         =   "user_page.frx":2AAA4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Page"
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
      Left            =   6720
      TabIndex        =   0
      Top             =   600
      Width           =   7980
   End
   Begin VB.Image Image1 
      Height          =   12285
      Left            =   0
      Picture         =   "user_page.frx":38E30
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22755
   End
End
Attribute VB_Name = "user_page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rw As New ADODB.Recordset

Private Sub Command1_Click()

End Sub

Private Sub exit_Click()
Unload Me
user_login.Show
WindowState = 2
End Sub

Private Sub profile_Click()
On Error GoTo errmsg
With Adodc1.Recordset
.MoveFirst
.Find "u_name='" & u_name.Text & "'"
If .EOF = True Then
prof.Show
WindowState = 2
Unload Me
u_name = Adodc1.Recordset("u_name")
u_id = Adodc1.Recordset("u_id")
cadd = Adodc1.Recordset("cadd")
cpno = Adodc1.Recordset("cpno")
Selec = Adodc1.Recordset("milk")
pass = Adodc1.Recordset("pass")
acno = Adodc1.Recordset("acno")
End If
End With
End Sub

