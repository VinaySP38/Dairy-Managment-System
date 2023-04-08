VERSION 5.00
Begin VB.Form admin_page 
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
   Begin VB.CommandButton Command1 
      Caption         =   "Modify Rate"
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
      Picture         =   "admin_page.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton payment 
      Caption         =   "Payment Details"
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
      Left            =   5880
      Picture         =   "admin_page.frx":E38C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4440
      Picture         =   "admin_page.frx":1C718
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton mod 
      Caption         =   "Modify Customer"
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
      Left            =   5880
      Picture         =   "admin_page.frx":2AAA4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton deposit 
      Caption         =   "Deposit Milk"
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
      Picture         =   "admin_page.frx":38E30
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Page"
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
      Left            =   5880
      TabIndex        =   1
      Top             =   720
      Width           =   9750
   End
   Begin VB.Image Image1 
      Height          =   12285
      Left            =   -120
      Picture         =   "admin_page.frx":471BC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22875
   End
End
Attribute VB_Name = "admin_page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rw As New ADODB.Recordset
Private Sub deposit_Click()
Billing.Show
WindowState = 2
End Sub

Private Sub exit_Click()
Unload Me
Welcome.Show
WindowState = 2
End Sub

Private Sub report_Click()
payment.Show
WindowState = 2
End Sub

Private Sub modify_Click()
modify.Show
WindowState = 2
End Sub

Private Sub mod_Click()
modify.Show
WindowState = 2
End Sub
