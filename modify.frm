VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form modify 
   Caption         =   "modify"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "modify.frx":0000
      Height          =   855
      Left            =   12720
      TabIndex        =   20
      Top             =   4800
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1508
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "u_name"
         Caption         =   "u_name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "u_id"
         Caption         =   "u_id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cadd"
         Caption         =   "cadd"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "cpno"
         Caption         =   "cpno"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "milk"
         Caption         =   "milk"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "pass"
         Caption         =   "pass"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "acno"
         Caption         =   "acno"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   12720
      Top             =   3480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"modify.frx":0015
      OLEDBString     =   $"modify.frx":00A7
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from reg"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton delete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Picture         =   "modify.frx":0139
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8640
      Width           =   2175
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Picture         =   "modify.frx":E4C5
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8640
      Width           =   2175
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Modify"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Picture         =   "modify.frx":1C851
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Picture         =   "modify.frx":2ABDD
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      Picture         =   "modify.frx":38F69
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ComboBox Selec 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "modify.frx":472F5
      Left            =   5640
      List            =   "modify.frx":472FF
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox acno 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox pass 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox cpno 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox cadd 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox u_id 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox u_name 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modify Details"
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
      Left            =   5040
      TabIndex        =   14
      Top             =   240
      Width           =   11760
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank A/c No. :"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      TabIndex        =   6
      Top             =   6600
      Width           =   2550
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      TabIndex        =   5
      Top             =   6000
      Width           =   1905
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Milk Type :"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      TabIndex        =   4
      Top             =   5400
      Width           =   1965
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No. :"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      TabIndex        =   3
      Top             =   4800
      Width           =   2085
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      TabIndex        =   2
      Top             =   4200
      Width           =   1650
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User_id :"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      TabIndex        =   1
      Top             =   3600
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      TabIndex        =   0
      Top             =   3000
      Width           =   2025
   End
   Begin VB.Image Image1 
      Height          =   12285
      Left            =   0
      Picture         =   "modify.frx":47311
      Stretch         =   -1  'True
      Top             =   0
      Width           =   22755
   End
End
Attribute VB_Name = "modify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rw As New ADODB.Recordset

Private Sub cmdadd_Click()
Dim id As Integer
Dim id1 As String
u_name.Enabled = True
u_id.Enabled = True
cadd.Enabled = True
cpno.Enabled = True
Selec.Enabled = True
pass.Enabled = True
acno.Enabled = True
u_name = ""
u_id = ""
cadd = ""
cpno = ""
pass = ""
acno = ""
u_name.SetFocus
Adodc1.Recordset.AddNew
Exit Sub
errmsg:
MsgBox Err.Description
End Sub

Private Sub cmdexit_Click()
Unload Me
admin_page.Show
End Sub

Private Sub cmdfind_Click()
Dim id1 As String
Adodc1.Refresh
id1 = InputBox("Enter the User ID to be modified")
Adodc1.Recordset.Find "u_id='" & id1 & "'"
If Adodc1.Recordset.EOF Then
MsgBox " Customer id not found"
Else
u_name = Adodc1.Recordset("u_name")
u_id = Adodc1.Recordset("u_id")
cadd = Adodc1.Recordset("cadd")
cpno = Adodc1.Recordset("cpno")
Selec = Adodc1.Recordset("milk")
pass = Adodc1.Recordset("pass")
acno = Adodc1.Recordset("acno")
End If
Exit Sub
errmsg:
MsgBox Err.Description
End Sub
Private Sub cpno_Click()
If Not Len(cpno) = 10 Then
MsgBox "enter a valid phone no"
cpno = ""
End If
End Sub

Private Sub cmdsave_Click()
On Error GoTo errmsg
If Not Len(cpno) = 10 Then
MsgBox "enter a valid phone no"
cpno = ""
End If
If (u_name.Text = "" Or u_id.Text = "" Or cadd.Text = "" Or cpno.Text = "" Or Selec.Text = "" Or pass.Text = "" Or acno.Text = "") Then
MsgBox "Please fill the empty fields"
Else
Adodc1.Recordset.Fields("u_name") = u_name.Text
Adodc1.Recordset.Fields("u_id") = u_id.Text
Adodc1.Recordset.Fields("cadd") = cadd.Text
Adodc1.Recordset.Fields("cpno") = cpno.Text
Adodc1.Recordset.Fields("milk") = Selec.Text
Adodc1.Recordset.Fields("pass") = pass.Text
Adodc1.Recordset.Fields("acno") = acno.Text
Adodc1.Recordset.Update
MsgBox "User details modified sucessfully"
End If
Exit Sub
errmsg:
MsgBox Err.Description
End Sub

Private Sub delete_Click()
On Error GoTo errmsg
Dim confirm As Integer
Adodc1.Refresh
Dim bno As String
bno = u_id.Text
confirm = MsgBox("Do you really want to delete this supplier y/n?", vbYesNo + vbInformation)
If confirm = vbYes Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find "u_id='" & bno & "'"
Adodc1.Recordset.Delete
DataGrid1.Refresh
MsgBox "This supplier is deleted successfully"
Else
MsgBox "Sorry wrong supplier id"
End If
Exit Sub
errmsg:
MsgBox Err.Description
End Sub

Private Sub form_load()
u_name = ""
u_id = ""
cadd = ""
cpno = ""
pass = ""
acno = ""
u_name.Enabled = False
u_id.Enabled = False
cadd.Enabled = False
cpno.Enabled = False
Selec.Enabled = False
pass.Enabled = False
acno.Enabled = False
End Sub

Private Sub cpno_keypress(keyascii As Integer)
If keyascii >= 48 And keyascii <= 57 Or keyascii = 8 Then
Else
keyascii = 0
MsgBox " Dear Admin Please Enter Digits Only"
End If
End Sub

Private Sub u_name_keypress(keyascii As Integer)
If keyascii >= 65 And keyascii <= 90 Or keyascii >= 97 And keyascii >= 65 And keyascii <= 122 Or keyascii = 8 Or keyascii = 32 Then
Else
keyascii = 0
MsgBox " Enter only alphabets"
End If
End Sub

Private Sub Selec_Change()
combo1.Clear
On Error GoTo errmsg
Adodc1.Refresh
With Adodc1.Recordset
.MoveFirst
While Not .EOF
combo1.AddItem .Fields("milk")
.MoveNext
Wend
End With
Exit Sub
errmsg:
MsgBox Err.Description
End Sub

