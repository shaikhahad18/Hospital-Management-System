VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form9 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Dashboard For Administrator"
   ClientHeight    =   11010
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   20370
   LinkTopic       =   "Form9"
   Moveable        =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Height          =   735
      Left            =   9000
      Picture         =   "DAshboard_admin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "namea"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.CommandButton Command6 
      DownPicture     =   "DAshboard_admin.frx":076E
      Height          =   700
      Left            =   120
      Picture         =   "DAshboard_admin.frx":1521
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   3800
   End
   Begin VB.CommandButton Command5 
      Height          =   700
      Left            =   120
      Picture         =   "DAshboard_admin.frx":22D4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   3800
   End
   Begin VB.CommandButton Command4 
      Height          =   700
      Left            =   135
      Picture         =   "DAshboard_admin.frx":3114
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4875
      Width           =   3800
   End
   Begin VB.CommandButton Command3 
      Height          =   700
      Left            =   120
      Picture         =   "DAshboard_admin.frx":4237
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   3800
   End
   Begin VB.CommandButton Command2 
      Height          =   700
      Left            =   135
      Picture         =   "DAshboard_admin.frx":5056
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2925
      Width           =   3800
   End
   Begin VB.CommandButton Command1 
      Height          =   700
      Left            =   135
      Picture         =   "DAshboard_admin.frx":623F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1995
      Width           =   3800
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   14520
      Top             =   8040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\HMS\HMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\HMS\HMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "hos"
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
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "You are logged in as"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   15360
      TabIndex        =   11
      Top             =   10330
      Width           =   1995
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit Hospital Name Details."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      DataField       =   "namea"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   7
      Top             =   840
      Width           =   11655
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   4320
      Y1              =   0
      Y2              =   10920
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17375
      TabIndex        =   0
      Top             =   10330
      Width           =   3135
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form6.Show
Form3.Hide
Form4.Hide
Form5.Hide
Form10.Hide
Form12.Hide
Form6.cp.Enabled = True

End Sub

Private Sub Command13_Click()


End Sub

Private Sub Command2_Click()
Form8.Command2.Visible = True
Form3.Label2.Visible = True

Form3.Show
Form6.Hide

Form4.Hide
Form5.Hide
Form10.Hide
Form12.Hide
End Sub

Private Sub Command3_Click()
Form4.Show
Form6.Hide
Form5.Hide
Form10.Hide
Form12.Hide
Form3.Hide
End Sub

Private Sub Command4_Click()
Form5.Show
Form4.Hide
Form16.Hide
Form10.Hide
Form12.Hide
Form3.Hide
End Sub

Private Sub Command5_Click()
Form10.Show
Form4.Hide
Form5.Hide
Form6.Hide
Form12.Hide
Form3.Hide
End Sub

Private Sub Command6_Click()
Form12.Show
Form4.Hide
Form5.Hide
Form6.Hide
Form10.Hide
Form3.Hide
End Sub

Private Sub Command7_Click()

Label6.Visible = False
Text1.Visible = False
Command7.Visible = False
Adodc1.Recordset.Update
MsgBox "Changes Saved Successfully", vbOKOnly + vbInformation, "Changed"
Adodc1.Refresh

End Sub

Private Sub Form_Load()
Dim adovv As New ADODB.Connection
Dim Rado As New ADODB.Recordset
Dim constr As String
Dim a As String
a = Form1.Text1.Text
constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adovv.ConnectionString = constr
adovv.Open
Rado.Source = "SELECT * FROM login where [Username]='" & a & "'"
Rado.CursorType = adOpenForwardOnly
Rado.ActiveConnection = adovv
Rado.Open
Do While Not Rado.EOF
     Label1.Caption = Rado.Fields("lname").Value

     Rado.MoveNext
Loop
Rado.Close
Set Rado = Nothing
adovv.Close
Set adovv = Nothing
If Label1.Caption = "" Then
MsgBox "Invalid login", vbCritical, "Error!"
End
End If

End Sub

Private Sub Label2_Click()
If MsgBox("Are you sure want to Logout?", vbYesNo + vbQuestion, "Are you sure?") = vbNo Then
Else
End
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Are you sure want to exit", vbYesNo + vbQuestion, "Confirm?") = vbYes Then
End
Else
Cancel = 1

End If
End Sub

Private Sub Label4_Click()
Label6.Visible = True
Text1.Visible = True
Command7.Visible = True
End Sub

Private Sub Label5_Click()


End Sub
