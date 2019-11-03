VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Dashboard"
   ClientHeight    =   10710
   ClientLeft      =   3765
   ClientTop       =   450
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   Moveable        =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "dashboard.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7200
      Width           =   3300
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "dashboard.frx":09C1
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8040
      Width           =   3300
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "dashboard.frx":1848
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8880
      Width           =   3300
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   17400
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "dashboard.frx":2170
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   3300
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "dashboard.frx":2BEA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   3300
   End
   Begin VB.CommandButton Command8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "dashboard.frx":3824
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   3300
   End
   Begin VB.CommandButton Command7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "dashboard.frx":4776
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   3300
   End
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "dashboard.frx":55D3
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   3300
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Picture         =   "dashboard.frx":646B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   3300
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   15240
      Top             =   8160
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\DAIRY.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\DAIRY.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "firm"
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
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   19320
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      DataField       =   "namea"
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6240
      TabIndex        =   9
      Top             =   240
      Width           =   11055
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   4320
      Y1              =   0
      Y2              =   10920
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   17375
      TabIndex        =   1
      Top             =   10300
      Width           =   3000
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   15360
      TabIndex        =   0
      Top             =   10300
      Width           =   2000
   End
   Begin VB.Image Image1 
      Height          =   10695
      Left            =   0
      Picture         =   "dashboard.frx":73FB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20490
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset



Private Sub ad_Click(Index As Integer)
Form4.Show

End Sub

Private Sub ar_Click(Index As Integer)
Form5.Show

End Sub

Private Sub Command1_Click()
Form7.Show
Form15.Hide
Form17.Hide
Form14.Hide
Form8.Hide
Form13.Hide




End Sub

Private Sub Command2_Click()
Form13.Show
Form15.Hide
Form17.Hide
Form14.Hide
Form8.Hide
Form7.Hide

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Form15.Show
Form7.Hide
Form17.Hide
Form14.Hide
Form8.Hide
Form13.Hide

End Sub

Private Sub Command5_Click()
Form11.Show

End Sub

Private Sub Command6_Click()
Form8.Show
Form15.Hide
Form17.Hide
Form14.Hide
Form7.Hide
Form13.Hide


End Sub

Private Sub Command7_Click()

Form14.Show
Form15.Hide
Form17.Hide
Form7.Hide
Form8.Hide
Form13.Hide

End Sub

Private Sub Command8_Click()
Form17.Show
Form15.Hide
Form7.Hide
Form14.Hide
Form8.Hide
Form13.Hide


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
     Label21.Caption = Rado.Fields("lname").Value
     Label1.Caption = Rado.Fields("logintxt").Value
     Rado.MoveNext
Loop
Rado.Close
Set Rado = Nothing
adovv.Close
Set adovv = Nothing


End Sub



Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Are you sure want to exit", vbYesNo + vbQuestion, "You will be logged out.") = vbYes Then
End
Else
Cancel = 1

End If
End Sub

Private Sub Label2_Click()
Form3.Hide


End Sub

