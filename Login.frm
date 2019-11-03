VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Login"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1200
      Top             =   7440
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from login"
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
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   10080
      Picture         =   "Login.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7000
      Width           =   2200
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      Picture         =   "Login.frx":0D72
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7000
      Width           =   2200
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      IMEMode         =   3  'DISABLE
      Left            =   7575
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   5400
      Width           =   6000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   7560
      TabIndex        =   3
      Top             =   4200
      Width           =   6000
   End
   Begin VB.Label Label7 
      Height          =   615
      Left            =   1560
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      Height          =   855
      Left            =   600
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Caps Lock is on"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   13800
      Picture         =   "Login.frx":157A
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Forgot Password?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   8640
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   5400
      Width           =   6000
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   4200
      Width           =   6000
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   8640
      Picture         =   "Login.frx":2646
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3000
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please Enter Username", vbOKOnly + vbCritical, "Error!"
Else
If Text2.Text = "" Then
MsgBox "Please Enter Password", vbOKOnly + vbCritical, "Error!"
Text2.SetFocus

Else
Adodc1.RecordSource = "select * from login where Username='" + Text1.Text + "' and Password='" + Text2.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Login failed,Try Again..!!!", vbCritical, "Please Enter correct Username and Password"
Text2.Text = ""
Text2.SetFocus
Else
If Text2.Text = "123456" Then
If Label6.Caption = "D" Then
Form6.Show
Form11.Show
Form1.Hide
ElseIf Label6.Caption = "R" Then
Form3.Show
Form11.Show
Form1.Hide
ElseIf Label6.Caption = "A" Then
Form9.Show
Form11.Show
Form1.Hide
Else
MsgBox "Invalid Login", vbCritical, "Error!"
End
End If
Else
If Label6.Caption = "D" Then
Form6.Show
Form1.Hide
ElseIf Label6.Caption = "R" Then
Form3.Show
Form1.Hide
ElseIf Label6.Caption = "A" Then
Form9.Show
Form1.Hide
Else
MsgBox "Invalid Login", vbCritical, "Error!"
End
End If
End If




Form1.Hide
End If
End If
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus

End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
Form1.BackColor = vbWhite

End Sub

Private Sub Form_Unload(Cancel As Integer)
End

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.PasswordChar = ""
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.PasswordChar = "*"
End Sub

Private Sub Label4_Click()
Form2.Show
Form1.Hide

End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.Value = True

End If
End Sub

Private Sub Text1_LostFocus()
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
     
     Label6.Caption = Rado.Fields("logintxt").Value
     Rado.MoveNext
Loop
Rado.Close
Set Rado = Nothing
adovv.Close
Set adovv = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.Value = True

End If

If Text2.Text = "" Then
Image2.Visible = False

Else
Image2.Visible = True
End If






 If GetKeyState(vbKeyCapital) = 0 Then
        Label5.Visible = False
        
    Else
        Label5.Visible = True
        
    End If

End Sub

Private Sub Timer1_Timer()
Label6.Caption = time


End Sub
