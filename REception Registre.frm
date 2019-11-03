VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receptionist Register"
   ClientHeight    =   10995
   ClientLeft      =   4365
   ClientTop       =   435
   ClientWidth     =   16095
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10995
   ScaleWidth      =   16095
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   12120
      TabIndex        =   34
      Text            =   "R"
      Top             =   6840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   4560
      Picture         =   "REception Registre.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7440
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   7080
      Picture         =   "REception Registre.frx":076E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7395
      Width           =   1800
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   9600
      Picture         =   "REception Registre.frx":0F51
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7395
      Width           =   1800
   End
   Begin VB.TextBox Text1 
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
      Height          =   400
      Left            =   3420
      TabIndex        =   14
      Top             =   2715
      Width           =   4200
   End
   Begin VB.TextBox Text2 
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
      Height          =   400
      Left            =   3420
      TabIndex        =   13
      Top             =   4320
      Width           =   4200
   End
   Begin VB.TextBox Text3 
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
      Height          =   400
      Left            =   3420
      TabIndex        =   12
      Top             =   5115
      Width           =   4200
   End
   Begin VB.TextBox Text4 
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
      Height          =   400
      Left            =   3420
      TabIndex        =   11
      Top             =   5925
      Width           =   840
   End
   Begin VB.TextBox Text5 
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
      Height          =   400
      Left            =   5760
      TabIndex        =   10
      Top             =   5925
      Width           =   855
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11040
      TabIndex        =   9
      Top             =   1920
      Width           =   4200
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11040
      TabIndex        =   8
      Top             =   2760
      Width           =   4200
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11040
      TabIndex        =   7
      Top             =   3525
      Width           =   4200
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11040
      TabIndex        =   6
      Top             =   4320
      Width           =   4200
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11040
      TabIndex        =   5
      Top             =   5925
      Width           =   4200
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11040
      TabIndex        =   4
      Top             =   5160
      Width           =   4200
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3420
      TabIndex        =   3
      Top             =   1920
      Width           =   4200
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Text            =   "123456"
      Top             =   6840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
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
      Height          =   405
      Left            =   3420
      TabIndex        =   1
      Text            =   "Select"
      Top             =   3480
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   -360
      Top             =   7320
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
      CommandType     =   8
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
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   33
      Top             =   5520
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add New Receptionist"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   32
      Top             =   600
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "REC ID"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   675
      TabIndex        =   31
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Full Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   675
      TabIndex        =   30
      Top             =   2715
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Qualification"
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
      Left            =   675
      TabIndex        =   29
      Top             =   5115
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   28
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Experience"
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
      Left            =   8760
      TabIndex        =   27
      Top             =   5925
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mobile No"
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
      Left            =   8760
      TabIndex        =   26
      Top             =   2715
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email ID "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   675
      TabIndex        =   25
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   675
      TabIndex        =   24
      Top             =   5925
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Adhar Number"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   23
      Top             =   3525
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username for Login"
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
      Left            =   8760
      TabIndex        =   22
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DOB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   21
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gender"
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
      Left            =   675
      TabIndex        =   20
      Top             =   3525
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "To"
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
      Left            =   4680
      TabIndex        =   19
      Top             =   5925
      Width           =   615
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(ddmmyyy)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function isEmail(email As String) As Boolean
Dim myAt As Integer
Dim myDot As Integer
Dim myDotDot As Integer

isEmail = True
myAt = InStr(1, email, "@", vbTextCompare)
myDot = InStr(myAt + 2, email, ".", vbTextCompare)
myDotDot = InStr(myAt + 2, email, "..", vbTextCompare)
If myAt = 0 Or myDot = 0 Or Not myDotDot = 0 Or Right(email, 1) = "." Then isEmail = False
End Function
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text4.Text = "" Or Label16.Caption = "User ID already exists try another one!" Then
   MsgBox "Please Fill All Fields Properly", vbInformation, "HMS"
Else
 If MsgBox("Are you sure want to add this Record?", vbYesNo, "Sure?") = vbNo Then
 Else
 
 On Error GoTo ICanDealWithThis
 Dim cn As New ADODB.Connection
 Dim cmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
 cn.ConnectionString = strConn
 cn.Open

 strSQL = "INSERT INTO rec([rid],[fname],[gender],[email],[qual],[timea],[timeb],[mobile],[adhar],[addr],[experience],[dob]) VALUES('" & Text12.Text & "','" & Text1.Text & "','" & Combo1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "','" & Text6.Text & "')"
Command4.Value = True

 cmd.CommandText = strSQL
 cmd.CommandType = adCmdText
 cmd.ActiveConnection = cn
 cmd.Execute


 
 Set cmd = Nothing
 cn.Close
 Set cn = Nothing
 
 Exit Sub
ICanDealWithThis:
 MsgBox "Something went wrong!,Please try filling all details properly", vbCritical + vbOKOnly, "Error!"

End If
End If
End Sub

Private Sub Command3_Click()
Me.Hide

End Sub

Private Sub Command4_Click()
On Error GoTo ICanDealWithThis
 Dim cn As New ADODB.Connection
 Dim cmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
 cn.ConnectionString = strConn
 cn.Open

 strSQL = "INSERT INTO login([Username],[lname],[Password],[dob],[logintxt]) VALUES('" & Text11.Text & "','" & Text1.Text & "','" & Text13.Text & "','" & Text6.Text & "','" & Text14.Text & "')"

 cmd.CommandText = strSQL
 cmd.CommandType = adCmdText
 cmd.ActiveConnection = cn
 cmd.Execute

 MsgBox "Record Added Successfully The Default Password is 123456", vbInformation, "Added"
 
 Set cmd = Nothing
 cn.Close
 Set cn = Nothing
 
 Exit Sub
 
ICanDealWithThis:
MsgBox "Something went wrong!,Please try filling all details properly", vbCritical + vbOKOnly, "Error!"
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
Combo1.AddItem "Male"
Combo1.AddItem "Female"
Dim adoConn1 As New ADODB.Connection
Dim adoRS1 As New ADODB.Recordset
Dim strConn1 As String
strConn1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn1.ConnectionString = strConn1
adoConn1.Open
adoRS1.Source = "SELECT * FROM rec"
adoRS1.CursorType = adOpenForwardOnly
adoRS1.ActiveConnection = adoConn1
adoRS1.Open
Do While Not adoRS1.EOF
     Text12.Text = adoRS1.Fields("rid").Value + 1
     adoRS1.MoveNext
Loop
adoRS1.Close
Set adoRS1 = Nothing
adoConn1.Close
Set adoConn1 = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
  
    Case 0 To 9
       MsgBox "Only Characters are allowed", vbOKOnly + vbCritical, "Error!"
        KeyAscii = 0
    End Select
End Sub

Private Sub Text11_LostFocus()
If Text11.Text = "" Then
Else

Adodc1.RecordSource = "Select * from login where Username='" + Text11.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Label16.Caption = "User ID is available you can use it "
Label16.ForeColor = &H8000&
Else
Label16.Caption = "User ID already exists try another one!"


Label16.ForeColor = &HFF&


End If
End If
End Sub

Private Sub Text2_LostFocus()
Dim ans As Boolean
If isEmail(Text2.Text) = True Then
Else
MsgBox "Invalid Email id", vbOKOnly + vbCritical, "Error!"
Text2.SetFocus

End If
End Sub

Private Sub Text3_LostFocus()
Text3.Text = UCase(Text3.Text)

End Sub

Private Sub Text4_LostFocus()
Text4.Text = UCase(Text4.Text)
End Sub

Private Sub Text5_LostFocus()
Text5.Text = UCase(Text5.Text)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
    Case "a" To "z"
        MsgBox "Only Numbers are allowed", vbOKOnly + vbCritical, "Error!"
        KeyAscii = 0
    Case "A" To "Z"
        MsgBox "Only Numbers are allowed", vbOKOnly + vbCritical, "Error!"
        KeyAscii = 0
    End Select
End Sub

Private Sub Text6_LostFocus()
If Len(Text6.Text) = 8 Then
Else
MsgBox "Please Enter valid Date of Birth! ", vbCritical + vbOKOnly, "Error!"


End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
    Case "a" To "z"
        MsgBox "Only Numbers are allowed", vbOKOnly + vbCritical, "Error!"
        KeyAscii = 0
    Case "A" To "Z"
        MsgBox "Only Numbers are allowed", vbOKOnly + vbCritical, "Error!"
        KeyAscii = 0
    End Select
End Sub

Private Sub Text7_LostFocus()
If Len(Text7.Text) = 10 Then
Else
MsgBox "Please Enter valid Mobile Number! ", vbCritical + vbOKOnly, "Error!"


End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
    Case "a" To "z"
        MsgBox "Only Numbers are allowed", vbOKOnly + vbCritical, "Error!"
        KeyAscii = 0
    Case "A" To "Z"
        MsgBox "Only Numbers are allowed", vbOKOnly + vbCritical, "Error!"
        KeyAscii = 0
    End Select
End Sub

Private Sub Text8_LostFocus()
If Len(Text8.Text) = 12 Then
Else
MsgBox "Please Enter valid aadhar Number! ", vbCritical + vbOKOnly, "Error!"


End If
End Sub
