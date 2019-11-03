VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form13 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add To Bill"
   ClientHeight    =   10935
   ClientLeft      =   4365
   ClientTop       =   435
   ClientWidth     =   16035
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   16035
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   8400
      Picture         =   "Add_to_bill.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7440
      Width           =   2200
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   5520
      Picture         =   "Add_to_bill.frx":07A1
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7440
      Width           =   2200
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   2400
      Picture         =   "Add_to_bill.frx":0F1D
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
      Width           =   2200
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
      Height          =   375
      Left            =   9720
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   5640
      Width           =   3615
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
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   4440
      Width           =   3615
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
      Height          =   495
      Left            =   11160
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
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
      Left            =   2760
      TabIndex        =   0
      Text            =   "Select"
      Top             =   1560
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   3360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   81985539
      CurrentDate     =   43270
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add to Bill"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   15
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label logchk 
      Height          =   615
      Left            =   12720
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amount"
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
      Left            =   2520
      TabIndex        =   7
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date"
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
      Left            =   2520
      TabIndex        =   4
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   14040
      Picture         =   "Add_to_bill.frx":168B
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Or Enter Patient No"
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
      Left            =   7800
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Patient"
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
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Text1.Text = ""
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM p_register where [pname]='" & Combo1.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text1.Text = adoRS.Fields("p_id").Value

     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
End Sub

Private Sub Command1_Click()
If MsgBox("Are You Sure want to add This Record?", vbYesNo + vbQuestion, "Confirm?") = vbNo Then
Else
Dim cn As New ADODB.Connection
 Dim cmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
 cn.ConnectionString = strConn
 cn.Open

 strSQL = "INSERT INTO add_Bill([P_name],[P_no],[ddate],[description],[Amount],[Added by]) VALUES('" & Combo1.Text & "','" & Text1.Text & "','" & Text4.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & logchk.Caption & "')"


 cmd.CommandText = strSQL
 cmd.CommandType = adCmdText
 cmd.ActiveConnection = cn
 cmd.Execute

MsgBox "Record Added Successfully", vbOKOnly + vbInformation, "Done!"

 
 Set cmd = Nothing
 cn.Close
 Set cn = Nothing
 
 Exit Sub
ICanDealWithThis:
 MsgBox "Something went wrong!,Please try filling all details properly", vbCritical + vbOKOnly, "Error!"
End If
End Sub

Private Sub Command2_Click()
Combo1.Text = "Select"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
Text4.Text = Format(DTPicker1.Value, "dd/mm/yyyy")
End Sub

Private Sub DTPicker1_Change()
Text4.Text = Format(DTPicker1.Value, "dd/mm/yyyy")
End Sub

Private Sub DTPicker1_Click()
Text4.Text = Format(DTPicker1.Value, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
logchk.Caption = Form3.Label21.Caption
If logchk.Caption = "" Then
MsgBox "Invalid Login", vbCritical + vbOKOnly, "Error!"
End
End If

DTPicker1.Value = Now
Text4.Text = Format(DTPicker1.Value, "dd/mm/yyyy")

'combo1 code
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM p_register where st='A'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Combo1.AddItem (adoRS.Fields("pname").Value)

     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
End Sub

Private Sub Image1_Click()
Combo1.Text = ""
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM p_register where [p_id]='" & Text1.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
If adoRS.EOF Then
MsgBox "Invalid Patient Number Please Check", vbOKOnly + vbCritical, "Error!"
Else
Do While Not adoRS.EOF
   Combo1.Text = adoRS.Fields("pname").Value
   
     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo1.Text = ""
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM p_register where [p_id]='" & Text1.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
If adoRS.EOF Then
MsgBox "Invalid Patient Number Please Check", vbOKOnly + vbCritical, "Error!"
Else
Do While Not adoRS.EOF
   Combo1.Text = adoRS.Fields("pname").Value
   
     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
    Case "a" To "z"
        MsgBox "Only Numbers are allowed", vbOKOnly + vbCritical, "Error!"
        KeyAscii = 0
    Case "A" To "Z"
        MsgBox "Only Numbers are allowed", vbOKOnly + vbCritical, "Error!"
        KeyAscii = 0
    End Select
End Sub
