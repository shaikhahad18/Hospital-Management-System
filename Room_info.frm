VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Information"
   ClientHeight    =   10980
   ClientLeft      =   4365
   ClientTop       =   435
   ClientWidth     =   16065
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   16065
   Begin VB.CommandButton Command3 
      Height          =   650
      Left            =   7200
      Picture         =   "Room_info.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8580
      Width           =   2000
   End
   Begin VB.CommandButton Command2 
      Height          =   650
      Left            =   9960
      Picture         =   "Room_info.frx":07A1
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8580
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Height          =   650
      Left            =   4320
      Picture         =   "Room_info.frx":0FF9
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8580
      Visible         =   0   'False
      Width           =   2000
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
      Height          =   405
      Left            =   8000
      TabIndex        =   12
      Top             =   7320
      Width           =   5000
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
      Height          =   405
      Left            =   8000
      TabIndex        =   11
      Top             =   6480
      Width           =   5000
   End
   Begin VB.TextBox Text5 
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
      Left            =   8000
      TabIndex        =   10
      Top             =   5640
      Width           =   5000
   End
   Begin VB.TextBox Text4 
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
      Left            =   8000
      TabIndex        =   9
      Top             =   4800
      Width           =   5000
   End
   Begin VB.TextBox Text3 
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
      Left            =   8000
      TabIndex        =   8
      Top             =   4080
      Width           =   5000
   End
   Begin VB.TextBox Text2 
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
      Left            =   8000
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   5000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7680
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Room Information"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   17
      Top             =   360
      Width           =   7335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Patient Name"
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
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Type"
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
      Left            =   4680
      TabIndex        =   6
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status"
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
      TabIndex        =   5
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Floor"
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
      Left            =   4680
      TabIndex        =   4
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Price"
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
      Left            =   4680
      TabIndex        =   3
      Top             =   7440
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Building"
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
      Left            =   4680
      TabIndex        =   2
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   9840
      Picture         =   "Room_info.frx":1895
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter room No"
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
      Left            =   5160
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
If MsgBox("Are you sure want to Delete This Record?This action can't be undone", vbYesNo + vbQuestion, "Confirm?") = vbNo Then
Else

Dim adoCon As New ADODB.Connection
 Dim adcmd As New ADODB.Command
 Dim strconnnn As String, stSQL As String

 ' Open a Connection object
  strconnnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"

  adoCon.ConnectionString = strconnnn

  adoCon.Open

  ' Define a query string
stSQL = "DELETE FROM room WHERE[Room_no]='" & Text1.Text & "'"


 ' Set up the Command object
 adcmd.CommandText = stSQL
 adcmd.CommandType = adCmdText

 adcmd.ActiveConnection = adoCon

 adcmd.Execute
MsgBox "All the information of This room is deleted Suceessfully", vbOKOnly + vbInformation, "Done!"

 ' Tidy up
 Set adcmd = Nothing
 adoCon.Close
 Set adoCon = Nothing
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Image1_Click()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""


Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM room where [Room_no]='" & Text1.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
If adoRS.EOF Then
MsgBox "Invalid Room Number Please Check", vbOKOnly + vbCritical, "Error!"
Else
Do While Not adoRS.EOF
   Text3.Text = adoRS.Fields("Building").Value
   Text7.Text = adoRS.Fields("perday").Value
   Text6.Text = adoRS.Fields("TYPE").Value
   Text5.Text = adoRS.Fields("status").Value
   Text4.Text = adoRS.Fields("Floor").Value
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
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""


Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM room where [Room_no]='" & Text1.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
If adoRS.EOF Then
MsgBox "Invalid Room Number Please Check", vbOKOnly + vbCritical, "Error!"
Else
Do While Not adoRS.EOF
   Text3.Text = adoRS.Fields("Building").Value
   Text7.Text = adoRS.Fields("perday").Value
   Text6.Text = adoRS.Fields("TYPE").Value
   Text5.Text = adoRS.Fields("status").Value
   Text4.Text = adoRS.Fields("Floor").Value
     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
End If
End If

End Sub

Private Sub Text5_Change()
If Text5.Text = "BUSY" Then
Label8.Visible = True
Text2.Visible = True
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM p_register where [room_no]='" & Text1.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
   Text2.Text = adoRS.Fields("pname").Value
     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
Else
Label8.Visible = False
Text2.Visible = False
End If
End Sub
