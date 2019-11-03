VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form15 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Registration"
   ClientHeight    =   10965
   ClientLeft      =   4365
   ClientTop       =   435
   ClientWidth     =   16065
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10965
   ScaleWidth      =   16065
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   13920
      TabIndex        =   38
      Text            =   "A"
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Height          =   600
      Left            =   11400
      Picture         =   "Patient_reg.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   9000
      Width           =   2000
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7800
      Top             =   10200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
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
      RecordSource    =   ""
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
      Caption         =   "Status"
      Height          =   375
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   9840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   600
      Left            =   8880
      Picture         =   "Patient_reg.frx":07FD
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   9000
      Width           =   2000
   End
   Begin VB.ComboBox Combo5 
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
      Left            =   8880
      TabIndex        =   34
      Text            =   "SELECT"
      Top             =   4080
      Width           =   3495
   End
   Begin VB.TextBox Text11 
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
      Left            =   10440
      TabIndex        =   31
      Text            =   " "
      Top             =   7560
      Width           =   2055
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   10440
      TabIndex        =   30
      Text            =   " SELECT"
      Top             =   6840
      Width           =   2055
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   10440
      TabIndex        =   29
      Text            =   "SELECT"
      Top             =   6120
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   3000
      TabIndex        =   24
      Text            =   "SELECT"
      Top             =   6120
      Width           =   4095
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
      Left            =   3000
      TabIndex        =   23
      Text            =   " "
      Top             =   9480
      Width           =   4095
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
      Left            =   3000
      TabIndex        =   22
      Text            =   " "
      Top             =   8880
      Width           =   4095
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
      Left            =   3000
      TabIndex        =   21
      Text            =   " "
      Top             =   8280
      Width           =   4095
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
      Left            =   3000
      TabIndex        =   20
      Text            =   " "
      Top             =   7560
      Width           =   4095
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
      Height          =   400
      Left            =   3000
      TabIndex        =   19
      Text            =   " "
      Top             =   6840
      Width           =   4095
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
      Height          =   400
      Left            =   3000
      TabIndex        =   18
      Text            =   " "
      Top             =   5400
      Width           =   4095
   End
   Begin VB.TextBox Text3 
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
      Left            =   3000
      TabIndex        =   17
      Text            =   " "
      Top             =   4680
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   8880
      TabIndex        =   7
      Text            =   " SELECT"
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox Text2 
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
      Left            =   3000
      TabIndex        =   5
      Text            =   " "
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text1 
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
      Left            =   3000
      TabIndex        =   4
      Text            =   " "
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dr Treating"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   33
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/day"
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
      Left            =   12600
      TabIndex        =   32
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label19 
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
      Height          =   495
      Left            =   8880
      TabIndex        =   28
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Room No"
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
      Left            =   8880
      TabIndex        =   27
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label16 
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
      Height          =   495
      Left            =   8880
      TabIndex        =   26
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Room Information"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   25
      Top             =   5280
      Width           =   3735
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "StatusofDisease"
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
      Left            =   1200
      TabIndex        =   16
      Top             =   9480
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Disease"
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
      Left            =   1560
      TabIndex        =   15
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label Label12 
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
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mobile"
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
      Left            =   1560
      TabIndex        =   13
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Age"
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
      Left            =   1560
      TabIndex        =   12
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label9 
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
      Left            =   1560
      TabIndex        =   11
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name"
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
      Left            =   1560
      TabIndex        =   10
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PID"
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
      Left            =   1560
      TabIndex        =   9
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Patient's Information "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Room Type"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label4 
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
      Left            =   1560
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reg_no"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Registration ID"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Patient Registration"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim datedii As Integer


Private Sub Combo1_Change()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
a = Combo1.Text
Combo4.Clear

strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM room where status='AVL' and TYPE ='" & a & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Combo4.AddItem (adoRS.Fields("Room_no").Value)
    
     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
End Sub

Private Sub Combo1_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
a = Combo1.Text
Combo4.Clear

strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM room where status='AVL' and TYPE ='" & a & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Combo4.AddItem (adoRS.Fields("Room_no").Value)
    
     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
End Sub

Private Sub Combo3_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
a = Combo1.Text
ac = Combo3.Text

Combo4.Clear

strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM room where status='AVL' and TYPE ='" & a & "' and Building='" & ac & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Combo4.AddItem (adoRS.Fields("Room_no").Value)
    
     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
End Sub

Private Sub Combo4_Change()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM room where [Room_no]='" & Combo4.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text11.Text = adoRS.Fields("perday").Value
     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
End Sub

Private Sub Combo4_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM room where [Room_no]='" & Combo4.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text11.Text = adoRS.Fields("perday").Value
     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
End Sub

Private Sub Command1_Click()
 If MsgBox("Are you sure want to add this Record?", vbYesNo, "Sure?") = vbNo Then
 Else
 
 
 Dim cn As New ADODB.Connection
 Dim cmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
 cn.ConnectionString = strConn
 cn.Open

 strSQL = "INSERT INTO p_register([reg_no],[ddate],[p_id],[pname],[gender],[age],[phone],[address],[disease],[status],[r_type],[dr_name],[building],[room_no],[price],[st]) VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Combo2.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Combo1.Text & "','" & Combo5.Text & "','" & Combo3.Text & "','" & Combo4.Text & "','" & Text11.Text & "','" & Text10.Text & "')"


 cmd.CommandText = strSQL
 cmd.CommandType = adCmdText
 cmd.ActiveConnection = cn
 cmd.Execute
Command2.Value = True


 
 Set cmd = Nothing
 cn.Close
 Set cn = Nothing
 
 Exit Sub
ICanDealWithThis:

 MsgBox "Something went wrong!,Please try filling all details properly", vbCritical + vbOKOnly, "Error!"
End If
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "Select * from room where Room_no='" + Combo4.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Unable to find this room details try other", vbOKOnly, "Error!"
End If
Adodc1.Recordset.Fields("status") = "BUSY"
Adodc1.Recordset.Update
MsgBox "Record Added Successfully", vbInformation + vbOKOnly, "Done!"
End Sub

Private Sub Command3_Click()
If MsgBox("All the unsaved work will be lost", vbYesNo + vbQuestion, "Confirmation") = vbNo Then
Else
Form15.Hide
End If


End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
Adodc1.Visible = False
Combo2.AddItem "Male"
Combo2.AddItem "Female"
Combo1.AddItem "VIP"
Combo1.AddItem "NORMAL"
Combo1.AddItem "MEDIUM"
Combo3.AddItem "A"
Combo3.AddItem "B"
Combo3.AddItem "C"
Text2.Text = Format(Now, "dd/mm/YYYY")
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
a = Combo1.Text
Combo4.Clear

strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM room where status='AVL' and TYPE ='" & a & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Combo4.AddItem (adoRS.Fields("Room_no").Value)
    
     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing

Dim adoConn1 As New ADODB.Connection
Dim adoRS1 As New ADODB.Recordset
Dim strConn1 As String
strConn1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn1.ConnectionString = strConn1
adoConn1.Open
adoRS1.Source = "SELECT * FROM doctor"
adoRS1.CursorType = adOpenForwardOnly
adoRS1.ActiveConnection = adoConn1
adoRS1.Open
Do While Not adoRS1.EOF
     Combo5.AddItem (adoRS1.Fields("fname").Value)

     adoRS1.MoveNext
Loop
adoRS1.Close
Set adoRS1 = Nothing
adoConn1.Close
Set adoConn1 = Nothing

Dim adoConn2 As New ADODB.Connection
Dim adoRS2 As New ADODB.Recordset
Dim strConn2 As String
strConn2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn2.ConnectionString = strConn2
adoConn2.Open
adoRS2.Source = "SELECT * FROM p_register"
adoRS2.CursorType = adOpenForwardOnly
adoRS2.ActiveConnection = adoConn2
adoRS2.Open
Do While Not adoRS2.EOF
     Text1.Text = adoRS2.Fields("reg_no").Value + 1
     Text3.Text = adoRS2.Fields("p_id").Value + 1
     adoRS2.MoveNext
Loop
adoRS2.Close
Set adoRS2 = Nothing
adoConn2.Close
Set adoConn2 = Nothing



End Sub

Private Sub Image1_Click()

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
  
    Case 0 To 9
       MsgBox "Only Characters are allowed", vbOKOnly + vbCritical, "Error!"
        KeyAscii = 0
    End Select
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
If Len(Text6.Text) = 10 Then
Else
MsgBox "Please Enter valid Mobile Number! ", vbCritical + vbOKOnly, "Error!"

End If
End Sub
