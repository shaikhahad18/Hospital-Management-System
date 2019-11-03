VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form14 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Checkout/Discharge"
   ClientHeight    =   10980
   ClientLeft      =   4365
   ClientTop       =   435
   ClientWidth     =   16065
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   16065
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   480
      Top             =   9720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "status chage"
      Height          =   495
      Left            =   12240
      TabIndex        =   29
      Top             =   10080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   615
      Left            =   120
      TabIndex        =   28
      Text            =   "YES"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Payment is received."
      Height          =   285
      Left            =   4200
      TabIndex        =   27
      Top             =   9000
      Width           =   6255
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   600
      Left            =   4080
      Picture         =   "discharge_checkout.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9700
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   12840
      Top             =   9240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "room"
      Height          =   525
      Left            =   9720
      TabIndex        =   24
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Height          =   600
      Left            =   7080
      Picture         =   "discharge_checkout.frx":076E
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9700
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   5160
      TabIndex        =   22
      Top             =   8160
      Width           =   4095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   11280
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
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
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "discharge_checkout.frx":0F0F
      Height          =   2415
      Left            =   10440
      TabIndex        =   19
      Top             =   4560
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4260
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   10560
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Height          =   500
      Left            =   5160
      TabIndex        =   17
      Top             =   6840
      Width           =   4000
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   500
      Left            =   5160
      TabIndex        =   14
      Top             =   6000
      Width           =   4000
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      Height          =   500
      Left            =   5160
      TabIndex        =   12
      Top             =   5160
      Width           =   4000
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      Height          =   500
      Left            =   5160
      TabIndex        =   10
      Top             =   4320
      Width           =   4000
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   500
      Left            =   5160
      TabIndex        =   6
      Top             =   2520
      Width           =   4000
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   2160
      TabIndex        =   2
      Text            =   "Select"
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10920
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   3360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   81657859
      CurrentDate     =   43270
   End
   Begin VB.Label Label13 
      Height          =   495
      Left            =   11640
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Grand Total"
      Height          =   615
      Left            =   1800
      TabIndex        =   21
      Top             =   8160
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "See Details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9480
      TabIndex        =   20
      Top             =   6960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(Medicines,Services etc.)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Others"
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Room Charges"
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Room Charge/Day"
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Days"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Discharge Date"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date of admitted"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Patient"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Or Enter Patient No"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   13800
      Picture         =   "discharge_checkout.frx":0F24
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Discharge"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ac As Integer

Dim dd As Integer

Private Sub Check1_Click()
If Check1.Value = 1 Then

Command3.Enabled = True
Else
Command3.Enabled = False
End If
End Sub

Private Sub Combo1_Change()
Adodc1.RecordSource = "SELECT  sum(Amount) As Total FROM add_Bill where P_name='" + Combo1.Text + "' And P_no='" + Text1.Text + "'"
Adodc1.Refresh

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
Text2.Text = adoRS.Fields("ddate").Value
Text4.Text = adoRS.Fields("price").Value

     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
Text5.Text = Val(Text3.Text) * Val(Text4.Text)
End Sub

Private Sub Combo1_Click()
Adodc1.RecordSource = "SELECT  sum(Amount) As Total FROM add_Bill where P_name='" + Combo1.Text + "' And P_no='" + Text1.Text + "'"
Adodc1.Refresh
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
Text2.Text = adoRS.Fields("ddate").Value
Text4.Text = adoRS.Fields("price").Value

     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
Text5.Text = Val(Text3.Text) * Val(Text4.Text)
End Sub

Private Sub Command1_Click()
Unload Me




End Sub

Private Sub Command2_Click()
Adodc2.RecordSource = "Select * from room where Room_no='" + Text1.Text + "'"
Adodc2.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Unable to find this room details try other", vbOKOnly, "Error!"
End If
Adodc2.Recordset.Fields("status") = "AVL"
Adodc2.Recordset.Update
Command4.Value = True

End Sub

Private Sub Command3_Click()
If MsgBox("Are you sure want to add this Record?", vbYesNo, "Sure?") = vbNo Then
 Else
 
 
 Dim cn As New ADODB.Connection
 Dim cmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
 cn.ConnectionString = strConn
 cn.Open

 strSQL = "INSERT INTO discharge([p_no],[pname],[ddate],[amount],[paid],[c_by]) VALUES('" & Text1.Text & "','" & Combo1.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Label13.Caption & "')"


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

Private Sub Command4_Click()
Adodc3.RecordSource = "Select * from p_register where p_id='" + Text1.Text + "'"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
MsgBox "Unable to find this room details try other", vbOKOnly, "Error!"
End If
Adodc3.Recordset.Fields("st") = "D"
Adodc3.Recordset.Update
MsgBox "Record Added Successfully", vbInformation + vbOKOnly, "Done!"
End Sub

Private Sub DTPicker1_Change()
Text7.Text = Format(DTPicker1.Value, "dd/mm/yyyy")
Text3.Text = CLng(CDate(Text7.Text) - CDate(Text2.Text) + 1)
Text5.Text = Val(Text3.Text) * Val(Text4.Text)
End Sub

Private Sub Form_Load()
Text5.Text = ""
Label13.Caption = Form3.Label21.Caption
If Label13.Caption = "" Then
MsgBox "Invalid login", vbCritical + vbOKOnly, "Error!"
End
Else
End If

DTPicker1.Value = Now
Text7.Text = Format(DTPicker1.Value, "dd/mm/yyyy")
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"


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
ac = Val(Text1.Text)
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM p_register where [p_id]='" & ac & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
If adoRS.EOF Then
Command2.Value = True

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

Private Sub Label11_Click()
Form16.Show
End Sub

Private Sub Text1_Change()
Adodc1.RecordSource = "SELECT  sum(Amount) As Total FROM add_Bill where P_name='" + Combo1.Text + "' And P_no='" + Text1.Text + "'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
Text6.Text = DataGrid1.Text
End Sub

Private Sub Text2_Change()
DTPicker1.Enabled = True

Text3.Text = CLng(CDate(Text7.Text) - CDate(Text2.Text) + 1)
If Text2.Text = "0" Then

Else
Text5.Text = Val(Text3.Text) * Val(Text4.Text)

End If
End Sub

Private Sub Text3_Change()
Text3.Text = CLng(CDate(Text7.Text) - CDate(Text2.Text) + 1)
Label11.Visible = True
If Val(Text3.Text) <= 0 Then
MsgBox "The Date of Discharge must be greater then Date of Admission", vbCritical + vbOKOnly, "Error!"
Else
End If
End Sub

Private Sub Text5_Change()
If Text5.Text = "0" Then
Text5.Text = ""
End If
Text8.Text = Val(Text5.Text) + Val(Text6.Text)
End Sub

Private Sub Text7_Change()

Text5.Text = Val(Text3.Text) * Val(Text4.Text)
End Sub
