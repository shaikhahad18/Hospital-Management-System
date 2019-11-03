VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Shedules for Doctors"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   19230
   LinkTopic       =   "Form6"
   Moveable        =   0   'False
   ScaleHeight     =   4854.828
   ScaleMode       =   0  'User
   ScaleWidth      =   34270.79
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   8760
      Visible         =   0   'False
      Width           =   5000
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   6840
      Picture         =   "schedules for doctor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   1740
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "I'm Busy"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   6240
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   615
      Left            =   10800
      TabIndex        =   0
      Top             =   8880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc adogrid 
      Height          =   375
      Left            =   240
      Top             =   8640
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\HMS\HMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\HMS\HMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from apt"
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2880
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
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
      Format          =   81395715
      CurrentDate     =   43271
      MinDate         =   43271
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   14880
      Top             =   7800
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "schedules for doctor.frx":076E
      Height          =   2415
      Left            =   2160
      TabIndex        =   11
      Top             =   3720
      Width           =   13950
      _ExtentX        =   24606
      _ExtentY        =   4260
      _Version        =   393216
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "ano"
         Caption         =   "A_no"
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
         DataField       =   "pname"
         Caption         =   "Patients Name"
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
         DataField       =   "timea"
         Caption         =   "Time"
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
         DataField       =   "pd"
         Caption         =   "Problem Description"
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
         DataField       =   "mobile"
         Caption         =   "Mobile"
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
         DataField       =   "bdb"
         Caption         =   "Booked By"
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
         DataField       =   "ddate"
         Caption         =   "Date"
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
            ColumnWidth     =   1203.332
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4705.219
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.371
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4517.293
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3501.887
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   3902.998
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   4357.657
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   2160
      Picture         =   "schedules for doctor.frx":0784
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2655
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
      Height          =   390
      Left            =   16080
      TabIndex        =   10
      Top             =   9630
      Width           =   2040
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
      Height          =   1215
      Left            =   5040
      TabIndex        =   9
      Top             =   840
      Width           =   12735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "When you are busy no any appointment booked for you.."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   6960
      Width           =   6015
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
      Height          =   405
      Left            =   18135
      TabIndex        =   3
      Top             =   9630
      Width           =   2415
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search By Date"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Menu ac 
      Caption         =   "Accounts"
      Begin VB.Menu cps 
         Caption         =   "Change Password"
         Shortcut        =   ^P
      End
      Begin VB.Menu cp 
         Caption         =   "Close"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu g 
         Caption         =   "Logout"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If MsgBox("Are you sure want to change status?", vbYesNo, "Sure?") = vbNo Then
 Else
Dim a As String
a = Text1.Text
If a = "" Then
  MsgBox "Please Enter Name", vbInformation, "Error!"
Else

 Dim adoCon As New ADODB.Connection
 Dim adcmd As New ADODB.Command
 Dim strconnnn As String, stSQL As String

 ' Open a Connection object
  strconnnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"

  adoCon.ConnectionString = strconnnn

  adoCon.Open

  ' Define a query string
stSQL = "DELETE FROM statust WHERE[dname]='" & a & "'"


 ' Set up the Command object
 adcmd.CommandText = stSQL
 adcmd.CommandType = adCmdText

 adcmd.ActiveConnection = adoCon

 adcmd.Execute


 ' Tidy up
 Set adcmd = Nothing
 adoCon.Close
 Set adoCon = Nothing

End If



  Dim AB As String
  
  If Check1.Value = 1 Then
  AB = "NO"
  Else
  AB = "YES"
  End If
  
  
 On Error GoTo ICanDealWithThis
 Dim cn As New ADODB.Connection
 Dim cmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
 cn.ConnectionString = strConn
 cn.Open

 strSQL = "INSERT INTO statust([dname],[status1]) VALUES('" & Text1.Text & "','" & AB & "')"
MsgBox "status changed successfully", vbInformation, "Confirmed!"

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

End Sub

Private Sub Command2_Click()
Me.Hide

End Sub

Private Sub cp_Click()
Form6.Hide

End Sub

Private Sub Command3_Click()
Form11.Show
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub cps_Click()
Form11.Show
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
Text2.Text = Format(DTPicker1.Value, "dd/mm/yyyy")

End Sub

Private Sub DTPicker1_Change()
Text2.Text = Format(DTPicker1.Value, "dd/mm/yyyy")
adogrid.RecordSource = "Select * from apt where ad='" + Label1.Caption + "' and ddate='" + Text2.Text + "'"
adogrid.Refresh
If adogrid.Recordset.EOF Then

Else
adogrid.Caption = adogrid.RecordSource
End If
End Sub


Private Sub Form_Load()
DTPicker1.Value = Now
adogrid.RecordSource = "Select * from apt where ad='" + Label1.Caption + "' and ddate='" + Text2.Text + "'"
adogrid.Refresh
If adogrid.Recordset.EOF Then

Else
adogrid.Caption = adogrid.RecordSource
End If

Text2.Text = Format(DTPicker1.Value, "dd/mm/yyyy")
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
Text1.Text = Label1.Caption



Text2.Text = Format(DTPicker1.Value, "dd/mm/yyyy")

Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM statust where [dname]='" & Text1.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Text3.Text = adoRS.Fields("status1").Value

     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
If Text3.Text = "NO" Then
Check1.Value = 1
End If
If Label1.Caption = "" Then
MsgBox "Invalid Login!", vbCritical + vbExclamation, "Error!"
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

Private Sub log_Click()
Unload Me
End Sub

Private Sub g_Click()
Unload Me
End Sub

Private Sub Image1_Click()
If MsgBox("Do you want Excel Report?", vbYesNo + vbQuestion, "Confirmation!") = vbNo Then
Else
adogrid.RecordSource = "Select * from apt where ad='" + Label1.Caption + "'"
adogrid.Refresh
If adogrid.Recordset.EOF Then

Else
adogrid.Caption = adogrid.RecordSource
End If

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
 
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add
   Set oSheet = oBook.Worksheets(1)
   On Error GoTo errcode
 
   With oBook.Worksheets("sheet1").Rows(1)
        .Font.Bold = True
        For j = 0 To DataGrid1.Columns.Count - 1
            Worksheets("sheet1").Cells(1, j + 1).Value = DataGrid1.Columns(j).Caption
        Next j
   End With
 
   oSheet.Range("A2").CopyFromRecordset adogrid.Recordset
   
 
   oBook.SaveAs
   oBook.Close
   
   oExcel.Quit
   MsgBox "Data Exported Successfully", vbInformation, "Done!"
   Exit Sub
errcode:
   MsgBox Err.Description, , Err.Source
End If
End Sub
