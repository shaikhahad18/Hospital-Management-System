VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Appointment"
   ClientHeight    =   10965
   ClientLeft      =   4365
   ClientTop       =   435
   ClientWidth     =   16065
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10965
   ScaleWidth      =   16065
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Send confirmation sms to patient"
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
      Left            =   4560
      TabIndex        =   32
      Top             =   7680
      Width           =   3615
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   30
      Left            =   14040
      TabIndex        =   31
      Top             =   8160
      Width           =   135
      ExtentX         =   238
      ExtentY         =   53
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1560
      Picture         =   "appointments.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8400
      Width           =   2000
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10200
      Picture         =   "appointments.frx":08A3
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8400
      Width           =   2000
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7200
      Picture         =   "appointments.frx":0F64
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8400
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4320
      Picture         =   "appointments.frx":1747
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8400
      Width           =   2000
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   495
      Left            =   11160
      TabIndex        =   25
      Top             =   7440
      Visible         =   0   'False
      Width           =   2655
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
      Height          =   495
      Left            =   6960
      TabIndex        =   24
      Top             =   4005
      Width           =   2295
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
      Height          =   500
      Left            =   11355
      TabIndex        =   22
      Top             =   1215
      Width           =   4500
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
      Height          =   500
      Left            =   3240
      TabIndex        =   21
      Top             =   1215
      Width           =   4500
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
      Height          =   500
      Left            =   11355
      TabIndex        =   20
      Top             =   375
      Width           =   4500
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   615
      Left            =   13800
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
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
      Height          =   500
      Left            =   3255
      TabIndex        =   17
      Top             =   375
      Width           =   4500
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
      Left            =   360
      TabIndex        =   2
      Text            =   "Select Doctor"
      Top             =   3000
      Width           =   3855
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
      Left            =   5280
      TabIndex        =   1
      Text            =   "Select"
      Top             =   3000
      Width           =   3735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   12960
      TabIndex        =   0
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
      Format          =   16580611
      CurrentDate     =   43270
   End
   Begin MSAdodcLib.Adodc adogrid 
      Height          =   375
      Left            =   17880
      Top             =   9000
      Visible         =   0   'False
      Width           =   16000
      _ExtentX        =   28231
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "appointments.frx":1EB5
      Height          =   2175
      Left            =   0
      TabIndex        =   18
      Top             =   5160
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   22
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
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
         DataField       =   "ad"
         Caption         =   "Doctor Assigned"
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   7680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   7800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(Check doctor's Schedule and availability status before booking appointment)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   1200
      TabIndex        =   29
      Top             =   4680
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Time"
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
      TabIndex        =   23
      Top             =   4005
      Width           =   1935
   End
   Begin VB.Label Label14 
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
      Left            =   13200
      TabIndex        =   16
      Top             =   2205
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Patient's Name"
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
      Left            =   8160
      TabIndex        =   15
      Top             =   375
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Appointment Number"
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
      Left            =   255
      TabIndex        =   14
      Top             =   375
      Width           =   2415
   End
   Begin VB.Label Label3 
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
      Height          =   615
      Left            =   255
      TabIndex        =   13
      Top             =   1215
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Problem Description"
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
      Left            =   8160
      TabIndex        =   12
      Top             =   1215
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Doctor Name"
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
      Left            =   285
      TabIndex        =   11
      Top             =   2205
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Specilization"
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
      Left            =   5400
      TabIndex        =   10
      Top             =   2205
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Available"
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
      Left            =   9960
      TabIndex        =   9
      Top             =   2205
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "YES/NO"
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
      Left            =   9960
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Available Time"
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
      Left            =   285
      TabIndex        =   7
      Top             =   4005
      Width           =   1455
   End
   Begin VB.Label Label10 
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
      Left            =   2040
      TabIndex        =   6
      Top             =   4005
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TO"
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
      Left            =   3000
      TabIndex        =   5
      Top             =   4005
      Width           =   975
   End
   Begin VB.Label Label12 
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
      Left            =   3960
      TabIndex        =   4
      Top             =   4005
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Other AppointMents"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   6000
      Width           =   2055
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connect As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Declare Function InternetGetConnectedState Lib _
    "wininet" (ByRef dwflags As Long, ByVal dwReserved As _
    Long) As Long


Private Sub Combo1_Click()
Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM doctor where [fname]='" & Combo1.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Label10.Caption = adoRS.Fields("timea").Value
Label12.Caption = adoRS.Fields("timeb").Value
Combo2.Text = adoRS.Fields("specilization").Value
     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing



'////////////////////
If Combo2.Text = "ALL" Then

adogrid.RecordSource = "Select * from apt"
adogrid.Refresh
If adogrid.Recordset.EOF Then

Else
adogrid.Caption = adogrid.RecordSource
End If

Else

adogrid.RecordSource = "Select * from apt where ad='" + Combo1.Text + "' and ddate='" + Text2.Text + "'"
adogrid.Refresh
If adogrid.Recordset.EOF Then

Else
adogrid.Caption = adogrid.RecordSource
End If

End If


'///////////////////////////////////////////////////////////////////////////////
Dim adoConn3 As New ADODB.Connection
Dim adoRS3 As New ADODB.Recordset
Dim strConn3 As String
strConn3 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn3.ConnectionString = strConn
adoConn3.Open
adoRS3.Source = "SELECT * FROM statust where [dname]='" & Combo1.Text & "'"
adoRS3.CursorType = adOpenForwardOnly
adoRS3.ActiveConnection = adoConn3
adoRS3.Open
Do While Not adoRS3.EOF
     Label8.Caption = adoRS3.Fields("status1").Value
     adoRS3.MoveNext
Loop
adoRS3.Close
Set adoRS3 = Nothing
adoConn3.Close
Set adoConn3 = Nothing

End Sub

Private Sub Combo2_Click()
Dim adoConn As New ADODB.Connection
Combo1.Clear

Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM doctor where [specilization]='" & Combo2.Text & "'"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
    
Combo1.AddItem (adoRS.Fields("fname").Value)
     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing
 If Combo2.Text = "ALL" Then
 Combo1.Clear
 Combo1.Text = "Select"
 Dim adoConn2 As New ADODB.Connection
Dim adoRS2 As New ADODB.Recordset
Dim strConn2 As String
strConn2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn2.ConnectionString = strConn
adoConn2.Open
adoRS2.Source = "SELECT * FROM doctor"
adoRS2.CursorType = adOpenForwardOnly
adoRS2.ActiveConnection = adoConn2
adoRS2.Open
Do While Not adoRS2.EOF
     Combo1.AddItem (adoRS2.Fields("fname").Value)

     adoRS2.MoveNext
Loop
adoRS2.Close
Set adoRS2 = Nothing
adoConn2.Close
Set adoConn2 = Nothing
End If
If Combo2.Text = "ALL" Then

adogrid.RecordSource = "Select * from apt"
adogrid.Refresh
If adogrid.Recordset.EOF Then

Else
adogrid.Caption = adogrid.RecordSource
End If

Else

adogrid.RecordSource = "Select * from apt where ad='" + Combo1.Text + "'"
adogrid.Refresh
If adogrid.Recordset.EOF Then

Else
adogrid.Caption = adogrid.RecordSource
End If

End If



End Sub

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Combo1.Text = "Select Doctor" Then
   MsgBox "Please Fill All Fields Properly", vbInformation, "HMS"
Else
 If Label8.Caption = "NO" Or Label8.Caption = "No" Then
 MsgBox "Hi, Please check this doctor is unavailable", vbCritical + vbOKOnly, "Error!"
 Else
 If MsgBox("Are you sure want to add this Appointment to follwing Doctor?", vbYesNo, "Sure?") = vbNo Then
 Else
 
 On Error GoTo ICanDealWithThis
 Dim cn As New ADODB.Connection
 Dim cmd As New ADODB.Command
 Dim strConn As String, strSQL As String

 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
 cn.ConnectionString = strConn
 cn.Open

 strSQL = "INSERT INTO apt([ano],[pname],[timea],[ad],[pd],[mobile],[bdb],[ddate]) VALUES('" & Text1.Text & "','" & Text3.Text & "','" & Text6.Text & "','" & Combo1.Text & "','" & Text5.Text & "','" & Text4.Text & "','" & Text7.Text & "','" & Text2.Text & "')"
MsgBox "Appointment Added successfully to the Doctors schedule", vbInformation, "Confirmed!"
Dim a, b, c, d, e, f, g, url As String
Dim aptno As String
Dim mno, nm, time, drn, datea As String
mno = Text4.Text
nm = Text3.Text
time = Text6.Text
drn = Combo1.Text
datea = Text2.Text
aptno = Text1.Text


'http://newsms.designhost.in/index.php/smsapi/httpapi
'/?uname=anjalidemo&password=123456&sender=DZNHST&
'receiver=7972636456&route=TA&msgtype=1&sms=Hi,Mr.KAustubh Kulkarni
a = "http://newsms.designhost.in/index.php/smsapi/httpapi/?uname=anjalidemo&password=123456&sender=DZNHST&receiver="
b = "&route=TA&msgtype=1&sms="
c = "Hi,"
d = " your appointment with "
e = " is booked for "
f = " at "
g = " and appointment no is "
url = a + mno + b + c + nm + d + drn + e + datea + f + time + g + aptno
If Check1.Value = 1 Then

WebBrowser1.Navigate url
If WebBrowser1.Offline Then
MsgBox "Error!"
Else
MsgBox "SENT"
End If
Else
End If

adogrid.Refresh


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
End If
End Sub

Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Command4_Click()
If MsgBox("Are sure want to cancel this appointment?", vbYesNo + vbQuestion, "Confirm?") = vbNo Then
Else
adogrid.Recordset.Delete
MsgBox "Appointment Cancelled successfully", vbInformation + vbOKOnly, "Done!"
End If

End Sub

Private Sub Command5_Click()

End Sub

Private Sub DTPicker1_Change()
Text2.Text = Format(DTPicker1.Value, "dd/mm/yyyy")
If Combo2.Text = "ALL" Then

adogrid.RecordSource = "Select * from apt"
adogrid.Refresh
If adogrid.Recordset.EOF Then

Else
adogrid.Caption = adogrid.RecordSource
End If

Else

adogrid.RecordSource = "Select * from apt where ad='" + Combo1.Text + "' and ddate='" + Text2.Text + "'"
adogrid.Refresh
If adogrid.Recordset.EOF Then

Else
adogrid.Caption = adogrid.RecordSource
End If

End If
End Sub

Private Sub DTPicker1_Click()
Text2.Text = Format(DTPicker1.Value, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
If InternetGetConnectedState(0, 0) = 1 Then
        Shape2.Visible = True
        Shape1.Visible = False
    Else
        Shape1.Visible = True
        Shape2.Visible = False
    End If
DTPicker1.Value = Now
Text2.Text = Format(DTPicker1.Value, "dd/mm/yyyy")
Text7.Text = Form3.Label21.Caption


Combo2.AddItem "ALL"

Dim adoConn As New ADODB.Connection
Dim adoRS As New ADODB.Recordset
Dim strConn As String
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn.ConnectionString = strConn
adoConn.Open
adoRS.Source = "SELECT * FROM doctor"
adoRS.CursorType = adOpenForwardOnly
adoRS.ActiveConnection = adoConn
adoRS.Open
Do While Not adoRS.EOF
     Combo1.AddItem (adoRS.Fields("fname").Value)

     adoRS.MoveNext
Loop
adoRS.Close
Set adoRS = Nothing
adoConn.Close
Set adoConn = Nothing

'for combo 2
Dim adoConn1 As New ADODB.Connection
Dim adoRS1 As New ADODB.Recordset
Dim strconn11 As String
strconn11 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoConn1.ConnectionString = strconn11
adoConn1.Open
adoRS1.Source = "SELECT * FROM doctor"
adoRS1.CursorType = adOpenForwardOnly
adoRS1.ActiveConnection = adoConn1
adoRS1.Open
Do While Not adoRS1.EOF
     Combo2.AddItem (adoRS1.Fields("specilization").Value)

     adoRS1.MoveNext
Loop
adoRS1.Close
Set adoRS1 = Nothing
adoConn1.Close
Set adoConn1 = Nothing








Dim adoconn4 As New ADODB.Connection
Dim adors4 As New ADODB.Recordset
Dim strConn4 As String
strConn4 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\HMS.mdb;Persist Security Info=False"
adoconn4.ConnectionString = strConn4
adoconn4.Open
adors4.Source = "SELECT * FROM apt"
adors4.CursorType = adOpenForwardOnly
adors4.ActiveConnection = adoconn4
adors4.Open
Do While Not adors4.EOF
     Text1.Text = adors4.Fields("ano").Value + 1
     adors4.MoveNext
Loop
adors4.Close
Set adors4 = Nothing
adoconn4.Close
Set adoconn4 = Nothing
End Sub




Private Sub Label8_Change()
If Label8.Caption = "NO" Then
Label8.ForeColor = vbRed
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Select Case Chr(KeyAscii)
    Case "a" To "z"
        MsgBox "Only Numbers are allowed", vbOKOnly + vbCritical, "Error!"
        KeyAscii = 0
    Case "A" To "Z"
        MsgBox "Only Numbers are allowed", vbOKOnly + vbCritical, "Error!"
        KeyAscii = 0
    End Select
End Sub

Private Sub Text4_LostFocus()
If Len(Text4.Text) = 10 Then
Else
MsgBox "Please Enter valid Mobile Number! ", vbCritical + vbOKOnly, "Error!"
Text4.SetFocus

End If
End Sub

Private Sub Text6_GotFocus()
Label16.Visible = True

End Sub

Private Sub Text6_LostFocus()
Text6.Text = UCase(Text6.Text)
End Sub
