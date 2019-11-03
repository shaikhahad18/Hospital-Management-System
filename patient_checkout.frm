VERSION 5.00
Begin VB.Form Form16 
   Caption         =   " "
   ClientHeight    =   10965
   ClientLeft      =   4380
   ClientTop       =   450
   ClientWidth     =   16065
   LinkTopic       =   "Form16"
   ScaleHeight     =   10965
   ScaleWidth      =   16065
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   10800
      TabIndex        =   35
      Text            =   " "
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Height          =   495
      Left            =   10800
      TabIndex        =   34
      Text            =   " "
      Top             =   7800
      Width           =   3855
   End
   Begin VB.TextBox Text14 
      Height          =   495
      Left            =   10800
      TabIndex        =   33
      Text            =   " "
      Top             =   7080
      Width           =   3855
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   10800
      TabIndex        =   32
      Text            =   " "
      Top             =   6360
      Width           =   3855
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   10800
      TabIndex        =   31
      Text            =   " "
      Top             =   5640
      Width           =   3855
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   10800
      TabIndex        =   30
      Text            =   " "
      Top             =   4920
      Width           =   3855
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   10800
      TabIndex        =   29
      Text            =   " "
      Top             =   3480
      Width           =   3855
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   10800
      TabIndex        =   28
      Text            =   " "
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   10800
      TabIndex        =   27
      Text            =   " "
      Top             =   2040
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5040
      TabIndex        =   17
      Text            =   " "
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   5040
      TabIndex        =   16
      Text            =   " "
      Top             =   6360
      Width           =   3615
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   5040
      TabIndex        =   15
      Text            =   " "
      Top             =   5640
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   5040
      TabIndex        =   14
      Text            =   " "
      Top             =   4920
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   5040
      TabIndex        =   13
      Text            =   " "
      Top             =   4200
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5040
      TabIndex        =   12
      Text            =   " "
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Text            =   " "
      Top             =   2040
      Width           =   3615
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Text            =   " "
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   13200
      Picture         =   "patient_checkout.frx":0000
      Top             =   9240
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   11640
      Picture         =   "patient_checkout.frx":05C7
      Top             =   9240
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   9960
      Picture         =   "patient_checkout.frx":0BCC
      Top             =   9240
      Width           =   1200
   End
   Begin VB.Label Label18 
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   26
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "MAS Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   25
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Stautus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   24
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Building"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   23
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Unit Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   22
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Room Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   21
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Room No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   20
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Date Out"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   19
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Date In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   18
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Disease"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   " PID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Input Patient's Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Patient CheckOut"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
