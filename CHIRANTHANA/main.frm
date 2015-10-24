VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Main 
   Caption         =   "MAIN"
   ClientHeight    =   10950
   ClientLeft      =   1875
   ClientTop       =   1605
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Begin VB.CommandButton Command8 
      Caption         =   "REPORTS"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10320
      TabIndex        =   6
      Top             =   5760
      Width           =   2500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "RESULT  ENTRY"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   12600
      TabIndex        =   5
      Top             =   4440
      Width           =   2500
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   12000
      Top             =   8880
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000B&
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   14160
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   3
      Top             =   7320
      Width           =   2145
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000B&
      Caption         =   "NEW UPDATES"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8040
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   2
      Top             =   4440
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000B&
      Caption         =   "NEW EVENT"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   5760
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   1
      Top             =   5880
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "NEW SCHOOL"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   3480
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   0
      Top             =   4440
      Width           =   2500
   End
   Begin VB.Image Image1 
      Height          =   1305
      Left            =   4920
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "CHIRANTANA"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6840
      TabIndex        =   4
      Top             =   1320
      Width           =   7005
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Main.Hide
    School.Show
End Sub
Private Sub Command2_Click()
    Main.Hide
    Event123.Show
End Sub

Private Sub Command3_Click()
    Main.Hide
    Update.Show
End Sub

Private Sub Command4_Click()
    Main.Hide
    Form10.Show
End Sub

Private Sub Command5_Click()
    Main.Hide
    Form7.Show
End Sub

Private Sub Command8_Click()
    Main.Hide
    Form8.Show
End Sub

Private Sub Form_Load()
    Image1.Picture = LoadPicture("E:\logo\logo.jpg")
End Sub

