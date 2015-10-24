VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Event123 
   BackColor       =   &H8000000B&
   Caption         =   "NEW EVENT"
   ClientHeight    =   8925
   ClientLeft      =   2460
   ClientTop       =   1605
   ClientWidth     =   16050
   LinkTopic       =   "Form3"
   ScaleHeight     =   8925
   ScaleWidth      =   16050
   Begin VB.CommandButton Command6 
      Caption         =   "CHECK"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   12960
      Top             =   1080
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
      Connect         =   "Provider=MSDAORA.1;Password=123;User ID=system;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=123;User ID=system;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "system"
      Password        =   "123"
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H8000000C&
      Caption         =   "DETAILS"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   7080
      TabIndex        =   11
      Top             =   6960
      Width           =   2000
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000C&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   11160
      TabIndex        =   9
      Top             =   5640
      Width           =   2000
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000C&
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   8520
      TabIndex        =   8
      Top             =   5640
      Width           =   2000
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000C&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   5760
      TabIndex        =   7
      Top             =   5640
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   3240
      TabIndex        =   6
      Top             =   5640
      Width           =   2000
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6840
      TabIndex        =   2
      Top             =   4440
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6840
      TabIndex        =   1
      Top             =   3240
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6840
      TabIndex        =   0
      ToolTipText     =   "Enter event id and check"
      Top             =   2040
      Width           =   2000
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      Caption         =   "ENTER   EVENT   DETAILS"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   3240
      TabIndex        =   10
      Top             =   480
      Width           =   7800
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "                                                MAX_PARTICIPENTS"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3840
      TabIndex        =   5
      Top             =   4320
      Width           =   1995
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "                                                     EVENT NAME"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3840
      TabIndex        =   4
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "                                                    EVENT NUMBER"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3840
      TabIndex        =   3
      Top             =   1920
      Width           =   1995
   End
End
Attribute VB_Name = "Event123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 Dim max_eno As Integer
 Dim str1 As String
 Dim var As String
 Dim str2 As String
   
If Text1.Text = " " Or Text2.Text = " " Or Text3.Text = "" Then
    MsgBox "Enter Valid Data"
Else
    sql = "insert into event values('" & UCase(Text1.Text) & "','" & UCase(Text2.Text) & "','" & Val(Text3.Text) & "')"
    Set rs_eve = conn.Execute(sql)
    MsgBox "New Event Inserted"
End If

End Sub

Private Sub Command2_Click()
    Event123.Hide
    Form5.Show
End Sub

Private Sub Command3_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text1.SetFocus
End Sub

Private Sub Command4_Click()
    Main.Show
End Sub

Private Sub Command5_Click()
    Event123.Hide
    Form2.Show
End Sub

Private Sub Command6_Click()
    Set rs1 = conn.Execute("select * from event where eno = '" & UCase(Text1.Text) & "'")
 
 If rs1.EOF = True Then
    MsgBox ("Valid event id, Proceed to insert new event")
 Else
   Text1.Text = ""
   Text1.SetFocus
   MsgBox ("Choose different event id")
 End If

End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub

