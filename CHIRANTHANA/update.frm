VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Update 
   Caption         =   "NEW UPDATES"
   ClientHeight    =   8565
   ClientLeft      =   3405
   ClientTop       =   1890
   ClientWidth     =   14325
   BeginProperty Font 
      Name            =   "Dutch801 XBd BT"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   14325
   Begin VB.CommandButton Command7 
      Caption         =   "SHOW RESULTS"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   16
      Top             =   4320
      Width           =   2175
   End
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
      Left            =   8160
      TabIndex        =   15
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
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
      Left            =   4200
      TabIndex        =   14
      Top             =   7440
      Width           =   2000
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000018&
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
      Left            =   5880
      TabIndex        =   0
      Text            =   "Text3"
      ToolTipText     =   "Enter school id and check"
      Top             =   1440
      Width           =   1500
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.CommandButton Command1 
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
      Left            =   1440
      TabIndex        =   12
      Top             =   7440
      Width           =   2000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
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
      Left            =   10320
      TabIndex        =   11
      Top             =   7440
      Width           =   2000
   End
   Begin VB.CommandButton Command4 
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
      Left            =   7680
      TabIndex        =   10
      Top             =   7440
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
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
      Left            =   5880
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   3360
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   4440
      Width           =   3855
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Text            =   "Combo3"
      Top             =   5520
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
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
      Left            =   5880
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label7 
      Caption         =   "                                    STUDENT NO"
      BeginProperty Font 
         Name            =   "Dutch801 XBd BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   13
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "                                     CLASS"
      BeginProperty Font 
         Name            =   "Dutch801 XBd BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   9
      Top             =   3360
      Width           =   1500
   End
   Begin VB.Label Label3 
      Caption         =   "                                     SCHOOL NAME"
      BeginProperty Font 
         Name            =   "Dutch801 XBd BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   8
      Top             =   4320
      Width           =   1500
   End
   Begin VB.Label Label5 
      Caption         =   "                                       EVENT"
      BeginProperty Font 
         Name            =   "Dutch801 XBd BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   7
      Top             =   5400
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "                                    STUDENT  NAME"
      BeginProperty Font 
         Name            =   "Dutch801 XBd BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label6 
      Caption         =   "ENTER   STUDENT   DETAILS"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
Dim var As String
Dim max_sno As Integer
Dim m1 As String
'Dim m2 As String
Dim m3 As String
Dim m4 As String
Dim m5 As String


m3 = "select sid from school where sname = '" & Combo1.Text & "' "
Set rs1 = conn.Execute(m3)

Combo1.Text = rs1(0)

m1 = "insert into student(st_no,st_name,class,sid) values('" & UCase(Text3.Text) & "','" & UCase(Text1.Text) & "','" & UCase(Text2.Text) & "','" & UCase(Combo1.Text) & "')"
Set rs1 = conn.Execute(m1)


m4 = "select eno from event where ename = '" & Combo3.Text & "' "
Set rs_eve = conn.Execute(m4)
Combo3.Text = rs_eve(0)

m5 = "insert into participants(st_no,eno) values('" & UCase(Text3.Text) & "','" & UCase(Combo3.Text) & "')"
Set rs_eve = conn.Execute(m5)
MsgBox ("SUCCESSFULL")

'm2 = "select count(st_no) from student"
 'Set rs1 = conn.Execute(m2)
'max_sno = Val(rs1(0))
'var = "STD"
'max_sno = max_sno + 1
'Text3.Text = var & max_sno




End Sub



Private Sub Command3_Click()
Update.Hide
Main.Show
End Sub

Private Sub Command4_Click()
    Text1.Text = ""
    Text3.Text = ""
    Text2.Text = ""
    Combo1.Text = "Select School"
    'Combo2.Text = "Select event"
    Combo3.Text = "Select category"
    Text3.SetFocus
End Sub

Private Sub Command5_Click()
Update.Hide

Form4.Show

End Sub

Private Sub Command6_Click()
 Set rs1 = conn.Execute("select * from student where st_no = '" & UCase(Text3.Text) & "'")
    If rs1.EOF = True Then
        MsgBox ("Valid student id, proceed to insert new student")
        'Text1.SetFocus
         Else
        Text3.Text = ""
        Text3.SetFocus
        MsgBox ("Choose different student id")
    End If
End Sub

Private Sub Command7_Click()
Form9.Show
End Sub

Private Sub Form_Load()
    Text3.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Combo1.Text = "Select School"
    Combo3.Text = "Select Event"

    Adodc1.ConnectionString = ""
    Dim str As String
    str = "select sname from school"
    Set rs_scl = conn.Execute(str)

While rs_scl.EOF = False
    Combo1.AddItem rs_scl!sname
    rs_scl.MoveNext
Wend
 
    Dim str3 As String
    str3 = "select ename from event"
    Set rs_scl = conn.Execute(str3)

While rs_scl.EOF = False
    Combo3.AddItem rs_scl!ename
    rs_scl.MoveNext
Wend

End Sub

