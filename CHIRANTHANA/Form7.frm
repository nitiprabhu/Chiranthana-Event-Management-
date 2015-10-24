VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "8"
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14790
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   ScaleHeight     =   10170
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo4 
      Height          =   390
      Left            =   5880
      TabIndex        =   12
      Top             =   4800
      Width           =   3855
   End
   Begin VB.ComboBox Combo3 
      Height          =   390
      Left            =   5880
      TabIndex        =   11
      Top             =   2400
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      Height          =   390
      Left            =   5880
      TabIndex        =   9
      Top             =   1200
      Width           =   3855
   End
   Begin VB.ComboBox Combo2 
      Height          =   390
      Left            =   5880
      TabIndex        =   4
      Top             =   3720
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
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
      Height          =   800
      Left            =   4680
      TabIndex        =   3
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
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
      Height          =   800
      Left            =   1320
      TabIndex        =   2
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
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
      Height          =   800
      Left            =   8040
      TabIndex        =   1
      Top             =   7320
      Width           =   2415
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
      Height          =   800
      Left            =   11520
      TabIndex        =   0
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "                                                  EVENT NAME"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3000
      TabIndex        =   10
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Label Label6 
      Caption         =   "RESULT ENTRY"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "                                                 STUDENT NAME"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3120
      TabIndex        =   7
      Top             =   2280
      Width           =   1995
   End
   Begin VB.Label Label4 
      Caption         =   "                                                SCHOOL NAME"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3120
      TabIndex        =   6
      Top             =   3480
      Width           =   1995
   End
   Begin VB.Label Label5 
      Caption         =   "                                                PLACE OBTAINED"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3120
      TabIndex        =   5
      Top             =   4680
      Width           =   1995
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_LostFocus()
Dim st As String
Dim st1 As String
Dim st3 As String



st1 = "select distinct(s1.st_name) from student s1,participants p,event e where e.eno=p.eno and s1.st_no=p.st_no and ename = '" & Combo1.Text & "'"
Set rs_res = conn.Execute(st1)

While rs_res.EOF = False
   'Combo1.AddItem rs_sclcat(0)
     Combo3.AddItem rs_res!st_name
     rs_res.MoveNext
 Wend
 
 
End Sub





Private Sub Combo3_LostFocus()
Dim sn As String

sn = "select s.sname from school s,student s1 where s1.sid=s.sid and s1.st_name = '" & UCase(Combo3.Text) & "' "
Set rs_1 = conn.Execute(sn)
While rs_1.EOF = False
   'Combo1.AddItem rs_sclcat(0)
     Combo2.AddItem rs_1!sname
     rs_1.MoveNext
Wend

End Sub

Private Sub Command1_Click()
    Form7.Hide
    Main.Show
End Sub

Private Sub Command2_Click()
    Dim sclid As String
    Dim id1 As Integer
    Dim m5 As String
    Dim id2 As String
    Dim id3 As String
    Dim stid As String
    Dim str3 As String

    stid = "select st_no from student where st_name = '" & UCase(Combo3.Text) & "' "
    Set rs_eve = conn.Execute(stid)

    id3 = rs_eve(0)
    sclid = "select sid from school where sname = '" & UCase(Combo2.Text) & "' "
    Set rs_eve = conn.Execute(sclid)
    s = rs_eve(0)
    m5 = "select eno from event where ename = '" & UCase(Combo1.Text) & "' "
    Set rs_eve = conn.Execute(m5)
    id2 = rs_eve(0)
    str3 = "insert into result values('" & UCase(id3) & "','" & UCase(id2) & "','" & UCase(s) & "','" & Val(Combo4.Text) & "')"
    Set rs_eve = conn.Execute(str3)
    MsgBox "Successfull"
    Combo3.Text = "Select Student"
    Combo1.Text = "Select Event"
    Combo2.Text = "Select School"
    Combo4.Text = " "

    Combo2.Clear
    Combo3.Clear
End Sub

Private Sub Command3_Click()
    Form7.Hide
    Form6.Show
End Sub

Private Sub Command4_Click()
    Combo3.Text = "Select Student"
    Combo1.Text = "Select Event"
    Combo2.Text = "Select School"
    Combo4.Text = ""
End Sub

Private Sub Form_Activate()
    Combo3.Text = "Select Student "
    Combo1.Text = "Select Event"
    Combo2.Text = "Select School"
    Combo4.Text = ""
End Sub

Private Sub Form_Load()
    With Combo4
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
    End With

    Dim str5 As String
    str5 = "select ename from event"
    Set rs1 = conn.Execute(str5)

While rs1.EOF = False
    Combo1.AddItem rs1!ename
    rs1.MoveNext
Wend

End Sub

