VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form8"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14925
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   14925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
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
      Height          =   500
      Left            =   6480
      TabIndex        =   11
      Top             =   6960
      Width           =   2000
   End
   Begin VB.CommandButton Command4 
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
      Height          =   500
      Left            =   5400
      TabIndex        =   10
      Top             =   5400
      Width           =   2000
   End
   Begin VB.CommandButton Command3 
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
      Height          =   500
      Left            =   10560
      TabIndex        =   9
      Top             =   4320
      Width           =   2000
   End
   Begin VB.CommandButton Command2 
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
      Height          =   500
      Left            =   10560
      TabIndex        =   8
      Top             =   3120
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
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
      Height          =   500
      Index           =   0
      Left            =   10560
      TabIndex        =   7
      Top             =   1920
      Width           =   2000
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Select  Category To Display Category Wise School List"
      Top             =   4440
      Width           =   4500
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5400
      TabIndex        =   5
      ToolTipText     =   "Select Event Name To Display Event Wise Student List"
      Top             =   3240
      Width           =   4500
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5400
      TabIndex        =   4
      ToolTipText     =   "Select School Name To Display School Wise Student List"
      Top             =   2040
      Width           =   4500
   End
   Begin VB.Label Label5 
      Caption         =   "REPORTS"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   12
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "WINNERS LIST"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2400
      TabIndex        =   3
      Top             =   5520
      Width           =   1800
   End
   Begin VB.Label Label3 
      Caption         =   "                                          CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2400
      TabIndex        =   2
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "                                           EVENT NAME"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2400
      TabIndex        =   1
      Top             =   3120
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "                                           SCHOOL NAME"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2400
      TabIndex        =   0
      Top             =   1920
      Width           =   1800
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    Form8.Hide
    Dim mm3 As String
    Dim mm As String
    mm3 = "select sid from school where sname = '" & Combo1.Text & "' "
    Set rs_rp = conn.Execute(mm3)

    mm = rs_rp(0)
Load DataEnvironment1
    DataEnvironment1.Command4 (mm)
    DataReport4.Sections("section4").Controls("label5").Caption = Combo1.Text
    DataReport4.Show
Unload DataEnvironment1
    Form8.Show
    Combo1.Text = "select school"
End Sub

Private Sub Command2_Click()
    Form8.Hide
    Dim mm4 As String
    Dim mm5 As String
    mm4 = "select eno from event where ename = '" & Combo2.Text & "' "
    Set rs_rp = conn.Execute(mm4)
    mm5 = rs_rp(0)
Load DataEnvironment1
    DataEnvironment1.Command5 (mm5)
    DataReport5.Sections("section4").Controls("label4").Caption = Combo2.Text
    DataReport5.Show
Unload DataEnvironment1
    Form8.Show
    Combo2.Text = "select event"
End Sub

Private Sub Command3_Click()
    Form8.Hide
    Dim mm6 As String
    Dim mm7 As String
    mm6 = "select cid from category where descrip = '" & Combo3.Text & "' "
    Set rs_rp = conn.Execute(mm6)

    mm7 = rs_rp(0)
Load DataEnvironment1
    DataEnvironment1.Command6 (mm7)
    DataReport6.Sections("section4").Controls("label5").Caption = UCase(Combo3.Text)
    DataReport6.Show
Unload DataEnvironment1
    Form8.Show
    Combo3.Text = "select category"
End Sub

Private Sub Command4_Click()
    Form8.Hide
Load DataEnvironment1
    DataEnvironment1.Command1
    DataReport1.Show
Unload DataEnvironment1
    Form8.Show
End Sub

Private Sub Command5_Click()
    Form8.Hide
    Main.Show
End Sub

Private Sub Form_Load()
    Combo1.Text = "select school"
    Combo2.Text = "select event"
    Combo3.Text = "select category"
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
    Combo2.AddItem rs_scl!ename
    rs_scl.MoveNext
Wend
   
    Dim str8 As String
    str8 = "select descrip from category"
    Set rs1 = conn.Execute(str8)

While rs1.EOF = False
    Combo3.AddItem rs1!descrip
    rs1.MoveNext
Wend

End Sub
