VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Event delete"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14970
   LinkTopic       =   "Form5"
   ScaleHeight     =   8490
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
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
      Left            =   7560
      TabIndex        =   3
      Top             =   4080
      Width           =   2700
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4080
      TabIndex        =   2
      Top             =   4080
      Width           =   2700
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
      Left            =   7080
      TabIndex        =   0
      Text            =   "Select Event Name"
      ToolTipText     =   "Select Event Name To Delete"
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "                                                                   EVENT NAME"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   2400
      Width           =   2415
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim eve As String
    Dim eve1 As String
    Dim txt As String

    eve = "select eno from event where ename = '" & Combo1.Text & "' "
    Set rs1 = conn.Execute(eve)
    txt = rs1(0)
    eve1 = " delete from event where eno = '" & txt & "'"
    Set rs1 = conn.Execute(eve1)
    MsgBox ("Event Deleted")
    Form5.Hide
    Event123.Show
End Sub

Private Sub Command2_Click()
    Form5.Hide
    Event123.Show
End Sub

Private Sub Form_Load()
    Combo1.Text = "Select event Name"
    Dim str3 As String
    str3 = "select ename from event"
    Set rs_scl = conn.Execute(str3)

While rs_scl.EOF = False
    Combo1.AddItem rs_scl!ename
    rs_scl.MoveNext
Wend

End Sub

