VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Del School"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14265
   LinkTopic       =   "Form3"
   ScaleHeight     =   6840
   ScaleWidth      =   14265
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6240
      TabIndex        =   3
      Text            =   "Select School Name "
      ToolTipText     =   "Select School Name To Delete"
      Top             =   2400
      Width           =   4215
   End
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
      Left            =   7680
      TabIndex        =   2
      Top             =   4680
      Width           =   1800
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
      Left            =   4560
      TabIndex        =   1
      Top             =   4680
      Width           =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "                                              SCHOOL NAME"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   2160
      Width           =   1800
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim s2 As String
    Dim testmsg As Integer
    Dim scl As String
    Dim scl1 As String
    Dim txt As String

    testmsg = MsgBox("Delete This School Parmanently ?", 1, "Test message")
If testmsg = 1 Then
        scl = "select sid from school where sname = '" & Combo1.Text & "' "
        Set rs1 = conn.Execute(scl)
        txt = rs1(0)
        scl1 = " delete from school where sid = '" & txt & "'"
        Set rs1 = conn.Execute(scl1)
        MsgBox "School deleted"
        School.Show
        Form3.Hide
Else
        MsgBox "School Not Deleted"
        Form3.Hide
        School.Show
End If

End Sub

Private Sub Command2_Click()
    Form3.Hide
    School.Show
End Sub

Private Sub Form_Load()
    Dim str As String
    str = "select sname from school"
    Set rs1 = conn.Execute(str)

While rs1.EOF = False
    Combo1.AddItem rs1!sname
    rs1.MoveNext
Wend

End Sub

