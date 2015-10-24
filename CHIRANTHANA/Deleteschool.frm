VERSION 5.00
Begin VB.Form Deleteschool 
   Caption         =   "Delete school"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form3"
   ScaleHeight     =   8640
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7320
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "School name"
      Height          =   855
      Left            =   3240
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "SID"
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "Deleteschool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
SQL = "delete form school where sid='Text1.Text'"
Set rs1 = conn.Execute(SQL)
 MsgBox "School deleted"
End Sub

Private Sub Form_Load()
 Dim str As String
str = "select sname from school"
Set rs1 = conn.Execute(str)

While rs1.EOF = False
   'Combo1.AddItem rs_cat(0)
     Combo1.AddItem rs1!sname
    rs1.MoveNext
 Wend

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Text1_Change()
Dim s1 As String
s1 = "select sid from school where sname = 'combo1.Text'"
Set rs1 = conn.Execute(s1)
Text1.Text = rs1(0)

End Sub
