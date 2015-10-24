VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20190
   FillColor       =   &H80000014&
   ForeColor       =   &H80000013&
   LinkTopic       =   "Form10"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Form10.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      TabIndex        =   6
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "NEW USER"
      DownPicture     =   "Form10.frx":1D012
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12960
      TabIndex        =   5
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&LOGIN"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   4
      Top             =   6600
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Dutch801 XBd BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   8640
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4920
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Dutch801 XBd BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   2
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "     PASSWORD"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   5040
      Width           =   2355
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "    USER NAME"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3240
      Width           =   2355
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim str3 As String
    Dim str2 As String
    Dim str4 As String
    Dim str1 As String

If Text1.Text = " " Or Text2.Text = " " Then
        MsgBox "Enter valid user name and password"
        Form10.Show
ElseIf Text1.Text = "jnnce" Then
        str4 = "select password from login where uname='jnnce'"
        Set rs1 = conn.Execute(str4)
        str1 = rs1(0)
    If str1 = Text2.Text Then
        Form10.Hide
        Main.Show
    Else
        MsgBox "INVALID PASSWORD"
        Text2.Text = ""
    End If
Else
        str3 = "select uname from login"
        Set rs1 = conn.Execute(str3)

    While rs1.EOF = False
        If Text1.Text = rs1(0) Then
            str2 = "1"
        End If
            rs1.MoveNext
    Wend

    If str2 = "1" Then
        str4 = "select password from login where uname='" & Text1.Text & "'"
        Set rs1 = conn.Execute(str4)
        str1 = rs1(0)
            If str1 = Text2.Text Then
                Main.Show
                Form10.Hide
                Event123.Command2.Visible = False
                School.Command2.Visible = False
                
            Else
                MsgBox "Invalid Password"
                Text2.Text = ""
            End If
    Else
        MsgBox "User not found"
    End If
End If

End Sub

Private Sub Command2_Click()
    Form10.Hide
    Form11.Show
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Activate()
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
End Sub

Private Sub Form_Load()

    conn.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=system;Password = 123"
    sql = "select * from school"
   
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic
    
End Sub

