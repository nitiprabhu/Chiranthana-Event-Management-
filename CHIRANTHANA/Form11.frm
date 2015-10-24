VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18390
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   18390
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   840
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=login"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "login"
      OtherAttributes =   ""
      UserName        =   "system"
      Password        =   "123"
      RecordSource    =   "login"
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
      Caption         =   "CREATE"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7080
      TabIndex        =   1
      ToolTipText     =   "Enter Full Name"
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   7080
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Enter Password"
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label Label1 
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
      Height          =   435
      Left            =   3480
      TabIndex        =   3
      Top             =   2760
      Width           =   2355
   End
   Begin VB.Label Label2 
      Caption         =   "    PASSWORD"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3480
      TabIndex        =   2
      Top             =   4320
      Width           =   2355
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim u1 As String
    Dim str1 As String
    Dim strt2 As Integer

If Text1.Text = "" Or Text2.Text = "" Then
        MsgBox "Enter valid user name and password"
        Form11.Show
Else
        str1 = "select uname from login"
        Set rs1 = conn.Execute(str1)
    While rs1.EOF = False
        If Text1.Text = rs1(0) Then
            strt2 = 1
        End If
            rs1.MoveNext
    Wend

    If strt2 = 1 Then
        MsgBox "User already exists"
        Form11.Show
    Else
        u1 = "insert into login values('" & Text1.Text & "','" & Text2.Text & "')"
        Set rs_2 = conn.Execute(u1)
        MsgBox ("New User Created ! Login With Your User Name And Password")
        Form11.Hide
        Form10.Show
    End If
End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
End Sub

