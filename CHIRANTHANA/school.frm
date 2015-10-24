VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form School 
   Caption         =   "NEW SCHOOL "
   ClientHeight    =   10950
   ClientLeft      =   2265
   ClientTop       =   1410
   ClientWidth     =   20250
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Begin VB.CommandButton Command5 
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
      Left            =   11760
      TabIndex        =   14
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      ToolTipText     =   "Enter school ID and check"
      Top             =   2640
      Width           =   4125
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000018&
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
      Left            =   7440
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   6000
      Width           =   4050
   End
   Begin VB.CommandButton Details 
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
      Left            =   7560
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7440
      Width           =   2000
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   2880
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "DSN=school"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "school"
      OtherAttributes =   ""
      UserName        =   "system"
      Password        =   "123"
      RecordSource    =   "SCHOOL"
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
      Left            =   10320
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
      Width           =   2000
   End
   Begin VB.CommandButton Command3 
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
      Left            =   13080
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7440
      Width           =   2000
   End
   Begin VB.CommandButton Command2 
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
      Left            =   4800
      TabIndex        =   8
      Top             =   7440
      Width           =   2000
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
      Left            =   1920
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   2000
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000018&
      DataField       =   "SLOCATION"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7440
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   4920
      Width           =   4050
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      DataField       =   "SNAME"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7440
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   3840
      Width           =   4050
   End
   Begin VB.Label Label5 
      Caption         =   "ENTER    SCHOOL   DETAILS"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   11
      Top             =   840
      Width           =   8175
   End
   Begin VB.Label Label4 
      Caption         =   "                                      CATEGORY"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      Top             =   5880
      Width           =   1500
   End
   Begin VB.Label Label3 
      Caption         =   "                                       LOCATION"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      Top             =   4800
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "                                     SCHOOL NAME"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   4
      Top             =   3720
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "                                      SCHOOL ID"
      BeginProperty Font 
         Name            =   "Dutch801 Rm BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   3
      Top             =   2520
      Width           =   1500
   End
End
Attribute VB_Name = "School"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim catid As String
    Dim id As Integer
If Text4.Text = " " Or Text3.Text = " " Or Text2.Text = " " Or Combo1.Text = " " Then
    MsgBox "ENTER  VALID  DATA"
Else
    catid = "select cid from category where descrip = '" & UCase(Combo1.Text) & "' "
    Set rs1 = conn.Execute(catid)
    id1 = rs1(0)
    sql = "insert into school values('" & UCase(Text4.Text) & "','" & UCase(Text2.Text) & "','" & UCase(Text3.Text) & "'," & id1 & ")"
    Set rs1 = conn.Execute(sql)
    MsgBox "New School Inserted"
    Text2.Text = ""
    Text3.Text = ""
    Combo1.Text = ""
    Text4.Text = ""
    Text4.SetFocus
End If

End Sub

Private Sub Command2_Click()
    School.Hide
    Form3.Show
End Sub

Private Sub Command3_Click()
    School.Hide
    Main.Show
End Sub

Private Sub Command4_Click()
    Text4.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Combo1.Text = ""
    Text4.SetFocus
    
End Sub

Private Sub Command5_Click()
    Set rs1 = conn.Execute("select * from school where sid = '" & UCase(Text4.Text) & "'")

If rs1.EOF = True Then
    MsgBox ("Valid school id, proceed to insert new school")
Else
    Text4.Text = ""
    Text4.SetFocus
    MsgBox ("Choose different school id")
End If

End Sub

Private Sub Details_Click()
    School.Hide
    Form1.Show
End Sub

Private Sub Form_Load()
    Text4.Text = " "
    Text2.Text = " "
    Text3.Text = " "
    Combo1.Text = " "
    Dim str As String
    str = "select descrip from category"
    Set rs1 = conn.Execute(str)

While rs1.EOF = False
    Combo1.AddItem rs1!descrip
    rs1.MoveNext
Wend
 
End Sub

