VERSION 5.00
Begin VB.Form Main 
   Caption         =   "MAIN"
   ClientHeight    =   8940
   ClientLeft      =   1875
   ClientTop       =   1605
   ClientWidth     =   16515
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   16515
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      Height          =   800
      Left            =   11760
      TabIndex        =   3
      Top             =   3840
      Width           =   2500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NEW UPDATES"
      Height          =   800
      Left            =   8400
      TabIndex        =   2
      Top             =   3840
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEW EVENT"
      Height          =   800
      Left            =   5160
      TabIndex        =   1
      Top             =   3840
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEW SCHOOL"
      Height          =   800
      Left            =   1920
      TabIndex        =   0
      Top             =   3960
      Width           =   2500
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   855
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   11415
   End
   Begin VB.Label Label1 
      Caption         =   "CHIRANTANA"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   7000
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
Form4.Show
End Sub

Private Sub Command4_Click()
End
End Sub

