VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "RESULT ENTRY"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   15225
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   6960
      TabIndex        =   13
      Top             =   5520
      Width           =   2235
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
      Left            =   12360
      TabIndex        =   12
      Top             =   7080
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
      Left            =   8880
      TabIndex        =   11
      Top             =   7080
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
      Left            =   2160
      TabIndex        =   10
      Top             =   7080
      Width           =   2415
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
      Left            =   5520
      TabIndex        =   9
      Top             =   7080
      Width           =   2415
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
      Left            =   6960
      TabIndex        =   8
      Text            =   "Select School"
      Top             =   4440
      Width           =   3500
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
      Left            =   6960
      TabIndex        =   7
      Text            =   "Select Event"
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   6960
      TabIndex        =   1
      Top             =   1920
      Width           =   2235
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   6960
      TabIndex        =   0
      Top             =   840
      Width           =   2235
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
      Left            =   3960
      TabIndex        =   6
      Top             =   5520
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
      Left            =   3960
      TabIndex        =   5
      Top             =   4320
      Width           =   1995
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
      Left            =   3960
      TabIndex        =   4
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label Label2 
      Caption         =   "                                                  CLASS"
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
      Left            =   3960
      TabIndex        =   3
      Top             =   1920
      Width           =   1995
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
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   1995
   End
End
Attribute VB_Name = "form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo3_Change()

End Sub


