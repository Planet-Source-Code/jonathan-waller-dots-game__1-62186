VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   4380
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FormInstructions.frx":0000
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Please report any bugs to ebdyandjonathan@hotmail.com"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   48
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
