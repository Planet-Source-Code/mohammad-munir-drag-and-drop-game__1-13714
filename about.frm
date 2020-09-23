VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0FF&
   Caption         =   "About me"
   ClientHeight    =   1740
   ClientLeft      =   2085
   ClientTop       =   3255
   ClientWidth     =   7650
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   7650
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4920
      Top             =   600
   End
   Begin VB.Label exitnow 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "someone@munir.to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Valuable Comments at :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mohammad Munir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Made by :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub exitnow_Click()
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
exitnow.ForeColor = &HFF0000
End Sub
Private Sub exitnow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
exitnow.ForeColor = &HFF&
End Sub
Private Sub Timer1_Timer()
If Label4.Left >= 5040 Then
Label4.Left = 240

End If
Label4.Left = Label4.Left + 26
End Sub
