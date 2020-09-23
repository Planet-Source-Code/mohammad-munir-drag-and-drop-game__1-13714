VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Deag and Drop"
   ClientHeight    =   6675
   ClientLeft      =   1725
   ClientTop       =   855
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8310
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   6120
      Top             =   5040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   6000
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3960
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Drag the City and Drop it into the Correct Country"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   6000
      Width           =   5295
   End
   Begin VB.Label label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Japan 
      BackColor       =   &H000080FF&
      Caption         =   "Japan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label USA 
      BackColor       =   &H0080FF80&
      Caption         =   "USA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Canada 
      BackColor       =   &H00FF00FF&
      Caption         =   "Canada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label SaudiaArabia 
      BackColor       =   &H0000C000&
      Caption         =   "Saudia Arabia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label France 
      BackColor       =   &H0000C0C0&
      Caption         =   "France"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Pakistan 
      BackColor       =   &H000000FF&
      Caption         =   "Pakistan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Paris"
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
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
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Jeddah"
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Index           =   4
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Karachi"
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Houston"
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Tokyo"
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Toronto"
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Mohammad Munir
' Someone@Munir.to


Private oldArray() As Integer, newArray() As Integer, maxN As Integer
                                       
Private Sub ListLabels()
Dim i As Integer, i1 As Integer, i0 As Integer
i0 = newArray(0)
Label1(i0).Move 640, 240
For i = 1 To maxN
i1 = newArray(i)
Label1(i1).Move Label1(i0).Left, Label1(i0).Top + 300 + Label1(i0).Height
i0 = i1
Next i
End Sub



Private Sub Canada_DragDrop(Source As Control, X As Single, Y As Single)
Select Case Source

Case Label1(0)

Label1(0).Visible = False
Canada.BorderStyle = 1
Label4.Caption = Label4.Caption + 1
Winner
End Select
End Sub


Private Sub Command2_Click()
Label1(0).Enabled = True
Label1(1).Enabled = True
Label1(2).Enabled = True
Label1(3).Enabled = True
Label1(4).Enabled = True
Label1(5).Enabled = True

label2.Visible = True
label2.Caption = 15
Label4.Visible = True
Label4.Caption = 0
Timer1.Enabled = True

Command2.Visible = False
End Sub


Private Sub Form_Load()

maxN = 5
ReDim oldArray(maxN) As Integer, newArray(maxN) As Integer
Dim selectN As Integer
For i = 0 To maxN
oldArray(i) = i

Next i
Randomize
For i = maxN To 0 Step -1
selectN = Int((i + 1) * Rnd)
newArray(maxN - i) = oldArray(selectN)
For j = selectN To i - 1
oldArray(j) = oldArray(j + 1)
Next j
Next i
ListLabels
''''''
''''''

Label1(0).DragIcon = LoadPicture(App.Path & "\drop1pg.ico")
Label1(1).DragIcon = LoadPicture(App.Path & "\drop1pg.ico")
Label1(2).DragIcon = LoadPicture(App.Path & "\drop1pg.ico")
Label1(3).DragIcon = LoadPicture(App.Path & "\drop1pg.ico")
Label1(4).DragIcon = LoadPicture(App.Path & "\drop1pg.ico")
Label1(5).DragIcon = LoadPicture(App.Path & "\drop1pg.ico")
Label4.Caption = 0
End Sub







Private Sub France_DragDrop(Source As Control, X As Single, Y As Single)
Select Case Source

Case Label1(5)

Label1(5).Visible = False
France.BorderStyle = 1
Label4.Caption = Label4.Caption + 1
Winner
End Select

End Sub


Private Sub Japan_DragDrop(Source As Control, X As Single, Y As Single)
Select Case Source

Case Label1(1)

Label1(1).Visible = False
Japan.BorderStyle = 1
Label4.Caption = Label4.Caption + 1
Winner
End Select
End Sub

Private Sub Pakistan_DragDrop(Source As Control, X As Single, Y As Single)
Select Case Source

Case Label1(3)

Label1(3).Visible = False
Pakistan.BorderStyle = 1
Label4.Caption = Label4.Caption + 1
Winner
End Select

End Sub


Private Sub SaudiaArabia_DragDrop(Source As Control, X As Single, Y As Single)
Select Case Source

Case Label1(4)

Label1(4).Visible = False
SaudiaArabia.BorderStyle = 1
Label4.Caption = Label4.Caption + 1
Winner
End Select
End Sub

Private Sub Timer1_Timer()
If label2.Caption <= 0 Then
   MsgBox "Sorry, You drop only " & Label4.Caption & " Cities", vbCritical, "L O S T"
   Timer1.Enabled = False
   label2.Visible = False
End
End If

label2.Caption = label2.Caption - 1
End Sub

Private Sub Timer2_Timer()
If Label3.Visible = True Then
   Label3.Visible = False
Else
   Label3.Visible = True
End If
End Sub

Private Sub USA_DragDrop(Source As Control, X As Single, Y As Single)
Select Case Source

Case Label1(2)

Label1(2).Visible = False
USA.BorderStyle = 1
Label4.Caption = Label4.Caption + 1
Winner
End Select
End Sub

Private Sub Winner()
If Label1(0).Visible = False And Label1(1).Visible = False And Label1(2).Visible = False And Label1(3).Visible = False And Label1(4).Visible = False And Label1(5).Visible = False Then
   
label2.Visible = False
Label4.Visible = False
Timer1.Enabled = False
Timer2.Enabled = False
Label3.Visible = False

If label2.Caption <= 4 Then

MsgBox "You are Winner", vbInformation, "Good"

ElseIf label2.Caption <= 8 Then
MsgBox "You are Winner", vbInformation, "Very Good"

ElseIf label2.Caption > 8 Then
MsgBox "You are Winner", vbInformation, "Excellent"

End If
Form1.Enabled = False
Form2.Show
End If
End Sub
