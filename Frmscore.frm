VERSION 5.00
Begin VB.Form Frmscore 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form3"
   Picture         =   "Frmscore.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   978.594
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   5
      Top             =   6480
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Input your name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   6495
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6360
      TabIndex        =   3
      Top             =   4455
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Your score is"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3390
      TabIndex        =   2
      Top             =   4455
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "CONGRATURATIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   3480
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "GAME OVER"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3105
      TabIndex        =   0
      Top             =   1875
      Width           =   6015
   End
End
Attribute VB_Name = "Frmscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Unload Form1
  Form1.Timer2.Enabled = False
  flgquit = True
  Label4.Caption = score
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Playername = UCase(Text1.Text)
   If Playername = "" Then Playername = "NONE"
   Player(11) = Playername
   hiscore(11) = score
   For Mainloopctr = 1 To 10
       For loopctr = 1 To 10
           If hiscore(loopctr) < hiscore(loopctr + 1) Then
              Playertemp = Player(loopctr)
              Player(loopctr) = Player(loopctr + 1)
              Player(loopctr + 1) = Playertemp
              scoretemp = hiscore(loopctr)
              hiscore(loopctr) = hiscore(loopctr + 1)
              hiscore(loopctr + 1) = scoretemp
           End If
       Next loopctr
   Next Mainloopctr
   Open "hiscore.dat" For Output As #1
   For loopctr = 1 To 10
       Write #1, Player(loopctr), hiscore(loopctr)
   Next loopctr
   Close #1
   Frmtop10.Show
End If
End Sub

