VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "Starfield.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   360
   End
   Begin VB.Timer Tmr_upgr 
      Interval        =   1000
      Left            =   1560
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   360
   End
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   0
      ScaleHeight     =   601
      ScaleMode       =   0  'User
      ScaleWidth      =   807
      TabIndex        =   0
      Top             =   0
      Width           =   12105
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   6135
         Left            =   2760
         ScaleHeight     =   6075
         ScaleWidth      =   6435
         TabIndex        =   29
         Top             =   1200
         Visible         =   0   'False
         Width           =   6495
         Begin VB.PictureBox Picture23 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4440
            Picture         =   "Starfield.frx":104A
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   64
            Top             =   5400
            Width           =   360
         End
         Begin VB.PictureBox Picture22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4440
            Picture         =   "Starfield.frx":11D4
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   59
            Top             =   4920
            Width           =   360
         End
         Begin VB.PictureBox Picture21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4440
            Picture         =   "Starfield.frx":135E
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   58
            Top             =   4440
            Width           =   360
         End
         Begin VB.PictureBox Picture20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   120
            Picture         =   "Starfield.frx":14E8
            ScaleHeight     =   390
            ScaleWidth      =   390
            TabIndex        =   43
            Top             =   120
            Width           =   390
         End
         Begin VB.PictureBox Picture19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   0
            Picture         =   "Starfield.frx":1D4A
            ScaleHeight     =   555
            ScaleWidth      =   675
            TabIndex        =   42
            Top             =   720
            Width           =   675
         End
         Begin VB.PictureBox Picture18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            Picture         =   "Starfield.frx":3134
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   41
            Top             =   2880
            Width           =   495
         End
         Begin VB.PictureBox Picture17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   240
            Picture         =   "Starfield.frx":3E5A
            ScaleHeight     =   450
            ScaleWidth      =   345
            TabIndex        =   40
            Top             =   3720
            Width           =   345
         End
         Begin VB.PictureBox Picture16 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            Picture         =   "Starfield.frx":470C
            ScaleHeight     =   495
            ScaleWidth      =   600
            TabIndex        =   39
            Top             =   2040
            Width           =   600
         End
         Begin VB.PictureBox Picture15 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4380
            Picture         =   "Starfield.frx":56C6
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   38
            Top             =   210
            Width           =   360
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4380
            Picture         =   "Starfield.frx":5850
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   37
            Top             =   1650
            Width           =   360
         End
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4380
            Picture         =   "Starfield.frx":59DA
            ScaleHeight     =   330
            ScaleWidth      =   630
            TabIndex        =   36
            Top             =   2610
            Width           =   630
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4380
            Picture         =   "Starfield.frx":651C
            ScaleHeight     =   330
            ScaleWidth      =   630
            TabIndex        =   35
            Top             =   2130
            Width           =   630
         End
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4440
            Picture         =   "Starfield.frx":705E
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   34
            Top             =   3960
            Width           =   360
         End
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4380
            Picture         =   "Starfield.frx":71E8
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   33
            Top             =   690
            Width           =   360
         End
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4380
            Picture         =   "Starfield.frx":7372
            ScaleHeight     =   330
            ScaleWidth      =   360
            TabIndex        =   32
            Top             =   1170
            Width           =   360
         End
         Begin VB.PictureBox Picture13 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   4380
            Picture         =   "Starfield.frx":74FC
            ScaleHeight     =   330
            ScaleWidth      =   510
            TabIndex        =   31
            Top             =   3090
            Width           =   510
         End
         Begin VB.PictureBox Picture14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   240
            Picture         =   "Starfield.frx":7E2E
            ScaleHeight     =   540
            ScaleWidth      =   450
            TabIndex        =   30
            Top             =   4440
            Width           =   450
         End
         Begin VB.Label Label19 
            BackColor       =   &H00000000&
            Caption         =   "TOP SCORE  "
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4920
            TabIndex        =   65
            Top             =   5520
            Width           =   1455
         End
         Begin VB.Label Label18 
            BackColor       =   &H00000000&
            Caption         =   "EVERY 1500 POINTS EXTRA LIVE BONUS"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   360
            TabIndex        =   63
            Top             =   5520
            Width           =   3255
         End
         Begin VB.Label Label16 
            BackColor       =   &H00000000&
            Caption         =   "SOUNDS  ON/OFF"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4920
            TabIndex        =   61
            Top             =   5040
            Width           =   1455
         End
         Begin VB.Label Label15 
            BackColor       =   &H00000000&
            Caption         =   "MUSIC  ON/OFF"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4920
            TabIndex        =   60
            Top             =   4560
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "RESTORE SHIELD"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   840
            TabIndex        =   57
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            Caption         =   "POWERUP GUN"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   56
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackColor       =   &H00000000&
            Caption         =   "POWERUP GUN COUNTDOWN TIMER (Powerup gun has limit in time)"
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   840
            TabIndex        =   55
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label Label4 
            BackColor       =   &H00000000&
            Caption         =   "LIVES"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   54
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label5 
            BackColor       =   &H00000000&
            Caption         =   "SHIELD"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   53
            Top             =   3840
            Width           =   615
         End
         Begin VB.Label Label6 
            BackColor       =   &H00000000&
            Caption         =   "MOVE LEFT"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4860
            TabIndex        =   52
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label Label7 
            BackColor       =   &H00000000&
            Caption         =   "MOVE RIGHT"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4860
            TabIndex        =   51
            Top             =   690
            Width           =   1095
         End
         Begin VB.Label Label8 
            BackColor       =   &H00000000&
            Caption         =   "MOVE UP"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4860
            TabIndex        =   50
            Top             =   1170
            Width           =   1335
         End
         Begin VB.Label Label9 
            BackColor       =   &H00000000&
            Caption         =   "MOVE DOWN"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4860
            TabIndex        =   49
            Top             =   1650
            Width           =   1335
         End
         Begin VB.Label Label10 
            BackColor       =   &H00000000&
            Caption         =   "FIRE"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5100
            TabIndex        =   48
            Top             =   2130
            Width           =   735
         End
         Begin VB.Label Label11 
            BackColor       =   &H00000000&
            Caption         =   "PAUSE GAME"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5100
            TabIndex        =   47
            Top             =   2610
            Width           =   1095
         End
         Begin VB.Label Label12 
            BackColor       =   &H00000000&
            Caption         =   "HELP  ON/OFF "
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4920
            TabIndex        =   46
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label13 
            BackColor       =   &H00000000&
            Caption         =   "END GAME"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4980
            TabIndex        =   45
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label14 
            BackColor       =   &H00000000&
            Caption         =   "SCORE"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   44
            Top             =   4680
            Width           =   615
         End
      End
      Begin VB.Timer Tmrkey 
         Interval        =   1
         Left            =   240
         Top             =   1320
      End
      Begin VB.Timer Tmr_countdown 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3840
         Top             =   240
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3240
         Top             =   240
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         ScaleHeight     =   585
         ScaleWidth      =   12105
         TabIndex        =   2
         Top             =   8400
         Width           =   12135
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   4320
            Picture         =   "Starfield.frx":8B60
            ScaleHeight     =   540
            ScaleWidth      =   450
            TabIndex        =   28
            Top             =   0
            Width           =   480
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   18
            Left            =   10665
            Picture         =   "Starfield.frx":9892
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   24
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   17
            Left            =   10545
            Picture         =   "Starfield.frx":9B44
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   23
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   16
            Left            =   10425
            Picture         =   "Starfield.frx":9DF6
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   22
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   15
            Left            =   10305
            Picture         =   "Starfield.frx":A0A8
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   21
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   14
            Left            =   10185
            Picture         =   "Starfield.frx":A35A
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   20
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   13
            Left            =   10065
            Picture         =   "Starfield.frx":A60C
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   19
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   12
            Left            =   9945
            Picture         =   "Starfield.frx":A8BE
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   18
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   11
            Left            =   9825
            Picture         =   "Starfield.frx":AB70
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   17
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   10
            Left            =   9705
            Picture         =   "Starfield.frx":AE22
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   16
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   9
            Left            =   9585
            Picture         =   "Starfield.frx":B0D4
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   15
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   8
            Left            =   9465
            Picture         =   "Starfield.frx":B386
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   14
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   7
            Left            =   9345
            Picture         =   "Starfield.frx":B638
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   13
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   6
            Left            =   9225
            Picture         =   "Starfield.frx":B8EA
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   12
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   5
            Left            =   9105
            Picture         =   "Starfield.frx":BB9C
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   11
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   4
            Left            =   8985
            Picture         =   "Starfield.frx":BE4E
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   10
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   3
            Left            =   8865
            Picture         =   "Starfield.frx":C100
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   9
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   2
            Left            =   8745
            Picture         =   "Starfield.frx":C3B2
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   8
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   1
            Left            =   8625
            Picture         =   "Starfield.frx":C664
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   7
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox bar 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   390
            Index           =   0
            Left            =   8505
            Picture         =   "Starfield.frx":C916
            ScaleHeight     =   26
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   6
            Top             =   30
            Width           =   120
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   8040
            Picture         =   "Starfield.frx":CBC8
            ScaleHeight     =   450
            ScaleWidth      =   345
            TabIndex        =   5
            Top             =   0
            Width           =   375
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   2160
            Picture         =   "Starfield.frx":D47A
            ScaleHeight     =   495
            ScaleWidth      =   600
            TabIndex        =   4
            Top             =   0
            Width           =   630
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   240
            Picture         =   "Starfield.frx":E434
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   3
            Top             =   0
            Width           =   525
         End
         Begin VB.Label Label17 
            BackColor       =   &H00000000&
            Caption         =   "F1=Help"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   11160
            TabIndex        =   62
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Lblscore 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "0"
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
            Left            =   4920
            TabIndex        =   27
            Top             =   0
            Width           =   3015
         End
         Begin VB.Label Countdown 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   26
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Piclive 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "5"
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
            Left            =   840
            TabIndex        =   25
            Top             =   0
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox PicScreenBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   0
      ScaleHeight     =   601
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   807
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   12105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim result As Long, pressed As Boolean, flgmusic As Boolean
Dim ClockFlashState As Boolean, flgf1 As Boolean
Private Const SRCCOPY = &HCC0020


Private Sub Form_Load()
Dim lngReturnResult As Long
InitDirectX
NUM = 1
BackxPos = 3200

MDirX.Play_Sound "music.wav", True, True, 0
flgmusic = True

Lblscore.Caption = score

BackTile1 = GenerateDC(App.Path & "\space.bmp")
stars = True
Form1.Timer3.Enabled = True
Form1.Timer2.Enabled = True
flgf1 = False
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyLeft Then sLeft = 1
If KeyCode = vbKeyRight Then sRight = 1
If KeyCode = vbKeyUp Then sUp = 1
If KeyCode = vbKeyDown Then sDown = 1
If KeyCode = vbKeyControl Then Firing = 1 'Timer1.Enabled = True
If KeyCode = &H13 Then
   Timer2.Enabled = False
   frmPause.Show
End If
If KeyCode = vbKeyM Then
   If flgmusic = True Then flgmusic = False Else flgmusic = True
End If
If KeyCode = vbKeyS Then
   If flgsound = True Then flgsound = False Else flgsound = True
End If
If KeyCode = vbKeyF12 Then
   Timer2.Enabled = False
   Frmtop10.Show
End If
If KeyCode = vbKeyF1 Then
   If flgf1 = False Then
      Timer2.Enabled = False
      Picture2.Visible = True
   End If
   If flgf1 = True Then
      Picture2.Visible = False
      Timer2.Enabled = True
   End If
End If
If KeyCode = vbKeyEscape Then
   GameActive = 0
   Open "Hiscore.dat" For Input As #1
   For loopctr = 1 To 10
       Input #1, Player(loopctr), hiscore(loopctr)
   Next loopctr
   Close #1
   For loopctr = 1 To 10
       If score > hiscore(loopctr) Then foundsw = 1
   Next loopctr
   If foundsw = 1 Then
      Frmscore.Show
   Else
      Unload Me
      Timer2.Enabled = False
      FrmEnd.Show
   End If
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then sLeft = 0
If KeyCode = vbKeyRight Then sRight = 0
If KeyCode = vbKeyUp Then sUp = 2
If KeyCode = vbKeyDown Then sDown = 2
If KeyCode = vbKeyControl Then
   Firing = 0
   Timer1.Enabled = False
End If
If KeyCode = vbKeyF1 Then
   If flgf1 = False Then flgf1 = True Else flgf1 = False
End If
If KeyCode = vbKeyM Then
   If flgmusic = True Then MDirX.Play_Sound "music.wav", True, True, 0
   If flgmusic = False Then MDirX.Play_Sound "null.wav", False, True, 0
End If
End Sub

Private Sub Timer2_Timer()

OldShipX = ShipX
OldShipY = ShipY

'Do all of the game engine stuff...

MoveAndDrawBack
If stars = True Then DrawStars
VelocityCode
If stage = 1 Then badguy1
If stage = 2 Then badguy2
If stage = 3 Then badguy3
If stage = 4 Then badguy4
If stage = 5 Then badguy5
If stage = 6 Then badguy7
If stage = 7 Then badguy6
If stage = 8 Then mine

FireBullets
GoodGuyStuff

Lblscore.Caption = score

If score > bonus Then
   lives = lives + 1
   Piclive.Caption = lives
   bonus = bonus + 1500
End If
End Sub
Private Sub Timer1_Timer()
Timer1.Interval = Timer1.Interval + 1
Firing = 1
If Timer1.Interval > 3 Then
   Firing = 0
   Timer1.Interval = 1
End If
End Sub

Private Sub Timer3_Timer()
Timer3.Interval = Timer3.Interval + 1
If Timer3.Interval = 10 Then
      stage = 1
      BuildBadGuys
End If
If Timer3.Interval = 20 Then
   If stage = 1 Then Timer3.Interval = Timer3.Interval - 1
   If stage = 0 Then
      stage = 2
      BuildBadGuys
   End If
End If
If Timer3.Interval = 30 Then
   If stage = 2 Then Timer3.Interval = Timer3.Interval - 1
   If stage = 0 Then
      stage = 3
      BuildBadGuys
   End If
End If
If Timer3.Interval = 40 Then
   If stage = 3 Then Timer3.Interval = Timer3.Interval - 1
   If stage = 0 Then
      stage = 4
      BuildBadGuys
   End If
End If
If Timer3.Interval = 50 Then
   If stage = 4 Then Timer3.Interval = Timer3.Interval - 1
   If stage = 0 Then
      stage = 5
      BuildBadGuys
    End If
End If
If Timer3.Interval = 60 Then
   If stage = 5 Then Timer3.Interval = Timer3.Interval - 1
   If stage = 0 Then
      stage = 6
      BuildBadGuys
    End If
End If
If Timer3.Interval = 70 Then
   If stage = 6 Then Timer3.Interval = Timer3.Interval - 1
   If stage = 0 Then
      stage = 7
      BuildBadGuys
    End If
End If
If Timer3.Interval = 80 Then
   If stage = 7 Then Timer3.Interval = Timer3.Interval - 1
   If stage = 0 Then
      stage = 8
      BuildBadGuys
    End If
End If
End Sub

Private Sub Tmr_countdown_Timer()
tempo2 = tempo2 - 1
Countdown.Caption = Str(tempo2)
If tempo2 = 0 Then
   powerup = 0
   Tmr_countdown.Enabled = False
   tempo1 = 0
   Tmr_upgr.Enabled = True
End If
End Sub

Private Sub Tmr_upgr_Timer()
tempo1 = tempo1 + 1
If tempo1 > 120 Then
   flgpower = 1
   tempo1 = 0
   Tmr_upgr.Enabled = False
End If
End Sub

Private Sub Tmrkey_Timer()
If sUp = 1 Then
   Form2.Picture1.Picture = Form2.PicU(ind).Picture
   Form2.Picture2.Picture = Form2.PicUM(ind).Picture
   ShipW = Form2.Picture1.Picture.Width
   ShipH = Form2.Picture1.Picture.Height
   If ind < 3 Then ind = ind + 1
End If
If sDown = 1 Then
   Form2.Picture1.Picture = Form2.PicD(ind).Picture
   Form2.Picture2.Picture = Form2.PicDM(ind).Picture
   ShipW = Form2.Picture1.Picture.Width
   ShipH = Form2.Picture1.Picture.Height
   If ind < 3 Then ind = ind + 1
End If
If sUp = 2 Then
   Form2.Picture1.Picture = Form2.PicU(ind).Picture
   Form2.Picture2.Picture = Form2.PicUM(ind).Picture
   ShipW = Form2.Picture1.Picture.Width
   ShipH = Form2.Picture1.Picture.Height
   If ind > 0 Then ind = ind - 1
   If ind = 0 Then sUp = 0
End If
If sDown = 2 Then
   Form2.Picture1.Picture = Form2.PicD(ind).Picture
   Form2.Picture2.Picture = Form2.PicDM(ind).Picture
   ShipW = Form2.Picture1.Picture.Width
   ShipH = Form2.Picture1.Picture.Height
   If ind > 0 Then ind = ind - 1
   If ind = 0 Then sDown = 0
End If

End Sub

