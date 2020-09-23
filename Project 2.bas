Attribute VB_Name = "Game_Logic"
'ALL VARIABLES:
Public iWidth As Integer
Public iHeight As Integer
Public Bad1 As BadGuy
Public ShipX As Single, ShipY As Single
Public ShipW As Integer, ShipH As Integer
Public OldShipX As Single, OldShipY As Single
Public KeyState As Byte
Public sLeft As Byte
Public sRight As Byte
Public sUp As Byte
Public sDown As Byte
Public sVelocity As Single
Public sVel As Single
Public Firing As Byte
Public BulletsActivated As Byte
Public CurLevel As Level
Public Exploding As Byte
Public ExplodingFrame As Byte
Public Explosions(0 To 9) As PointXY
Public BadGuys() As BadGuy
Public BadGuyNum As Byte
Public Bulletl1(0 To NumOfBullets) As bullet
Public Bulletr1(0 To NumOfBullets) As bullet
Public Bulletl2(0 To NumOfBullets) As bullet
Public Bulletr2(0 To NumOfBullets) As bullet
Public StarArray(0 To NumOfStars) As Star
Public fboom As Integer
Public flgr As Integer
Public lives As Integer
Public stage As Integer
Public mission As Integer
Public TempCalc As Byte
Public Tempclac As Byte
Public score As Long
Public numsec As Integer
Public bullets As Boolean
Public oldX As Single
Public oldY As Single
Public frame As Integer
Public thrust As Integer
Public ind As Integer
Public indice As Integer
Public mode As Integer
Public posi As Integer
Public stars As Boolean
Public planets As planet
Public upgr1 As upgr, upgr2 As upgr
Public flgplanets As Byte, flgupgr As Byte, flgpower As Byte
Public numplanets As Byte, powerup As Byte, power As Byte
Public tempo1 As Byte, tempo2 As Byte
Public seq As Byte
Public appo As Boolean, flgf1 As Boolean, flgquit As Boolean
Public bonus As Long
Public angle  As Single
Public curX As Single, curY As Single
Public deltax As Double, deltay As Double
Public bulletframe As Integer
Public i As Integer, rec As Integer
Public Colo As RGB, INIFound As Boolean
Public Const NumLines = 24
Public Playername As String
Public Playertemp As String
Public scoretemp As Long
Public Mainloopctr As Byte
Public loopctr As Byte
Public foundsw As Byte
Public Player(11) As String
Public hiscore(11) As Long
Public flgsound As Boolean

Public Sub InitializeGameEngine()

    Form1.WindowState = 2
    Randomize
    
    For X = 0 To NumOfStars
        BuildNewStar1 (X)
    Next X
    
    TempCalc = 0
    Tempclac = 0

    score = 0
    lives = 5
    bonus = 1500
    bulletframe = 0
    fboom = 1
    OldShipX = 0
    OldShipY = 0
    ShipX = 10
    ShipY = (Form1.PicMain.ScaleHeight / 2) - 50
    sVel = 1
    sVelocity = 50
    BulletsActivated = 0
    Health = 19
    BuildBadGuys
    DrawHealthBar
    Form2.Picture1.Picture = Form2.PicD(0).Picture
    Form2.Picture2.Picture = Form2.PicDM(0).Picture
    ShipW = Form2.Picture1.Picture.Width
    ShipH = Form2.Picture1.Picture.Height
    numplanets = 0
    flgplanets = 0
    flgupgr = 0
    powerup = 0
    power = 0
    flgpower = 0
    upgr1.frame = 0
    upgr2.frame = 0
    thrust = 0
    tempo1 = 0
    tempo2 = 60
    flgquit = False
    flgsound = True
    foundsw = 0
End Sub

Public Sub DrawStars()
'Draw the stars to their buffer
For X = 0 To NumOfStars
    StarArray(X).X = StarArray(X).X - StarArray(X).SPEED
    If StarArray(X).X < 0 Then BuildNewStar (X)
    SetPixelV Form1.PicScreenBuffer.hdc, StarArray(X).X, StarArray(X).Y, RGB(StarArray(X).bright, StarArray(X).bright, StarArray(X).bright)
Next X
End Sub

Public Sub FireBullets()
Dim i As Integer, result As Integer

If Firing = 1 Then
If flgsound = True Then MDirX.Play_Sound "fire.wav", False, True, 1
'    BitBlt Form1.PicScreenBuffer.hdc, ShipX + 14, ShipY + 4, 7, 12, Form2.Picshotm.hdc, 9 * bulletframe, 0, vbMergePaint
'    BitBlt Form1.PicScreenBuffer.hdc, ShipX + 14, ShipY + 4, 7, 12, Form2.Picshot.hdc, 9 * bulletframe, 0, vbSrcAnd
'    BitBlt Form1.PicScreenBuffer.hdc, ShipX + 31, ShipY + 4, 4, 12, Form2.Picshotm.hdc, 9 * bulletframe, 0, vbMergePaint
'    BitBlt Form1.PicScreenBuffer.hdc, ShipX + 31, ShipY + 4, 4, 12, Form2.Picshot.hdc, 9 * bulletframe, 0, vbSrcAnd
    bulletframe = bulletframe + 1
    If bulletframe > 2 Then bulletframe = 0
End If
'Bullets 1
For X = 0 To NumOfBullets
If Firing = 1 And BulletsActivated <= 1 And Bulletl1(X).activated = 0 And Bulletr1(X).activated = 0 Then
   Bulletl1(X).activated = 1
   Bulletr1(X).activated = 1
   If ind = 0 Then
      Bulletl1(X).X = ShipX + 15
      Bulletl1(X).Y = ShipY + 15
      Bulletr1(X).X = ShipX + 15
      Bulletr1(X).Y = ShipY + 26
   End If
   If ind = 1 Then
      Bulletl1(X).X = ShipX + 15
      Bulletl1(X).Y = ShipY + 14
      Bulletr1(X).X = ShipX + 15
      Bulletr1(X).Y = ShipY + 36
   End If
   If ind = 2 Then
      Bulletl1(X).X = ShipX + 15
      Bulletl1(X).Y = ShipY + 21
      Bulletr1(X).X = ShipX + 15
      Bulletr1(X).Y = ShipY + 53
   End If
   If ind = 3 Then
      Bulletl1(X).X = ShipX + 15
      Bulletl1(X).Y = ShipY + 24
      Bulletr1(X).X = ShipX + 15
      Bulletr1(X).Y = ShipY + 61
   End If
   
   BulletsActivated = BulletsActivated + 1
End If

If Bulletl1(X).activated = 1 Then
    Bulletl1(X).X = Bulletl1(X).X + BulletSpeed
    If Bulletl1(X).X > Form1.PicMain.ScaleWidth Then Bulletl1(X).activated = 0
    BitBlt Form1.PicScreenBuffer.hdc, Bulletl1(X).X, Bulletl1(X).Y, 16, 8, Form2.PicBulletM.hdc, 0, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hdc, Bulletl1(X).X, Bulletl1(X).Y, 16, 8, Form2.PicBullet.hdc, 0, 0, vbSrcAnd

    For Y = 0 To CurLevel.NumOfBadGuys
      If BadGuys(Y).activated = 1 And BadGuys(Y).Exploding = 0 And Bulletl1(X).activated = 1 Then
         If Abs((BadGuys(Y).X + (BadGuys(Y).xsize / 2)) - (Bulletl1(X).X + 0.5)) < (BadGuys(Y).xsize / 2) And Abs((BadGuys(Y).Y + (BadGuys(Y).ysize / 2)) - (Bulletl1(X).Y + 2)) < (BadGuys(Y).ysize / 2) Then
            BadGuys(Y).Damage = BadGuys(Y).Damage + 1
            Bulletl1(X).activated = 0
            BitBlt Form1.PicScreenBuffer.hdc, Bulletl1(X).X, Bulletl1(X).Y - 12, 12, 12, Form2.PicHITM.hdc, 12 * BadGuys(Y).frame, 0, vbPatInvert
            BitBlt Form1.PicScreenBuffer.hdc, Bulletl1(X).X, Bulletl1(X).Y - 12, 12, 12, Form2.PicHIT.hdc, 12 * BadGuys(Y).frame, 0, vbSrcPaint
            BadGuys(Y).frame = BadGuys(Y).frame + 1
            If BadGuys(Y).frame > 6 Then BadGuys(Y).frame = 0
         End If
      End If
    Next Y
    
End If
If Bulletr1(X).activated = 1 Then
    Bulletr1(X).X = Bulletr1(X).X + BulletSpeed
    If Bulletr1(X).X > Form1.PicMain.ScaleWidth Then Bulletr1(X).activated = 0
    BitBlt Form1.PicScreenBuffer.hdc, Bulletr1(X).X, Bulletr1(X).Y, 16, 8, Form2.PicBulletM.hdc, 0, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hdc, Bulletr1(X).X, Bulletr1(X).Y, 16, 8, Form2.PicBullet.hdc, 0, 0, vbSrcAnd

    For Y = 0 To CurLevel.NumOfBadGuys
       If BadGuys(Y).activated = 1 And BadGuys(Y).Exploding = 0 And Bulletr1(X).activated = 1 Then
         If Abs((BadGuys(Y).X + (BadGuys(Y).xsize / 2)) - (Bulletr1(X).X + 0.5)) < (BadGuys(Y).xsize / 2) And Abs((BadGuys(Y).Y + (BadGuys(Y).ysize / 2)) - (Bulletr1(X).Y + 2)) < (BadGuys(Y).ysize / 2) Then
            BadGuys(Y).Damage = BadGuys(Y).Damage + 1
            Bulletr1(X).activated = 0
            BitBlt Form1.PicScreenBuffer.hdc, Bulletr1(X).X, Bulletr1(X).Y - 12, 12, 12, Form2.PicHITM.hdc, 12 * BadGuys(Y).frame, 0, vbPatInvert
            BitBlt Form1.PicScreenBuffer.hdc, Bulletr1(X).X, Bulletr1(X).Y - 12, 12, 12, Form2.PicHIT.hdc, 12 * BadGuys(Y).frame, 0, vbSrcPaint
            BadGuys(Y).frame = BadGuys(Y).frame + 1
            If BadGuys(Y).frame > 6 Then BadGuys(Y).frame = 0
         End If
       End If
    Next Y
    
End If

11 Next X

'Don't allow any more bullets to be created
BulletsActivated = 0


If powerup = 1 Then
'Powerup Bullets
For X = 0 To NumOfBullets
If Firing = 1 And BulletsActivated <= 1 And Bulletl2(X).activated = 0 And Bulletr2(X).activated = 0 Then
   Bulletl2(X).activated = 1
   Bulletr2(X).activated = 1
   If ind = 0 Then
      Bulletl2(X).X = ShipX + 15
      Bulletl2(X).Y = ShipY + 15
      Bulletr2(X).X = ShipX + 15
      Bulletr2(X).Y = ShipY + 26
   End If
   If ind = 1 Then
      Bulletl2(X).X = ShipX + 15
      Bulletl2(X).Y = ShipY + 14
      Bulletr2(X).X = ShipX + 15
      Bulletr2(X).Y = ShipY + 36
   End If
   If ind = 2 Then
      Bulletl2(X).X = ShipX + 15
      Bulletl2(X).Y = ShipY + 21
      Bulletr2(X).X = ShipX + 15
      Bulletr2(X).Y = ShipY + 53
   End If
   If ind = 3 Then
      Bulletl2(X).X = ShipX + 15
      Bulletl2(X).Y = ShipY + 24
      Bulletr2(X).X = ShipX + 15
      Bulletr2(X).Y = ShipY + 61
   End If
      
   BulletsActivated = BulletsActivated + 1
End If

If Bulletl2(X).activated = 1 Then
    Bulletl2(X).X = Bulletl2(X).X + BulletSpeed
    Bulletl2(X).Y = Bulletl2(X).Y - Sin(36 * Radians) * 50
    If Bulletl2(X).X > Form1.PicMain.ScaleWidth Then Bulletl2(X).activated = 0
    BitBlt Form1.PicScreenBuffer.hdc, Bulletl2(X).X, Bulletl2(X).Y, 16, 8, Form2.PicBulletM.hdc, 0, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hdc, Bulletl2(X).X, Bulletl2(X).Y, 16, 8, Form2.PicBullet.hdc, 0, 0, vbSrcAnd

    For Y = 0 To CurLevel.NumOfBadGuys
      If BadGuys(Y).activated = 1 And BadGuys(Y).Exploding = 0 And Bulletl2(X).activated = 1 Then
         If Abs((BadGuys(Y).X + (BadGuys(Y).xsize / 2)) - (Bulletl2(X).X + 0.5)) < (BadGuys(Y).xsize / 2) And Abs((BadGuys(Y).Y + (BadGuys(Y).ysize / 2)) - (Bulletl2(X).Y + 2)) < (BadGuys(Y).ysize / 2) Then
            BadGuys(Y).Damage = BadGuys(Y).Damage + 1
            Bulletl2(X).activated = 0
            BitBlt Form1.PicScreenBuffer.hdc, Bulletl2(X).X, Bulletl2(X).Y - 12, 12, 12, Form2.PicHITM.hdc, 12 * BadGuys(Y).frame, 0, vbPatInvert
            BitBlt Form1.PicScreenBuffer.hdc, Bulletl2(X).X, Bulletl2(X).Y - 12, 12, 12, Form2.PicHIT.hdc, 12 * BadGuys(Y).frame, 0, vbSrcPaint
            BadGuys(Y).frame = BadGuys(Y).frame + 1
            If BadGuys(Y).frame > 6 Then BadGuys(Y).frame = 0
         End If
      End If
    Next Y
    
End If
If Bulletr2(X).activated = 1 Then
    Bulletr2(X).X = Bulletr2(X).X + BulletSpeed
    Bulletr2(X).Y = Bulletr2(X).Y + Sin(36 * Radians) * 50
    If Bulletr2(X).X > Form1.PicMain.ScaleWidth Then Bulletr2(X).activated = 0
    BitBlt Form1.PicScreenBuffer.hdc, Bulletr2(X).X, Bulletr2(X).Y, 16, 8, Form2.PicBulletM.hdc, 0, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hdc, Bulletr2(X).X, Bulletr2(X).Y, 16, 8, Form2.PicBullet.hdc, 0, 0, vbSrcAnd

    For Y = 0 To CurLevel.NumOfBadGuys
       If BadGuys(Y).activated = 1 And BadGuys(Y).Exploding = 0 And Bulletr2(X).activated = 1 Then
         If Abs((BadGuys(Y).X + (BadGuys(Y).xsize / 2)) - (Bulletr2(X).X + 0.5)) < (BadGuys(Y).xsize / 2) And Abs((BadGuys(Y).Y + (BadGuys(Y).ysize / 2)) - (Bulletr2(X).Y + 2)) < (BadGuys(Y).ysize / 2) Then
            BadGuys(Y).Damage = BadGuys(Y).Damage + 1
            Bulletr2(X).activated = 0
            BitBlt Form1.PicScreenBuffer.hdc, Bulletr2(X).X, Bulletr2(X).Y - 12, 12, 12, Form2.PicHITM.hdc, 12 * BadGuys(Y).frame, 0, vbPatInvert
            BitBlt Form1.PicScreenBuffer.hdc, Bulletr2(X).X, Bulletr2(X).Y - 12, 12, 12, Form2.PicHIT.hdc, 12 * BadGuys(Y).frame, 0, vbSrcPaint
            BadGuys(Y).frame = BadGuys(Y).frame + 1
            If BadGuys(Y).frame > 6 Then BadGuys(Y).frame = 0
         End If
       End If
    Next Y
    
End If

22 Next X
End If

'Don't allow any more bullets to be created
BulletsActivated = 0

End Sub

Public Sub ScrollShip()
    'Scrolling across edges of screen
    If ShipX > PicMain.ScaleWidth Then ShipX = 0
    If ShipX < -51 Then ShipX = PicMain.ScaleWidth
    If ShipY > PicMain.ScaleHeight Then ShipY = 0
    If ShipY < -44 Then ShipY = PicMain.ScaleHeight
End Sub

Public Sub VelocityCode()

'Movement Up
If sUp = 1 Then
sVel = sVel - VAcc
If sVel < -50 Then sVel = -50
ShipY = ShipY + sVel
If ShipY < 0 Then
   ShipY = 0
   sVel = 0
End If
End If

'Movement Down
If sDown = 1 Then
sVel = sVel + VAcc
If sVel > 50 Then sVel = 50
ShipY = ShipY + sVel
If ShipY > Form1.PicMain.ScaleHeight - 130 Then
   ShipY = Form1.PicMain.ScaleHeight - 130
   sVel = 0
End If
End If

'Vertical Deceleration
If sUp = 0 And sDown = 0 And sVel <> 0 Then
'Form2.Picture1.Picture = Form2.PicD(0).Picture
'Form2.Picture2.Picture = Form2.PicDM(0).Picture
If sVel > 0 Then
sVel = sVel - VDel
If sVel <= 0 Then sVel = 0
Else
sVel = sVel + VDel
If sVel >= 0 Then sVel = 0
End If
If ShipY > Form1.PicMain.ScaleHeight - 130 Then
   ShipY = Form1.PicMain.ScaleHeight - 130
   sVel = 0
ElseIf ShipY < 0 Then
   ShipY = 0
   sVel = 0
Else
   ShipY = ShipY + sVel
End If
End If

'Movement Left
If sLeft = 1 Then
    sVelocity = sVelocity - HAcc
    If sVelocity < -50 Then sVelocity = -50
    ShipX = ShipX + sVelocity
End If

'Movement Right
If sRight = 1 Then
    sVelocity = sVelocity + HAcc
    If sVelocity > 50 Then sVelocity = 50
    ShipX = ShipX + sVelocity
End If

If ShipX >= Form1.PicMain.ScaleWidth - 120 Then
    ShipX = Form1.PicMain.ScaleWidth - 120
    sVelocity = 0
End If

If ShipX <= 0 Then
    ShipX = 0
    sVelocity = 0
End If

'Horizontal Deceleration
If ShipX > 0 And ShipX < Form1.PicMain.ScaleWidth - 120 Then
 If sRight = 0 And sLeft = 0 And sVelocity <> 0 Then
   If sVelocity > 0 Then
        sVelocity = sVelocity - HDel
        If sVelocity <= 0 Then sVelocity = 0
    Else
        sVelocity = sVelocity + HDel
        If sVelocity >= 0 Then sVelocity = 0
    End If
    ShipX = ShipX + sVelocity
 End If
End If

If sVelocity = 0 And sUp = 0 And sDown = 0 Then
    Form2.Picture1.Picture = Form2.PicD(0).Picture
    Form2.Picture2.Picture = Form2.PicDM(0).Picture
End If

End Sub

Public Sub BuildNewStar1(ByVal ArrayVal As Integer)
    StarArray(ArrayVal).X = Rnd * Form1.PicMain.ScaleWidth
    StarArray(ArrayVal).Y = Rnd * Form1.PicMain.ScaleHeight
    StarArray(ArrayVal).bright = Rnd * 255
    StarArray(ArrayVal).SPEED = Rnd * 5 + 50
End Sub

Public Sub BuildNewStar(ByVal ArrayVal As Integer)
    StarArray(ArrayVal).X = Form1.PicMain.ScaleWidth
    StarArray(ArrayVal).Y = Rnd * Form1.PicMain.ScaleHeight
    StarArray(ArrayVal).bright = Rnd * 200 + 55
    StarArray(ArrayVal).SPEED = Rnd * 5 + 50
End Sub

Public Sub GoodGuyStuff()

    If Health <= 0 Then Exploding = 1
    If Exploding = 1 Then
       Firing = 0
       Form1.Timer1.Enabled = False
    End If
    If Exploding = 1 And ExplodingFrame = 0 Then
        For q = 0 To 5
            Explosions(q).X = ShipX + Int(Rnd * 51)
            Explosions(q).Y = ShipY + Int(Rnd * 44)
        Next q
        ExplodingFrame = 1
        If flgsound = True Then MDirX.Play_Sound "shipxpl.wav", False, True, 1
    End If
    If Exploding = 1 And ExplodingFrame <> 0 Then
        For q = 0 To 5
          If q <= 1 Then
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).X, Explosions(q).Y, 120, 121, Form2.PicExplodem.hdc, 120 * ExplodingFrame, 0, vbPatInvert
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).X, Explosions(q).Y, 120, 121, Form2.PicExplode.hdc, 120 * ExplodingFrame, 0, vbSrcPaint
          ElseIf q <= 3 Then
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).X, Explosions(q).Y, 160, 137, Form2.PicExplode1M.hdc, 160 * ExplodingFrame, 0, vbPatInvert
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).X, Explosions(q).Y, 160, 137, Form2.PicExplode1.hdc, 160 * ExplodingFrame, 0, vbSrcPaint
          Else
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).X, Explosions(q).Y, 120, 121, Form2.PicExplode2M.hdc, 120 * ExplodingFrame, 0, vbPatInvert
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).X, Explosions(q).Y, 120, 121, Form2.PicExplode2.hdc, 120 * ExplodingFrame, 0, vbSrcPaint
          End If
        Next q
        ExplodingFrame = ExplodingFrame + 1
    End If

    If ExplodingFrame >= 14 Then
         Exploding = 0
         Health = 19
         lives = lives - 1
         Form1.Piclive.Caption = lives
         DrawHealthBar
         ExplodingFrame = 0
         powerup = 0
         Form1.Tmr_countdown.Enabled = False
         tempo2 = 60
         Form1.Countdown.Caption = 0
         Form1.Tmr_upgr.Enabled = True
         Form2.Picture1.Picture = Form2.PicFlash.Picture
         If lives < 0 Then
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
               Unload Form1
               Form1.Timer2.Enabled = False
               FrmEnd.Show
            End If
         End If
    End If

    If Exploding = 0 Then
        BitBlt Form1.PicScreenBuffer.hdc, ShipX, ShipY, ShipW, ShipH, Form2.Picture2.hdc, 0, 0, vbMergePaint
        BitBlt Form1.PicScreenBuffer.hdc, ShipX, ShipY, ShipW, ShipH, Form2.Picture1.hdc, 0, 0, vbSrcAnd
        BitBlt Form1.PicScreenBuffer.hdc, ShipX - 7, ShipY + (Form2.Picture1.ScaleHeight / 2) - 8, 9, 19, Form2.PicthrustSm.hdc, 0, 19 * thrust, vbPatInvert
        BitBlt Form1.PicScreenBuffer.hdc, ShipX - 7, ShipY + (Form2.Picture1.ScaleHeight / 2) - 8, 9, 19, Form2.PicthrustS.hdc, 0, 19 * thrust, vbSrcPaint
        thrust = thrust + 1
        If thrust > 3 Then thrust = 0
    End If
    
BitBlt Form1.PicMain.hdc, 0, 0, BufferWidth, BufferHeight, Form1.PicScreenBuffer.hdc, 0, 0, vbSrcCopy

End Sub

Public Sub BuildBadGuys()

If stage = 1 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 5
        CurLevel.Velocity = 50
        CurLevel.Damagelimit = 20
        CurLevel.BulletSpeed = 100
        CurLevel.OddsOfFiring = 2
End If
If stage = 2 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 4
        CurLevel.Velocity = 60
        CurLevel.Damagelimit = 20
        CurLevel.BulletSpeed = 100
        CurLevel.OddsOfFiring = 2
End If
If stage = 3 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 5
        CurLevel.Velocity = 50
        CurLevel.Damagelimit = 10
        CurLevel.BulletSpeed = 100
        CurLevel.OddsOfFiring = 2
End If
If stage = 4 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 4
        CurLevel.Velocity = 50
        CurLevel.Damagelimit = 20
        CurLevel.BulletSpeed = 80
        CurLevel.OddsOfFiring = 2
End If
If stage = 5 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 5
        CurLevel.Velocity = 50
        CurLevel.Damagelimit = 20
        CurLevel.BulletSpeed = 80
        CurLevel.OddsOfFiring = 2
End If
If stage = 6 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 4
        CurLevel.Velocity = 60
        CurLevel.Damagelimit = 20
        CurLevel.BulletSpeed = 100
        CurLevel.OddsOfFiring = 2
End If
If stage = 7 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 5
        CurLevel.Velocity = 40
        CurLevel.Damagelimit = 10
        CurLevel.BulletSpeed = 80
        CurLevel.OddsOfFiring = 2
End If
If stage = 8 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 50
        CurLevel.Velocity = 25
        CurLevel.Damagelimit = 10
End If
TempCalc = 0
Tempclac = 0
mode = Rnd * 4
bullets = False
numsec = 0
frame = 0
appo = False

ReDim BadGuys(0 To CurLevel.NumOfBadGuys) As BadGuy

For X = 0 To CurLevel.NumOfBadGuys
     Randomize
     BadGuys(X).activated = 1
     BadGuys(X).X = Form1.PicMain.ScaleWidth
     BadGuys(X).Y = Rnd * (Form1.PicMain.ScaleHeight - 180)
     BadGuys(X).DstX = Form1.PicMain.ScaleWidth + 100
     BadGuys(X).DstY = Rnd * (Form1.PicMain.ScaleHeight - 180)
     BadGuys(X).indice = Rnd * CurLevel.NumOfBadGuys
     BadGuys(X).Velocity = Rnd * CurLevel.Velocity
     If BadGuys(X).Velocity <= 20 Then BadGuys(X).Velocity = 20
     BadGuys(X).Damage = CurLevel.Damage
     BadGuys(X).Exploding = 0
     BadGuys(X).frame = 0
     BadGuys(X).ExplodingFrame = 0
     BadGuys(X).boomed = False
     BadGuys(X).shot = False
     BadGuys(X).inverse = False
Next X

End Sub

Public Sub MoveAndDrawBack()

'Get the exact Xposition relative to the background tiles
If BackxPos > 0 Then BitBlt Form1.PicScreenBuffer.hdc, 0, 0, ViewportWidth, ViewportHeight, BackTile1, Tilewidth - BackxPos, 0, vbSrcCopy
If flgplanets = 1 Then
   If numplanets = 3 Then
      BitBlt Form1.PicScreenBuffer.hdc, planets.X, planets.Y, 117, 117, Form2.Picplanets3.hdc, 0, 0, vbMergePaint
      BitBlt Form1.PicScreenBuffer.hdc, planets.X, planets.Y, 117, 117, Form2.Picplanets(numplanets).hdc, 0, 0, vbSrcAnd
   ElseIf numplanets = 2 Then
      BitBlt Form1.PicScreenBuffer.hdc, planets.X, planets.Y, 99, 99, Form2.Picplanets2.hdc, 0, 0, vbMergePaint
      BitBlt Form1.PicScreenBuffer.hdc, planets.X, planets.Y, 99, 99, Form2.Picplanets(numplanets).hdc, 0, 0, vbSrcAnd
   ElseIf numplanets = 9 Then
      BitBlt Form1.PicScreenBuffer.hdc, planets.X, planets.Y, 117, 117, Form2.Picplanets9.hdc, 0, 0, vbMergePaint
      BitBlt Form1.PicScreenBuffer.hdc, planets.X, planets.Y, 117, 117, Form2.Picplanets(numplanets).hdc, 0, 0, vbSrcAnd
   ElseIf numplanets = 5 Then
      BitBlt Form1.PicScreenBuffer.hdc, planets.X, planets.Y, 99, 99, Form2.Picplanets5.hdc, 0, 0, vbMergePaint
      BitBlt Form1.PicScreenBuffer.hdc, planets.X, planets.Y, 99, 99, Form2.Picplanets(numplanets).hdc, 0, 0, vbSrcAnd
   Else
      BitBlt Form1.PicScreenBuffer.hdc, planets.X, planets.Y, 61, 66, Form2.Picplanets(numplanets).hdc, 0, 0, vbSrcCopy
   End If
   planets.X = planets.X - 2
End If
If flgupgr = 1 Then
   Set upgr1.mask = Form2.upgrMask
   BitBlt Form1.PicScreenBuffer.hdc, upgr1.X, upgr1.Y, 32, 30, Form2.Picupgr1M.hdc, 32 * upgr1.frame, 0, vbMergePaint
   BitBlt Form1.PicScreenBuffer.hdc, upgr1.X, upgr1.Y, 32, 30, Form2.Picupgr1.hdc, 32 * upgr1.frame, 0, vbSrcAnd
   upgr1.frame = upgr1.frame + 1
   If upgr1.frame > 19 Then upgr1.frame = 0
   upgr1.X = upgr1.X - 5
   If CollisionDetect(ShipX, ShipY, Form2.Pictm, upgr1.X, upgr1.Y, upgr1.mask, Form2.PicTemp) And Exploding = 0 Then
      If flgsound = True Then MDirX.Play_Sound "shield.wav", False, True, 4
      Health = 19
      DrawHealthBar
      flgupgr = 0
   End If
   If upgr1.X < -30 Then flgupgr = 0
End If
If power = 1 Then
   Set upgr2.mask = Form2.PowerMask
   BitBlt Form1.PicScreenBuffer.hdc, upgr2.X, upgr2.Y, 46, 37, Form2.PicpowerM.hdc, 46 * upgr2.frame, 0, vbMergePaint
   BitBlt Form1.PicScreenBuffer.hdc, upgr2.X, upgr2.Y, 46, 37, Form2.Picpower.hdc, 46 * upgr2.frame, 0, vbSrcAnd
   upgr2.frame = upgr2.frame + 1
   If upgr2.frame > 14 Then upgr2.frame = 0
   upgr2.X = upgr2.X - 5
   flgpower = 0
   If CollisionDetect(ShipX, ShipY, Form2.Pictm, upgr2.X, upgr2.Y, upgr2.mask, Form2.PicTemp) And Exploding = 0 Then
      If flgsound = True Then MDirX.Play_Sound "power.wav", False, True, 4
      powerup = 1
      power = 0
      tempo2 = 60
      Form1.Tmr_countdown.Enabled = True
   End If
   If upgr2.X < -45 Then
      power = 0
      Form1.Tmr_upgr.Enabled = True
   End If
End If

If BackxPos <= 800 Then
   BackxPos = 1600
   If numplanets < 10 Then
      flgplanets = 1
      numplanets = numplanets + 1
      planets.X = 900
      planets.Y = Rnd * (Form1.PicMain.ScaleHeight - 180)
   Else
      flgplanets = 0
   End If
End If

BackxPos = BackxPos - 1
End Sub
Public Sub cleanup()
  DeleteGeneratedDC BackTile1
End Sub
Public Function GenerateDC(FileName As String) As Long
Dim dc As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
dc = CreateCompatibleDC(ByVal 0&)

If dc < 1 Then
    GenerateDC = 0
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC dc
    GenerateDC = 0
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject dc, hBitmap

'Return the device context
GenerateDC = dc

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Public Function DeleteGeneratedDC(dc As Long) As Long

If dc > 0 Then
    DeleteGeneratedDC = DeleteDC(dc)
Else
    DeleteGeneratedDC = 0
End If

End Function
Public Sub AngleRadians()

    deltax = (ShipX + Form2.Picture1.ScaleWidth / 2) - curX
    deltay = (ShipY + Form2.Picture1.ScaleHeight / 2) - curY
    
    If deltax = 0 Then      'Vertical
        If deltay < 0 Then
            angle = Pi / 2
        Else
'            Angle = Pi * 2
            angle = Pi * 1.5
        End If
    
    ElseIf deltay = 0 Then  'Horizontal
        If deltax >= 0 Then
            angle = 0
        Else
            angle = Pi
        End If
    
    Else
        angle = Atn(Abs(deltay / deltax))
        
        If deltax >= 0 And deltay >= 0 Then
            angle = (Pi * 2) - angle
            
        ElseIf deltax < 0 And deltay >= 0 Then
            angle = Pi + angle
            
        ElseIf deltax < 0 And deltay < 0 Then
            angle = Pi - angle
            
        End If
        
    End If
    angle = angle * (180# / Pi)
End Sub

Public Sub Blend(Destination As Object, source As Object, Amount As Integer, X, Y, X2, Y2)
AlphaBlending Destination.hdc, X, Y, X2, Y2, source.hdc, X, Y, X2, Y2, Amount
End Sub
