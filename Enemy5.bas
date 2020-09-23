Attribute VB_Name = "Enemy5"
Public Sub badguy5()

For X = 0 To CurLevel.NumOfBadGuys
If BadGuys(X).activated = 0 Then GoTo 10
    Set BadGuys(X).PicT = Form2.Alien3(BadGuys(X).indice)
    Set BadGuys(X).mask = Form2.Alien3M(BadGuys(X).indice)
    If BadGuys(X).indice = 0 Then
       BadGuys(X).bulletlxpos = 23
       BadGuys(X).bulletlypos = 9
       BadGuys(X).bulletrxpos = 23
       BadGuys(X).bulletrypos = 9
    End If
    If BadGuys(X).indice = 1 Then
       BadGuys(X).bulletlxpos = 16
       BadGuys(X).bulletlypos = 26
       BadGuys(X).bulletrxpos = 16
       BadGuys(X).bulletrypos = 10
    End If
    If BadGuys(X).indice = 2 Then
       BadGuys(X).bulletlxpos = 15
       BadGuys(X).bulletlypos = 49
       BadGuys(X).bulletrxpos = 15
       BadGuys(X).bulletrypos = 22
    End If
    If BadGuys(X).indice = 3 Then
       BadGuys(X).bulletlxpos = 14
       BadGuys(X).bulletlypos = 68
       BadGuys(X).bulletrxpos = 14
       BadGuys(X).bulletrypos = 30
    End If
    If BadGuys(X).indice = 4 Then
       BadGuys(X).bulletlxpos = 13
       BadGuys(X).bulletlypos = 78
       BadGuys(X).bulletrxpos = 13
       BadGuys(X).bulletrypos = 34
    End If
    If BadGuys(X).indice = 5 Then
       BadGuys(X).bulletlxpos = 12
       BadGuys(X).bulletlypos = 85
       BadGuys(X).bulletrxpos = 12
       BadGuys(X).bulletrypos = 37
    End If
    If BadGuys(X).indice = 6 Then
       BadGuys(X).bulletlxpos = 15
       BadGuys(X).bulletlypos = 79
       BadGuys(X).bulletrxpos = 15
       BadGuys(X).bulletrypos = 34
    End If
    If BadGuys(X).indice = 7 Then
       BadGuys(X).bulletlxpos = 14
       BadGuys(X).bulletlypos = 68
       BadGuys(X).bulletrxpos = 14
       BadGuys(X).bulletrypos = 30
    End If
    If BadGuys(X).indice = 8 Then
       BadGuys(X).bulletlxpos = 14
       BadGuys(X).bulletlypos = 51
       BadGuys(X).bulletrxpos = 14
       BadGuys(X).bulletrypos = 23
    End If
    If BadGuys(X).indice = 9 Then
       BadGuys(X).bulletlxpos = 14
       BadGuys(X).bulletlypos = 30
       BadGuys(X).bulletrxpos = 14
       BadGuys(X).bulletrypos = 13
    End If
    If BadGuys(X).indice = 10 Then
       BadGuys(X).bulletlxpos = 18
       BadGuys(X).bulletlypos = 9
       BadGuys(X).bulletrxpos = 18
       BadGuys(X).bulletrypos = 9
    End If
    If BadGuys(X).indice = 11 Then
       BadGuys(X).bulletlxpos = 10
       BadGuys(X).bulletlypos = 25
       BadGuys(X).bulletrxpos = 10
       BadGuys(X).bulletrypos = 10
    End If
    If BadGuys(X).indice = 12 Then
       BadGuys(X).bulletlxpos = 6
       BadGuys(X).bulletlypos = 20
       BadGuys(X).bulletrxpos = 6
       BadGuys(X).bulletrypos = 45
    End If
    If BadGuys(X).indice = 13 Then
       BadGuys(X).bulletlxpos = 9
       BadGuys(X).bulletlypos = 63
       BadGuys(X).bulletrxpos = 9
       BadGuys(X).bulletrypos = 29
    End If
    If BadGuys(X).indice = 14 Then
       BadGuys(X).bulletlxpos = 10
       BadGuys(X).bulletlypos = 75
       BadGuys(X).bulletrxpos = 10
       BadGuys(X).bulletrypos = 34
    End If
    If BadGuys(X).indice = 15 Then
       BadGuys(X).bulletlxpos = 14
       BadGuys(X).bulletlypos = 82
       BadGuys(X).bulletrxpos = 14
       BadGuys(X).bulletrypos = 36
    End If
    If BadGuys(X).indice = 16 Then
       BadGuys(X).bulletlxpos = 12
       BadGuys(X).bulletlypos = 88
       BadGuys(X).bulletrxpos = 12
       BadGuys(X).bulletrypos = 38
    End If
    If BadGuys(X).indice = 17 Then
       BadGuys(X).bulletlxpos = 14
       BadGuys(X).bulletlypos = 75
       BadGuys(X).bulletrxpos = 14
       BadGuys(X).bulletrypos = 33
    End If
    If BadGuys(X).indice = 18 Then
       BadGuys(X).bulletlxpos = 12
       BadGuys(X).bulletlypos = 63
       BadGuys(X).bulletrxpos = 12
       BadGuys(X).bulletrypos = 28
    End If
    If BadGuys(X).indice = 19 Then
       BadGuys(X).bulletlxpos = 8
       BadGuys(X).bulletlypos = 49
       BadGuys(X).bulletrxpos = 8
       BadGuys(X).bulletrypos = 22
    End If
    
    BadGuys(X).xsize = Form2.Alien3(BadGuys(X).indice).ScaleWidth
    BadGuys(X).ysize = Form2.Alien3(BadGuys(X).indice).ScaleHeight

    BadGuys(X).oldX = BadGuys(X).X
    BadGuys(X).oldY = BadGuys(X).Y
    If BadGuys(X).activated = 1 And BadGuys(X).Exploding = 0 Then
       If mode = 0 Or mode = 4 Then BadGuys(X).X = BadGuys(X).X - BadGuys(X).Velocity
       If mode = 1 Then
          If X > 0 And numsec < 10 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
          If X > 1 And numsec < 20 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
          If X > 2 And numsec < 30 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
          If X > 3 And numsec < 40 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
          If X > 4 And numsec < 50 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
          If X > 5 And numsec < 60 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
          BadGuys(X).X = BadGuys(X).X - BadGuys(X).Velocity
       End If
      If mode = 2 Then
         BadGuys(X).Velocity = 20
         If BadGuys(X).X < Form1.PicMain.ScaleLeft + 400 Then BadGuys(X).inverse = True
    
         If X = 0 And numsec < 1 Then
            BadGuys(X).Y = Form1.PicMain.ScaleTop - 300
            BadGuys(X).X = Form1.PicMain.ScaleWidth
         End If
         If X > 0 And numsec < 5 Then
            BadGuys(X).Y = Form1.PicMain.ScaleTop - 300
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If X > 1 And numsec < 10 Then
            BadGuys(X).Y = Form1.PicMain.ScaleTop - 300
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If X > 2 And numsec < 15 Then
            BadGuys(X).Y = Form1.PicMain.ScaleTop - 300
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If X > 3 And numsec < 20 Then
            BadGuys(X).Y = Form1.PicMain.ScaleTop - 300
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If X > 4 And numsec < 25 Then
            BadGuys(X).Y = Form1.PicMain.ScaleTop - 300
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If X > 5 And numsec < 30 Then
            BadGuys(X).Y = Form1.PicMain.ScaleTop - 300
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If BadGuys(X).inverse = True Then
            BadGuys(X).X = BadGuys(X).X + BadGuys(X).Velocity
            BadGuys(X).Y = BadGuys(X).Y - Sin(15 * Radians) + BadGuys(X).Velocity
         Else
            BadGuys(X).X = BadGuys(X).X - BadGuys(X).Velocity
            BadGuys(X).Y = BadGuys(X).Y + Sin(45 * Radians) + BadGuys(X).Velocity
         End If
      End If
      If mode = 3 Then
         BadGuys(X).Velocity = 20
         If BadGuys(X).X < Form1.PicMain.ScaleLeft + 400 Then BadGuys(X).inverse = True
         If X = 0 And numsec < 1 Then
            BadGuys(X).Y = Form1.PicMain.ScaleHeight + 200
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If X > 0 And numsec < 5 Then
            BadGuys(X).Y = Form1.PicMain.ScaleHeight + 200
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If X > 1 And numsec < 10 Then
            BadGuys(X).Y = Form1.PicMain.ScaleHeight + 200
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If X > 2 And numsec < 15 Then
            BadGuys(X).Y = Form1.PicMain.ScaleHeight + 200
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If X > 3 And numsec < 20 Then
            BadGuys(X).Y = Form1.PicMain.ScaleHeight + 200
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If X > 4 And numsec < 25 Then
            BadGuys(X).Y = Form1.PicMain.ScaleHeight + 200
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If X > 5 And numsec < 30 Then
            BadGuys(X).Y = Form1.PicMain.ScaleHeight + 200
            BadGuys(X).X = Form1.PicMain.ScaleWidth + 150
         End If
         If BadGuys(X).inverse = True Then
            BadGuys(X).X = BadGuys(X).X + BadGuys(X).Velocity
            BadGuys(X).Y = BadGuys(X).Y - Sin(15 * Radians) - BadGuys(X).Velocity
         Else
            BadGuys(X).X = BadGuys(X).X - BadGuys(X).Velocity
            BadGuys(X).Y = BadGuys(X).Y - Sin(45 * Radians) - BadGuys(X).Velocity
         End If
      End If
      BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).X, BadGuys(X).Y, BadGuys(X).xsize, BadGuys(X).ysize, BadGuys(X).mask.hdc, 0, 0, vbMergePaint
      BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).X, BadGuys(X).Y, BadGuys(X).xsize, BadGuys(X).ysize, BadGuys(X).PicT.hdc, 0, 0, vbSrcAnd
    End If
    If BadGuys(X).Damage > CurLevel.Damagelimit Then
       BadGuys(X).Exploding = 1
       If flgsound = True Then MDirX.Play_Sound "xplode1.wav", False, True, 2
    End If
   If CollisionDetect(ShipX, ShipY, Form2.Pictm, BadGuys(X).X, BadGuys(X).Y, BadGuys(X).mask, Form2.PicTemp) And Exploding = 0 Then
      If flgsound = True Then MDirX.Play_Sound "xplode1.wav", False, True, 2
      Exploding = 1
      BadGuys(X).Exploding = 1
   End If
BadGuys(X).indice = BadGuys(X).indice + 1
If BadGuys(X).indice > 19 Then BadGuys(X).indice = 0

   
If BadGuys(X).activated = 1 Then
   BadGuys(X).Firing = Int(Rnd * CurLevel.OddsOfFiring)
Else
   BadGuys(X).Firing = 0
End If

'Bullets
If mode <> 2 And mode <> 3 Then
If BadGuys(X).X < (Form1.PicMain.ScaleWidth / 2) + 100 Then
For Y = 0 To 0
If BadGuys(X).Firing = 1 And BadGuys(X).BulletsActivated <= 1 Then
If BadGuys(X).activated > 0 Then
If BadGuys(X).Bulletl(Y).activated = 0 Then
   BadGuys(X).Bulletl(Y).activated = 1
   BadGuys(X).Bulletl(Y).X = BadGuys(X).X - BadGuys(X).bulletlxpos
   BadGuys(X).Bulletl(Y).Y = BadGuys(X).Y + BadGuys(X).bulletlypos + 10
   BadGuys(X).shot = True
End If
If BadGuys(X).Bulletr(Y).activated = 0 Then
   BadGuys(X).Bulletr(Y).activated = 1
   BadGuys(X).Bulletr(Y).X = BadGuys(X).X - BadGuys(X).bulletrxpos
   BadGuys(X).Bulletr(Y).Y = BadGuys(X).Y + BadGuys(X).bulletrypos
   BadGuys(X).shot = True
End If
BadGuys(X).BulletsActivated = BadGuys(X).BulletsActivated + 1
End If
End If
11 Next Y
End If
End If


If (BadGuys(X).Y > Rnd * (Form1.PicMain.ScaleHeight - 180) And mode = 2) Or _
   (BadGuys(X).Y < Rnd * (Form1.PicMain.ScaleHeight - 180) And mode = 3) Then

For Y = 0 To 0
If BadGuys(X).Firing = 1 And BadGuys(X).BulletsActivated <= 1 Then
If BadGuys(X).activated > 0 Then
If BadGuys(X).Bulletl(Y).activated = 0 Then
   BadGuys(X).Bulletl(Y).activated = 1
   BadGuys(X).Bulletl(Y).X = BadGuys(X).X - BadGuys(X).bulletlxpos
   BadGuys(X).Bulletl(Y).Y = BadGuys(X).Y + BadGuys(X).bulletlypos + 10
   BadGuys(X).shot = True
End If
If BadGuys(X).Bulletr(Y).activated = 0 Then
   BadGuys(X).Bulletr(Y).activated = 1
   BadGuys(X).Bulletr(Y).X = BadGuys(X).X - BadGuys(X).bulletrxpos
   BadGuys(X).Bulletr(Y).Y = BadGuys(X).Y + BadGuys(X).bulletrypos
   BadGuys(X).shot = True
End If
BadGuys(X).BulletsActivated = BadGuys(X).BulletsActivated + 1
End If
End If
22 Next Y
End If

If BadGuys(X).shot = True Then
   If flgsound = True Then MDirX.Play_Sound "missile.wav", False, True, 3
   BadGuys(X).shot = False
End If

10 Next X
    
'************************************************************
'Firing bullets
For X = 0 To CurLevel.NumOfBadGuys
For Y = 0 To 0
If BadGuys(X).Bulletl(Y).activated = 1 Then
bullets = True
BadGuys(X).Bulletl(Y).X = BadGuys(X).Bulletl(Y).X - CurLevel.BulletSpeed
If BadGuys(X).Bulletl(Y).X < 0 Then
   BadGuys(X).Bulletl(Y).activated = 0
   bullets = False
End If
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).Bulletl(Y).X, BadGuys(X).Bulletl(Y).Y, 152, 12, Form2.PicMissileM.hdc, 0, 12 * frame, vbMergePaint
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).Bulletl(Y).X, BadGuys(X).Bulletl(Y).Y, 152, 12, Form2.PicMissile.hdc, 0, 12 * frame, vbSrcAnd
frame = frame + 1
If frame > 10 Then frame = 0
If BadGuys(X).Bulletl(Y).X + 24 > ShipX And BadGuys(X).Bulletl(Y).X < ShipX + Form2.Picture1.ScaleWidth And _
   Abs((BadGuys(X).Bulletl(Y).Y - 25) - ShipY) < 18 Then
   If flgsound = True Then MDirX.Play_Sound "hit.wav", False, True, 3
   Form2.Picture1.Picture = Form2.PicFlash.Picture
   Health = Health - 1
   UpdateHealth
   BadGuys(X).Bulletl(Y).activated = 0
   bullets = False
End If
End If

If BadGuys(X).Bulletr(Y).activated = 1 Then
bullets = True
BadGuys(X).Bulletr(Y).X = BadGuys(X).Bulletr(Y).X - CurLevel.BulletSpeed
If BadGuys(X).Bulletr(Y).X < 0 Then
   BadGuys(X).Bulletr(Y).activated = 0
   bullets = False
End If
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).Bulletr(Y).X, BadGuys(X).Bulletr(Y).Y, 152, 12, Form2.PicMissileM.hdc, 0, 12 * frame, vbMergePaint
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).Bulletr(Y).X, BadGuys(X).Bulletr(Y).Y, 152, 12, Form2.PicMissile.hdc, 0, 12 * frame, vbSrcAnd
frame = frame + 1
If frame > 10 Then frame = 0
If BadGuys(X).Bulletr(Y).X + 24 > ShipX And BadGuys(X).Bulletr(Y).X < ShipX + Form2.Picture1.ScaleWidth And _
   Abs((BadGuys(X).Bulletr(Y).Y - 25) - ShipY) < 18 Then
   If flgsound = True Then MDirX.Play_Sound "hit.wav", False, True, 3
   Form2.Picture1.Picture = Form2.PicFlash.Picture
   Health = Health - 1
   UpdateHealth
   BadGuys(X).Bulletr(Y).activated = 0
   bullets = False
End If
End If

Next Y
''If comment BadGuys shoot once only
'BadGuys(X).BulletsActivated = 0
Next X

'************************************************************
'Exploding
For X = 0 To CurLevel.NumOfBadGuys
    If BadGuys(X).activated = 1 And (BadGuys(X).X < -BadGuys(X).xsize) Then
          Tempclac = Tempclac + 1
          BadGuys(X).activated = 0
    End If
    If BadGuys(X).activated = 1 And _
       (BadGuys(X).Y > Form1.PicMain.ScaleHeight + 400) And mode = 2 Then
          Tempclac = Tempclac + 1
          BadGuys(X).activated = 0
    End If
    If BadGuys(X).activated = 1 And _
       (BadGuys(X).Y < Form1.PicMain.ScaleTop - 400) And mode = 3 Then
          Tempclac = Tempclac + 1
          BadGuys(X).activated = 0
    End If
    If BadGuys(X).Exploding = 1 Then
       BadGuys(X).activated = 0
       BadGuys(X).boomed = True
       If Health < 9 And flgupgr = 0 Then
          flgupgr = 1
          upgr1.X = BadGuys(X).X
          upgr1.Y = BadGuys(X).Y
       End If
       If flgpower = 1 Then
          power = 1
          upgr2.X = BadGuys(X).X
          upgr2.Y = BadGuys(X).Y
       End If
       For Y = 0 To 9
          BadGuys(X).particle(Y).activated = 1
          BadGuys(X).particle(Y).X = BadGuys(X).X + BadGuys(X).xsize / 2 + Int(Rnd * (Y * 10))
          BadGuys(X).particle(Y).Y = BadGuys(X).Y + BadGuys(X).ysize / 2 + Int(Rnd * (Y * 10))
       Next Y

       If fboom = 1 Then
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).X, BadGuys(X).Y, 120, 121, Form2.PicExplodem.hdc, 120 * BadGuys(X).ExplodingFrame, 0, vbPatInvert
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).X, BadGuys(X).Y, 120, 121, Form2.PicExplode.hdc, 120 * BadGuys(X).ExplodingFrame, 0, vbSrcPaint
       ElseIf fboom = 2 Then
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).X, BadGuys(X).Y, 160, 137, Form2.PicExplode1M.hdc, 160 * BadGuys(X).ExplodingFrame, 0, vbPatInvert
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).X, BadGuys(X).Y, 160, 137, Form2.PicExplode1.hdc, 160 * BadGuys(X).ExplodingFrame, 0, vbSrcPaint
       Else
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).X, BadGuys(X).Y, 120, 121, Form2.PicExplode2M.hdc, 120 * BadGuys(X).ExplodingFrame, 0, vbPatInvert
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).X, BadGuys(X).Y, 120, 121, Form2.PicExplode2.hdc, 120 * BadGuys(X).ExplodingFrame, 0, vbSrcPaint
       End If
       BadGuys(X).ExplodingFrame = BadGuys(X).ExplodingFrame + 1
       If BadGuys(X).ExplodingFrame = 13 Then
          BadGuys(X).Exploding = 0
          score = score + 50
          TempCalc = TempCalc + 1
          fboom = fboom + 1
          If fboom > 3 Then fboom = 1
       End If
    End If
Next X
   
For X = 0 To CurLevel.NumOfBadGuys
If BadGuys(X).boomed = True Then
   For Y = 0 To 9
   If BadGuys(X).particle(Y).activated = 1 Then
      BadGuys(X).particle(Y).X = BadGuys(X).particle(Y).X + Cos((36 * Y) * Radians) * 20
      BadGuys(X).particle(Y).Y = BadGuys(X).particle(Y).Y + Sin((36 * Y) * Radians) * 20
      If (BadGuys(X).particle(Y).X >= Form1.PicMain.ScaleWidth) Or _
         (BadGuys(X).particle(Y).X <= Form1.PicMain.ScaleLeft) Or _
         (BadGuys(X).particle(Y).Y >= Form1.PicMain.ScaleHeight) Or _
         (BadGuys(X).particle(Y).Y <= Form1.PicMain.ScaleTop) Then
         BadGuys(X).particle(Y).activated = 0
      End If
'      If Y < 3 Then
'         BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).particle(Y).X, BadGuys(X).particle(Y).Y, 16, 16, Form2.Picpart1m.hdc, 16 * BadGuys(X).frame, 0, vbMergePaint
'         BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).particle(Y).X, BadGuys(X).particle(Y).Y, 16, 16, Form2.Picpart1.hdc, 16 * BadGuys(X).frame, 0, vbSrcAnd
'      End If
'      If Y < 5 Then
'         BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).particle(Y).X, BadGuys(X).particle(Y).Y, 16, 16, Form2.Picpart2m.hdc, 16 * BadGuys(X).frame, 0, vbMergePaint
'         BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).particle(Y).X, BadGuys(X).particle(Y).Y, 16, 16, Form2.Picpart2.hdc, 16 * BadGuys(X).frame, 0, vbSrcAnd
'      End If
'      If Y <= 9 Then
         BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).particle(Y).X, BadGuys(X).particle(Y).Y, 16, 16, Form2.Picpart3m.hdc, 16 * BadGuys(X).frame, 0, vbMergePaint
         BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).particle(Y).X, BadGuys(X).particle(Y).Y, 16, 16, Form2.Picpart3.hdc, 16 * BadGuys(X).frame, 0, vbSrcAnd
'      End If
      BadGuys(X).frame = BadGuys(X).frame + 1
      If BadGuys(X).frame = 39 Then BadGuys(X).frame = 0
  End If
  Next Y
End If
20 Next X

If Tempclac + TempCalc >= CurLevel.NumOfBadGuys + 1 And bullets = False Then
   stage = 0
End If
numsec = numsec + 1
End Sub


