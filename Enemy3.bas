Attribute VB_Name = "enemy3"
Public Sub badguy3()

For X = 0 To CurLevel.NumOfBadGuys
If BadGuys(X).activated = 0 Then GoTo 10
    Set BadGuys(X).PicT = Form2.Alien6(BadGuys(X).indice)
    Set BadGuys(X).mask = Form2.Alien6M(BadGuys(X).indice)
    BadGuys(X).xsize = Form2.Alien6(BadGuys(X).indice).ScaleWidth
    BadGuys(X).ysize = Form2.Alien6(BadGuys(X).indice).ScaleHeight
    BadGuys(X).bulletcxpos = BadGuys(X).xsize / 2
    BadGuys(X).bulletcypos = BadGuys(X).ysize / 2

    BadGuys(X).oldX = BadGuys(X).X
    BadGuys(X).oldY = BadGuys(X).Y
    If BadGuys(X).activated = 1 And BadGuys(X).Exploding = 0 Then
      If mode = 1 Or mode = 3 Then BadGuys(X).X = BadGuys(X).X - BadGuys(X).Velocity
      If mode = 0 Or mode = 2 Or mode = 4 Then
        If X > 0 And numsec < 10 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
        If X > 1 And numsec < 20 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
        If X > 2 And numsec < 30 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
        If X > 3 And numsec < 40 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
        If X > 4 And numsec < 50 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
        BadGuys(X).X = BadGuys(X).X - BadGuys(X).Velocity
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
   If BadGuys(X).indice > 5 Then BadGuys(X).indice = 0

'Bullets
If BadGuys(X).activated = 1 Then
   BadGuys(X).Firing = Int(Rnd * CurLevel.OddsOfFiring)
Else
   BadGuys(X).Firing = 0
End If

If BadGuys(X).X < (Form1.PicMain.ScaleWidth / 2) + 100 Then
For Y = 0 To 0
If BadGuys(X).Firing = 1 And BadGuys(X).BulletsActivated <= 1 Then
If BadGuys(X).activated > 0 Then
If BadGuys(X).Bulletc(Y).activated = 0 Then
   BadGuys(X).Bulletc(Y).activated = 1
   BadGuys(X).Bulletc(Y).X = BadGuys(X).X - BadGuys(X).bulletcxpos
   BadGuys(X).Bulletc(Y).Y = BadGuys(X).Y + BadGuys(X).bulletcypos + 10
   BadGuys(X).shot = True
   curX = BadGuys(X).Bulletc(Y).X
   curY = BadGuys(X).Bulletc(Y).Y
   AngleRadians
   BadGuys(X).Bulletc(Y).angle = angle
End If
BadGuys(X).BulletsActivated = BadGuys(X).BulletsActivated + 1
End If
End If
11 Next Y
End If

If BadGuys(X).shot = True Then
   If flgsound = True Then MDirX.Play_Sound "laser.wav", False, True, 3
   BadGuys(X).shot = False
End If

10 Next X
 

'Firing bullets
For X = 0 To CurLevel.NumOfBadGuys
For Y = 0 To 0
If BadGuys(X).Bulletc(Y).activated = 1 Then
bullets = True

BadGuys(X).Bulletc(Y).X = BadGuys(X).Bulletc(Y).X + (CurLevel.BulletSpeed * Cos(BadGuys(X).Bulletc(Y).angle * Radians))
BadGuys(X).Bulletc(Y).Y = BadGuys(X).Bulletc(Y).Y - (CurLevel.BulletSpeed * Sin(BadGuys(X).Bulletc(Y).angle * Radians))

If (BadGuys(X).Bulletc(Y).X > Form1.PicMain.ScaleWidth) Or _
   (BadGuys(X).Bulletc(Y).X < Form1.PicMain.ScaleLeft) Or _
   (BadGuys(X).Bulletc(Y).Y > Form1.PicMain.ScaleHeight) Or _
   (BadGuys(X).Bulletc(Y).Y < Form1.PicMain.ScaleTop) Then
   BadGuys(X).Bulletc(Y).activated = 0
   bullets = False
End If
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).Bulletc(Y).X, BadGuys(X).Bulletc(Y).Y, 15, 15, Form2.PicBBulletM.hdc, 15 * frame, 0, vbPatInvert
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).Bulletc(Y).X, BadGuys(X).Bulletc(Y).Y, 15, 15, Form2.PicBBullet.hdc, 15 * frame, 0, vbSrcPaint
frame = frame + 1
If frame > 3 Then frame = 0
If BadGuys(X).Bulletc(Y).X + 7 > ShipX And BadGuys(X).Bulletc(Y).X < (ShipX + Form2.Picture1.ScaleWidth) And _
   (BadGuys(X).Bulletc(Y).Y + 7) > ShipY And (BadGuys(X).Bulletc(Y).Y + 7) < (ShipY + Form2.Picture1.ScaleHeight) Then
   If flgsound = True Then MDirX.Play_Sound "hit.wav", False, True, 3
   Form2.Picture1.Picture = Form2.PicFlash.Picture
   Health = Health - 1
   UpdateHealth
   BadGuys(X).Bulletc(Y).activated = 0
   bullets = False
End If
End If

Next Y

Next X

'Exploding
For X = 0 To CurLevel.NumOfBadGuys
    If BadGuys(X).activated = 1 And BadGuys(X).X < -BadGuys(X).xsize Then
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
   
'Explosion Particles
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
