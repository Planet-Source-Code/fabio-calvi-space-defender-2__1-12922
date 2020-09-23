Attribute VB_Name = "Mines"
Public Sub mine()

For X = 0 To CurLevel.NumOfBadGuys
If BadGuys(X).activated = 0 Then GoTo 10
    Set BadGuys(X).PicT = Form2.PicMines
    Set BadGuys(X).mask = Form2.PicMinesm
    BadGuys(X).xsize = 64
    BadGuys(X).ysize = 64
    BadGuys(X).oldX = BadGuys(X).X
    BadGuys(X).oldY = BadGuys(X).Y
    If BadGuys(X).activated = 1 And BadGuys(X).Exploding = 0 Then
       If X > 2 And numsec < 10 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 5 And numsec < 20 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 8 And numsec < 30 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 11 And numsec < 40 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 14 And numsec < 50 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 17 And numsec < 60 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 20 And numsec < 70 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 23 And numsec < 80 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 26 And numsec < 90 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 30 And numsec < 100 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 33 And numsec < 110 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 36 And numsec < 120 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 40 And numsec < 130 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 43 And numsec < 140 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
       If X > 46 And numsec < 150 Then BadGuys(X).X = Form1.PicMain.ScaleWidth + 200
     
      BadGuys(X).X = BadGuys(X).X - BadGuys(X).Velocity

      BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).X, BadGuys(X).Y, BadGuys(X).xsize, BadGuys(X).ysize, BadGuys(X).mask.hdc, 64 * BadGuys(X).frame, 0, vbMergePaint
      BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).X, BadGuys(X).Y, BadGuys(X).xsize, BadGuys(X).ysize, BadGuys(X).PicT.hdc, 64 * BadGuys(X).frame, 0, vbSrcAnd
      BadGuys(X).frame = BadGuys(X).frame + 1
      If BadGuys(X).frame > 49 Then BadGuys(X).frame = 0
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

10 Next X
    

For X = 0 To CurLevel.NumOfBadGuys
    If BadGuys(X).activated = 1 And BadGuys(X).X < -BadGuys(X).xsize Then
       Tempclac = Tempclac + 1
       BadGuys(X).activated = 0
    End If
    If BadGuys(X).Exploding = 1 Then
       BadGuys(X).activated = 0
       BadGuys(X).boomed = True
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
          score = score + 20
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
      If Y >= 0 And Y < 2 Then
         BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).particle(Y).X, BadGuys(X).particle(Y).Y, 32, 32, Form2.Picmine1m.hdc, 32 * BadGuys(X).frame, 0, vbMergePaint
         BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).particle(Y).X, BadGuys(X).particle(Y).Y, 32, 32, Form2.Picmine1.hdc, 32 * BadGuys(X).frame, 0, vbSrcAnd
         BadGuys(X).frame = BadGuys(X).frame + 1
         If BadGuys(X).frame > 39 Then BadGuys(X).frame = 0
      End If
      If Y >= 2 And Y < 5 Then
         BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).particle(Y).X, BadGuys(X).particle(Y).Y, 16, 16, Form2.Picmine2m.hdc, 16 * BadGuys(X).frame, 0, vbMergePaint
         BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).particle(Y).X, BadGuys(X).particle(Y).Y, 16, 16, Form2.Picmine2.hdc, 16 * BadGuys(X).frame, 0, vbSrcAnd
         BadGuys(X).frame = BadGuys(X).frame + 1
         If BadGuys(X).frame > 29 Then BadGuys(X).frame = 0
      End If
      If Y >= 5 Then
         BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).particle(Y).X, BadGuys(X).particle(Y).Y, 16, 16, Form2.Picmine3m.hdc, 16 * BadGuys(X).frame, 0, vbMergePaint
         BitBlt Form1.PicScreenBuffer.hdc, BadGuys(X).particle(Y).X, BadGuys(X).particle(Y).Y, 16, 16, Form2.Picmine3.hdc, 16 * BadGuys(X).frame, 0, vbSrcAnd
         BadGuys(X).frame = BadGuys(X).frame + 1
         If BadGuys(X).frame > 29 Then BadGuys(X).frame = 0
      End If
  
  End If
  Next Y
End If
20 Next X

If Tempclac + TempCalc >= CurLevel.NumOfBadGuys + 1 And bullets = False Then
   Form1.Timer3.Interval = 1
   stage = 1
End If
numsec = numsec + 1
End Sub

