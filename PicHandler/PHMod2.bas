Attribute VB_Name = "PHMod2"
Public Color&, TempCol&, MidX%, MidY%, Pct%, kk!
Public EchoX%, EchoY%, EchoNr%, EchoRed%, LineColor&()

Public Sub SelectAll()
FMain.Shape1.Visible = False
Xcor0 = 0
Xcor1 = FMain.Pic1.Width - 1
Ycor0 = 0
Ycor1 = FMain.Pic1.Height - 1
SetCoordinates
FMain.Toolbar1.Buttons(3).Enabled = False
End Sub

Public Sub ReadColor(Rx1%, Ry1%, Rx2%, Ry2%)
On Error Resume Next
Screen.MousePointer = 11
FMain.Label2.Caption = "Reading colors..."
DoEvents
FMain.PB1.Value = 0
FMain.PB1.Min = 0
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 '- 1
For yy = Ry1 To Ry2 '- 1
Color = GetPixel(FMain.Pic1.hdc, xx, yy)
R(xx, yy) = Color Mod 256&
G(xx, yy) = ((Color And &HFF00) / 256&) Mod 256&
B(xx, yy) = (Color And &HFF0000) / 65536
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
FMain.PB1.Value = 0
End Sub

Public Sub KillComp(Rx1%, Ry1%, Rx2%, Ry2%, Comp%)
On Error Resume Next
Dim Mask&
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
If Comp = 0 Then Mask = &HFFFF00
If Comp = 1 Then Mask = &HFF00FF
If Comp = 2 Then Mask = &HFFFF&
For xx = Rx1 To Rx2
For yy = Ry1 To Ry2
SetPixel FMain.Pic1.hdc, xx, yy, (GetPixel(FMain.Pic1.hdc, xx, yy) And Mask)
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub ColorComp(Rx1, Ry1, Rx2, Ry2, Rpct!, Gpct!, Bpct!)
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
Rpct = 1 + (Rpct / 100)
Gpct = 1 + (Gpct / 100)
Bpct = 1 + (Bpct / 100)
For xx = Rx1 To Rx2
For yy = Ry1 To Ry2
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) * Rpct, G(xx, yy) * Gpct, B(xx, yy) * Bpct)
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub SwapComp(Rx1%, Ry1%, Rx2%, Ry2%, Comp%)
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
Rpct = 1 + (Rpct / 100)
Gpct = 1 + (Gpct / 100)
Bpct = 1 + (Bpct / 100)
For xx = Rx1 To Rx2
For yy = Ry1 To Ry2
If Comp = 0 Then SetPixel FMain.Pic1.hdc, xx, yy, RGB(B(xx, yy), G(xx, yy), R(xx, yy))
If Comp = 1 Then SetPixel FMain.Pic1.hdc, xx, yy, RGB(B(xx, yy), R(xx, yy), G(xx, yy))
If Comp = 2 Then SetPixel FMain.Pic1.hdc, xx, yy, RGB(G(xx, yy), B(xx, yy), R(xx, yy))
If Comp = 3 Then SetPixel FMain.Pic1.hdc, xx, yy, RGB(G(xx, yy), R(xx, yy), B(xx, yy))
If Comp = 4 Then SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), B(xx, yy), G(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub PhotoNeg(Rx1%, Ry1%, Rx2%, Ry2%)
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2
For yy = Ry1 To Ry2
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) Xor 255, B(xx, yy) Xor 255, G(xx, yy) Xor 255)
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub PhotoNegComp(Rx1%, Ry1%, Rx2%, Ry2%, Comp%)
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2
For yy = Ry1 To Ry2
If Comp = 0 Then SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) Xor 255, G(xx, yy), B(xx, yy))
If Comp = 1 Then SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy) Xor 255, B(xx, yy))
If Comp = 2 Then SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy) Xor 255)
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub GreyColor(Rx1%, Ry1%, Rx2%, Ry2%)  'grey
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2
For yy = Ry1 To Ry2
    R(xx, yy) = R(xx, yy) * 0.3 + G(xx, yy) * 0.59 + B(xx, yy) * 0.11
    If R(xx, yy) > 255 Then R(xx, yy) = 255
    If R(xx, yy) < 0 Then R(xx, yy) = 0
    G(xx, yy) = R(xx, yy)
    B(xx, yy) = R(xx, yy)
    SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub ContrastPic(Rx1, Ry1, Rx2, Ry2, Rpct!)
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2
For yy = Ry1 To Ry2
    If R(xx, yy) > 127 Then
    R(xx, yy) = R(xx, yy) + Rpct
    Else
    R(xx, yy) = R(xx, yy) - Rpct
    End If
    If G(xx, yy) > 127 Then
    G(xx, yy) = G(xx, yy) + Rpct
    Else
    G(xx, yy) = G(xx, yy) - Rpct
    End If
    If B(xx, yy) > 127 Then
    B(xx, yy) = B(xx, yy) + Rpct
    Else
    B(xx, yy) = B(xx, yy) - Rpct
    End If
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub FlipX() 'flip horizontal
FMain.Pic1.PaintPicture FMain.Pic1, FMain.Pic1.Width, 0, -FMain.Pic1.Width, FMain.Pic1.Height
FMain.Pic1.Refresh
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Label2.Caption = "Done !"
End Sub

Public Sub FlipY() 'flip vertical
FMain.Pic1.PaintPicture FMain.Pic1, 0, FMain.Pic1.Height, FMain.Pic1.Width, -FMain.Pic1.Height
FMain.Pic1.Refresh
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Label2.Caption = "Done !"
End Sub

Public Sub MirrorX() 'mirror x
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Pic1.PaintPicture FMain.Pic1, FMain.Pic1.Width, 0, -FMain.Pic1.Width / 2, FMain.Pic1.Height, 0, 0, FMain.Pic1.Width / 2
FMain.Pic1.Refresh
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Label2.Caption = "Done !"
End Sub

Public Sub MirrorXRev() 'mirror x
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Pic1.PaintPicture FMain.Pic1, FMain.Pic1.Width, 0, -FMain.Pic1.Width, FMain.Pic1.Height
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Pic1.PaintPicture FMain.Pic1, FMain.Pic1.Width, 0, -FMain.Pic1.Width / 2, FMain.Pic1.Height, 0, 0, FMain.Pic1.Width / 2
FMain.Pic1.Refresh
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Label2.Caption = "Done !"
End Sub

Public Sub MirrorY() 'mirror y
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Pic1.PaintPicture FMain.Pic1, 0, FMain.Pic1.Height, FMain.Pic1.Width, -FMain.Pic1.Height / 2, 0, 0, FMain.Pic1.Width, FMain.Pic1.Height / 2
FMain.Pic1.Refresh
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Label2.Caption = "Done !"
End Sub

Public Sub MirrorYRev() 'mirror y
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Pic1.PaintPicture FMain.Pic1, 0, FMain.Pic1.Height, FMain.Pic1.Width, -FMain.Pic1.Height
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Pic1.PaintPicture FMain.Pic1, 0, FMain.Pic1.Height, FMain.Pic1.Width, -FMain.Pic1.Height / 2, 0, 0, FMain.Pic1.Width, FMain.Pic1.Height / 2
FMain.Pic1.Refresh
FMain.Pic1.Picture = FMain.Pic1.Image
FMain.Label2.Caption = "Done !"
End Sub

Public Sub EmbossPicture(Rx1%, Ry1%, Rx2%, Ry2%) 'emboss
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
R(xx, yy) = (Abs(R(xx, yy) - R(xx + 1, yy + 1) + 128))
G(xx, yy) = (Abs(G(xx, yy) - G(xx + 1, yy + 1) + 128))
B(xx, yy) = (Abs(B(xx, yy) - B(xx + 1, yy + 1) + 128))
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub NeonPicture(Rx1%, Ry1%, Rx2%, Ry2%) 'emboss
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
R(xx, yy) = (Abs(R(xx - 1, yy) + R(xx, yy) - R(xx + 1, yy) - R(xx + 2, yy) + 32))
G(xx, yy) = (Abs(G(xx - 1, yy) + G(xx, yy) - G(xx + 1, yy) - G(xx + 2, yy) + 32))
B(xx, yy) = (Abs(B(xx - 1, yy) + B(xx, yy) - B(xx + 1, yy) - B(xx + 2, yy) + 32))
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub EngravePicture(Rx1%, Ry1%, Rx2%, Ry2%) 'engrave
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
R(xx, yy) = (Abs(R(xx + 1, yy + 1) - R(xx, yy) + 128))
G(xx, yy) = (Abs(G(xx + 1, yy + 1) - G(xx, yy) + 148))
B(xx, yy) = (Abs(B(xx + 1, yy + 1) - B(xx, yy) + 128))
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub BeginProcess()
ReadColor 0, 0, FMain.Pic1.Width - 1, FMain.Pic1.Height - 1
FMain.Label2.Caption = "Adjusting image..."
DoEvents
FMain.PB1.Value = 0
FMain.PB1.Min = 0
End Sub

Public Sub EndProcess()
FMain.PB1.Value = 0
FMain.Pic1.Refresh
FMain.Label2.Caption = "Done !"
Screen.MousePointer = 1
End Sub

Public Sub HoldRed(Rx1, Ry1, Rx2, Ry2) 'Hold red
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
If R(xx, yy) < 128 Then
R(xx, yy) = (Abs(R(xx, yy) - R(xx + 1, yy + 1) + 128))
G(xx, yy) = (Abs(G(xx, yy) - G(xx + 1, yy + 1) + 128))
B(xx, yy) = (Abs(B(xx, yy) - B(xx + 1, yy + 1) + 128))
End If
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub HoldGreen(Rx1, Ry1, Rx2, Ry2) 'Hold green
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
If G(xx, yy) < 128 Then
R(xx, yy) = (Abs(R(xx, yy) - R(xx + 1, yy + 1) + 128))
G(xx, yy) = (Abs(G(xx, yy) - G(xx + 1, yy + 1) + 128))
B(xx, yy) = (Abs(B(xx, yy) - B(xx + 1, yy + 1) + 128))
End If
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub HoldBlue(Rx1, Ry1, Rx2, Ry2) 'Hold blue
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
If B(xx, yy) < 128 Then
R(xx, yy) = (Abs(R(xx, yy) - R(xx + 1, yy + 1) + 128))
G(xx, yy) = (Abs(G(xx, yy) - G(xx + 1, yy + 1) + 128))
B(xx, yy) = (Abs(B(xx, yy) - B(xx + 1, yy + 1) + 128))
End If
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub BlurPicture(Rx1%, Ry1%, Rx2%, Ry2%) 'blur
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 + 1 To Rx2 - 1
For yy = Ry1 + 1 To Ry2 - 1
R(xx, yy) = (Abs(R(xx - 1, yy - 1) + R(xx - 1, yy) + R(xx - 1, yy + 1) + R(xx, yy - 1) + R(xx, yy) + R(xx, yy + 1) + R(xx + 1, yy - 1) + R(xx + 1, yy) + R(xx + 1, yy + 1))) / 9
G(xx, yy) = (Abs(G(xx - 1, yy - 1) + G(xx - 1, yy) + G(xx - 1, yy + 1) + G(xx, yy - 1) + G(xx, yy) + G(xx, yy + 1) + G(xx + 1, yy - 1) + G(xx + 1, yy) + G(xx + 1, yy + 1))) / 9
B(xx, yy) = (Abs(B(xx - 1, yy - 1) + B(xx - 1, yy) + B(xx - 1, yy + 1) + B(xx, yy - 1) + B(xx, yy) + B(xx, yy + 1) + B(xx + 1, yy - 1) + B(xx + 1, yy) + B(xx + 1, yy + 1))) / 9
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub BlurPictureMore(Rx1%, Ry1%, Rx2%, Ry2%)  'blur more
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 + 2 To Rx2 - 1
For yy = Ry1 + 2 To Ry2 - 1
R(xx, yy) = (Abs(R(xx - 2, yy - 2) + R(xx - 2, yy - 1) + R(xx - 2, yy) + R(xx - 2, yy + 1) + R(xx - 2, yy + 2) + R(xx - 1, yy - 2) + R(xx - 1, yy - 1) + R(xx - 1, yy) + R(xx - 1, yy + 1) + R(xx - 1, yy + 2) + R(xx, yy - 2) + R(xx, yy - 1) + R(xx, yy) + R(xx, yy + 1) + R(xx, yy + 2) + R(xx + 1, yy - 2) + R(xx + 1, yy - 1) + R(xx + 1, yy) + R(xx + 1, yy + 1) + R(xx + 1, yy + 2) + R(xx + 2, yy - 2) + R(xx + 2, yy - 1) + R(xx + 2, yy) + R(xx + 2, yy + 1) + R(xx + 2, yy + 2))) / 25
G(xx, yy) = (Abs(G(xx - 2, yy - 2) + G(xx - 2, yy - 1) + G(xx - 2, yy) + G(xx - 2, yy + 1) + G(xx - 2, yy + 2) + G(xx - 1, yy - 2) + G(xx - 1, yy - 1) + G(xx - 1, yy) + G(xx - 1, yy + 1) + G(xx - 1, yy + 2) + G(xx, yy - 2) + G(xx, yy - 1) + G(xx, yy) + G(xx, yy + 1) + G(xx, yy + 2) + G(xx + 1, yy - 2) + G(xx + 1, yy - 1) + G(xx + 1, yy) + G(xx + 1, yy + 1) + G(xx + 1, yy + 2) + G(xx + 2, yy - 2) + G(xx + 2, yy - 1) + G(xx + 2, yy) + G(xx + 2, yy + 1) + G(xx + 2, yy + 2))) / 25
B(xx, yy) = (Abs(B(xx - 2, yy - 2) + B(xx - 2, yy - 1) + B(xx - 2, yy) + B(xx - 2, yy + 1) + B(xx - 2, yy + 2) + B(xx - 1, yy - 2) + B(xx - 1, yy - 1) + B(xx - 1, yy) + B(xx - 1, yy + 1) + B(xx - 1, yy + 2) + B(xx, yy - 2) + B(xx, yy - 1) + B(xx, yy) + B(xx, yy + 1) + B(xx, yy + 2) + B(xx + 1, yy - 2) + B(xx + 1, yy - 1) + B(xx + 1, yy) + B(xx + 1, yy + 1) + B(xx + 1, yy + 2) + B(xx + 2, yy - 2) + B(xx + 2, yy - 1) + B(xx + 2, yy) + B(xx + 2, yy + 1) + B(xx + 2, yy + 2))) / 25
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub SharpenPicture(Rx1%, Ry1%, Rx2%, Ry2%)   'sharpen
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 + 1 To Rx2 - 1
For yy = Ry1 + 1 To Ry2 - 1
R(xx, yy) = R(xx, yy) + 0.5 * (R(xx, yy) - R(xx - 1, yy - 1))
G(xx, yy) = G(xx, yy) + 0.5 * (G(xx, yy) - G(xx - 1, yy - 1))
B(xx, yy) = B(xx, yy) + 0.5 * (B(xx, yy) - B(xx - 1, yy - 1))
            If R(xx, yy) > 255 Then R(xx, yy) = 255
            If R(xx, yy) < 0 Then R(xx, yy) = 0
            If G(xx, yy) > 255 Then G(xx, yy) = 255
            If G(xx, yy) < 0 Then G(xx, yy) = 0
            If B(xx, yy) > 255 Then B(xx, yy) = 255
            If B(xx, yy) < 0 Then B(xx, yy) = 0
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub DiffusePic(Rx1%, Ry1%, Rx2%, Ry2%, Diffuse%) 'diffuse
Dim tt%, tt1%
On Error Resume Next
tt = Diffuse * 10
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
tt1 = (Rnd * tt) - 2
R(xx, yy) = Abs(R(xx, yy) + tt1)
G(xx, yy) = Abs(G(xx, yy) + tt1)
B(xx, yy) = Abs(B(xx, yy) + tt1)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub ErodePic(Rx1%, Ry1%, Rx2%, Ry2%, Erode%) 'erode
On Error Resume Next
Pct = Erode * 8
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
R(xx, yy) = Abs(R(xx, yy) Xor Pct)
G(xx, yy) = Abs(G(xx, yy) Xor Pct)
B(xx, yy) = Abs(B(xx, yy) Xor Pct)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub BlowPic(Rx1%, Ry1%, Rx2%, Ry2%, Blow%) 'blow
On Error Resume Next
Pct = Blow
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
R(xx, yy) = Abs(R(xx, yy) Xor (R(xx, yy) / Pct))
G(xx, yy) = Abs(G(xx, yy) Xor (G(xx, yy) / Pct))
B(xx, yy) = Abs(B(xx, yy) Xor (B(xx, yy) / Pct))
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub AddNoise(Rx1%, Ry1%, Rx2%, Ry2%) 'addnoise
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 + 1 To Rx2 - 1
For yy = Ry1 + 1 To Ry2 - 1
R(xx, yy) = ((Rnd * R(xx, yy)) + R(xx, yy)) / 2
G(xx, yy) = ((Rnd * G(xx, yy)) + G(xx, yy)) / 2
B(xx, yy) = ((Rnd * B(xx, yy)) + B(xx, yy)) / 2
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub FogPic(Rx1%, Ry1%, Rx2%, Ry2%, Fog%) 'fog
Dim tt1%
On Error Resume Next
Pct = Fog
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For yy = Ry1 To Ry2 - 1
tt1 = (Rnd * Pct) - 2
For xx = Rx1 To Rx2 - 1
R(xx, yy) = Abs(R(xx, yy) + tt1)
G(xx, yy) = Abs(G(xx, yy) + tt1)
B(xx, yy) = Abs(B(xx, yy) + tt1)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next xx
FMain.PB1.Value = yy - Ry1
Next yy
EndProcess
End Sub

Public Sub FreezePic(Rx1%, Ry1%, Rx2%, Ry2%, Freeze!) 'freeze
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 + 1 To Rx2 - 1
For yy = Ry1 + 1 To Ry2 - 1
R(xx, yy) = Abs((R(xx, yy) - G(xx, yy) - B(xx, yy)) * Freeze)
G(xx, yy) = Abs((G(xx, yy) - B(xx, yy) - R(xx, yy)) * Freeze)
B(xx, yy) = Abs((B(xx, yy) - R(xx, yy) - G(xx, yy)) * Freeze)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub BnW(Rx1%, Ry1%, Rx2%, Ry2%, BW%) 'B & W
Dim BWColor&
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 + 1 To Rx2 - 1
For yy = Ry1 + 1 To Ry2 - 1
    If R(xx, yy) < BW And G(xx, yy) < BW And B(xx, yy) < BW Then
    BWColor = 0
    Else
    BWColor = &HFFFFFF
    End If
SetPixel FMain.Pic1.hdc, xx, yy, BWColor
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub Effect0(Rx1%, Ry1%, Rx2%, Ry2%, Eff%)
On Error Resume Next
Dim C&
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For yy = Ry1 To Ry2 - 1
For xx = Rx1 To Rx2 - 1
Select Case Eff
Case 0
    G(xx, yy) = (R(xx, yy) + G(xx, yy)) / 2
    R(xx, yy) = G(xx, yy)
    B(xx, yy) = B(xx, yy) * (Atn(G(xx, yy)) * 2)
Case 1
    G(xx, yy) = (R(xx, yy) + G(xx, yy)) / 2
    R(xx, yy) = G(xx, yy)
Case 2
    R(xx, yy) = (R(xx, yy) + B(xx, yy)) / 2
    B(xx, yy) = R(xx, yy)
Case 3
    G(xx, yy) = (G(xx, yy) + B(xx, yy)) / 2
    B(xx, yy) = G(xx, yy)
Case 4
    G(xx, yy) = (B(xx, yy) + G(xx, yy)) / 2
    B(xx, yy) = G(xx, yy)
    R(xx, yy) = R(xx, yy) * (Atn(G(xx, yy)) * 2)
Case 5
    B(xx, yy) = (B(xx, yy) + R(xx, yy)) / 2
    R(xx, yy) = B(xx, yy)
    G(xx, yy) = G(xx, yy) * (Atn(R(xx, yy)) * 2)
Case 6
    B(xx, yy) = Sin(B(xx, yy)) * B(xx, yy)
    R(xx, yy) = Sin(R(xx, yy)) * R(xx, yy)
    G(xx, yy) = Sin(G(xx, yy)) * G(xx, yy)
Case 7
    C = (R(xx, yy) + G(xx, yy) + B(xx, yy)) / 12
    B(xx, yy) = Abs(Not (G(xx, yy) + C))
    R(xx, yy) = Abs(Not (B(xx, yy) + C))
    G(xx, yy) = Abs(Not (R(xx, yy) + C))
Case 8
    B(xx, yy) = G(xx, yy)
    G(xx, yy) = R(xx, yy)
Case 9
    R(xx, yy) = R(xx, yy) / 2
    B(xx, yy) = G(xx, yy) / 2
    G(xx, yy) = R(xx, yy)
Case 10
    R(xx, yy) = R(xx, yy)
    B(xx, yy) = G(xx, yy) / 2
    G(xx, yy) = R(xx, yy) / 2
Case 11
    R(xx, yy) = R(xx, yy) + Abs(Sin(R(xx, yy)) * 64)
    G(xx, yy) = G(xx, yy) + Abs(Sin(G(xx, yy)) * 64)
    B(xx, yy) = B(xx, yy) + Abs(Sin(B(xx, yy)) * 64)
End Select
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next xx
FMain.PB1.Value = yy - Ry1
Next yy
EndProcess
End Sub

Public Sub Brown(Rx1%, Ry1%, Rx2%, Ry2%, Brown%) 'brown
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
R(xx, yy) = Abs(G(xx, yy) * B(xx, yy)) / Brown
G(xx, yy) = Abs(B(xx, yy) * R(xx, yy)) / 256
B(xx, yy) = Abs(R(xx, yy) * G(xx, yy)) / 256
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub Liquid(Rx1%, Ry1%, Rx2%, Ry2%) 'liquid
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
R(xx, yy) = ((G(xx, yy) - B(xx, yy)) ^ 2) / 125
G(xx, yy) = ((R(xx, yy) - B(xx, yy)) ^ 2) / 125
B(xx, yy) = ((R(xx, yy) - G(xx, yy)) ^ 2) / 125
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub Yellow(Rx1%, Ry1%, Rx2%, Ry2%) 'yellow
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
B(xx, yy) = ((G(xx, yy) - R(xx, yy)) ^ 2) / 125
R(xx, yy) = ((G(xx, yy) - B(xx, yy)) ^ 2) / 125
G(xx, yy) = ((B(xx, yy) + R(xx, yy)) ^ 2) / 125
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub Charcoal(Rx1%, Ry1%, Rx2%, Ry2%) 'charcoal
Dim tCol&
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
            R(xx, yy) = Abs(R(xx, yy) * (G(xx, yy) - B(xx, yy) + G(xx, yy) + R(xx, yy))) / 256
            G(xx, yy) = Abs(R(xx, yy) * (B(xx, yy) - G(xx, yy) + B(xx, yy) + R(xx, yy))) / 256
            B(xx, yy) = Abs(G(xx, yy) * (B(xx, yy) - G(xx, yy) + B(xx, yy) + R(xx, yy))) / 256
            tCol = RGB(R(xx, yy), G(xx, yy), B(xx, yy))
            R(xx, yy) = Abs(tCol Mod 256)
            G(xx, yy) = Abs((tCol \ 256) Mod 256)
            B(xx, yy) = Abs(tCol \ 256 \ 256)
            R(xx, yy) = (R(xx, yy) + G(xx, yy) + B(xx, yy)) / 3
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), R(xx, yy), R(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub DarkMoon(Rx1%, Ry1%, Rx2%, Ry2%) 'dark moon
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
R(xx, yy) = Abs(R(xx, yy) - 64)
G(xx, yy) = Abs(R(xx, yy) - 64)
B(xx, yy) = Abs(R(xx, yy) - 64)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub TotalEclipse(Rx1%, Ry1%, Rx2%, Ry2%) 'eclipse
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
R(xx, yy) = Abs(G(xx, yy) - 64)
G(xx, yy) = Abs(G(xx, yy) - 64)
B(xx, yy) = Abs(G(xx, yy) - 64)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub PurpleRain(Rx1%, Ry1%, Rx2%, Ry2%) 'purple
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
R(xx, yy) = Abs(G(xx, yy) + R(xx, yy) / 2)
G(xx, yy) = Abs(B(xx, yy) + G(xx, yy) / 2)
B(xx, yy) = Abs(R(xx, yy) + B(xx, yy) / 2)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub Spooky(Rx1%, Ry1%, Rx2%, Ry2%) 'Spooky
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
G(xx, yy) = Abs(R(xx, yy) + G(xx, yy) / 2)
B(xx, yy) = Abs(R(xx, yy) + B(xx, yy) / 2)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub UnReal(Rx1%, Ry1%, Rx2%, Ry2%) 'unreal
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
If (G(xx, yy) = 0) Or (B(xx, yy) = 0) Then
    G(xx, yy) = 1
    B(xx, yy) = 1
End If
        R(xx, yy) = Abs(Sin(Atn(G(xx, yy) / B(xx, yy))) * 125 + 20)
        G(xx, yy) = Abs(Sin(Atn(R(xx, yy) / B(xx, yy))) * 125 + 20)
        B(xx, yy) = Abs(Sin(Atn(R(xx, yy) / G(xx, yy))) * 125 + 20)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub Flame(Rx1%, Ry1%, Rx2%, Ry2%) 'flame
Dim C As Long
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
    C = (R(xx, yy) + G(xx, yy) + B(xx, yy)) / 3
        If R(xx, yy) > B(xx, yy) Then
            R(xx, yy) = Abs(R(xx, yy) + C)
            B(xx, yy) = Abs(B(xx, yy) - C)
        End If
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub Aquarel(Rx1%, Ry1%, Rx2%, Ry2%) 'aquarel
On Error Resume Next
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
If R(xx, yy) < 128 And G(xx, yy) < 128 And B(xx, yy) < 128 Then
R(xx, yy) = 2 * R(xx, yy): G(xx, yy) = 2 * G(xx, yy): B(xx, yy) = 2 * B(xx, yy)
Else
R(xx, yy) = R(xx, yy) / 2: G(xx, yy) = G(xx, yy) / 2: B(xx, yy) = B(xx, yy) / 2
End If
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub Blinds(Rx1, Ry1, Rx2, Ry2, Blinds%, Reverse As Boolean) 'hor blinds
Dim rt%
On Error Resume Next
BeginProcess
Pct = Blinds
FMain.PB1.Max = Ry2 - Ry1
If Reverse = False Then
rt = 0
Else
rt = Pct
End If
For yy = Ry1 To Ry2 - 1
For xx = Rx1 To Rx2 - 1
R(xx, yy) = R(xx, yy) - (rt * R(xx, yy) / Pct)
G(xx, yy) = G(xx, yy) - (rt * G(xx, yy) / Pct)
B(xx, yy) = B(xx, yy) - (rt * B(xx, yy) / Pct)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next xx
If Reverse = False Then
    rt = rt + 1
    If rt = Pct Then rt = 0
Else
    rt = rt - 1
    If rt = 0 Then rt = Pct
End If
FMain.PB1.Value = yy - Ry1
Next yy
EndProcess
End Sub

Public Sub Blinds2(Rx1, Ry1, Rx2, Ry2, Blinds%, Reverse As Boolean) 'vert blinds
Dim rt%
On Error Resume Next
BeginProcess
Pct = Blinds
FMain.PB1.Max = Rx2 - Rx1
If Reverse = False Then
rt = 0
Else
rt = Pct
End If
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
R(xx, yy) = R(xx, yy) - (rt * R(xx, yy) / Pct)
G(xx, yy) = G(xx, yy) - (rt * G(xx, yy) / Pct)
B(xx, yy) = B(xx, yy) - (rt * B(xx, yy) / Pct)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
If Reverse = False Then
    rt = rt + 1
    If rt = Pct Then rt = 0
Else
    rt = rt - 1
    If rt = 0 Then rt = Pct
End If
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub Blinds3(Rx1, Ry1, Rx2, Ry2, Blinds%) 'hor bump blinds
Dim rt%, Rtt As Boolean
On Error Resume Next
BeginProcess
Pct = Blinds
FMain.PB1.Max = Ry2 - Ry1
rt = 0
Rtt = False
For yy = Ry1 To Ry2 - 1
For xx = Rx1 To Rx2 - 1
R(xx, yy) = R(xx, yy) - (rt * R(xx, yy) / Pct)
G(xx, yy) = G(xx, yy) - (rt * G(xx, yy) / Pct)
B(xx, yy) = B(xx, yy) - (rt * B(xx, yy) / Pct)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next xx
    If Rtt = False Then
    rt = rt + 2
    Else
    rt = rt - 2
    End If
        If rt >= Pct Then Rtt = True
        If rt <= 0 Then Rtt = False
FMain.PB1.Value = yy - Ry1
Next yy
EndProcess
End Sub

Public Sub Blinds4(Rx1, Ry1, Rx2, Ry2, Blinds%) 'bump vert blinds
Dim rt%, Rtt As Boolean
On Error Resume Next
BeginProcess
Pct = Blinds
FMain.PB1.Max = Rx2 - Rx1
rt = 0
Rtt = False
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
R(xx, yy) = R(xx, yy) - (rt * R(xx, yy) / Pct)
G(xx, yy) = G(xx, yy) - (rt * G(xx, yy) / Pct)
B(xx, yy) = B(xx, yy) - (rt * B(xx, yy) / Pct)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
    If Rtt = False Then
    rt = rt + 2
    Else
    rt = rt - 2
    End If
        If rt >= Pct Then Rtt = True
        If rt <= 0 Then Rtt = False
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub HLines(Rx1, Ry1, Rx2, Ry2, Dist%, AB!, LCol&)
On Error Resume Next
Dim Lr&, Lg&, Lb&
Lr = LCol Mod 256&
Lg = ((LCol And &HFF00) / 256&) Mod 256&
Lb = (LCol And &HFF0000) / 65536
BeginProcess
AB = AB / 10
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1 Step Dist
R(xx, yy) = (R(xx, yy) * (1 - AB)) + (Lr * AB)
G(xx, yy) = (G(xx, yy) * (1 - AB)) + (Lg * AB)
B(xx, yy) = (B(xx, yy) * (1 - AB)) + (Lb * AB)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub VLines(Rx1, Ry1, Rx2, Ry2, Dist%, AB!, LCol&)
On Error Resume Next
Dim Lr&, Lg&, Lb&
Lr = LCol Mod 256&
Lg = ((LCol And &HFF00) / 256&) Mod 256&
Lb = (LCol And &HFF0000) / 65536
BeginProcess
AB = AB / 10
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1 Step Dist
For yy = Ry1 To Ry2 - 1
R(xx, yy) = (R(xx, yy) * (1 - AB)) + (Lr * AB)
G(xx, yy) = (G(xx, yy) * (1 - AB)) + (Lg * AB)
B(xx, yy) = (B(xx, yy) * (1 - AB)) + (Lb * AB)
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub Squares(Rx1%, Ry1%, Rx2%, Ry2%, Dist%, AB!, LCol&)
On Error Resume Next
If Lr = 0 And Lg = 0 And Lb = 0 Then Lr = 1 'not black!
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For yy = Ry1 To Ry2 - 1 Step Dist
FMain.TempMem.Line (Rx1, yy)-(Rx2, yy), LCol
Next yy
For xx = Rx1 To Rx2 - 1 Step Dist
FMain.TempMem.Line (xx, Ry1)-(xx, Ry2), LCol
Next xx
MixObject Rx1, Ry1, Rx2, Ry2, AB, LCol
EndProcess
End Sub

Public Sub AddBoxes(Rx1%, Ry1%, Rx2%, Ry2%, Dist%, AB!, LCol&) 'add boxes
On Error Resume Next
Dim ttt%
ttt = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = 0 To (FMain.Pic1.Width / (Dist * 2))
FMain.TempMem.Line (ttt, ttt)-(FMain.Pic1.Width - ttt, FMain.Pic1.Height - ttt), LCol, B
If FMain.Pic1.Width - ttt - ttt < Dist Then Exit For
ttt = ttt + Dist
Next xx
MixObject Rx1, Ry1, Rx2, Ry2, AB, LCol
EndProcess
End Sub

Public Sub AddCircles(Rx1%, Ry1%, Rx2%, Ry2%, Dist%, AB!, LCol&) 'add circles
On Error Resume Next
Dim ttt%
ttt = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = 0 To (FMain.Pic1.Width * 2) / Dist
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), ttt, LCol
If ttt > Int(Sqr(2 * ((FMain.Pic1.Width / 2) ^ 2))) Then Exit For
ttt = ttt + Dist
Next xx
MixObject Rx1, Ry1, Rx2, Ry2, AB, LCol
EndProcess
End Sub

Public Sub AddDiaRLines(Rx1%, Ry1%, Rx2%, Ry2%, Dist%, AB!, LCol&)
On Error Resume Next
Dim ttt%
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
For xx = 0 To (FMain.Pic1.Width / Dist) * 3
FMain.TempMem.Line (0, ttt)-(ttt, 0), LCol
ttt = ttt + Dist
Next xx
MixObject Rx1, Ry1, Rx2, Ry2, AB, LCol
EndProcess
End Sub

Public Sub AddDiaLLines(Rx1%, Ry1%, Rx2%, Ry2%, Dist%, AB!, LCol&)
On Error Resume Next
Dim ttt%
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
For xx = 0 To (FMain.Pic1.Width / Dist) * 3
FMain.TempMem.Line (0, FMain.Pic1.Height - ttt)-(FMain.Pic1.Width, (2 * FMain.Pic1.Height) - ttt), LCol
ttt = ttt + Dist
Next xx
MixObject Rx1, Ry1, Rx2, Ry2, AB, LCol
EndProcess
End Sub

Public Sub AddCrossLines(Rx1%, Ry1%, Rx2%, Ry2%, Dist%, AB!, LCol&)
On Error Resume Next
Dim ttt%
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
For xx = 0 To (FMain.Pic1.Width / Dist) * 3
FMain.TempMem.Line (0, ttt)-(ttt, 0), LCol
ttt = ttt + Dist
Next xx
ttt = 0
For xx = 0 To (FMain.Pic1.Width / Dist) * 3
FMain.TempMem.Line (0, FMain.Pic1.Height - ttt)-(FMain.Pic1.Width, (2 * FMain.Pic1.Height) - ttt), LCol
ttt = ttt + Dist
Next xx
MixObject Rx1, Ry1, Rx2, Ry2, AB, LCol
EndProcess
End Sub

Public Sub MixObject(Rx1%, Ry1%, Rx2%, Ry2%, AB!, LCol&) 'mix with object
On Error Resume Next
Dim Lr&, Lg&, Lb&
Lr = LCol Mod 256&
Lg = ((LCol And &HFF00) / 256&) Mod 256&
Lb = (LCol And &HFF0000) / 65536
AB = AB / 10
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1
    Color = GetPixel(FMain.TempMem.hdc, xx, yy)
    If Color <> 0 Then
    Lr = Color Mod 256&
    Lg = ((Color And &HFF00) / 256&) Mod 256&
    Lb = (Color And &HFF0000) / 65536
    R(xx, yy) = (R(xx, yy) * (1 - AB)) + (Lr * AB)
    G(xx, yy) = (G(xx, yy) * (1 - AB)) + (Lg * AB)
    B(xx, yy) = (B(xx, yy) * (1 - AB)) + (Lb * AB)
    SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
    End If
    Next yy
    FMain.PB1.Value = xx - Rx1
    Next xx
End Sub

Public Sub SinusLineX(Rx1%, Ry1%, Rx2%, Ry2%, AB!, Wave%, Ampl%, LCol&, Dist%, Eff%)
On Error Resume Next
Dim Degree As Single, k!
Dim Lr&, Lg&, Lb&
Lr = LCol Mod 256&
Lg = ((LCol And &HFF00) / 256&) Mod 256&
Lb = (LCol And &HFF0000) / 65536
AB = AB / 10
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
For xx = Rx1 To Rx2 - 1
For yy = Ry1 To Ry2 - 1 Step Dist
Degree = xx * (Wave) / 180 * 3.14
If Eff = 0 Then k = Cos(Degree) * Ampl
If Eff = 1 Then k = Abs(Cos(Degree) * Ampl)
If Eff = 2 Then k = -Abs(Cos(Degree) * Ampl)
Color = GetPixel(FMain.Pic1.hdc, xx, k + yy)
R(xx, yy) = Color Mod 256&
G(xx, yy) = ((Color And &HFF00) / 256&) Mod 256&
B(xx, yy) = (Color And &HFF0000) / 65536
R(xx, yy) = (R(xx, yy) * (1 - AB)) + (Lr * AB)
G(xx, yy) = (G(xx, yy) * (1 - AB)) + (Lg * AB)
B(xx, yy) = (B(xx, yy) * (1 - AB)) + (Lb * AB)
SetPixel FMain.Pic1.hdc, xx, k + yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub SinusLineY(Rx1%, Ry1%, Rx2%, Ry2%, AB!, Wave%, Ampl%, LCol&, Dist%, Eff%)
On Error Resume Next
Dim Degree As Single, k!
Dim Lr&, Lg&, Lb&
Lr = LCol Mod 256&
Lg = ((LCol And &HFF00) / 256&) Mod 256&
Lb = (LCol And &HFF0000) / 65536
AB = AB / 10
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
For yy = Ry1 To Ry2 - 1
For xx = Rx1 To Rx2 - 1 Step Dist
Degree = yy * (Wave) / 180 * 3.14
If Eff = 0 Then k = Cos(Degree) * Ampl
If Eff = 1 Then k = Abs(Cos(Degree) * Ampl)
If Eff = 2 Then k = -Abs(Cos(Degree) * Ampl)
Color = GetPixel(FMain.Pic1.hdc, k + xx, yy)
R(xx, yy) = Color Mod 256&
G(xx, yy) = ((Color And &HFF00) / 256&) Mod 256&
B(xx, yy) = (Color And &HFF0000) / 65536
R(xx, yy) = (R(xx, yy) * (1 - AB)) + (Lr * AB)
G(xx, yy) = (G(xx, yy) * (1 - AB)) + (Lg * AB)
B(xx, yy) = (B(xx, yy) * (1 - AB)) + (Lb * AB)
SetPixel FMain.Pic1.hdc, k + xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next xx
FMain.PB1.Value = xx - Rx1
Next yy
EndProcess
End Sub

Public Sub SBorder(Dist%, AB!, LCol&, Redu As Boolean)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
FMain.Tempmem2.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.Tempmem2.Picture = FMain.Pic1.Image
FMain.PB1.Max = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
For xx = 0 To Dist - 1
FMain.TempMem.Line (xx, xx)-(FMain.TempMem.Width - 1 - xx, FMain.TempMem.Height - 1 - xx), LCol, B
FMain.PB1.Value = xx
Next xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
If Redu = True Then
FMain.Pic1.PaintPicture FMain.Tempmem2, Dist, Dist, FMain.Pic1.Width - (2 * Dist), FMain.Pic1.Height - (2 * Dist)
End If
EndProcess
End Sub

Public Sub GBorder1(Dist%, AB!, LCol&, Scol&, Redu As Boolean)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, Ri, Gi, Bi
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist
Gi = (Gg1 - Gg) / Dist
Bi = (Gb1 - Gb) / Dist
FMain.Tempmem2.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.Tempmem2.Picture = FMain.Pic1.Image
FMain.PB1.Max = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
For xx = 0 To Dist - 1
FMain.TempMem.Line (xx, xx)-(FMain.TempMem.Width - 1 - xx, FMain.TempMem.Height - 1 - xx), RGB(Gr, Gg, Gb), B
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + Bi
FMain.PB1.Value = xx
Next xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
If Redu = True Then
FMain.Pic1.PaintPicture FMain.Tempmem2, Dist, Dist, FMain.Pic1.Width - (2 * Dist), FMain.Pic1.Height - (2 * Dist)
End If
EndProcess
End Sub

Public Sub GBorder2(Dist%, AB!, LCol&, Scol&, Redu As Boolean)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, Ri, Gi, Bi
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist * 2
Gi = (Gg1 - Gg) / Dist * 2
Bi = (Gb1 - Gb) / Dist * 2
FMain.Tempmem2.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.Tempmem2.Picture = FMain.Pic1.Image
FMain.PB1.Max = Dist / 2
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
BeginProcess
For xx = 0 To (Dist / 2)
FMain.TempMem.Line (xx, xx)-(FMain.TempMem.Width - 1 - xx, FMain.TempMem.Height - 1 - xx), RGB(Gr, Gg, Gb), B
FMain.TempMem.Line (Dist - xx, Dist - xx)-(FMain.TempMem.Width - 1 - Dist + xx, FMain.TempMem.Height - 1 - Dist + xx), RGB(Gr, Gg, Gb), B
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + Bi
FMain.PB1.Value = xx
Next xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
If Redu = True Then
FMain.Pic1.PaintPicture FMain.Tempmem2, Dist, Dist, FMain.Pic1.Width - (2 * Dist), FMain.Pic1.Height - (2 * Dist)
End If
EndProcess
End Sub

Public Sub CBorder(Dist%, AB!, LCol&)
On Error Resume Next
Dim Ra!, Rr%
If LCol = 0 Then LCol = &H10101
FMain.PB1.Max = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
FMain.TempMem.DrawWidth = 2
Ra = FMain.Pic1.Height / FMain.Pic1.Width
    If FMain.Pic1.Height < FMain.Pic1.Width Then
    Rr = (FMain.Pic1.Width / 2) + Dist
    Else
    Rr = (FMain.Pic1.Height / 2) + Dist
    End If
BeginProcess
For xx = 0 To Dist - 1
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Rr - xx, LCol, , , Ra
FMain.PB1.Value = xx
Next xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
FMain.TempMem.DrawWidth = 1
EndProcess
End Sub

Public Sub GCBorder1(Dist%, AB!, LCol&, Scol&)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, Ri, Gi, Bi
Dim Ra!, Rr%
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist
Gi = (Gg1 - Gg) / Dist
Bi = (Gb1 - Gb) / Dist
FMain.PB1.Max = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
FMain.TempMem.DrawWidth = 2
Ra = FMain.Pic1.Height / FMain.Pic1.Width
    If FMain.Pic1.Height < FMain.Pic1.Width Then
    Rr = (FMain.Pic1.Width / 2) + Dist
    Else
    Rr = (FMain.Pic1.Height / 2) + Dist
    End If
BeginProcess
For xx = 0 To Dist - 1
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Rr - xx, RGB(Gr, Gg, Gb), , , Ra
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + Bi
FMain.PB1.Value = xx
Next xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
FMain.TempMem.DrawWidth = 1
EndProcess
End Sub

Public Sub GCBorder2(Dist%, AB!, LCol&, Scol&)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, Ri, Gi, Bi
Dim Ra!, Rr%
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist * 2
Gi = (Gg1 - Gg) / Dist * 2
Bi = (Gb1 - Gb) / Dist * 2
FMain.PB1.Max = Dist / 2
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
FMain.TempMem.DrawWidth = 2
Ra = FMain.Pic1.Height / FMain.Pic1.Width
    If FMain.Pic1.Height < FMain.Pic1.Width Then
    Rr = (FMain.Pic1.Width / 2) + Dist
    Else
    Rr = (FMain.Pic1.Height / 2) + Dist
    End If
BeginProcess
For xx = 0 To Dist / 2
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Rr - xx, RGB(Gr, Gg, Gb), , , Ra
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Rr - Dist + xx, RGB(Gr, Gg, Gb), , , Ra
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + Bi
FMain.PB1.Value = xx
Next xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
FMain.TempMem.DrawWidth = 1
EndProcess
End Sub

Public Sub MixSolid(AB!, LCol&)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = LCol
BeginProcess
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub MixGradient1(AB!, LCol&, Scol&, CH%)
On Error Resume Next
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, H%, Ri, Gi, Bi, Dist%
If CH = 0 Then Dist = FMain.Pic1.Height
If CH = 1 Then Dist = FMain.Pic1.Width
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist
Gi = (Gg1 - Gg) / Dist
Bi = (Gb1 - Gb) / Dist
FMain.PB1.Max = Dist
BeginProcess
For xx = 0 To Dist - 1
If CH = 0 Then FMain.TempMem.Line (0, xx)-(FMain.TempMem.Width - 1, xx), RGB(Gr, Gg, Gb)
If CH = 1 Then FMain.TempMem.Line (xx, 0)-(xx, FMain.TempMem.Height - 1), RGB(Gr, Gg, Gb)
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + Bi
FMain.PB1.Value = xx
Next xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub MixGradient2(AB!, LCol&, Scol&, CH%)
On Error Resume Next
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, H%, Ri, Gi, Bi, Dist%
If CH = 0 Then Dist = FMain.Pic1.Height
If CH = 1 Then Dist = FMain.Pic1.Width
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist * 2
Gi = (Gg1 - Gg) / Dist * 2
Bi = (Gb1 - Gb) / Dist * 2
FMain.PB1.Max = Dist / 2
BeginProcess
For xx = 0 To Dist / 2
If CH = 0 Then
FMain.TempMem.Line (0, xx)-(FMain.TempMem.Width - 1, xx), RGB(Gr, Gg, Gb)
FMain.TempMem.Line (0, Dist - 1 - xx)-(FMain.TempMem.Width - 1, Dist - 1 - xx), RGB(Gr, Gg, Gb)
Else
FMain.TempMem.Line (xx, 0)-(xx, FMain.TempMem.Height - 1), RGB(Gr, Gg, Gb)
FMain.TempMem.Line (Dist - 1 - xx, 0)-(Dist - 1 - xx, FMain.TempMem.Height - 1), RGB(Gr, Gg, Gb)
End If
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + Bi
FMain.PB1.Value = xx
Next xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub MixBoxGradient1(AB!, LCol&, Scol&)
On Error Resume Next
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, H%, Ri, Gi, Bi, Dist%
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
If FMain.Pic1.Width < FMain.Pic1.Height Then
Dist = FMain.Pic1.Width / 2
Else
Dist = FMain.Pic1.Height / 2
End If
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist
Gi = (Gg1 - Gg) / Dist
Bi = (Gb1 - Gb) / Dist
FMain.PB1.Max = Dist
BeginProcess
For xx = 0 To Dist - 1
FMain.TempMem.Line (xx, xx)-(FMain.TempMem.Width - 1 - xx, FMain.TempMem.Height - 1 - xx), RGB(Gr, Gg, Gb), B
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + Bi
FMain.PB1.Value = xx
Next xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub MixBoxGradient2(AB!, LCol&, Scol&)
On Error Resume Next
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, H%, Ri, Gi, Bi, Dist%
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
If FMain.Pic1.Width < FMain.Pic1.Height Then
Dist = FMain.Pic1.Width / 2
Else
Dist = FMain.Pic1.Height / 2
End If
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
Ri = (Gr1 - Gr) / Dist * 2
Gi = (Gg1 - Gg) / Dist * 2
Bi = (Gb1 - Gb) / Dist * 2
FMain.PB1.Max = Dist / 2
BeginProcess
For xx = 0 To Dist / 2
FMain.TempMem.Line (xx, xx)-(FMain.TempMem.Width - 1 - xx, FMain.TempMem.Height - 1 - xx), RGB(Gr, Gg, Gb), B
FMain.TempMem.Line (Dist - xx, Dist - xx)-(FMain.TempMem.Width - 1 - Dist + xx, FMain.TempMem.Height - 1 - Dist + xx), RGB(Gr, Gg, Gb), B
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + Bi
FMain.PB1.Value = xx
Next xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub GCircle1(AB!, LCol&, Scol&)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, Ri, Gi, Bi, Dist%
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
    Dist = Sqr(((FMain.Pic1.Width / 2) ^ 2) + ((FMain.Pic1.Height / 2) ^ 2))
Ri = (Gr1 - Gr) / Dist
Gi = (Gg1 - Gg) / Dist
Bi = (Gb1 - Gb) / Dist
FMain.PB1.Max = Dist
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
FMain.TempMem.DrawWidth = 2
BeginProcess
For xx = 0 To Dist - 1
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Dist - xx, RGB(Gr, Gg, Gb) ', , , Ra
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + Bi
FMain.PB1.Value = xx
Next xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
FMain.TempMem.DrawWidth = 1
EndProcess
End Sub

Public Sub GCircle2(AB!, LCol&, Scol&)
On Error Resume Next
If LCol = 0 Then LCol = &H10101
If Scol = 0 Then Scol = &H10101
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, Ri, Gi, Bi, Dist%
Gr = LCol Mod 256&
Gg = ((LCol And &HFF00) / 256&) Mod 256&
Gb = (LCol And &HFF0000) / 65536
Gr1 = Scol Mod 256&
Gg1 = ((Scol And &HFF00) / 256&) Mod 256&
Gb1 = (Scol And &HFF0000) / 65536
    Dist = Sqr(((FMain.Pic1.Width / 2) ^ 2) + ((FMain.Pic1.Height / 2) ^ 2))
Ri = (Gr1 - Gr) / Dist * 2
Gi = (Gg1 - Gg) / Dist * 2
Bi = (Gb1 - Gb) / Dist * 2
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Set FMain.TempMem = Nothing
FMain.TempMem.BackColor = 0
FMain.TempMem.DrawWidth = 2
BeginProcess
FMain.PB1.Max = Dist / 2
For xx = 0 To Dist / 2
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), Dist - xx, RGB(Gr, Gg, Gb) ', , , Ra
FMain.TempMem.Circle (FMain.Pic1.Width / 2, FMain.Pic1.Height / 2), xx, RGB(Gr, Gg, Gb) ', , , Ra
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + Bi
FMain.PB1.Value = xx
Next xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
FMain.TempMem.DrawWidth = 1
EndProcess
End Sub

Public Sub MixPic(AB!, Op As Boolean)
On Error Resume Next
BeginProcess
Dim LCol&
LCol = 0
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.TempMem.BackColor = 0
If Op = True Then
FMain.TempMem.PaintPicture FPicture.Pic2, 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
Else
FMain.TempMem.PaintPicture FPicture.Pic2, (FMain.Pic1.Width - FPicture.Pic2.Width) / 2, (FMain.Pic1.Height - FPicture.Pic2.Height) / 2, FPicture.Pic2.Width, FPicture.Pic2.Height
End If
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub MixPattern(AB!)
BeginProcess
Dim LCol&
LCol = 0
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.TempMem.BackColor = 0
For xx = 0 To FMain.Pic1.Width / FPicture.Pic2.Width
For yy = 0 To FMain.Pic1.Height / FPicture.Pic2.Height
FMain.TempMem.PaintPicture FPicture.Pic2, xx * FPicture.Pic2.Width, yy * FPicture.Pic2.Height, FPicture.Pic2.Width, FPicture.Pic2.Height
Next yy, xx
MixObject 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, AB, LCol
EndProcess
End Sub

Public Sub Echo(ENr%, ERed%, EX%, EY%) 'echo picture
Dim EchoW&, EchoH&
Dim EchoLeft%, EchoTop%, Phase%
On Error Resume Next
FMain.Label2.Caption = "": DoEvents
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.TempMem.Picture = FMain.Pic1.Image
EchoW = FMain.TempMem.Width - 1
EchoH = FMain.TempMem.Height - 1
Phase = 0
FMain.PB1.Max = ENr - 1
For xx = 0 To ENr - 1
If FEcho.Check1.Value = 1 Then Phase = xx
EchoW = EchoW * (100 - ERed) / 100
EchoH = EchoH * (100 - ERed) / 100
EchoLeft = (FMain.TempMem.Width / 2) - (EchoW / 2) + ((Phase + 1) * EX)
EchoTop = (FMain.TempMem.Height / 2) - (EchoH / 2) + ((Phase + 1) * EY)
FMain.Pic1.PaintPicture FMain.TempMem, EchoLeft, EchoTop, EchoW, EchoH
FMain.PB1.Value = xx
Next xx
EndProcess
End Sub

Public Sub Mozaic(Rx1, Ry1, Rx2, Ry2, Br%) 'mozaic
Dim Br2%, MozaicColor&, qq%, pp%
On Error Resume Next
Br2 = Int(Br / 2)
FMain.PB1.Max = Rx2 - Rx1
BeginProcess
For xx = Rx1 To Rx2 Step Br
For yy = Ry1 To Ry2 Step Br
MozaicColor = GetPixel(FMain.Pic1.hdc, xx + Br2, yy + Br2)
    For qq = xx To xx + Br - 1
    For pp = yy To yy + Br - 1
    R(qq, pp) = MozaicColor
    Next pp, qq
Next yy, xx
For xx = Rx1 To Rx2
For yy = Ry1 To Ry2
SetPixel FMain.Pic1.hdc, xx, yy, R(xx, yy)
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub Mozaic2(Rx1, Ry1, Rx2, Ry2, Br%) 'mozaic2
Dim Br2%, MozaicColor&, qq%, pp%, R1&, G1&, B1&
On Error Resume Next
Br2 = Int(Br / 2)
BeginProcess
FMain.PB1.Max = Rx2 - Rx1
For xx = Rx1 To Rx2 Step Br
For yy = Ry1 To Ry2 Step Br
Color = GetPixel(FMain.Pic1.hdc, xx + Br2, yy + Br2)
R1 = Color Mod 256&
G1 = ((Color And &HFF00) / 256&) Mod 256&
B1 = (Color And &HFF0000) / 65536
    For qq = xx To xx + Br - 1
    For pp = yy To yy + Br - 1
    If qq = xx Or pp = yy Or qq = xx + Br - 1 Or pp = yy + Br - 1 Then
        R(qq, pp) = R(qq, pp) - ((Rnd * 10) - 5)
        If R(qq, pp) < 0 Then R(qq, pp) = 0
        G(qq, pp) = G(qq, pp) - ((Rnd * 10) - 5)
        If G(qq, pp) < 0 Then G(qq, pp) = 0
        B(qq, pp) = B(qq, pp) - ((Rnd * 10) - 5)
        If B(qq, pp) < 0 Then B(qq, pp) = 0
    Else
    R(qq, pp) = R1
    G(qq, pp) = G1
    B(qq, pp) = B1
    End If
    Next pp, qq
Next yy, xx
For xx = Rx1 To Rx2
For yy = Ry1 To Ry2
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy), G(xx, yy), B(xx, yy))
Next yy
FMain.PB1.Value = xx - Rx1
Next xx
EndProcess
End Sub

Public Sub EffectX(Strength%, Wave As Single, Eff%) 'wave x
On Error Resume Next
Dim Degree As Single, k!
ReDim LineColor(FMain.Pic1.Width)
BeginProcess
Wave = Wave / 10
FMain.PB1.Max = FMain.Pic1.Height - 1
For yy = 0 To FMain.Pic1.Height - 1
Degree = (yy * Wave / 180 * 3.14)
If Eff = 0 Then k = (Cos(Degree) * Strength)
If Eff = 1 Then k = Abs(Cos(Degree) * Strength)
GetColorsX
For xx = 0 To FMain.Pic1.ScaleWidth
If k < 0 Then
SetPixel FMain.Pic1.hdc, xx, yy, LineColor(FMain.Pic1.ScaleWidth - 1)
Else
SetPixel FMain.Pic1.hdc, xx, yy, LineColor(0)
End If
Next xx
For xx = 0 To FMain.Pic1.ScaleWidth - 1
SetPixel FMain.Pic1.hdc, xx + k, yy, LineColor(xx)
Next xx
FMain.PB1.Value = yy
Next yy
EndProcess
End Sub

Public Sub EffectY(Strength%, Wave As Single, Eff%) 'wave y
On Error Resume Next
Dim k!
ReDim LineColor(FMain.Pic1.Height)
BeginProcess
Wave = Wave / 10
FMain.PB1.Max = FMain.Pic1.Width - 1
For xx = 0 To FMain.Pic1.Width - 1
If Eff = 0 Then k = (Cos(xx * Wave / 180 * 3.14) * Strength)
If Eff = 1 Then k = Abs(Cos(xx * Wave / 180 * 3.14) * Strength)
GetColorsY
For yy = 0 To FMain.Pic1.ScaleHeight
If k < 0 Then
SetPixel FMain.Pic1.hdc, xx, yy, LineColor(OB.ScaleHeight - 1)
Else
SetPixel FMain.Pic1.hdc, xx, yy, LineColor(0)
End If
Next yy
For yy = 0 To FMain.Pic1.ScaleHeight - 1
SetPixel FMain.Pic1.hdc, xx, yy + k, LineColor(yy)
Next yy
FMain.PB1.Value = xx
Next xx
EndProcess
End Sub

Private Sub GetColorsX()
Dim tt%
For tt = 0 To FMain.Pic1.Width - 1
LineColor(tt) = GetPixel(FMain.Pic1.hdc, tt, yy)
Next tt
End Sub

Private Sub GetColorsY()
Dim tt%
For tt = 0 To FMain.Pic1.Height - 1
LineColor(tt) = GetPixel(FMain.Pic1.hdc, xx, tt)
Next tt
End Sub

Public Sub KillColXGrad1(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border left 1
kk = 1
On Error Resume Next
FMain.PB1.Max = FMain.Pic1.Width / 4
For xx = Rx1 To Rx2 / 4
kk = (1 - (xx / FMain.Pic1.ScaleWidth) * 4)
For yy = Ry1 To Ry2 - 1
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) - (R(xx, yy) * kk), G(xx, yy) - (G(xx, yy) * kk), B(xx, yy) - (B(xx, yy) * kk))
Next yy
FMain.PB1.Value = xx
Next xx
EndProcess
End Sub

Public Sub KillColXGrad2(Rx1%, Ry1%, Rx2%, Ry2%)  'grad border left 2
kk = 1
On Error Resume Next
FMain.PB1.Max = FMain.Pic1.Width / 2
For xx = Rx1 To Rx2 / 2
kk = (1 - (xx / FMain.Pic1.ScaleWidth) * 2)
For yy = Ry1 To Ry2 - 1
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) - (R(xx, yy) * kk), G(xx, yy) - (G(xx, yy) * kk), B(xx, yy) - (B(xx, yy) * kk))
Next yy
FMain.PB1.Value = xx
Next xx
EndProcess
End Sub

Public Sub KillColXGrad3(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border left 3
kk = 1
On Error Resume Next
FMain.PB1.Max = FMain.Pic1.Width - 1
For xx = Rx1 To Rx2 - 1
kk = 1 - (xx / FMain.Pic1.ScaleWidth)
For yy = Ry1 To Ry2 - 1
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) - (R(xx, yy) * kk), G(xx, yy) - (G(xx, yy) * kk), B(xx, yy) - (B(xx, yy) * kk))
Next yy
FMain.PB1.Value = xx
Next xx
EndProcess
End Sub

Public Sub KillColXGradRev1(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border right 1
On Error Resume Next
FMain.PB1.Max = Rx2 - 1 - (Rx2 / 4 * 3)
For xx = Rx2 / 4 * 3 To Rx2 - 1
kk = (xx - (Rx2 / 4 * 3)) / (Rx2 / 4)
For yy = Ry1 To Ry2 - 1
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) - (R(xx, yy) * kk), G(xx, yy) - (G(xx, yy) * kk), B(xx, yy) - (B(xx, yy) * kk))
Next yy
FMain.PB1.Value = xx - (Rx2 / 4 * 3)
Next xx
EndProcess
End Sub

Public Sub KillColXGradRev2(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border right 2
On Error Resume Next
FMain.PB1.Max = Rx2 - 1 - (Rx2 / 2)
For xx = Rx2 / 2 To Rx2 - 1
kk = (xx - (Rx2 / 2)) / (Rx2 / 2)
For yy = Ry1 To Ry2 - 1
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) - (R(xx, yy) * kk), G(xx, yy) - (G(xx, yy) * kk), B(xx, yy) - (B(xx, yy) * kk))
Next yy
FMain.PB1.Value = xx - (Rx2 / 2)
Next xx
EndProcess
End Sub

Public Sub KillColXGradRev3(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border right 3
On Error Resume Next
FMain.PB1.Max = Rx2
For xx = Rx1 To Rx2 - 1
kk = xx / Rx2
For yy = Ry1 To Ry2 - 1
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) - (R(xx, yy) * kk), G(xx, yy) - (G(xx, yy) * kk), B(xx, yy) - (B(xx, yy) * kk))
Next yy
FMain.PB1.Value = xx
Next xx
EndProcess
End Sub

Public Sub KillColYGrad1(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border top 1
On Error Resume Next
FMain.PB1.Max = Ry2 / 4
For yy = Ry1 To Ry2 / 4
kk = (1 - (yy / Ry2) * 4) '/ 1.4
For xx = Rx1 To Rx2 - 1
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) - (R(xx, yy) * kk), G(xx, yy) - (G(xx, yy) * kk), B(xx, yy) - (B(xx, yy) * kk))
Next xx
FMain.PB1.Value = yy
Next yy
EndProcess
End Sub

Public Sub KillColYGrad2(Rx1%, Ry1%, Rx2%, Ry2%) 'grad border top 2
On Error Resume Next
FMain.PB1.Max = Ry2 / 2
For yy = Ry1 To Ry2 / 2
kk = (1 - (yy / Ry2) * 2)
For xx = Rx1 To Rx2 - 1
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) - (R(xx, yy) * kk), G(xx, yy) - (G(xx, yy) * kk), B(xx, yy) - (B(xx, yy) * kk))
Next xx
FMain.PB1.Value = yy
Next yy
EndProcess
End Sub

Public Sub KillColYGrad3(Rx1%, Ry1%, Rx2%, Ry2%)  'grad top border 3
On Error Resume Next
FMain.PB1.Max = Ry2
For yy = Ry1 To Ry2 - 1
kk = 1 - (yy / Ry2)
For xx = Rx1 To Rx2 - 1
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) - (R(xx, yy) * kk), G(xx, yy) - (G(xx, yy) * kk), B(xx, yy) - (B(xx, yy) * kk))
Next xx
FMain.PB1.Value = yy
Next yy
EndProcess
End Sub

Public Sub KillColYGradRev1(Rx1%, Ry1%, Rx2%, Ry2%)  'grad bottom border 1
On Error Resume Next
FMain.PB1.Max = Ry2 - 1 - (Ry2 / 4 * 3)
For yy = ((Ry2) / 4) * 3 To Ry2 - 1
kk = (yy - (Ry2 / 4 * 3)) / (Ry2 / 4)
For xx = Rx1 To Rx2 - 1
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) - (R(xx, yy) * kk), G(xx, yy) - (G(xx, yy) * kk), B(xx, yy) - (B(xx, yy) * kk))
Next xx
FMain.PB1.Value = yy - (Ry2 / 4 * 3)
Next yy
EndProcess
End Sub

Public Sub KillColYGradRev2(Rx1%, Ry1%, Rx2%, Ry2%)  'grad bottom border 2
On Error Resume Next
FMain.PB1.Max = Ry2 - 1 - (Ry2 / 2)
For yy = ((Ry2) / 2) To Ry2 - 1
kk = (yy - (Ry2 / 2)) / (Ry2 / 2)
For xx = Rx1 To Rx2 - 1
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) - (R(xx, yy) * kk), G(xx, yy) - (G(xx, yy) * kk), B(xx, yy) - (B(xx, yy) * kk))
Next xx
FMain.PB1.Value = yy - (Ry2 / 2)
Next yy
EndProcess
End Sub

Public Sub KillColYGradRev3(Rx1%, Ry1%, Rx2%, Ry2%)  'grad bottom border 3
On Error Resume Next
FMain.PB1.Max = Ry2 - 1
For yy = Ry1 To Ry2 - 1
kk = yy / Ry2
For xx = Rx1 To Rx2 - 1
SetPixel FMain.Pic1.hdc, xx, yy, RGB(R(xx, yy) - (R(xx, yy) * kk), G(xx, yy) - (G(xx, yy) * kk), B(xx, yy) - (B(xx, yy) * kk))
Next xx
FMain.PB1.Value = yy
Next yy
EndProcess
End Sub

Public Sub Tile(XTile%, YTile%)
On Error Resume Next
Dim TileX%, TileY%
FMain.PB1.Max = XTile - 1
TileX = Int(FMain.Pic1.Width / XTile)
TileY = Int(FMain.Pic1.Height / YTile)
FMain.Label2.Caption = "": DoEvents
FMain.TempMem.Move 0, 0, FMain.Pic1.Width, FMain.Pic1.Height
FMain.TempMem.Picture = FMain.Pic1.Image
For xx = 0 To XTile - 1
For yy = 0 To YTile - 1
FMain.Pic1.PaintPicture FMain.TempMem, xx * TileX, yy * TileY, TileX, TileY
Next yy
FMain.PB1.Value = xx
Next xx
EndProcess
End Sub
