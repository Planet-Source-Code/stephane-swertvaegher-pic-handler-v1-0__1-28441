Attribute VB_Name = "PHMod1"
Public R(), G(), B(), Tim%, PicFileName$, Temp$, FTitle$
Public xx%, yy%, Xcor0%, Xcor1%, Ycor0%, Ycor1%, XXX1%, XXX2%, YYY1%, YYY2%
Public Col%, Mix%
Public PicMem As Picture, Im As Picture
Public PicMem0(4) As Picture, MemCount%
Public OrWidth%, OrHeight%, Factor!
Public Scol(15)
Public Enum T3dFill
T3dF0
T3dF1
End Enum

Public Enum Borderstyle
T3dRaiseRaise
T3dRaiseInset
T3dInsetRaise
T3dInsetInset
T3dNone
End Enum
'API for translating system colors to 'normal' colors
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function EnumFonts Lib "gdi32" Alias "EnumFontsA" (ByVal hdc As Long, ByVal lpsz As String, ByVal lpFontEnumProc As Long, ByVal lParam As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Const LF_FACESIZE = 32
Public Const LOGPIXELSY = 90
Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lsngStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lsngPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE - 1) As Byte
End Type


Function EnumFontProc(ByVal lplf As Long, ByVal lptm As Long, ByVal dwType As Long, ByVal lpData As Long) As Long
    Dim LF As LOGFONT, FontName As String, ZeroPos As Long
    CopyMemory LF, ByVal lplf, LenB(LF)
    FontName = StrConv(LF.lfFaceName, vbUnicode)
    ZeroPos = InStr(1, FontName, Chr$(0))
    If ZeroPos > 0 Then FontName = Left$(FontName, ZeroPos - 1)
    FText.Combo1.AddItem FontName
    EnumFontProc = 1
End Function
   
Public Function T3D(Obj0 As Object, Obj As Object, Bev%, Optional Style3D As Borderstyle, Optional T3dFilled As T3dFill)
Dim R%, G%, B%, R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%
Dim FC&, T3Dxx%, SM%
On Error Resume Next

'global things
SM = Obj0.ScaleMode 'save scalemode
Obj0.ScaleMode = 3 'pixel
Obj.Borderstyle = 0 'no border
If IsMissing(Style3D) Then Style3D = 0
If Style3D > 4 Then Style3D = 3

'get formcolor
FC = Obj0.BackColor
'in case formcolor = systemcolor --> call the function RealColor
FC = RealColor(FC)
' convert to RGB
R = FC And &HFF
G = Int((FC And &HFF00&) / 256)
B = Int((FC And &HFF0000) / 65536)
'-------------------
If Style3D = 0 Then 'RaiseRaise
    R1 = R + 64
    If R1 > 255 Then R1 = 255
    R2 = R - 64
    If R2 < 0 Then R2 = 0
    R3 = R1
    R4 = R2
    G1 = G + 64
    If G1 > 255 Then G1 = 255
    G2 = G - 64
    If G2 < 0 Then G2 = 0
    G3 = G1
    G4 = G2
    B1 = B + 64
    If B1 > 255 Then B1 = 255
    B2 = B - 64
    If B2 < 0 Then B2 = 0
    B3 = B1
    B4 = B2
End If
'-------------------
If Style3D = 1 Then 'RaiseInset
    R1 = R + 64
    If R1 > 255 Then R1 = 255
    R2 = R - 64
    If R2 < 0 Then R2 = 0
    R4 = R1
    R3 = R2
    G1 = G + 64
    If G1 > 255 Then G1 = 255
    G2 = G - 64
    If G2 < 0 Then G2 = 0
    G4 = G1
    G3 = G2
    B1 = B + 64
    If B1 > 255 Then B1 = 255
    B2 = B - 64
    If B2 < 0 Then B2 = 0
    B4 = B1
    B3 = B2
End If
If Style3D = 2 Then 'InsetRaise
    R2 = R + 64
    If R2 > 255 Then R2 = 255
    R1 = R - 64
    If R1 < 0 Then R1 = 0
    R4 = R1
    R3 = R2
    G2 = G + 64
    If G2 > 255 Then G2 = 255
    G1 = G - 64
    If G1 < 0 Then G1 = 0
    G4 = G1
    G3 = G2
    B2 = B + 64
    If B2 > 255 Then B2 = 255
    B1 = B - 64
    If B1 < 0 Then B1 = 0
    B4 = B1
    B3 = B2
End If
If Style3D = 3 Then 'InsetInset
    R2 = R + 64
    If R2 > 255 Then R2 = 255
    R1 = R - 64
    If R1 < 0 Then R1 = 0
    R3 = R1
    R4 = R2
    G2 = G + 64
    If G2 > 255 Then G2 = 255
    G1 = G - 64
    If G1 < 0 Then G1 = 0
    G3 = G1
    G4 = G2
    B2 = B + 64
    If B2 > 255 Then B2 = 255
    B1 = B - 64
    If B1 < 0 Then B1 = 0
    B3 = B1
    B4 = B2
End If
If Style3D = 4 Then 'No Border
R1 = R: R2 = R: R3 = R: R4 = R
G1 = G: G2 = G: G3 = G: G4 = G
B1 = B: B2 = B: B3 = B: B4 = B
End If
Bev = Bev + 1
T3Dxx = Bev 'just in case Filled = 1

'Outer
If IsMissing(T3dFilled) Or T3dFilled = 0 Then
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Else
For Bev = T3Dxx To 1 Step -1 'in case T3DF1 (filled)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Next Bev
End If
'Inner
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left - 1, Obj.Top + Obj.Height + 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top - 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left + Obj.Width + 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)
    Obj0.Line (Obj.Left - 1, Obj.Top + Obj.Height + 1)-(Obj.Left + Obj.Width + 2, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)

Obj0.ScaleMode = SM 'restore original scalemode
End Function
  
  ' if System Color then translate to 'normal color'
  ' else, do nothing
  Public Function RealColor(ByVal Color As OLE_COLOR) As Long
     Dim Col As Long
     Col = TranslateColor(Color, 0, RealColor)
  End Function

Public Sub SetScrollBars()
With FMain
.Pic1.Move 0, 0
.VS1.Value = 0
.HS1.Value = 0
.VS1.Enabled = False
.HS1.Enabled = False
If .Pic1.Width > .PicX.Width - .VS1.Width Then .HS1.Enabled = True
If .Pic1.Height > .PicX.Height - .HS1.Height Then .VS1.Enabled = True
    If .VS1.Enabled = True Then
    .VS1.Max = .PicX.Height - .HS1.Height - .Pic1.Height
    .VS1.LargeChange = .Pic1.Height / 10
    End If
    If .HS1.Enabled = True Then
    .HS1.Max = .PicX.Width - .VS1.Width - .Pic1.Width
    .HS1.LargeChange = .Pic1.Width / 10
    End If
End With
End Sub

Public Sub SetPicInfo()
With FMain
.Label1.Caption = " Picture information:" & vbCr & vbCr
.Label1.Caption = .Label1.Caption & "  Original name: " & PicFileName & vbCr
.Label1.Caption = .Label1.Caption & "  Picture width: " & .Pic1.Width & vbCr
.Label1.Caption = .Label1.Caption & "  Picture height: " & .Pic1.Height & vbCr
ReDim R(.Pic1.Width, .Pic1.Height)
ReDim G(.Pic1.Width, .Pic1.Height)
ReDim B(.Pic1.Width, .Pic1.Height)
OrHeight = .Pic1.Height
OrWidth = .Pic1.Width
End With
End Sub

Public Sub SaveRedo()
For xx = 4 To 1 Step -1
FMain.TempMem = PicMem0(xx - 1)
Set PicMem0(xx) = FMain.TempMem.Image
Next xx
FMain.TempMem.Picture = FMain.Pic1.Image
Set PicMem0(0) = FMain.TempMem.Image
MemCount = MemCount + 1
If MemCount > 5 Then MemCount = 5
ShowMem
End Sub

Public Sub Redo()
FMain.Pic1 = PicMem0(0)
For xx = 0 To 3
Set PicMem0(xx) = PicMem0(xx + 1)
Next xx
Set PicMem0(4) = Nothing
MemCount = MemCount - 1
If MemCount = 0 Then MemCount = 0
ShowMem
SetPicInfo
End Sub
Public Sub ShowMem()
FMain.Toolbar1.Buttons(1).Enabled = False
For xx = 0 To 4
FMain.Image1(xx) = PicMem0(xx)
Next xx
If MemCount > 0 Then
FMain.Toolbar1.Buttons(1).Enabled = True
End If
End Sub

Public Sub ClearMem()
For xx = 0 To 4
Set PicMem0(xx) = Nothing
Next xx
ShowMem
End Sub

Public Sub SetCoordinates()
FMain.Label4.Caption = "Selection" & vbCr & Format(Xcor0, "000") & " X " & Format(Xcor1, "000") & " - " & Format(Ycor0, "000") & " X " & Format(Ycor1, "000")
Xcor1 = Xcor0 + Xcor1
Ycor1 = Ycor0 + Ycor1
End Sub
