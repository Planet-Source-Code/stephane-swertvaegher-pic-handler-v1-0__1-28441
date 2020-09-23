VERSION 5.00
Begin VB.Form FAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About the programmer..."
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4335
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1395
      Left            =   90
      Picture         =   "FAbout.frx":0000
      ScaleHeight     =   1335
      ScaleWidth      =   990
      TabIndex        =   2
      Top             =   135
      Width           =   1050
   End
   Begin VB.PictureBox FontPic2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   45
      Picture         =   "FAbout.frx":099A
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   767
      TabIndex        =   1
      Top             =   1845
      Visible         =   0   'False
      Width           =   11505
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   45
      Top             =   2745
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   1590
      Left            =   45
      ScaleHeight     =   106
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   277
      TabIndex        =   0
      Top             =   1620
      Width           =   4155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   1500
      Left            =   1350
      TabIndex        =   3
      Top             =   135
      Width           =   5325
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
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
      Height          =   1320
      Left            =   1215
      TabIndex        =   4
      Top             =   180
      Width           =   5460
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SY%(499)
Dim Phrase3$, xpos%, CH%, ypos%, tt%
Dim Maxx%, Maxx2%

Private Sub Form_Activate()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Phrase3 = "*" & FTitle & "*"
Phrase3 = UCase(Phrase3)
Open App.Path & "\wave" For Input As #1
For xx = 0 To 499
Input #1, SY(xx)
Next xx
Close #1
tt = 40
Label2.Caption = "Stephan Swertvaegher" & vbCr
Label2.Caption = Label2.Caption & "     Age: 41" & vbCr
Label2.Caption = Label2.Caption & "     Sex: Male" & vbCr
Label2.Caption = Label2.Caption & "     Country: Belgium, Europe" & vbCr
Label2.Caption = Label2.Caption & "     Status: married" & vbCr
Label2.Caption = Label2.Caption & "     Programming since 1984"

Label3.Caption = Label2.Caption
Label3.Move Label2.Left + 1, Label2.Top + 1, Label2.Width, Label2.Height
End Sub

Private Sub Timer1_Timer()
For xx = 1 To Len(Phrase3)
BitBlt Pic1.hdc, 12 + (xx * 13), SY(xx + tt - 1), 13, 22, FontPic2.hdc, 0, 0, SRCCOPY
Next xx
For xx = 1 To Len(Phrase3)
CH = Asc(Mid(Phrase3, xx, 1))
CH = CH - 32

BitBlt Pic1.hdc, 12 + (xx * 13), SY(xx + tt), 13, 22, FontPic2.hdc, CH * 13, 0, vbSrcCopy

Next xx
tt = tt + 1
If tt > 310 Then tt = 1
Pic1.Refresh
End Sub
