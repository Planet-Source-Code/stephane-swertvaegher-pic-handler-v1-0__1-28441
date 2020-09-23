VERSION 5.00
Begin VB.Form FPicture 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   466
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Mix with pattern"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1365
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   5130
      Width           =   3795
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Index           =   1
         Left            =   1395
         Max             =   10
         Min             =   1
         TabIndex        =   15
         Top             =   360
         Value           =   1
         Width           =   1635
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alpha blend"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   1
         Left            =   3105
         TabIndex        =   16
         Top             =   315
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mix with picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1365
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   5130
      Width           =   3795
      Begin VB.OptionButton Option1 
         Caption         =   "Mix ""as is"""
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   990
         Width           =   2130
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mix fit to screen"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   720
         Value           =   -1  'True
         Width           =   2130
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Index           =   0
         Left            =   1395
         Max             =   10
         Min             =   1
         TabIndex        =   9
         Top             =   360
         Value           =   1
         Width           =   1635
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   0
         Left            =   3105
         TabIndex        =   11
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alpha blend"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   315
         Width           =   1140
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   90
      TabIndex        =   7
      Top             =   135
      Width           =   1725
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   45
      TabIndex        =   6
      Top             =   585
      Width           =   1770
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   3195
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   4
      Top             =   3105
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   1845
      TabIndex        =   3
      Top             =   135
      Width           =   2130
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show me"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   6615
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply !"
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      Top             =   6615
      Width           =   870
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3105
      TabIndex        =   0
      Top             =   6615
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Picture: 000 X 000"
      Height          =   195
      Index           =   0
      Left            =   1215
      TabIndex        =   5
      Top             =   4770
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   1260
      Stretch         =   -1  'True
      Top             =   3150
      Width           =   1500
   End
End
Attribute VB_Name = "FPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click() 'apply
FMain.Pic1.Picture = PicMem
SaveRedo
FMain.Pic1 = Im
FPicture.Hide
End Sub

Private Sub Command2_Click() 'cancel
FMain.Pic1.Picture = PicMem
FPicture.Hide
End Sub

Private Sub Command3_Click() 'show me
FMain.Pic1.Picture = PicMem
Screen.MousePointer = 11
Select Case Mix
Case 0 'mix picture
MixPic HScroll1(0).Value, Option1(0).Value
Case 1 'mix pattern
MixPattern HScroll1(1).Value
End Select
Command1.Enabled = True
Set Im = FMain.Pic1.Image
Screen.MousePointer = 1
End Sub

Private Sub Dir1_Change()
Command3.Enabled = False
On Error Resume Next
File1.Path = Dir1.Path
File1.Selected(0) = True
For xx = 0 To File1.ListCount - 1
If File1.Selected(xx) = True Then Command3.Enabled = True
Next xx
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
On Error Resume Next
    Pic2.Picture = LoadPicture(File1.Path & "\" & File1.List(File1.ListIndex))
Image1.Picture = LoadPicture(File1.Path & "\" & File1.List(File1.ListIndex))
DimensionImage Pic2.Width, Pic2.Height
End Sub

Private Sub Form_Activate()
On Error Resume Next
Command3.Enabled = False
Command1.Enabled = False
For xx = 0 To 9
Frame1(xx).Visible = False
Next xx
Drive1.Enabled = False
Dir1.Enabled = False
If Mix = 0 Then
Drive1.Enabled = True
Dir1.Enabled = True
End If
If Mix = 1 Then
Dir1.Path = App.Path & "\Patterns"
End If
FPicture.Move FMain.Left, FMain.Top + 330, 4140, 7365
File1.Pattern = "*.bmp;*.gif;*.jpg;*.wmf"
File1.Selected(0) = True
For xx = 0 To File1.ListCount - 1
If File1.Selected(xx) = True Then Command3.Enabled = True
Next xx
HScroll1(0).Value = 5
HScroll1(1).Value = 5
Set PicMem = FMain.Pic1.Image
Frame1(Mix).Visible = True
End Sub

Private Sub DimensionImage(L%, B%)
Dim T As Single, newL%, newH%
Label1(0).Caption = "Picture: " & Format(L, "000") & " X " & Format(B, "000")
newL = L: newH = B
T = 1
Do While newL > 100 Or newH > 100
newL = Int(L / T)
newH = Int(B / T)
T = T + 0.1
Loop
Image1.Width = newL
Image1.Height = newH
Image1.Move ((100 - newL) / 2) + 88, ((100 - newH) / 2) + 210
End Sub

Private Sub Form_Load()
Image1.Move 88, 210
T3D FPicture, Image1, 5, T3dRaiseInset
Drive1.Drive = "C:\"
Dir1.Path = App.Path
File1.Path = App.Path
Option1(0).Value = True
End Sub

Private Sub HScroll1_Change(Index As Integer)
Label2(Index).Caption = Format(HScroll1(Index).Value / 10, "0.0")
End Sub
