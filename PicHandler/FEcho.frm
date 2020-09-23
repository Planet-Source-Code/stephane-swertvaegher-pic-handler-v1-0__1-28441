VERSION 5.00
Begin VB.Form FEcho 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Echo"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2490
      Left            =   90
      TabIndex        =   3
      Top             =   180
      Width           =   2895
      Begin VB.CommandButton Command5 
         Caption         =   "Default"
         Height          =   240
         Left            =   1575
         TabIndex        =   9
         Top             =   315
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Phasing off"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   315
         Width           =   1185
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   375
         Left            =   2295
         Max             =   1
         Min             =   49
         TabIndex        =   7
         Top             =   1080
         Value           =   1
         Width           =   240
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   375
         Left            =   2295
         Max             =   -100
         Min             =   100
         TabIndex        =   6
         Top             =   1980
         Width           =   240
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   375
         Left            =   2295
         Max             =   -100
         Min             =   100
         TabIndex        =   5
         Top             =   1530
         Width           =   240
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Left            =   2295
         Max             =   1
         Min             =   50
         TabIndex        =   4
         Top             =   630
         Value           =   1
         Width           =   240
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Left            =   1710
         TabIndex        =   17
         Top             =   1170
         Width           =   555
      End
      Begin VB.Label Lab3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reduced"
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
         Left            =   90
         TabIndex        =   16
         Top             =   1170
         Width           =   1545
      End
      Begin VB.Label Lab7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vertical center"
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
         Left            =   90
         TabIndex        =   15
         Top             =   2070
         Width           =   1545
      End
      Begin VB.Label Lab6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Horizontal center"
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
         Left            =   90
         TabIndex        =   14
         Top             =   1620
         Width           =   1545
      End
      Begin VB.Label Lab5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Number"
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
         Left            =   90
         TabIndex        =   13
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Left            =   1710
         TabIndex        =   12
         Top             =   2070
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Left            =   1710
         TabIndex        =   11
         Top             =   1620
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Left            =   1710
         TabIndex        =   10
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show me"
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   2835
      Width           =   825
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply!"
      Height          =   330
      Left            =   1035
      TabIndex        =   1
      Top             =   2835
      Width           =   825
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2160
      TabIndex        =   0
      Top             =   2835
      Width           =   825
   End
End
Attribute VB_Name = "FEcho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'show me
FMain.Pic1.Picture = PicMem
Echo VScroll1.Value, VScroll2.Value, VScroll3.Value, VScroll4.Value
Command2.Enabled = True
Set Im = FMain.Pic1.Image
End Sub

Private Sub Command2_Click() 'apply
FMain.Pic1.Picture = PicMem
SaveRedo
FMain.Pic1 = Im
FEcho.Hide
End Sub

Private Sub Command3_Click() 'cancel
FMain.Pic1.Picture = PicMem
FEcho.Hide
End Sub

Private Sub Form_Activate()
On Error Resume Next
Set PicMem = FMain.Pic1.Image
FEcho.Move 0, 330, 3195, 3570
Command2.Enabled = False
End Sub

Private Sub Form_Load()
FEcho.Move 0, 330, 3195, 3570
VScroll1.Value = 5
VScroll2.Value = 10
VScroll3.Value = 0
VScroll4.Value = 0
End Sub

Private Sub VScroll1_Change() 'number
Label2.Caption = Format(VScroll1.Value, "00")
End Sub

Private Sub VScroll2_Change()
Label3.Caption = Format(VScroll2.Value, "00") & "%"
End Sub

Private Sub VScroll3_Change()
Label4.Caption = Format(VScroll3.Value, "000")
End Sub

Private Sub VScroll4_Change()
Label5.Caption = Format(VScroll4.Value, "000")
End Sub
