VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PicHandler"
   ClientHeight    =   7920
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11880
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   528
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   225
      Pattern         =   "*bmp"
      TabIndex        =   17
      Top             =   4455
      Width           =   2265
   End
   Begin VB.PictureBox Tempmem2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   765
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   12
      Top             =   2115
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1395
      Top             =   2880
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   315
      Top             =   2835
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FMain.frx":046E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   60
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keyUndo"
            Object.ToolTipText     =   "Undo last action"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "keySelectAll"
            Object.ToolTipText     =   "No selection"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox TempMem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   315
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   8
      Top             =   2115
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   7560
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   900
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox PicX 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7305
      Left            =   4545
      ScaleHeight     =   487
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   487
      TabIndex        =   0
      Top             =   540
      Width           =   7305
      Begin VB.PictureBox Dummy 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   7080
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   5
         Top             =   7080
         Width           =   225
      End
      Begin VB.HScrollBar HS1 
         Height          =   240
         LargeChange     =   10
         Left            =   0
         TabIndex        =   3
         Top             =   7065
         Width           =   7080
      End
      Begin VB.VScrollBar VS1 
         Height          =   7080
         LargeChange     =   10
         Left            =   7065
         TabIndex        =   2
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox Pic1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   0
         ScaleHeight     =   1
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   1
         Top             =   0
         Width           =   15
      End
      Begin VB.Label Text 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   405
         TabIndex        =   14
         Top             =   1890
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Text 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   405
         TabIndex        =   13
         Top             =   1665
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   3  'Dot
         Height          =   915
         Left            =   360
         Top             =   405
         Width           =   825
      End
   End
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   2655
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   1500
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   2610
      TabIndex        =   19
      Top             =   4815
      Width           =   1590
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Picture: 0000 X 0000"
      Height          =   240
      Left            =   2610
      TabIndex        =   18
      Top             =   6525
      Width           =   1590
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Special map"
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
      Height          =   285
      Left            =   225
      TabIndex        =   15
      Top             =   4140
      Width           =   3930
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   90
      TabIndex        =   11
      Top             =   3390
      Width           =   4200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   800
      Y1              =   28
      Y2              =   28
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   800
      Y1              =   27
      Y2              =   27
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   4
      Left            =   3420
      Stretch         =   -1  'True
      Top             =   855
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   2610
      Stretch         =   -1  'True
      Top             =   855
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   2
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   855
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   990
      Stretch         =   -1  'True
      Top             =   855
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   180
      Stretch         =   -1  'True
      Top             =   855
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   7335
      Width           =   3840
   End
   Begin VB.Label Label1 
      Caption         =   " Picture Information:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1320
      Left            =   90
      TabIndex        =   4
      Top             =   1875
      Width           =   4200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   0
      X2              =   800
      Y1              =   4
      Y2              =   4
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   800
      Y1              =   3
      Y2              =   3
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Memory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1140
      Left            =   90
      TabIndex        =   10
      Top             =   540
      Width           =   4200
   End
   Begin VB.Label Label6 
      Height          =   3030
      Left            =   90
      TabIndex        =   16
      Top             =   4095
      Width           =   4200
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFiles 
         Caption         =   "Load picture..."
         Index           =   0
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Save picture"
         Index           =   1
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Save to special map"
         Index           =   2
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Print picture..."
         Index           =   4
      End
   End
   Begin VB.Menu mnuSelection 
      Caption         =   "Selection"
      Begin VB.Menu mnuSel 
         Caption         =   "Adjust selection"
         Index           =   0
      End
      Begin VB.Menu mnuSel 
         Caption         =   "No selection"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPicture 
      Caption         =   "Picture"
      Begin VB.Menu mnuPic 
         Caption         =   "Flip picture X"
         Index           =   0
      End
      Begin VB.Menu mnuPic 
         Caption         =   "Flip picture Y"
         Index           =   1
      End
      Begin VB.Menu mnuPic 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPic 
         Caption         =   "Mirror X"
         Index           =   3
      End
      Begin VB.Menu mnuPic 
         Caption         =   "Mirror X reversed"
         Index           =   4
      End
      Begin VB.Menu mnuPic 
         Caption         =   "Mirror Y"
         Index           =   5
      End
      Begin VB.Menu mnuPic 
         Caption         =   "Mirror Y reversed"
         Index           =   6
      End
   End
   Begin VB.Menu mnuColors 
      Caption         =   "Colors"
      Begin VB.Menu mnuCol 
         Caption         =   "Adjust Color..."
         Index           =   0
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Kill component"
         Index           =   2
         Begin VB.Menu mnuKill 
            Caption         =   "Kill red component"
            Index           =   0
         End
         Begin VB.Menu mnuKill 
            Caption         =   "Kill green component"
            Index           =   1
         End
         Begin VB.Menu mnuKill 
            Caption         =   "Kill blue component"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Swap component"
         Index           =   4
         Begin VB.Menu mnuSwap 
            Caption         =   "RGB --> BGR"
            Index           =   0
         End
         Begin VB.Menu mnuSwap 
            Caption         =   "RGB --> BRG"
            Index           =   1
         End
         Begin VB.Menu mnuSwap 
            Caption         =   "RGB --> GBR"
            Index           =   2
         End
         Begin VB.Menu mnuSwap 
            Caption         =   "RGB --> GRB"
            Index           =   3
         End
         Begin VB.Menu mnuSwap 
            Caption         =   "RGB --> RBG"
            Index           =   4
         End
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Brighten picture..."
         Index           =   6
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Contrast..."
         Index           =   7
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Photonegative"
         Index           =   9
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Invert component"
         Index           =   11
         Begin VB.Menu mnuInv 
            Caption         =   "Invert red"
            Index           =   0
         End
         Begin VB.Menu mnuInv 
            Caption         =   "Invert green"
            Index           =   1
         End
         Begin VB.Menu mnuInv 
            Caption         =   "Invert blue"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCol 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuCol 
         Caption         =   "Greyscale"
         Index           =   13
      End
   End
   Begin VB.Menu mnuFilters 
      Caption         =   "Filters"
      Begin VB.Menu mnuFilter 
         Caption         =   "Emboss"
         Index           =   0
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Emboss special"
         Index           =   1
         Begin VB.Menu mnuEmbossSpecial 
            Caption         =   "Hold red"
            Index           =   0
         End
         Begin VB.Menu mnuEmbossSpecial 
            Caption         =   "Hold green"
            Index           =   1
         End
         Begin VB.Menu mnuEmbossSpecial 
            Caption         =   "Hold blue"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Engrave"
         Index           =   3
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Neon"
         Index           =   5
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Blur picture"
         Index           =   7
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Blur picture more"
         Index           =   8
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Sharpen picture"
         Index           =   10
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Diffuse picture..."
         Index           =   12
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Erode picture..."
         Index           =   14
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Blow picture..."
         Index           =   15
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Add fog..."
         Index           =   16
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Add noise"
         Index           =   17
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Freeze picture"
         Index           =   19
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Freeze picture more"
         Index           =   20
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   21
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Black and white"
         Index           =   22
         Begin VB.Menu mnuBW 
            Caption         =   "B and W filter 1"
            Index           =   0
         End
         Begin VB.Menu mnuBW 
            Caption         =   "B and W filter 2"
            Index           =   1
         End
         Begin VB.Menu mnuBW 
            Caption         =   "B and W filter 3"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   23
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Soft Colors"
         Index           =   24
         Begin VB.Menu mnuSoft 
            Caption         =   "Soft red"
            Index           =   0
         End
         Begin VB.Menu mnuSoft 
            Caption         =   "Soft green"
            Index           =   1
         End
         Begin VB.Menu mnuSoft 
            Caption         =   "Soft orange"
            Index           =   2
         End
         Begin VB.Menu mnuSoft 
            Caption         =   "Soft yellow"
            Index           =   3
         End
         Begin VB.Menu mnuSoft 
            Caption         =   "Soft purple"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   25
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Hard colors"
         Index           =   26
         Begin VB.Menu mnuHard 
            Caption         =   "Hard red"
            Index           =   0
         End
         Begin VB.Menu mnuHard 
            Caption         =   "Hard green"
            Index           =   1
         End
         Begin VB.Menu mnuHard 
            Caption         =   "Hard blue"
            Index           =   2
         End
         Begin VB.Menu mnuHard 
            Caption         =   "Hard yellow"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuSpecialFilters 
      Caption         =   "Special filters"
      Begin VB.Menu mnuSpFil 
         Caption         =   "Brown"
         Index           =   0
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Dark Brown"
         Index           =   1
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Liquid"
         Index           =   2
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Yellow"
         Index           =   3
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Charcoal"
         Index           =   4
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Dark moon"
         Index           =   5
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Total eclipse"
         Index           =   6
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Purple rain"
         Index           =   7
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Spooky"
         Index           =   8
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Unreal"
         Index           =   9
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Flame"
         Index           =   10
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Aquarel"
         Index           =   11
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Spotted"
         Index           =   12
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Retro"
         Index           =   13
      End
      Begin VB.Menu mnuSpFil 
         Caption         =   "Wet paper"
         Index           =   14
      End
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "Effects"
      Begin VB.Menu mnuEff 
         Caption         =   "Horizontal blinds..."
         Index           =   0
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Vertical blinds..."
         Index           =   1
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Bumped hor. blinds"
         Index           =   2
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Bumped vert. blinds"
         Index           =   3
      End
      Begin VB.Menu mnuEff 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Horizontal lines"
         Index           =   5
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Vertical lines"
         Index           =   6
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Squares"
         Index           =   7
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Boxes"
         Index           =   8
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Circles"
         Index           =   9
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Diagonal right lines"
         Index           =   10
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Diagonal left lines"
         Index           =   11
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Crossed lines"
         Index           =   12
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Horizontal wave lines"
         Index           =   13
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Vertical wave lines"
         Index           =   14
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Horiz. abs. wave lines"
         Index           =   15
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Vertic. abs. wave lines"
         Index           =   16
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Hor. abs. wave lines rev."
         Index           =   17
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Vert. abs. wave lines rev."
         Index           =   18
      End
      Begin VB.Menu mnuEff 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnuEff 
         Caption         =   "Add border"
         Index           =   20
         Begin VB.Menu mnuBorder 
            Caption         =   "Solid border"
            Index           =   0
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "Solid border reduced"
            Index           =   1
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "Gradient border 1"
            Index           =   2
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "Grad. border 1 reduced"
            Index           =   3
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "Gradient border 2"
            Index           =   4
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "Grad. border 2 reduced"
            Index           =   5
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "Solid circular border"
            Index           =   7
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "Gradient circ. border 1"
            Index           =   8
         End
         Begin VB.Menu mnuBorder 
            Caption         =   "Gradient circ. border 2"
            Index           =   9
         End
      End
   End
   Begin VB.Menu mnuMixing 
      Caption         =   "Mixing"
      Begin VB.Menu mnuMix 
         Caption         =   "Mix with solid color"
         Index           =   0
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Mix with gradient 1"
         Index           =   1
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Mix with gradient 2"
         Index           =   2
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Mix with box gradient 1"
         Index           =   3
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Mix with box gradient 2"
         Index           =   4
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Mix with circular gradient 1"
         Index           =   5
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Mix with circular gradient 2"
         Index           =   6
      End
      Begin VB.Menu mnuMix 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Mix with picture"
         Index           =   8
      End
      Begin VB.Menu mnuMix 
         Caption         =   "Mix with pattern"
         Index           =   9
      End
   End
   Begin VB.Menu mnuDeformation 
      Caption         =   "Deformation"
      Begin VB.Menu mnuDef 
         Caption         =   "Echo picture"
         Index           =   0
      End
      Begin VB.Menu mnuDef 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuDef 
         Caption         =   "Mozaic"
         Index           =   2
      End
      Begin VB.Menu mnuDef 
         Caption         =   "Blurred mozaic"
         Index           =   3
      End
      Begin VB.Menu mnuDef 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuDef 
         Caption         =   "Wave X"
         Index           =   5
      End
      Begin VB.Menu mnuDef 
         Caption         =   "Abs. Wave X"
         Index           =   6
      End
      Begin VB.Menu mnuDef 
         Caption         =   "Wave Y"
         Index           =   7
      End
      Begin VB.Menu mnuDef 
         Caption         =   "Abs. Wave Y"
         Index           =   8
      End
      Begin VB.Menu mnuDef 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuDef 
         Caption         =   "Tile picture"
         Index           =   10
      End
   End
   Begin VB.Menu mnuEdges 
      Caption         =   "Edges"
      Begin VB.Menu mnuEdge 
         Caption         =   "Increase edges"
         Index           =   0
      End
      Begin VB.Menu mnuEdge 
         Caption         =   "Increase edges more"
         Index           =   1
      End
      Begin VB.Menu mnuEdge 
         Caption         =   "Partly increase edge"
         Index           =   2
         Begin VB.Menu mnuPEdge 
            Caption         =   "Left 1"
            Index           =   0
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Left 2"
            Index           =   1
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Left 3"
            Index           =   2
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Right 1"
            Index           =   4
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Right 2"
            Index           =   5
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Right 3"
            Index           =   6
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Top 1"
            Index           =   8
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Top 2"
            Index           =   9
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Top 3"
            Index           =   10
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Bottom 1"
            Index           =   12
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Bottom 2"
            Index           =   13
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Bottom 3"
            Index           =   14
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "-"
            Index           =   15
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Left Right 1"
            Index           =   16
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Left Right 2"
            Index           =   17
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "-"
            Index           =   18
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Top Bottom 1"
            Index           =   19
         End
         Begin VB.Menu mnuPEdge 
            Caption         =   "Top Bottom 2"
            Index           =   20
         End
      End
   End
   Begin VB.Menu mnuText 
      Caption         =   "Text"
      Begin VB.Menu mnuTxt 
         Caption         =   "Add text"
         Index           =   0
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuAb 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub File1_Click()
TempMem.Picture = LoadPicture(File1.Path & "\" & File1.List(File1.ListIndex))
Image2.Picture = LoadPicture(File1.Path & "\" & File1.List(File1.ListIndex))
SetImage2 TempMem.Width, TempMem.Height
End Sub

Private Sub SetImage2(L%, B%)
Dim T As Single, newL%, newH%
Label7.Caption = "Picture: " & Format(L, "000") & " X " & Format(B, "000")
newL = L: newH = B
T = 1
Do While newL > 100 Or newH > 100
newL = Int(L / T)
newH = Int(B / T)
T = T + 0.1
Loop
Image2.Width = newL
Image2.Height = newH
Image2.Move ((100 - newL) / 2) + 177, ((100 - newH) / 2) + 324
End Sub
Private Sub File1_DblClick()
Temp = MsgBox("Load picture " & File1.List(File1.ListIndex) & vbCr & " from the special map ?", vbYesNo + vbQuestion, FTitle)
If Temp = vbNo Then Exit Sub
Pic1.Picture = LoadPicture(File1.Path & "\" & File1.List(File1.ListIndex))
PicFileName = File1.List(File1.ListIndex)
SetScrollBars
SetPicInfo
DoEvents
SelectAll
MemCount = 0
Toolbar1.Buttons(1).Enabled = False
ClearMem
mnuColors.Enabled = True
mnuPicture.Enabled = True
mnuSelection.Enabled = True
mnuFiles(1).Enabled = True
mnuFiles(2).Enabled = True
mnuFiles(4).Enabled = True
mnuFilters.Enabled = True
mnuSpecialFilters.Enabled = True
mnuEffects.Enabled = True
mnuMixing.Enabled = True
mnuDeformation.Enabled = True
mnuEdges.Enabled = True
mnuText.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next
FTitle = "Pic-Handler V1.0"
Caption = FTitle
MkDir App.Path & "\Pictures"
File1.Path = App.Path & "\Pictures"
Set Shape1.Container = Pic1
Scol(0) = &H202020
Scol(1) = &H404040
Scol(2) = &H606060
Scol(3) = &H808080
Scol(4) = &HA0A0A0
Scol(5) = &HC0C0C0
Scol(6) = &HE0E0E0
Scol(7) = &HFFFFFF
Scol(8) = &HE0E0E0
Scol(9) = &HC0C0C0
Scol(10) = &HA0A0A0
Scol(11) = &H808080
Scol(12) = &H606060
Scol(13) = &H404040
Scol(14) = &H202020
Scol(15) = 0
FMain.Move 0, 0, 12000, 8610
PicX.Move 300, 36, 487, 487
Pic1.Move 0, 0
HS1.Enabled = False
VS1.Enabled = False
T3D FMain, PicX, 5, T3dRaiseInset
T3D FMain, Label1, 5, T3dRaiseInset
T3D FMain, Label3, 5, T3dRaiseInset
T3D FMain, Label4, 5, T3dRaiseInset
T3D FMain, Label6, 5, T3dRaiseInset
PB1.Value = 0
Toolbar1.Buttons(1).Enabled = False
ClearMem
mnuColors.Enabled = False
mnuPicture.Enabled = False
mnuSelection.Enabled = False
mnuFiles(1).Enabled = False
mnuFiles(2).Enabled = False
mnuFiles(4).Enabled = False
mnuFilters.Enabled = False
mnuSpecialFilters.Enabled = False
mnuEffects.Enabled = False
mnuMixing.Enabled = False
mnuDeformation.Enabled = False
mnuEdges.Enabled = False
mnuText.Enabled = False
SelectAll
For xx = 0 To 1
Set Text(xx).Container = Pic1
Text(xx).BackStyle = 0
Next xx
EnumFonts Printer.hdc, vbNullString, AddressOf EnumFontProc, 0
File1.Selected(0) = True
FMain.Show
FAbout.Show 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End '!!!!!!!!!!
End Sub

Private Sub HS1_Change()
Pic1.Left = HS1.Value
Pic1.SetFocus
End Sub

Private Sub HS1_GotFocus()
Dummy.SetFocus
End Sub

Private Sub mnuAb_Click()
FAbout.Show 1
End Sub

Private Sub mnuBorder_Click(Index As Integer)
Select Case Index
Case 0 'solid border
Col = 26
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 1 'solid border reduced
Col = 27
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 2 'gradient border 1
Col = 28
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 3 'gradient border 1 reduced
Col = 29
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 4 'gradient border 2
Col = 30
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 5 'gradient border 2 reduced
Col = 31
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 7 'solid circular border
Col = 32
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 8 'gradient circular border 1
Col = 33
FColor.Caption = "Effect - Borders"
FColor.Show 1
Case 9 'gradient circular border 1
Col = 34
FColor.Caption = "Effect - Borders"
FColor.Show 1
End Select
End Sub

Private Sub mnuBW_Click(Index As Integer)
SaveRedo
Select Case Index
Case 0
BnW Xcor0, Ycor0, Xcor1, Ycor1, 200
Case 1
BnW Xcor0, Ycor0, Xcor1, Ycor1, 150
Case 2
BnW Xcor0, Ycor0, Xcor1, Ycor1, 100
End Select
End Sub

Private Sub mnuCol_Click(Index As Integer)
Select Case Index
Case 0 'color comp
Col = 0
FColor.Caption = "Color"
FColor.Show 1
Case 6 'brighten
Col = 1
FColor.Caption = "Color"
FColor.Show 1
Case 7 'contrast
Col = 3
FColor.Caption = "Color"
FColor.Show 1
Case 9 'fotoneg.
SaveRedo
PhotoNeg Xcor0, Ycor0, Xcor1, Ycor1
Case 13 'grey
SaveRedo
GreyColor Xcor0, Ycor0, Xcor1, Ycor1
End Select
End Sub

Private Sub mnuDef_Click(Index As Integer)
FColor.Caption = "Deformation"
Select Case Index
Case 0
FEcho.Show 1
Case 2 'mozaic
Col = 42
FColor.Show 1
Case 3 'blurred mozaic
Col = 43
FColor.Show 1
Case 5 'wave X
Col = 44
FColor.Show 1
Case 6 'abs wave X
Col = 45
FColor.Show 1
Case 7 'wave Y
Col = 46
FColor.Show 1
Case 8 'abs wave Y
Col = 47
FColor.Show 1
Case 10 'tile pic
Col = 49
FColor.Show 1
End Select
End Sub

Private Sub mnuEdge_Click(Index As Integer)
Select Case Index
Case 0 'inc edges
SaveRedo
BeginProcess
KillColXGrad1 0, 0, Pic1.Width, Pic1.Height
KillColXGradRev1 0, 0, Pic1.Width, Pic1.Height
BeginProcess
KillColYGrad1 0, 0, Pic1.Width, Pic1.Height
KillColYGradRev1 0, 0, Pic1.Width, Pic1.Height
Case 1 'more inc edges
SaveRedo
BeginProcess
KillColXGrad2 0, 0, Pic1.Width, Pic1.Height
KillColXGradRev2 0, 0, Pic1.Width, Pic1.Height
BeginProcess
KillColYGrad2 0, 0, Pic1.Width, Pic1.Height
KillColYGradRev2 0, 0, Pic1.Width, Pic1.Height
End Select
End Sub

Private Sub mnuEff_Click(Index As Integer)
Select Case Index
Case 0 'H blinds
Col = 8
FColor.Caption = "Effects"
FColor.Show 1
Case 1 'V blinds
Col = 9
FColor.Caption = "Effects"
FColor.Show 1
Case 2 ' bump H blinds
Col = 10
FColor.Caption = "Effects"
FColor.Show 1
Case 3 ' bump V blinds
Col = 11
FColor.Caption = "Effects"
FColor.Show 1
Case 5 ' add hor lines
Col = 12
FColor.Caption = "Effects"
FColor.Show 1
Case 6 ' add ver lines
Col = 13
FColor.Caption = "Effects"
FColor.Show 1
Case 7 ' add squares
Col = 14
FColor.Caption = "Effects"
FColor.Show 1
Case 8 ' add squares
Col = 15
FColor.Caption = "Effects"
FColor.Show 1
Case 9 ' add squares
Col = 16
FColor.Caption = "Effects"
FColor.Show 1
Case 10 ' add dia R lines
Col = 17
FColor.Caption = "Effects"
FColor.Show 1
Case 11 ' add dia R lines
Col = 18
FColor.Caption = "Effects"
FColor.Show 1
Case 12 ' add crossed lines
Col = 19
FColor.Caption = "Effects"
FColor.Show 1
Case 13 ' add H wave lines
Col = 20
FColor.Caption = "Effects"
FColor.Show 1
Case 14 ' add V wave lines
Col = 21
FColor.Caption = "Effects"
FColor.Show 1
Case 15 ' add abs H wave lines
Col = 22
FColor.Caption = "Effects"
FColor.Show 1
Case 16 ' add abs V wave lines
Col = 23
FColor.Caption = "Effects"
FColor.Show 1
Case 17 ' add abs H wave lines reversed
Col = 24
FColor.Caption = "Effects"
FColor.Show 1
Case 18 ' add abs V wave lines reversed
Col = 25
FColor.Caption = "Effects"
FColor.Show 1
End Select
End Sub

Private Sub mnuEmbossSpecial_Click(Index As Integer)
SaveRedo
Select Case Index
Case 0 'holdred
HoldRed Xcor0, Ycor0, Xcor1, Ycor1
Case 1 'holdgreen
HoldGreen Xcor0, Ycor0, Xcor1, Ycor1
Case 2 'holdblue
HoldBlue Xcor0, Ycor0, Xcor1, Ycor1
End Select
End Sub

Private Sub mnuFiles_Click(Index As Integer)
On Error GoTo mnuFilesError
Select Case Index
Case 0 'open pic
CD1.Filter = "All picture files|*.bmp;*.jpg;*.gif;*.wmf;*.ico"
CD1.Flags = 2
CD1.ShowOpen
Pic1.Picture = LoadPicture(CD1.FileName)
PicFileName = CD1.FileTitle
SetScrollBars
SetPicInfo
DoEvents
SelectAll
MemCount = 0
Toolbar1.Buttons(1).Enabled = False
ClearMem
mnuColors.Enabled = True
mnuPicture.Enabled = True
mnuSelection.Enabled = True
mnuFiles(1).Enabled = True
mnuFiles(2).Enabled = True
mnuFiles(4).Enabled = True
mnuFilters.Enabled = True
mnuSpecialFilters.Enabled = True
mnuEffects.Enabled = True
mnuMixing.Enabled = True
mnuDeformation.Enabled = True
mnuEdges.Enabled = True
mnuText.Enabled = True
Case 1 'save picture
CD1.Filter = "Bitmap|*.bmp"
CD1.Flags = 2
CD1.FileName = 0
PicFileName = Left(PicFileName, Len(PicFileName) - 3) & "bmp"
CD1.FileName = PicFileName
CD1.ShowSave
PicFileName = CD1.FileTitle
SavePicture Pic1.Image, PicFileName
SetPicInfo
DoEvents
Case 2 'save to spec. map"
If LCase(Right(PicFileName, 3)) <> "bmp" Then
PicFileName = Left(PicFileName, Len(PicFileName) - 3) & "bmp"
End If
SavePicture Pic1.Image, App.Path & "\pictures\" & PicFileName
File1.Refresh
CD1.FileTitle = PicFileName
SetPicInfo
DoEvents
Case 4 'print
Temp = MsgBox("print in Landscape-orientation?", vbQuestion + vbYesNoCancel, FTitle)
If Temp = vbCancel Then MsgBox "cancel printing", , FTitle: Exit Sub
If Temp = vbYes Then Printer.Orientation = 2
If Temp = vbNo Then Printer.Orientation = 1
Printer.PaintPicture Pic1.Image, 0, 0
Printer.EndDoc
Printer.Orientation = 1
End Select
Exit Sub
mnuFilesError:
End Sub

Private Sub mnuFilter_Click(Index As Integer)
Select Case Index
Case 0 'emboss
SaveRedo
EmbossPicture Xcor0, Ycor0, Xcor1, Ycor1
Case 3 'engrave
SaveRedo
EngravePicture Xcor0, Ycor0, Xcor1, Ycor1
Case 5 'neon
SaveRedo
NeonPicture Xcor0, Ycor0, Xcor1, Ycor1
Case 7 'blur
SaveRedo
BlurPicture Xcor0, Ycor0, Xcor1, Ycor1
Case 8 'blur more
SaveRedo
BlurPictureMore Xcor0, Ycor0, Xcor1, Ycor1
Case 10 'sharpen
SaveRedo
SharpenPicture Xcor0, Ycor0, Xcor1, Ycor1
Case 12 'diffuse
Col = 4
FColor.Caption = "Filters"
FColor.Show 1
Case 14 'erode
Col = 5
FColor.Caption = "Filters"
FColor.Show 1
Case 15 'Blow
Col = 6
FColor.Caption = "Filters"
FColor.Show 1
Case 16 'fog
Col = 7
FColor.Caption = "Filters"
FColor.Show 1
Case 17 'noise
SaveRedo
AddNoise Xcor0, Ycor0, Xcor1, Ycor1
Case 19 'freeze
SaveRedo
FreezePic Xcor0, Ycor0, Xcor1, Ycor1, 1.1
Case 20 'freezemore
SaveRedo
FreezePic Xcor0, Ycor0, Xcor1, Ycor1, 1.5
End Select
End Sub

Private Sub mnuInv_Click(Index As Integer)
SaveRedo
PhotoNegComp Xcor0, Ycor0, Xcor1, Ycor1, Index
End Sub

Private Sub mnuKill_Click(Index As Integer)
SaveRedo
KillComp Xcor0, Ycor0, Xcor1, Ycor1, Index
End Sub

Private Sub mnuMix_Click(Index As Integer)
Select Case Index
Case 0 'mix solid color
Col = 35
FColor.Caption = "Mixing with color"
FColor.Show 1
Case 1 'mix gradient 1
Col = 36
FColor.Caption = "Mixing with color"
FColor.Show 1
Case 2 'mix gradient 2
Col = 37
FColor.Caption = "Mixing with color"
FColor.Show 1
Case 3 'mix box gradient 1
Col = 38
FColor.Caption = "Mixing with color"
FColor.Show 1
Case 4 'mix box gradient 1
Col = 39
FColor.Caption = "Mixing with color"
FColor.Show 1
Case 5 'mix circular gradient 1
Col = 40
FColor.Caption = "Mixing with color"
FColor.Show 1
Case 6 'mix circular gradient 2
Col = 41
FColor.Caption = "Mixing with color"
FColor.Show 1
'---------------
Case 8 'mix picture
Mix = 0
FPicture.Caption = "Graphic mixing"
FPicture.Show 1
Case 9 'mix pattern
Mix = 1
FPicture.Caption = "Graphic mixing"
FPicture.Show 1
End Select
End Sub

Private Sub mnuPEdge_Click(Index As Integer)
Select Case Index
Case 0 'inc edges L1
SaveRedo
BeginProcess
KillColXGrad1 0, 0, Pic1.Width, Pic1.Height
Case 1 'inc edges L2
SaveRedo
BeginProcess
KillColXGrad2 0, 0, Pic1.Width, Pic1.Height
Case 2 'inc edges L3
SaveRedo
BeginProcess
KillColXGrad3 0, 0, Pic1.Width, Pic1.Height
Case 4 'inc edges R1
SaveRedo
BeginProcess
KillColXGradRev1 0, 0, Pic1.Width, Pic1.Height
Case 5 'inc edges R2
SaveRedo
BeginProcess
KillColXGradRev2 0, 0, Pic1.Width, Pic1.Height
Case 6 'inc edges R3
SaveRedo
BeginProcess
KillColXGradRev3 0, 0, Pic1.Width, Pic1.Height
Case 8 'inc edges T1
SaveRedo
BeginProcess
KillColYGrad1 0, 0, Pic1.Width, Pic1.Height
Case 9 'inc edges T2
SaveRedo
BeginProcess
KillColYGrad2 0, 0, Pic1.Width, Pic1.Height
Case 10 'inc edges T3
SaveRedo
BeginProcess
KillColYGrad3 0, 0, Pic1.Width, Pic1.Height
Case 12 'inc edges B1
SaveRedo
BeginProcess
KillColYGradRev1 0, 0, Pic1.Width, Pic1.Height
Case 13 'inc edges B2
SaveRedo
BeginProcess
KillColYGradRev2 0, 0, Pic1.Width, Pic1.Height
Case 14 'inc edges B3
SaveRedo
BeginProcess
KillColYGradRev3 0, 0, Pic1.Width, Pic1.Height
Case 16 'inc edges L1 & R1
SaveRedo
BeginProcess
KillColXGrad1 0, 0, Pic1.Width, Pic1.Height
KillColXGradRev1 0, 0, Pic1.Width, Pic1.Height
Case 17 'inc edges L2& R2
SaveRedo
BeginProcess
KillColXGrad2 0, 0, Pic1.Width, Pic1.Height
KillColXGradRev2 0, 0, Pic1.Width, Pic1.Height
Case 19 'inc edges T1 & B1
SaveRedo
BeginProcess
KillColYGrad1 0, 0, Pic1.Width, Pic1.Height
KillColYGradRev1 0, 0, Pic1.Width, Pic1.Height
Case 20 'inc edges T2& B2
SaveRedo
BeginProcess
KillColYGrad2 0, 0, Pic1.Width, Pic1.Height
KillColYGradRev2 0, 0, Pic1.Width, Pic1.Height
End Select
End Sub

Private Sub mnuPic_Click(Index As Integer)
Select Case Index
Case 0 ' flip X
SaveRedo
FlipX
Case 1 'flip Y
SaveRedo
FlipY
Case 3 'Mirror X
SaveRedo
MirrorX
Case 4 'Mirror Y
SaveRedo
MirrorXRev
Case 5 'Mirror X
SaveRedo
MirrorY
Case 6 'Mirror Y
SaveRedo
MirrorYRev
End Select
End Sub

Private Sub mnuSel_Click(Index As Integer)
Select Case Index
Case 0 'adjust selection
If Shape1.Visible = False Then
MsgBox "No selection present", vbInformation, FTitle
Exit Sub
End If
Col = 48
FColor.Show 1
Case 1 'select all
SelectAll
End Select
End Sub

Private Sub mnuSoft_Click(Index As Integer)
SaveRedo
Select Case Index
Case 0
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 3
Case 1
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 1
Case 2
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 10
Case 3
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 9
Case 4
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 2
End Select
End Sub

Private Sub mnuHard_Click(Index As Integer)
SaveRedo
Select Case Index
Case 0
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 4
Case 1
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 5
Case 2
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 0
Case 3
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 8
End Select
End Sub

Private Sub mnuSpFil_Click(Index As Integer)
SaveRedo
Select Case Index
Case 0
Brown Xcor0, Ycor0, Xcor1, Ycor1, 128
Case 1
Brown Xcor0, Ycor0, Xcor1, Ycor1, 256
Case 2
Liquid Xcor0, Ycor0, Xcor1, Ycor1
Case 3
Yellow Xcor0, Ycor0, Xcor1, Ycor1
Case 4
Charcoal Xcor0, Ycor0, Xcor1, Ycor1
Case 5
DarkMoon Xcor0, Ycor0, Xcor1, Ycor1
Case 6
TotalEclipse Xcor0, Ycor0, Xcor1, Ycor1
Case 7
PurpleRain Xcor0, Ycor0, Xcor1, Ycor1
Case 8
Spooky Xcor0, Ycor0, Xcor1, Ycor1
Case 9
UnReal Xcor0, Ycor0, Xcor1, Ycor1
Case 10
Flame Xcor0, Ycor0, Xcor1, Ycor1
Case 11
Aquarel Xcor0, Ycor0, Xcor1, Ycor1
Case 12
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 6
Case 13
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 7
Case 14
Effect0 Xcor0, Ycor0, Xcor1, Ycor1, 11
End Select
End Sub

Private Sub mnuSwap_Click(Index As Integer)
SaveRedo
SwapComp Xcor0, Ycor0, Xcor1, Ycor1, Index
End Sub

Private Sub mnuTxt_Click(Index As Integer)
Select Case Index
Case 0 'add text
FText.Show 1
End Select
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Shape1.Visible = True
Xcor0 = x: Ycor0 = y
Xcor1 = 0: Ycor1 = 0
XXX1 = Xcor0: YYY1 = Ycor0
XXX2 = Xcor1: YYY2 = Ycor1
Shape1.Move Xcor0, Ycor0, Xcor1, Ycor1
SetCoordinates
FMain.Toolbar1.Buttons(3).Enabled = True
End Sub


Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
            XXX2 = x - XXX1: YYY2 = y - YYY1
            Xcor0 = XXX1: Ycor0 = YYY1
            Xcor1 = XXX2: Ycor1 = YYY2
    If XXX2 < 0 Then
    Xcor0 = XXX1 + XXX2
            If Xcor0 < 0 Then Xcor0 = 0
    Xcor1 = XXX1 - Xcor0
    End If
    If YYY2 < 0 Then
    Ycor0 = YYY1 + YYY2
            If Ycor0 < 0 Then Ycor0 = 0
    Ycor1 = YYY1 - Ycor0
    End If
        If Xcor0 + Xcor1 > Pic1.Width Then Xcor1 = Pic1.Width - Xcor0
        If Ycor0 + Ycor1 > Pic1.Height Then Ycor1 = Pic1.Height - Ycor0
        Shape1.Move Xcor0, Ycor0, Xcor1, Ycor1
        SetCoordinates
End If
End Sub

Private Sub Timer1_Timer()
If Shape1.Visible = False Then Exit Sub
Tim = (Tim + 1) And 15
Shape1.BorderColor = Scol(Tim)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "keyUndo"
Redo
Case "keySelectAll"
SelectAll
End Select
End Sub

Private Sub VS1_Change()
Pic1.Top = VS1.Value
Pic1.SetFocus
End Sub

Private Sub VS1_GotFocus()
Dummy.SetFocus
End Sub
