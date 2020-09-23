VERSION 5.00
Begin VB.Form FColor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   262
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Tile picture"
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
      Height          =   1005
      Index           =   49
      Left            =   135
      TabIndex        =   449
      Top             =   585
      Width           =   3645
      Begin VB.HScrollBar HScroll14 
         Height          =   195
         Index           =   7
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   451
         Top             =   630
         Value           =   2
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll14 
         Height          =   195
         Index           =   6
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   450
         Top             =   315
         Value           =   2
         Width           =   1635
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tyle X"
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
         Index           =   7
         Left            =   90
         TabIndex        =   455
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label50 
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
         Index           =   7
         Left            =   3015
         TabIndex        =   454
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label50 
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
         Index           =   6
         Left            =   3015
         TabIndex        =   453
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tyle Y"
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
         Index           =   6
         Left            =   90
         TabIndex        =   452
         Top             =   630
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjust selection"
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
      Height          =   1995
      Index           =   48
      Left            =   135
      TabIndex        =   435
      Top             =   45
      Width           =   3645
      Begin VB.CommandButton Command5 
         Caption         =   "Center"
         Height          =   330
         Left            =   2655
         TabIndex        =   448
         Top             =   1575
         Width           =   870
      End
      Begin VB.HScrollBar HScroll13 
         Height          =   195
         Index           =   3
         LargeChange     =   10
         Left            =   1305
         Max             =   200
         Min             =   1
         TabIndex        =   445
         Top             =   1260
         Value           =   2
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll13 
         Height          =   195
         Index           =   2
         LargeChange     =   10
         Left            =   1305
         Max             =   200
         Min             =   1
         TabIndex        =   442
         Top             =   945
         Value           =   2
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll13 
         Height          =   195
         Index           =   1
         LargeChange     =   10
         Left            =   1305
         Max             =   250
         Min             =   1
         TabIndex        =   437
         Top             =   630
         Value           =   2
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll13 
         Height          =   195
         Index           =   0
         LargeChange     =   10
         Left            =   1305
         Max             =   200
         Min             =   1
         TabIndex        =   436
         Top             =   315
         Value           =   2
         Width           =   1635
      End
      Begin VB.Label Label50 
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
         Index           =   3
         Left            =   3015
         TabIndex        =   447
         Top             =   1260
         Width           =   510
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Height"
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
         Index           =   3
         Left            =   90
         TabIndex        =   446
         Top             =   1260
         Width           =   1140
      End
      Begin VB.Label Label50 
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
         Index           =   2
         Left            =   3015
         TabIndex        =   444
         Top             =   945
         Width           =   510
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Width"
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
         Left            =   90
         TabIndex        =   443
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Start top"
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
         Left            =   90
         TabIndex        =   441
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label50 
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
         Index           =   1
         Left            =   3015
         TabIndex        =   440
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label50 
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
         Index           =   0
         Left            =   3015
         TabIndex        =   439
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Start left"
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
         Left            =   90
         TabIndex        =   438
         Top             =   315
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Absolute wave Y"
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
      Height          =   1095
      Index           =   47
      Left            =   135
      TabIndex        =   428
      Top             =   450
      Width           =   3645
      Begin VB.HScrollBar HScroll11 
         Height          =   195
         Index           =   3
         LargeChange     =   10
         Left            =   1305
         Max             =   200
         Min             =   1
         TabIndex        =   430
         Top             =   315
         Value           =   2
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll12 
         Height          =   195
         Index           =   3
         LargeChange     =   10
         Left            =   1305
         Max             =   250
         Min             =   1
         TabIndex        =   429
         Top             =   630
         Value           =   2
         Width           =   1635
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude"
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
         Index           =   39
         Left            =   90
         TabIndex        =   434
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label21 
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
         Index           =   3
         Left            =   3015
         TabIndex        =   433
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label22 
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
         Index           =   3
         Left            =   3015
         TabIndex        =   432
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave"
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
         Index           =   38
         Left            =   90
         TabIndex        =   431
         Top             =   630
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wave Y"
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
      Height          =   1095
      Index           =   46
      Left            =   135
      TabIndex        =   421
      Top             =   450
      Width           =   3645
      Begin VB.HScrollBar HScroll12 
         Height          =   195
         Index           =   2
         LargeChange     =   10
         Left            =   1305
         Max             =   250
         Min             =   1
         TabIndex        =   423
         Top             =   630
         Value           =   2
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll11 
         Height          =   195
         Index           =   2
         LargeChange     =   10
         Left            =   1305
         Max             =   200
         Min             =   1
         TabIndex        =   422
         Top             =   315
         Value           =   2
         Width           =   1635
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave"
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
         Index           =   37
         Left            =   90
         TabIndex        =   427
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label22 
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
         Index           =   2
         Left            =   3015
         TabIndex        =   426
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label21 
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
         Index           =   2
         Left            =   3015
         TabIndex        =   425
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude"
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
         Index           =   36
         Left            =   90
         TabIndex        =   424
         Top             =   315
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Absolute wave X"
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
      Height          =   1095
      Index           =   45
      Left            =   135
      TabIndex        =   414
      Top             =   450
      Width           =   3645
      Begin VB.HScrollBar HScroll11 
         Height          =   195
         Index           =   1
         LargeChange     =   10
         Left            =   1305
         Max             =   200
         Min             =   1
         TabIndex        =   416
         Top             =   315
         Value           =   2
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll12 
         Height          =   195
         Index           =   1
         LargeChange     =   10
         Left            =   1305
         Max             =   250
         Min             =   1
         TabIndex        =   415
         Top             =   630
         Value           =   2
         Width           =   1635
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude"
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
         Index           =   35
         Left            =   90
         TabIndex        =   420
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label21 
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
         Index           =   1
         Left            =   3015
         TabIndex        =   419
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label22 
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
         Index           =   1
         Left            =   3015
         TabIndex        =   418
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave"
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
         Index           =   34
         Left            =   90
         TabIndex        =   417
         Top             =   630
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wave X"
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
      Height          =   1095
      Index           =   44
      Left            =   135
      TabIndex        =   407
      Top             =   450
      Width           =   3645
      Begin VB.HScrollBar HScroll12 
         Height          =   195
         Index           =   0
         LargeChange     =   10
         Left            =   1305
         Max             =   250
         Min             =   1
         TabIndex        =   411
         Top             =   630
         Value           =   2
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll11 
         Height          =   195
         Index           =   0
         LargeChange     =   10
         Left            =   1305
         Max             =   200
         Min             =   1
         TabIndex        =   408
         Top             =   315
         Value           =   2
         Width           =   1635
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave"
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
         Index           =   33
         Left            =   90
         TabIndex        =   413
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label22 
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
         Index           =   0
         Left            =   3015
         TabIndex        =   412
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label21 
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
         Index           =   0
         Left            =   3015
         TabIndex        =   410
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude"
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
         Index           =   32
         Left            =   90
         TabIndex        =   409
         Top             =   315
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Blurred mozaic"
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
      Height          =   735
      Index           =   43
      Left            =   135
      TabIndex        =   403
      Top             =   450
      Width           =   3645
      Begin VB.HScrollBar HScroll10 
         Height          =   195
         Index           =   1
         Left            =   1305
         Max             =   50
         Min             =   2
         TabIndex        =   404
         Top             =   315
         Value           =   2
         Width           =   1635
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Width"
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
         Index           =   31
         Left            =   90
         TabIndex        =   406
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label20 
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
         Index           =   1
         Left            =   3015
         TabIndex        =   405
         Top             =   315
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mozaic"
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
      Height          =   735
      Index           =   42
      Left            =   135
      TabIndex        =   399
      Top             =   450
      Width           =   3645
      Begin VB.HScrollBar HScroll10 
         Height          =   195
         Index           =   0
         Left            =   1305
         Max             =   50
         Min             =   2
         TabIndex        =   400
         Top             =   315
         Value           =   2
         Width           =   1635
      End
      Begin VB.Label Label20 
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
         Index           =   0
         Left            =   3015
         TabIndex        =   402
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Width"
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
         Index           =   30
         Left            =   90
         TabIndex        =   401
         Top             =   315
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Circular gradient 2"
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
      Height          =   1410
      Index           =   41
      Left            =   135
      TabIndex        =   388
      Top             =   135
      Width           =   3645
      Begin VB.CommandButton Command4 
         Caption         =   "Swap colors"
         Height          =   330
         Index           =   11
         Left            =   2250
         TabIndex        =   391
         Top             =   990
         Width           =   1275
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   15
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   390
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.PictureBox Grad1 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   11
         Left            =   945
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   389
         Top             =   1035
         Width           =   1125
      End
      Begin VB.Label ColLabel1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Color"
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
         Index           =   11
         Left            =   90
         TabIndex        =   398
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Ficolor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   11
         Left            =   1440
         TabIndex        =   397
         Top             =   630
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   29
         Left            =   90
         TabIndex        =   396
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label17 
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
         Index           =   15
         Left            =   3015
         TabIndex        =   395
         Top             =   315
         Width           =   510
      End
      Begin VB.Label ColLabel2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Color"
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
         Index           =   11
         Left            =   1890
         TabIndex        =   394
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label SColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   11
         Left            =   3240
         TabIndex        =   393
         Top             =   630
         Width           =   285
      End
      Begin VB.Label Label18 
         Caption         =   "Gradient"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   392
         Top             =   1035
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Circular gradient 1"
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
      Height          =   1410
      Index           =   40
      Left            =   135
      TabIndex        =   377
      Top             =   135
      Width           =   3645
      Begin VB.PictureBox Grad1 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   10
         Left            =   945
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   380
         Top             =   1035
         Width           =   1125
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   14
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   379
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Swap colors"
         Height          =   330
         Index           =   10
         Left            =   2250
         TabIndex        =   378
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label18 
         Caption         =   "Gradient"
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   387
         Top             =   1035
         Width           =   780
      End
      Begin VB.Label SColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   10
         Left            =   3240
         TabIndex        =   386
         Top             =   630
         Width           =   285
      End
      Begin VB.Label ColLabel2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Color"
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
         Index           =   10
         Left            =   1890
         TabIndex        =   385
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Label17 
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
         Index           =   14
         Left            =   3015
         TabIndex        =   384
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label10 
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
         Index           =   28
         Left            =   90
         TabIndex        =   383
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Ficolor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   10
         Left            =   1440
         TabIndex        =   382
         Top             =   630
         Width           =   285
      End
      Begin VB.Label ColLabel1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Color"
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
         Index           =   10
         Left            =   90
         TabIndex        =   381
         Top             =   630
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Box gradient 2"
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
      Height          =   1410
      Index           =   39
      Left            =   135
      TabIndex        =   366
      Top             =   135
      Width           =   3645
      Begin VB.CommandButton Command4 
         Caption         =   "Swap colors"
         Height          =   330
         Index           =   9
         Left            =   2250
         TabIndex        =   369
         Top             =   990
         Width           =   1275
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   13
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   368
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.PictureBox Grad1 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   9
         Left            =   945
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   367
         Top             =   1035
         Width           =   1125
      End
      Begin VB.Label ColLabel1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Color"
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
         Index           =   9
         Left            =   90
         TabIndex        =   376
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Ficolor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   9
         Left            =   1440
         TabIndex        =   375
         Top             =   630
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   27
         Left            =   90
         TabIndex        =   374
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label17 
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
         Index           =   13
         Left            =   3015
         TabIndex        =   373
         Top             =   315
         Width           =   510
      End
      Begin VB.Label ColLabel2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Color"
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
         Index           =   9
         Left            =   1890
         TabIndex        =   372
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label SColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   9
         Left            =   3240
         TabIndex        =   371
         Top             =   630
         Width           =   285
      End
      Begin VB.Label Label18 
         Caption         =   "Gradient"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   370
         Top             =   1035
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Box gradient 1"
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
      Height          =   1410
      Index           =   38
      Left            =   135
      TabIndex        =   355
      Top             =   135
      Width           =   3645
      Begin VB.PictureBox Grad1 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   8
         Left            =   945
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   358
         Top             =   1035
         Width           =   1125
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   12
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   357
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Swap colors"
         Height          =   330
         Index           =   8
         Left            =   2250
         TabIndex        =   356
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label18 
         Caption         =   "Gradient"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   365
         Top             =   1035
         Width           =   780
      End
      Begin VB.Label SColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   8
         Left            =   3240
         TabIndex        =   364
         Top             =   630
         Width           =   285
      End
      Begin VB.Label ColLabel2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Color"
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
         Index           =   8
         Left            =   1890
         TabIndex        =   363
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Label17 
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
         Index           =   12
         Left            =   3015
         TabIndex        =   362
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label10 
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
         Index           =   26
         Left            =   90
         TabIndex        =   361
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Ficolor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   8
         Left            =   1440
         TabIndex        =   360
         Top             =   630
         Width           =   285
      End
      Begin VB.Label ColLabel1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Color"
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
         Index           =   8
         Left            =   90
         TabIndex        =   359
         Top             =   630
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gradient color 2"
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
      Height          =   1725
      Index           =   37
      Left            =   135
      TabIndex        =   343
      Top             =   135
      Width           =   3645
      Begin VB.CommandButton Command4 
         Caption         =   "Swap colors"
         Height          =   330
         Index           =   7
         Left            =   2250
         TabIndex        =   347
         Top             =   990
         Width           =   1275
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   11
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   346
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.PictureBox Grad1 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   7
         Left            =   945
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   345
         Top             =   1035
         Width           =   1125
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Vertical gradient"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   344
         Top             =   1395
         Width           =   1680
      End
      Begin VB.Label ColLabel1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Color"
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
         Index           =   7
         Left            =   90
         TabIndex        =   354
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Ficolor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   7
         Left            =   1440
         TabIndex        =   353
         Top             =   630
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   25
         Left            =   90
         TabIndex        =   352
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label17 
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
         Index           =   11
         Left            =   3015
         TabIndex        =   351
         Top             =   315
         Width           =   510
      End
      Begin VB.Label ColLabel2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Color"
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
         Index           =   7
         Left            =   1890
         TabIndex        =   350
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label SColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   7
         Left            =   3240
         TabIndex        =   349
         Top             =   630
         Width           =   285
      End
      Begin VB.Label Label18 
         Caption         =   "Gradient"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   348
         Top             =   1035
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gradient color 1"
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
      Height          =   1725
      Index           =   36
      Left            =   135
      TabIndex        =   331
      Top             =   135
      Width           =   3645
      Begin VB.CheckBox Check3 
         Caption         =   "Vertical gradient"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   342
         Top             =   1395
         Width           =   1680
      End
      Begin VB.PictureBox Grad1 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   6
         Left            =   945
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   334
         Top             =   1035
         Width           =   1125
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   10
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   333
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Swap colors"
         Height          =   330
         Index           =   6
         Left            =   2250
         TabIndex        =   332
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label18 
         Caption         =   "Gradient"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   341
         Top             =   1035
         Width           =   780
      End
      Begin VB.Label SColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   6
         Left            =   3240
         TabIndex        =   340
         Top             =   630
         Width           =   285
      End
      Begin VB.Label ColLabel2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Color"
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
         Index           =   6
         Left            =   1890
         TabIndex        =   339
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Label17 
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
         Index           =   10
         Left            =   3015
         TabIndex        =   338
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label10 
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
         Index           =   24
         Left            =   90
         TabIndex        =   337
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Ficolor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   6
         Left            =   1440
         TabIndex        =   336
         Top             =   630
         Width           =   285
      End
      Begin VB.Label ColLabel1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Color"
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
         Index           =   6
         Left            =   90
         TabIndex        =   335
         Top             =   630
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gradient circular border 2"
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
      Height          =   1725
      Index           =   34
      Left            =   135
      TabIndex        =   306
      Top             =   135
      Width           =   3645
      Begin VB.CommandButton Command4 
         Caption         =   "Swap colors"
         Height          =   330
         Index           =   5
         Left            =   2250
         TabIndex        =   324
         Top             =   1305
         Width           =   1275
      End
      Begin VB.HScrollBar HScroll8 
         Height          =   195
         Index           =   8
         LargeChange     =   10
         Left            =   1305
         Max             =   400
         Min             =   1
         TabIndex        =   309
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   8
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   308
         Top             =   630
         Value           =   1
         Width           =   1635
      End
      Begin VB.PictureBox Grad1 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   5
         Left            =   945
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   307
         Top             =   1350
         Width           =   1125
      End
      Begin VB.Label ColLabel1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Color"
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
         Index           =   5
         Left            =   90
         TabIndex        =   318
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label Ficolor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   5
         Left            =   1440
         TabIndex        =   317
         Top             =   945
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   22
         Left            =   90
         TabIndex        =   316
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   14
         Left            =   90
         TabIndex        =   315
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label16 
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
         Index           =   8
         Left            =   3015
         TabIndex        =   314
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label17 
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
         Index           =   8
         Left            =   3015
         TabIndex        =   313
         Top             =   630
         Width           =   510
      End
      Begin VB.Label ColLabel2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Color"
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
         Index           =   5
         Left            =   1890
         TabIndex        =   312
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label SColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   5
         Left            =   3240
         TabIndex        =   311
         Top             =   945
         Width           =   285
      End
      Begin VB.Label Label18 
         Caption         =   "Gradient"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   310
         Top             =   1350
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solid color"
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
      Height          =   960
      Index           =   35
      Left            =   135
      TabIndex        =   325
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   9
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   326
         Top             =   270
         Value           =   1
         Width           =   1635
      End
      Begin VB.Label Label17 
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
         Index           =   9
         Left            =   3015
         TabIndex        =   330
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label10 
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
         Index           =   23
         Left            =   90
         TabIndex        =   329
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   17
         Left            =   1305
         TabIndex        =   328
         Top             =   585
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mix Color"
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
         Index           =   17
         Left            =   90
         TabIndex        =   327
         Top             =   585
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solid circular border"
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
      Height          =   1320
      Index           =   32
      Left            =   135
      TabIndex        =   284
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll8 
         Height          =   195
         Index           =   6
         LargeChange     =   10
         Left            =   1305
         Max             =   400
         Min             =   1
         TabIndex        =   286
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   6
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   285
         Top             =   630
         Value           =   1
         Width           =   1635
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Border Color"
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
         Index           =   16
         Left            =   90
         TabIndex        =   292
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   16
         Left            =   1305
         TabIndex        =   291
         Top             =   945
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   20
         Left            =   90
         TabIndex        =   290
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   12
         Left            =   90
         TabIndex        =   289
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label16 
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
         Index           =   6
         Left            =   3015
         TabIndex        =   288
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label17 
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
         Index           =   6
         Left            =   3015
         TabIndex        =   287
         Top             =   630
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gradient circular border 1"
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
      Height          =   1725
      Index           =   33
      Left            =   135
      TabIndex        =   293
      Top             =   135
      Width           =   3645
      Begin VB.CommandButton Command4 
         Caption         =   "Swap colors"
         Height          =   330
         Index           =   4
         Left            =   2250
         TabIndex        =   323
         Top             =   1305
         Width           =   1275
      End
      Begin VB.PictureBox Grad1 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   4
         Left            =   945
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   296
         Top             =   1350
         Width           =   1125
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   7
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   295
         Top             =   630
         Value           =   1
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll8 
         Height          =   195
         Index           =   7
         LargeChange     =   10
         Left            =   1305
         Max             =   400
         Min             =   1
         TabIndex        =   294
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.Label Label18 
         Caption         =   "Gradient"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   305
         Top             =   1350
         Width           =   780
      End
      Begin VB.Label SColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   4
         Left            =   3240
         TabIndex        =   304
         Top             =   945
         Width           =   285
      End
      Begin VB.Label ColLabel2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Color"
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
         Index           =   4
         Left            =   1890
         TabIndex        =   303
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label Label17 
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
         Index           =   7
         Left            =   3015
         TabIndex        =   302
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label16 
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
         Index           =   7
         Left            =   3015
         TabIndex        =   301
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   13
         Left            =   90
         TabIndex        =   300
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label10 
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
         Index           =   21
         Left            =   90
         TabIndex        =   299
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Ficolor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   4
         Left            =   1440
         TabIndex        =   298
         Top             =   945
         Width           =   285
      End
      Begin VB.Label ColLabel1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Color"
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
         Index           =   4
         Left            =   90
         TabIndex        =   297
         Top             =   945
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gradient border 2 reduced"
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
      Height          =   1725
      Index           =   31
      Left            =   135
      TabIndex        =   271
      Top             =   135
      Width           =   3645
      Begin VB.CommandButton Command4 
         Caption         =   "Swap colors"
         Height          =   330
         Index           =   3
         Left            =   2250
         TabIndex        =   322
         Top             =   1305
         Width           =   1275
      End
      Begin VB.HScrollBar HScroll8 
         Height          =   195
         Index           =   5
         LargeChange     =   10
         Left            =   1305
         Max             =   400
         Min             =   1
         TabIndex        =   274
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   5
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   273
         Top             =   630
         Value           =   1
         Width           =   1635
      End
      Begin VB.PictureBox Grad1 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   3
         Left            =   945
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   272
         Top             =   1350
         Width           =   1125
      End
      Begin VB.Label ColLabel1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Color"
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
         Index           =   3
         Left            =   90
         TabIndex        =   283
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label Ficolor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   3
         Left            =   1440
         TabIndex        =   282
         Top             =   945
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   19
         Left            =   90
         TabIndex        =   281
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   11
         Left            =   90
         TabIndex        =   280
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label16 
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
         Index           =   5
         Left            =   3015
         TabIndex        =   279
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label17 
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
         Index           =   5
         Left            =   3015
         TabIndex        =   278
         Top             =   630
         Width           =   510
      End
      Begin VB.Label ColLabel2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Color"
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
         Index           =   3
         Left            =   1890
         TabIndex        =   277
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label SColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   3
         Left            =   3240
         TabIndex        =   276
         Top             =   945
         Width           =   285
      End
      Begin VB.Label Label18 
         Caption         =   "Gradient"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   275
         Top             =   1350
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gradient border 2"
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
      Height          =   1725
      Index           =   30
      Left            =   135
      TabIndex        =   258
      Top             =   135
      Width           =   3645
      Begin VB.CommandButton Command4 
         Caption         =   "Swap colors"
         Height          =   330
         Index           =   2
         Left            =   2250
         TabIndex        =   321
         Top             =   1305
         Width           =   1275
      End
      Begin VB.PictureBox Grad1 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   2
         Left            =   990
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   261
         Top             =   1350
         Width           =   1125
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   4
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   260
         Top             =   630
         Value           =   1
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll8 
         Height          =   195
         Index           =   4
         LargeChange     =   10
         Left            =   1305
         Max             =   400
         Min             =   1
         TabIndex        =   259
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.Label Label18 
         Caption         =   "Gradient"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   270
         Top             =   1350
         Width           =   780
      End
      Begin VB.Label SColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   2
         Left            =   3240
         TabIndex        =   269
         Top             =   945
         Width           =   285
      End
      Begin VB.Label ColLabel2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Color"
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
         Left            =   1890
         TabIndex        =   268
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label Label17 
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
         Index           =   4
         Left            =   3015
         TabIndex        =   267
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label16 
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
         Index           =   4
         Left            =   3015
         TabIndex        =   266
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   10
         Left            =   90
         TabIndex        =   265
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label10 
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
         Index           =   18
         Left            =   90
         TabIndex        =   264
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Ficolor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   2
         Left            =   1440
         TabIndex        =   263
         Top             =   945
         Width           =   285
      End
      Begin VB.Label ColLabel1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Color"
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
         Left            =   90
         TabIndex        =   262
         Top             =   945
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gradient border 1 reduced"
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
      Height          =   1725
      Index           =   29
      Left            =   135
      TabIndex        =   245
      Top             =   135
      Width           =   3645
      Begin VB.CommandButton Command4 
         Caption         =   "Swap colors"
         Height          =   330
         Index           =   1
         Left            =   2250
         TabIndex        =   320
         Top             =   1305
         Width           =   1275
      End
      Begin VB.HScrollBar HScroll8 
         Height          =   195
         Index           =   3
         LargeChange     =   10
         Left            =   1305
         Max             =   400
         Min             =   1
         TabIndex        =   248
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   3
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   247
         Top             =   630
         Value           =   1
         Width           =   1635
      End
      Begin VB.PictureBox Grad1 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   1
         Left            =   990
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   246
         Top             =   1350
         Width           =   1125
      End
      Begin VB.Label ColLabel1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Color"
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
         Left            =   90
         TabIndex        =   257
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label Ficolor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   1
         Left            =   1440
         TabIndex        =   256
         Top             =   945
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   17
         Left            =   90
         TabIndex        =   255
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   9
         Left            =   90
         TabIndex        =   254
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label16 
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
         Index           =   3
         Left            =   3015
         TabIndex        =   253
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label17 
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
         Index           =   3
         Left            =   3015
         TabIndex        =   252
         Top             =   630
         Width           =   510
      End
      Begin VB.Label ColLabel2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Color"
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
         Left            =   1890
         TabIndex        =   251
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label SColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   1
         Left            =   3240
         TabIndex        =   250
         Top             =   945
         Width           =   285
      End
      Begin VB.Label Label18 
         Caption         =   "Gradient"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   249
         Top             =   1350
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gradient border 1"
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
      Height          =   1725
      Index           =   28
      Left            =   135
      TabIndex        =   232
      Top             =   135
      Width           =   3645
      Begin VB.CommandButton Command4 
         Caption         =   "Swap colors"
         Height          =   330
         Index           =   0
         Left            =   2250
         TabIndex        =   319
         Top             =   1305
         Width           =   1275
      End
      Begin VB.PictureBox Grad1 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   0
         Left            =   945
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   71
         TabIndex        =   241
         Top             =   1350
         Width           =   1125
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   2
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   234
         Top             =   630
         Value           =   1
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll8 
         Height          =   195
         Index           =   2
         LargeChange     =   10
         Left            =   1305
         Max             =   400
         Min             =   1
         TabIndex        =   233
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.Label Label18 
         Caption         =   "Gradient"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   244
         Top             =   1350
         Width           =   780
      End
      Begin VB.Label SColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   3240
         TabIndex        =   243
         Top             =   945
         Width           =   285
      End
      Begin VB.Label ColLabel2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Second Color"
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
         Left            =   1890
         TabIndex        =   242
         Top             =   945
         Width           =   1320
      End
      Begin VB.Label Label17 
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
         Index           =   2
         Left            =   3015
         TabIndex        =   240
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label16 
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
         Index           =   2
         Left            =   3015
         TabIndex        =   239
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   8
         Left            =   90
         TabIndex        =   238
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label10 
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
         Index           =   16
         Left            =   90
         TabIndex        =   237
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Ficolor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   1440
         TabIndex        =   236
         Top             =   945
         Width           =   285
      End
      Begin VB.Label ColLabel1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "First Color"
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
         Left            =   90
         TabIndex        =   235
         Top             =   945
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solid border reduced"
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
      Height          =   1320
      Index           =   27
      Left            =   135
      TabIndex        =   223
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll8 
         Height          =   195
         Index           =   1
         LargeChange     =   10
         Left            =   1305
         Max             =   400
         Min             =   1
         TabIndex        =   225
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   1
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   224
         Top             =   630
         Value           =   1
         Width           =   1635
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Border Color"
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
         Index           =   15
         Left            =   90
         TabIndex        =   231
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   15
         Left            =   1305
         TabIndex        =   230
         Top             =   945
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   15
         Left            =   90
         TabIndex        =   229
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   7
         Left            =   90
         TabIndex        =   228
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label16 
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
         Left            =   3015
         TabIndex        =   227
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label17 
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
         Left            =   3015
         TabIndex        =   226
         Top             =   630
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solid border"
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
      Height          =   1320
      Index           =   26
      Left            =   135
      TabIndex        =   214
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll9 
         Height          =   195
         Index           =   0
         Left            =   1305
         Max             =   10
         Min             =   1
         TabIndex        =   220
         Top             =   630
         Value           =   1
         Width           =   1635
      End
      Begin VB.HScrollBar HScroll8 
         Height          =   195
         Index           =   0
         LargeChange     =   10
         Left            =   1305
         Max             =   400
         Min             =   1
         TabIndex        =   219
         Top             =   315
         Value           =   1
         Width           =   1635
      End
      Begin VB.Label Label17 
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
         Left            =   3015
         TabIndex        =   222
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label16 
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
         Left            =   3015
         TabIndex        =   221
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   6
         Left            =   90
         TabIndex        =   218
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label10 
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
         Index           =   14
         Left            =   90
         TabIndex        =   217
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   14
         Left            =   1305
         TabIndex        =   216
         Top             =   945
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Border Color"
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
         Index           =   14
         Left            =   90
         TabIndex        =   215
         Top             =   945
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Abs vert. wave lines reversed"
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
      Height          =   1905
      Index           =   25
      Left            =   135
      TabIndex        =   199
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   21
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   203
         Top             =   630
         Value           =   2
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   13
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   202
         Top             =   1260
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   195
         Index           =   5
         Left            =   1260
         Max             =   25
         Min             =   1
         TabIndex        =   201
         Top             =   945
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll7 
         Height          =   195
         Index           =   5
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   200
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.Label Label1 
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
         Index           =   26
         Left            =   2715
         TabIndex        =   213
         Top             =   630
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude"
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
         Index           =   13
         Left            =   90
         TabIndex        =   212
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Index           =   13
         Left            =   90
         TabIndex        =   211
         Top             =   1575
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   13
         Left            =   1260
         TabIndex        =   210
         Top             =   1575
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   13
         Left            =   90
         TabIndex        =   209
         Top             =   1260
         Width           =   1140
      End
      Begin VB.Label Label11 
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
         Index           =   13
         Left            =   2715
         TabIndex        =   208
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave"
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
         Index           =   5
         Left            =   90
         TabIndex        =   207
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label13 
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
         Index           =   5
         Left            =   2715
         TabIndex        =   206
         Top             =   945
         Width           =   555
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   5
         Left            =   90
         TabIndex        =   205
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label15 
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
         Index           =   5
         Left            =   2715
         TabIndex        =   204
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Abs hor. wave lines reversed"
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
      Height          =   1905
      Index           =   24
      Left            =   135
      TabIndex        =   184
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll7 
         Height          =   195
         Index           =   4
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   188
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   195
         Index           =   4
         Left            =   1260
         Max             =   25
         Min             =   1
         TabIndex        =   187
         Top             =   945
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   12
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   186
         Top             =   1260
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   20
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   185
         Top             =   630
         Value           =   2
         Width           =   1410
      End
      Begin VB.Label Label15 
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
         Index           =   4
         Left            =   2715
         TabIndex        =   198
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   4
         Left            =   90
         TabIndex        =   197
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label13 
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
         Index           =   4
         Left            =   2715
         TabIndex        =   196
         Top             =   945
         Width           =   555
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave"
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
         Index           =   4
         Left            =   90
         TabIndex        =   195
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label11 
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
         Index           =   12
         Left            =   2715
         TabIndex        =   194
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label10 
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
         Index           =   12
         Left            =   90
         TabIndex        =   193
         Top             =   1260
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   12
         Left            =   1260
         TabIndex        =   192
         Top             =   1575
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Index           =   12
         Left            =   90
         TabIndex        =   191
         Top             =   1575
         Width           =   1140
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude"
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
         Index           =   12
         Left            =   90
         TabIndex        =   190
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Index           =   25
         Left            =   2715
         TabIndex        =   189
         Top             =   630
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add abs vertical wave lines"
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
      Height          =   1905
      Index           =   23
      Left            =   135
      TabIndex        =   169
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   19
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   173
         Top             =   630
         Value           =   2
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   11
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   172
         Top             =   1260
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   195
         Index           =   3
         Left            =   1260
         Max             =   25
         Min             =   1
         TabIndex        =   171
         Top             =   945
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll7 
         Height          =   195
         Index           =   3
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   170
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.Label Label1 
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
         Index           =   24
         Left            =   2715
         TabIndex        =   183
         Top             =   630
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude"
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
         Index           =   11
         Left            =   90
         TabIndex        =   182
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Index           =   11
         Left            =   90
         TabIndex        =   181
         Top             =   1575
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   11
         Left            =   1260
         TabIndex        =   180
         Top             =   1575
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   11
         Left            =   90
         TabIndex        =   179
         Top             =   1260
         Width           =   1140
      End
      Begin VB.Label Label11 
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
         Index           =   11
         Left            =   2715
         TabIndex        =   178
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave"
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
         Index           =   3
         Left            =   90
         TabIndex        =   177
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label13 
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
         Index           =   3
         Left            =   2715
         TabIndex        =   176
         Top             =   945
         Width           =   555
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   3
         Left            =   90
         TabIndex        =   175
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label15 
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
         Index           =   3
         Left            =   2715
         TabIndex        =   174
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add abs horizontal wave lines"
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
      Height          =   1905
      Index           =   22
      Left            =   135
      TabIndex        =   154
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll7 
         Height          =   195
         Index           =   2
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   158
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   195
         Index           =   2
         Left            =   1260
         Max             =   25
         Min             =   1
         TabIndex        =   157
         Top             =   945
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   10
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   156
         Top             =   1260
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   18
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   155
         Top             =   630
         Value           =   2
         Width           =   1410
      End
      Begin VB.Label Label15 
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
         Index           =   2
         Left            =   2715
         TabIndex        =   168
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Left            =   90
         TabIndex        =   167
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label13 
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
         Index           =   2
         Left            =   2715
         TabIndex        =   166
         Top             =   945
         Width           =   555
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave"
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
         Left            =   90
         TabIndex        =   165
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label11 
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
         Index           =   10
         Left            =   2715
         TabIndex        =   164
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label10 
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
         Index           =   10
         Left            =   90
         TabIndex        =   163
         Top             =   1260
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   10
         Left            =   1260
         TabIndex        =   162
         Top             =   1575
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Index           =   10
         Left            =   90
         TabIndex        =   161
         Top             =   1575
         Width           =   1140
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude"
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
         Index           =   10
         Left            =   90
         TabIndex        =   160
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Index           =   23
         Left            =   2715
         TabIndex        =   159
         Top             =   630
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add vertical wave lines"
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
      Height          =   1905
      Index           =   21
      Left            =   135
      TabIndex        =   139
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   17
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   143
         Top             =   630
         Value           =   2
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   9
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   142
         Top             =   1260
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   195
         Index           =   1
         Left            =   1260
         Max             =   25
         Min             =   1
         TabIndex        =   141
         Top             =   945
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll7 
         Height          =   195
         Index           =   1
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   140
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.Label Label1 
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
         Index           =   22
         Left            =   2715
         TabIndex        =   153
         Top             =   630
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude"
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
         Index           =   9
         Left            =   90
         TabIndex        =   152
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Index           =   9
         Left            =   90
         TabIndex        =   151
         Top             =   1575
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   9
         Left            =   1260
         TabIndex        =   150
         Top             =   1575
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   9
         Left            =   90
         TabIndex        =   149
         Top             =   1260
         Width           =   1140
      End
      Begin VB.Label Label11 
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
         Index           =   9
         Left            =   2715
         TabIndex        =   148
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave"
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
         Left            =   90
         TabIndex        =   147
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label13 
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
         Index           =   1
         Left            =   2715
         TabIndex        =   146
         Top             =   945
         Width           =   555
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Left            =   90
         TabIndex        =   145
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label15 
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
         Index           =   1
         Left            =   2715
         TabIndex        =   144
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add horizontal wave lines"
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
      Height          =   1905
      Index           =   20
      Left            =   135
      TabIndex        =   124
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll7 
         Height          =   195
         Index           =   0
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   137
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll6 
         Height          =   195
         Index           =   0
         Left            =   1260
         Max             =   25
         Min             =   1
         TabIndex        =   134
         Top             =   945
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   8
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   126
         Top             =   1260
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   16
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   125
         Top             =   630
         Value           =   2
         Width           =   1410
      End
      Begin VB.Label Label15 
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
         Index           =   0
         Left            =   2715
         TabIndex        =   138
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Left            =   90
         TabIndex        =   136
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label13 
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
         Index           =   0
         Left            =   2715
         TabIndex        =   135
         Top             =   945
         Width           =   555
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave"
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
         Left            =   90
         TabIndex        =   133
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label11 
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
         Index           =   8
         Left            =   2715
         TabIndex        =   132
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label10 
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
         Index           =   8
         Left            =   90
         TabIndex        =   131
         Top             =   1260
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   8
         Left            =   1260
         TabIndex        =   130
         Top             =   1575
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Index           =   8
         Left            =   90
         TabIndex        =   129
         Top             =   1575
         Width           =   1140
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude"
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
         Index           =   8
         Left            =   90
         TabIndex        =   128
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Index           =   21
         Left            =   2715
         TabIndex        =   127
         Top             =   630
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add crossed lines"
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
      Index           =   19
      Left            =   135
      TabIndex        =   115
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   15
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   117
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   7
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   116
         Top             =   675
         Value           =   1
         Width           =   1410
      End
      Begin VB.Label Label1 
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
         Index           =   20
         Left            =   2715
         TabIndex        =   123
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   7
         Left            =   90
         TabIndex        =   122
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Index           =   7
         Left            =   90
         TabIndex        =   121
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   7
         Left            =   1260
         TabIndex        =   120
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   7
         Left            =   90
         TabIndex        =   119
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label11 
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
         Index           =   7
         Left            =   2715
         TabIndex        =   118
         Top             =   675
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add diagonal left lines"
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
      Index           =   18
      Left            =   135
      TabIndex        =   106
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   6
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   108
         Top             =   675
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   14
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   107
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.Label Label11 
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
         Index           =   6
         Left            =   2715
         TabIndex        =   114
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label10 
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
         Index           =   6
         Left            =   90
         TabIndex        =   113
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   6
         Left            =   1260
         TabIndex        =   112
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Index           =   6
         Left            =   90
         TabIndex        =   111
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   6
         Left            =   90
         TabIndex        =   110
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Index           =   19
         Left            =   2715
         TabIndex        =   109
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add diagonal right lines"
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
      Index           =   17
      Left            =   135
      TabIndex        =   97
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   13
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   99
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   5
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   98
         Top             =   675
         Value           =   1
         Width           =   1410
      End
      Begin VB.Label Label1 
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
         Index           =   18
         Left            =   2715
         TabIndex        =   105
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   5
         Left            =   90
         TabIndex        =   104
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Index           =   5
         Left            =   90
         TabIndex        =   103
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   5
         Left            =   1260
         TabIndex        =   102
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   5
         Left            =   90
         TabIndex        =   101
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label11 
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
         Index           =   5
         Left            =   2715
         TabIndex        =   100
         Top             =   675
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add circles"
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
      Index           =   16
      Left            =   135
      TabIndex        =   88
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   4
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   90
         Top             =   675
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   12
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   89
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.Label Label11 
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
         Index           =   4
         Left            =   2715
         TabIndex        =   96
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label10 
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
         Index           =   4
         Left            =   90
         TabIndex        =   95
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   4
         Left            =   1260
         TabIndex        =   94
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Index           =   4
         Left            =   90
         TabIndex        =   93
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   4
         Left            =   90
         TabIndex        =   92
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Index           =   17
         Left            =   2715
         TabIndex        =   91
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add boxes"
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
      Index           =   15
      Left            =   135
      TabIndex        =   79
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   11
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   81
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   3
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   80
         Top             =   675
         Value           =   1
         Width           =   1410
      End
      Begin VB.Label Label1 
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
         Index           =   16
         Left            =   2715
         TabIndex        =   87
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Index           =   3
         Left            =   90
         TabIndex        =   86
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Index           =   3
         Left            =   90
         TabIndex        =   85
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   3
         Left            =   1260
         TabIndex        =   84
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Index           =   3
         Left            =   90
         TabIndex        =   83
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label11 
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
         Index           =   3
         Left            =   2715
         TabIndex        =   82
         Top             =   675
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add squares"
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
      Index           =   14
      Left            =   135
      TabIndex        =   70
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   2
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   72
         Top             =   675
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   10
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   1
         TabIndex        =   71
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.Label Label11 
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
         Index           =   2
         Left            =   2715
         TabIndex        =   78
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label10 
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
         Left            =   90
         TabIndex        =   77
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   2
         Left            =   1260
         TabIndex        =   76
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Left            =   90
         TabIndex        =   75
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Left            =   90
         TabIndex        =   74
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Index           =   15
         Left            =   2715
         TabIndex        =   73
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add vertical lines"
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
      Index           =   13
      Left            =   135
      TabIndex        =   61
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   9
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   2
         SmallChange     =   2
         TabIndex        =   63
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   1
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   62
         Top             =   675
         Value           =   1
         Width           =   1410
      End
      Begin VB.Label Label1 
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
         Index           =   14
         Left            =   2715
         TabIndex        =   69
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Left            =   90
         TabIndex        =   68
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Left            =   90
         TabIndex        =   67
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   1
         Left            =   1260
         TabIndex        =   66
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label Label10 
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
         Left            =   90
         TabIndex        =   65
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label11 
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
         Index           =   1
         Left            =   2715
         TabIndex        =   64
         Top             =   675
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add horizontal lines"
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
      Index           =   12
      Left            =   135
      TabIndex        =   52
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll5 
         Height          =   195
         Index           =   0
         Left            =   1260
         Max             =   10
         Min             =   1
         TabIndex        =   59
         Top             =   675
         Value           =   1
         Width           =   1410
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   8
         LargeChange     =   10
         Left            =   1260
         Max             =   248
         Min             =   2
         SmallChange     =   2
         TabIndex        =   53
         Top             =   315
         Value           =   2
         Width           =   1410
      End
      Begin VB.Label Label11 
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
         Index           =   0
         Left            =   2715
         TabIndex        =   60
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label10 
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
         Index           =   0
         Left            =   90
         TabIndex        =   58
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   1260
         TabIndex        =   57
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Line Color"
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
         Left            =   90
         TabIndex        =   56
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distance"
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
         Left            =   90
         TabIndex        =   55
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Index           =   13
         Left            =   2715
         TabIndex        =   54
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bumped vertical blinds"
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
      Height          =   1185
      Index           =   11
      Left            =   135
      TabIndex        =   49
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   7
         LargeChange     =   10
         Left            =   315
         Max             =   248
         Min             =   2
         SmallChange     =   2
         TabIndex        =   50
         Top             =   450
         Value           =   2
         Width           =   2355
      End
      Begin VB.Label Label1 
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
         Index           =   12
         Left            =   2715
         TabIndex        =   51
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bumped horizontal blinds"
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
      Height          =   1185
      Index           =   10
      Left            =   135
      TabIndex        =   46
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   6
         LargeChange     =   10
         Left            =   315
         Max             =   248
         Min             =   2
         SmallChange     =   2
         TabIndex        =   47
         Top             =   450
         Value           =   2
         Width           =   2355
      End
      Begin VB.Label Label1 
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
         Index           =   11
         Left            =   2715
         TabIndex        =   48
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vertical blinds"
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
      Height          =   1185
      Index           =   9
      Left            =   135
      TabIndex        =   41
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   5
         LargeChange     =   10
         Left            =   315
         Max             =   249
         Min             =   1
         TabIndex        =   42
         Top             =   450
         Value           =   1
         Width           =   2355
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Reverse blinds"
         Height          =   240
         Left            =   315
         TabIndex        =   44
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         Index           =   10
         Left            =   2715
         TabIndex        =   43
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Horizontal blinds"
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
      Height          =   1185
      Index           =   8
      Left            =   135
      TabIndex        =   38
      Top             =   135
      Width           =   3645
      Begin VB.CheckBox Check2 
         Caption         =   "Reverse blinds"
         Height          =   195
         Left            =   315
         TabIndex        =   45
         Top             =   765
         Width           =   1860
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   4
         LargeChange     =   10
         Left            =   315
         Max             =   249
         Min             =   1
         TabIndex        =   39
         Top             =   450
         Value           =   1
         Width           =   2355
      End
      Begin VB.Label Label1 
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
         Index           =   9
         Left            =   2715
         TabIndex        =   40
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjust fog"
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
      Height          =   870
      Index           =   7
      Left            =   135
      TabIndex        =   35
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   3
         LargeChange     =   2
         Left            =   315
         Max             =   25
         Min             =   1
         TabIndex        =   36
         Top             =   450
         Value           =   1
         Width           =   2355
      End
      Begin VB.Label Label1 
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
         Index           =   8
         Left            =   2715
         TabIndex        =   37
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjust blowing"
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
      Height          =   870
      Index           =   6
      Left            =   135
      TabIndex        =   32
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   2
         LargeChange     =   2
         Left            =   315
         Max             =   10
         Min             =   1
         TabIndex        =   33
         Top             =   450
         Value           =   1
         Width           =   2355
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   2715
         TabIndex        =   34
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjust eroding"
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
      Height          =   870
      Index           =   5
      Left            =   135
      TabIndex        =   29
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   1
         LargeChange     =   2
         Left            =   315
         Max             =   16
         Min             =   1
         TabIndex        =   30
         Top             =   450
         Value           =   1
         Width           =   2355
      End
      Begin VB.Label Label1 
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
         Index           =   6
         Left            =   2715
         TabIndex        =   31
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjust diffusing"
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
      Height          =   870
      Index           =   4
      Left            =   135
      TabIndex        =   26
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll4 
         Height          =   195
         Index           =   0
         LargeChange     =   2
         Left            =   315
         Max             =   16
         Min             =   1
         TabIndex        =   27
         Top             =   450
         Value           =   1
         Width           =   2355
      End
      Begin VB.Label Label1 
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
         Index           =   5
         Left            =   2715
         TabIndex        =   28
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjust Contrast"
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
      Height          =   870
      Index           =   3
      Left            =   135
      TabIndex        =   23
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Index           =   4
         LargeChange     =   10
         Left            =   315
         Max             =   50
         Min             =   1
         TabIndex        =   24
         Top             =   450
         Value           =   1
         Width           =   2355
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000%"
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
         Index           =   4
         Left            =   2715
         TabIndex        =   25
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjust Brightness"
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
      Height          =   870
      Index           =   1
      Left            =   135
      TabIndex        =   13
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Index           =   3
         LargeChange     =   10
         Left            =   315
         Max             =   100
         Min             =   -100
         TabIndex        =   14
         Top             =   450
         Width           =   2355
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000%"
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
         Index           =   3
         Left            =   2715
         TabIndex        =   15
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjust"
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
      Index           =   2
      Left            =   135
      TabIndex        =   16
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll3 
         Height          =   195
         LargeChange     =   10
         Left            =   855
         Max             =   100
         Min             =   -100
         TabIndex        =   22
         Top             =   765
         Width           =   2085
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   195
         LargeChange     =   10
         Left            =   855
         Max             =   100
         Min             =   -100
         TabIndex        =   17
         Top             =   450
         Width           =   2085
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Width"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   405
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Left            =   2985
         TabIndex        =   19
         Top             =   765
         Width           =   555
      End
      Begin VB.Label Label3 
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
         Left            =   2970
         TabIndex        =   18
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2925
      TabIndex        =   12
      Top             =   2115
      Width           =   825
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply!"
      Height          =   330
      Left            =   1170
      TabIndex        =   11
      Top             =   2115
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show me"
      Height          =   330
      Left            =   180
      TabIndex        =   10
      Top             =   2115
      Width           =   825
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adjust Colors"
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
      Height          =   1410
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   3645
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Index           =   2
         LargeChange     =   10
         Left            =   630
         Max             =   100
         Min             =   -100
         TabIndex        =   8
         Top             =   990
         Width           =   2355
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Index           =   1
         LargeChange     =   10
         Left            =   630
         Max             =   100
         Min             =   -100
         TabIndex        =   6
         Top             =   720
         Width           =   2355
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Index           =   0
         LargeChange     =   10
         Left            =   630
         Max             =   100
         Min             =   -100
         TabIndex        =   1
         Top             =   450
         Width           =   2355
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000%"
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
         Height          =   240
         Index           =   2
         Left            =   3030
         TabIndex        =   9
         Top             =   990
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Index           =   1
         Left            =   3030
         TabIndex        =   7
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Blue"
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
         Height          =   240
         Left            =   45
         TabIndex        =   5
         Top             =   990
         Width           =   555
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   45
         TabIndex        =   4
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Red"
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
         Left            =   45
         TabIndex        =   3
         Top             =   450
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000%"
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
         Left            =   3030
         TabIndex        =   2
         Top             =   450
         Width           =   555
      End
   End
End
Attribute VB_Name = "FColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private AdS As Boolean

Private Sub Command1_Click() 'show me
FMain.Pic1.Picture = PicMem
Select Case Col
Case 0
ColorComp Xcor0, Ycor0, Xcor1, Ycor1, HScroll1(0).Value, HScroll1(1).Value, HScroll1(2).Value
Case 1
ColorComp Xcor0, Ycor0, Xcor1, Ycor1, HScroll1(3).Value, HScroll1(3).Value, HScroll1(3).Value
Case 3
ContrastPic Xcor0, Ycor0, Xcor1, Ycor1, HScroll1(4).Value
Case 4
DiffusePic Xcor0, Ycor0, Xcor1, Ycor1, HScroll4(0).Value
Case 5
ErodePic Xcor0, Ycor0, Xcor1, Ycor1, HScroll4(1).Value
Case 6
BlowPic Xcor0, Ycor0, Xcor1, Ycor1, 11 - HScroll4(2).Value
Case 7
FogPic Xcor0, Ycor0, Xcor1, Ycor1, HScroll4(3).Value * 15
Case 8
If Check2.Value = 0 Then
Blinds Xcor0, Ycor0, Xcor1, Ycor1, HScroll4(4).Value, False
Else
Blinds Xcor0, Ycor0, Xcor1, Ycor1, HScroll4(4).Value, True
End If
Case 9
If Check1.Value = 0 Then
Blinds2 Xcor0, Ycor0, Xcor1, Ycor1, HScroll4(5).Value, False
Else
Blinds2 Xcor0, Ycor0, Xcor1, Ycor1, HScroll4(5).Value, True
End If
Case 10
Blinds3 Xcor0, Ycor0, Xcor1, Ycor1, HScroll4(6).Value
Case 11
Blinds4 Xcor0, Ycor0, Xcor1, Ycor1, HScroll4(7).Value
Case 12 'add H lines
HLines Xcor0, Ycor0, Xcor1, Ycor1, HScroll4(8).Value, HScroll5(0).Value, Label9(0).BackColor
Case 13 'add V lines
VLines Xcor0, Ycor0, Xcor1, Ycor1, HScroll4(9).Value, HScroll5(1).Value, Label9(1).BackColor
Case 14 'add squares
Squares Xcor0, Ycor0, Xcor1, Ycor1, HScroll4(10).Value, HScroll5(2).Value, Label9(2).BackColor
Case 15 'add boxes
AddBoxes 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, HScroll4(11).Value, HScroll5(3).Value, Label9(3).BackColor
Case 16 'add circles
AddCircles 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, HScroll4(12).Value, HScroll5(4).Value, Label9(4).BackColor
Case 17 'add dia R lines
AddDiaRLines 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, HScroll4(13).Value, HScroll5(5).Value, Label9(5).BackColor
Case 18 'add dia R lines
AddDiaLLines 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, HScroll4(14).Value, HScroll5(6).Value, Label9(6).BackColor
Case 19 'add crossed lines
AddCrossLines 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, HScroll4(15).Value, HScroll5(7).Value, Label9(7).BackColor
Case 20 'add H wave lines
SinusLineX 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, HScroll5(8).Value, HScroll6(0).Value, HScroll4(16).Value, Label9(8).BackColor, HScroll7(0).Value, 0
Case 21 'add V wave lines
SinusLineY 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, HScroll5(9).Value, HScroll6(1).Value, HScroll4(17).Value, Label9(9).BackColor, HScroll7(1).Value, 0
Case 22 'add abs H wave lines
SinusLineX 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, HScroll5(10).Value, HScroll6(2).Value, HScroll4(18).Value, Label9(10).BackColor, HScroll7(2).Value, 1
Case 23 'add abs V wave lines
SinusLineY 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, HScroll5(11).Value, HScroll6(3).Value, HScroll4(19).Value, Label9(11).BackColor, HScroll7(3).Value, 1
Case 24 'add abs H wave lines reversed
SinusLineX 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, HScroll5(12).Value, HScroll6(4).Value, HScroll4(20).Value, Label9(12).BackColor, HScroll7(4).Value, 2
Case 25 'add abs V wave lines
SinusLineY 0, 0, FMain.Pic1.Width, FMain.Pic1.Height, HScroll5(13).Value, HScroll6(5).Value, HScroll4(21).Value, Label9(13).BackColor, HScroll7(5).Value, 2
'--------------------
Case 26 'solid border
SBorder HScroll8(0).Value, HScroll9(0).Value, Label9(14).BackColor, False
Case 27 'solid border reduced
SBorder HScroll8(1).Value, HScroll9(1).Value, Label9(15).BackColor, True
Case 28 'gradient border 1
GBorder1 HScroll8(2).Value, HScroll9(2).Value, Ficolor(0).BackColor, SColor(0).BackColor, False
Case 29 'gradient border 1 reduced
GBorder1 HScroll8(3).Value, HScroll9(3).Value, Ficolor(1).BackColor, SColor(1).BackColor, True
Case 30 'gradient border 2
GBorder2 HScroll8(4).Value, HScroll9(4).Value, Ficolor(2).BackColor, SColor(2).BackColor, False
Case 31 'gradient border 2 reduced
GBorder2 HScroll8(5).Value, HScroll9(5).Value, Ficolor(3).BackColor, SColor(3).BackColor, True
Case 32 'solid circular border
CBorder HScroll8(6).Value, HScroll9(6).Value, Label9(16).BackColor
Case 33 'gradient circular border 1
GCBorder1 HScroll8(7).Value, HScroll9(7).Value, Ficolor(4).BackColor, SColor(4).BackColor
Case 34 'gradient circular border 2
GCBorder2 HScroll8(8).Value, HScroll9(8).Value, Ficolor(5).BackColor, SColor(5).BackColor
'--------------------
Case 35 ' mix solid color
MixSolid HScroll9(9), Label9(17).BackColor
Case 36 ' mix gradient 1
MixGradient1 HScroll9(10), Ficolor(6).BackColor, SColor(6).BackColor, Check3(0).Value
Case 37 ' mix gradient 2
MixGradient2 HScroll9(11), Ficolor(7).BackColor, SColor(7).BackColor, Check3(1).Value
Case 38 ' mix box gradient 1
MixBoxGradient1 HScroll9(12), Ficolor(8).BackColor, SColor(8).BackColor
Case 39 ' mix box gradient 2
MixBoxGradient2 HScroll9(13), Ficolor(9).BackColor, SColor(9).BackColor
Case 40 ' mix circle gradient 2
GCircle1 HScroll9(14), Ficolor(10).BackColor, SColor(10).BackColor
Case 41 ' mix circle gradient 2
GCircle2 HScroll9(15), Ficolor(11).BackColor, SColor(11).BackColor
'--------------------
Case 42 'mozaic
Mozaic Xcor0, Ycor0, Xcor1, Ycor1, HScroll10(0).Value
Case 43 'blurred mozaic
Mozaic2 Xcor0, Ycor0, Xcor1, Ycor1, HScroll10(1).Value
Case 44 ' wave X
EffectX HScroll11(0).Value, HScroll12(0).Value, 0
Case 45 ' abs wave X
EffectX HScroll11(1).Value, HScroll12(1).Value, 1
Case 46 ' wave Y
EffectY HScroll11(2).Value, HScroll12(2).Value, 0
Case 47 ' abs wave Y
EffectY HScroll11(3).Value, HScroll12(3).Value, 1
Case 49 'tile pic
Tile HScroll14(6), HScroll14(7)
End Select
Command2.Enabled = True
Set Im = FMain.Pic1.Image
End Sub

Private Sub Command2_Click() 'apply
FMain.Pic1.Picture = PicMem
SaveRedo
FMain.Pic1 = Im
FColor.Hide
End Sub

Private Sub Command3_Click() 'cancel
FMain.Pic1.Picture = PicMem
FColor.Hide
End Sub

Private Sub ColLabel1_Click(Index As Integer)
Ficolor_Click Index
End Sub

Private Sub ColLabel2_Click(Index As Integer)
SColor_Click Index
End Sub

Private Sub Command4_Click(Index As Integer)
Dim TempCol&
TempCol = Ficolor(Index).BackColor
Ficolor(Index).BackColor = SColor(Index).BackColor
SColor(Index).BackColor = TempCol
SetGrad Index
End Sub

Private Sub Ficolor_Click(Index As Integer)
On Error GoTo FLabelExit
FMain.CD1.Flags = 3
FMain.CD1.Color = Ficolor(Index).BackColor
FMain.CD1.ShowColor
Ficolor(Index).BackColor = FMain.CD1.Color
SetGrad (Index)
FLabelExit:
End Sub

Private Sub HScroll10_Change(Index As Integer)
Label20(Index).Caption = Format(HScroll10(Index).Value, "00")
End Sub

Private Sub HScroll11_Change(Index As Integer)
Label21(Index).Caption = Format(HScroll11(Index).Value, "000")
End Sub

Private Sub HScroll12_Change(Index As Integer)
Label22(Index).Caption = Format(HScroll12(Index).Value / 10, "00.0")
End Sub

Private Sub HScroll13_Change(Index As Integer)
Label50(Index).Caption = Format(HScroll13(Index).Value, "000")
AdjustSelection
End Sub

Private Sub AdjustSelection()
If AdS = False Then Exit Sub
FMain.Shape1.Move HScroll13(0).Value, HScroll13(1).Value, HScroll13(2).Value, HScroll13(3).Value
Xcor0 = HScroll13(0).Value
Ycor0 = HScroll13(1).Value
Xcor1 = HScroll13(2).Value
Ycor1 = HScroll13(3).Value
SetCoordinates
End Sub

Private Sub HScroll14_Change(Index As Integer)
Label50(Index).Caption = Format(HScroll14(Index).Value, "00")
End Sub

Private Sub SColor_Click(Index As Integer)
On Error GoTo SLabelExit
FMain.CD1.Flags = 3
FMain.CD1.Color = SColor(Index).BackColor
FMain.CD1.ShowColor
SColor(Index).BackColor = FMain.CD1.Color
SetGrad (Index)
SLabelExit:
End Sub

Private Sub Form_Activate()
On Error Resume Next
For xx = 0 To 49
Frame1(xx).Visible = False
Next xx
Set PicMem = FMain.Pic1.Image
FColor.Move 0, 330, 4020, 2865
Command2.Enabled = False
Frame1(Col).Visible = True
If Col = 48 Then
AdS = False
Command1.Enabled = False
HScroll13(0).Min = -FMain.Pic1.Width
HScroll13(0).Max = 2 * FMain.Pic1.Width
HScroll13(0).Value = FMain.Shape1.Left
HScroll13(1).Min = -FMain.Pic1.Height
HScroll13(1).Max = 2 * FMain.Pic1.Height
HScroll13(1).Value = FMain.Shape1.Top
HScroll13(2).Min = 1
HScroll13(2).Max = FMain.Pic1.Width
HScroll13(2).Value = FMain.Shape1.Width
HScroll13(3).Min = 0
HScroll13(3).Max = FMain.Pic1.Height
HScroll13(3).Value = FMain.Shape1.Height
FMain.Shape1.Move HScroll13(0).Value, HScroll13(1).Value, HScroll13(2).Value, HScroll13(3).Value
AdS = True
Else
Command1.Enabled = True
End If

For xx = 0 To 3
HScroll1(xx).Value = 0
Next xx
HScroll1(xx).Value = 1
HScroll2.Min = 0
HScroll3.Min = 0
HScroll2.Value = 0
HScroll3.Value = 0
HScroll2.Max = FMain.Pic1.Width * 3
HScroll3.Max = FMain.Pic1.Height * 3
HScroll2.Value = FMain.Pic1.Width
HScroll3.Value = FMain.Pic1.Height
For xx = 0 To 3
HScroll4(xx).Value = 5
Next xx
For xx = 4 To 15
HScroll4(xx).Value = 10
Next xx
HScroll4(6).Value = 20
HScroll4(7).Value = 20
For xx = 16 To 21
HScroll4(xx).Value = 30
Next xx
For xx = 0 To 13
HScroll5(xx).Value = 5
Next xx
For xx = 0 To 5
HScroll6(xx).Value = 5
HScroll7(xx).Value = 20
Next xx
Check1.Value = 0
Check2.Value = 0
'-----------------------
For xx = 0 To 8
HScroll8(xx).Value = 20
Next xx
For xx = 0 To 15
HScroll9(xx).Value = 5
Next xx
Check3(0).Value = 0
Check3(1).Value = 0
Dim qq%
For qq = 0 To 11
SetGrad qq
Next qq
HScroll10(0).Value = 8
HScroll10(1).Value = 8
For xx = 0 To 3
HScroll11(xx).Value = 50
HScroll12(xx).Value = 10
Next xx
HScroll14(6).Value = 3
HScroll14(7).Value = 3
End Sub

Private Sub Form_Load()
FColor.Move 0, 330, 4020, 2535
End Sub

Private Sub HScroll1_Change(Index As Integer) 'add color
Label1(Index).Caption = Format(HScroll1(Index).Value, "000") & "%"
End Sub

Private Sub HScroll2_Change()
On Error Resume Next
Label3(0).Caption = Format(HScroll2.Value, "000")
Factor = HScroll2.Value / OrWidth
HScroll3.Value = OrHeight * Factor
End Sub

Private Sub HScroll3_Change()
On Error Resume Next
Label3(1).Caption = Format(HScroll3.Value, "000")
Factor = HScroll3.Value / OrHeight
HScroll2.Value = OrWidth * Factor
End Sub

Private Sub HScroll4_Change(Index As Integer)
Label1(Index + 5).Caption = Format(HScroll4(Index).Value, "00")
End Sub

Private Sub HScroll5_Change(Index As Integer)
Label11(Index).Caption = Format(HScroll5(Index).Value / 10, "0.0")
End Sub

Private Sub HScroll6_Change(Index As Integer)
Label13(Index).Caption = Format(HScroll6(Index).Value, "00")
End Sub

Private Sub HScroll7_Change(Index As Integer)
Label15(Index).Caption = Format(HScroll7(Index).Value, "000")
End Sub

Private Sub HScroll8_Change(Index As Integer)
Label16(Index).Caption = Format(HScroll8(Index).Value, "000")
End Sub

Private Sub HScroll9_Change(Index As Integer)
Label17(Index).Caption = Format(HScroll9(Index).Value / 10, "0.0")
End Sub

Private Sub Label8_Click(Index As Integer)
Label9_Click (Index)
End Sub

Private Sub Label9_Click(Index As Integer)
On Error GoTo Label9Exit
FMain.CD1.Flags = 3
FMain.CD1.Color = Label9(Index).BackColor
FMain.CD1.ShowColor
Label9(Index).BackColor = FMain.CD1.Color
Label9Exit:
End Sub

Private Sub SetGrad(Number%)
On Error Resume Next
Dim Gr!, Gg!, Gb!, Gr1&, Gg1&, Gb1&, H%, Ri, Gi, Bi
H = Grad1(Number).ScaleWidth
Gr = Ficolor(Number).BackColor Mod 256&
Gg = ((Ficolor(Number).BackColor And &HFF00) / 256&) Mod 256&
Gb = (Ficolor(Number).BackColor And &HFF0000) / 65536
Gr1 = SColor(Number).BackColor Mod 256&
Gg1 = ((SColor(Number).BackColor And &HFF00) / 256&) Mod 256&
Gb1 = (SColor(Number).BackColor And &HFF0000) / 65536
Ri = (Gr1 - Gr) / H
Gi = (Gg1 - Gg) / H
Bi = (Gb1 - Gb) / H
For xx = 0 To H - 1
Grad1(Number).Line (xx, 0)-(xx, Grad1(Number).ScaleHeight), RGB(Gr, Gg, Gb)
Gr = Gr + Ri
Gg = Gg + Gi
Gb = Gb + Bi
Next xx
Grad1(Number).Refresh
End Sub

