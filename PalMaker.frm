VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM Color Picker"
   ClientHeight    =   7815
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9660
   FillColor       =   &H80000000&
   Icon            =   "PalMaker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdWebBar 
      Caption         =   "Web Scroll Bar"
      Height          =   315
      Left            =   2790
      TabIndex        =   85
      Top             =   5385
      Width           =   1290
   End
   Begin VB.PictureBox PicScrollbar 
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   225
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   5850
      Width           =   3945
      Begin VB.CommandButton cmdCopy1 
         Height          =   375
         Left            =   2940
         Picture         =   "PalMaker.frx":23D2
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Copy CSS Code"
         Top             =   1065
         Width           =   900
      End
      Begin VB.PictureBox PicsBar 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   6
         Left            =   945
         ScaleHeight     =   225
         ScaleWidth      =   1350
         TabIndex        =   83
         Top             =   1620
         Width           =   1350
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trakbar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   6
            Left            =   60
            TabIndex        =   84
            Top             =   0
            Width           =   555
         End
      End
      Begin VB.PictureBox PicsBar 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   5
         Left            =   945
         ScaleHeight     =   225
         ScaleWidth      =   1350
         TabIndex        =   81
         Top             =   1365
         Width           =   1350
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dark Shadow"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   60
            TabIndex        =   82
            Top             =   0
            Width           =   990
         End
      End
      Begin VB.PictureBox PicsBar 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   4
         Left            =   945
         ScaleHeight     =   225
         ScaleWidth      =   1350
         TabIndex        =   79
         Top             =   1110
         Width           =   1350
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shadow"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   60
            TabIndex        =   80
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox PicsBar 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   3
         Left            =   945
         ScaleHeight     =   225
         ScaleWidth      =   1350
         TabIndex        =   77
         Top             =   855
         Width           =   1350
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3D Highlight"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   60
            TabIndex        =   78
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.PictureBox PicsBar 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   2
         Left            =   945
         ScaleHeight     =   225
         ScaleWidth      =   1350
         TabIndex        =   75
         Top             =   600
         Width           =   1350
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Highlight"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   60
            TabIndex        =   76
            Top             =   0
            Width           =   600
         End
      End
      Begin VB.PictureBox PicsBar 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   1
         Left            =   945
         ScaleHeight     =   225
         ScaleWidth      =   1350
         TabIndex        =   73
         Top             =   345
         Width           =   1350
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Face"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   60
            TabIndex        =   74
            Top             =   0
            Width           =   360
         End
      End
      Begin VB.PictureBox PicsBar 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   945
         ScaleHeight     =   225
         ScaleWidth      =   1350
         TabIndex        =   71
         Top             =   90
         Width           =   1350
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Arrrow"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   72
            Top             =   0
            Width           =   540
         End
      End
      Begin VB.PictureBox PicBar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   120
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   70
         Top             =   150
         Width           =   600
      End
      Begin VB.PictureBox PicHolder 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   120
         Picture         =   "PalMaker.frx":245D
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   69
         Top             =   150
         Width           =   600
      End
      Begin VB.CommandButton cmdSave3 
         Height          =   375
         Left            =   2940
         Picture         =   "PalMaker.frx":39A7
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Save"
         Top             =   570
         Width           =   900
      End
      Begin VB.CommandButton cmdOpen3 
         Height          =   375
         Left            =   2940
         Picture         =   "PalMaker.frx":3A9A
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Open"
         Top             =   60
         Width           =   900
      End
   End
   Begin VB.PictureBox PicPayLayout 
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   225
      ScaleHeight     =   1875
      ScaleWidth      =   3945
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   5850
      Width           =   3945
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1860
         Left            =   0
         ScaleHeight     =   1830
         ScaleWidth      =   2835
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   0
         Width           =   2865
         Begin VB.PictureBox PicWebStyle 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   330
            Index           =   3
            Left            =   0
            ScaleHeight     =   330
            ScaleWidth      =   2835
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   1500
            Width           =   2835
            Begin VB.Label lblWebStyle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Footer"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   60
               TabIndex        =   65
               Top             =   30
               Width           =   540
            End
         End
         Begin VB.PictureBox PicWebStyle 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1065
            Index           =   2
            Left            =   1200
            ScaleHeight     =   1065
            ScaleWidth      =   1635
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   420
            Width           =   1635
            Begin VB.Label lblWebStyle 
               BackStyle       =   0  'Transparent
               Caption         =   "Text, Text, Text, Text, Text, Text, Text, Text, Text"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   3
               Left            =   60
               TabIndex        =   64
               Top             =   45
               Width           =   1485
            End
         End
         Begin VB.PictureBox PicWebStyle 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1065
            Index           =   1
            Left            =   0
            ScaleHeight     =   1065
            ScaleWidth      =   1185
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   420
            Width           =   1185
            Begin VB.Label lblWebStyle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HpyerLinks"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   45
               TabIndex        =   63
               Top             =   315
               Width           =   945
            End
            Begin VB.Label lblWebStyle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HpyerLinks"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   45
               TabIndex        =   62
               Top             =   105
               Width           =   945
            End
         End
         Begin VB.PictureBox PicWebStyle 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   405
            Index           =   0
            Left            =   0
            ScaleHeight     =   405
            ScaleWidth      =   2835
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   0
            Width           =   2835
            Begin VB.Label lblWebStyle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Header"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   0
               Left            =   45
               TabIndex        =   61
               Top             =   0
               Width           =   1140
            End
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   2835
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Line Line3 
            X1              =   1185
            X2              =   1185
            Y1              =   405
            Y2              =   1485
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   2835
            Y1              =   405
            Y2              =   405
         End
      End
      Begin VB.CommandButton cmdOpen 
         Height          =   375
         Left            =   2940
         Picture         =   "PalMaker.frx":3C13
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Open"
         Top             =   60
         Width           =   900
      End
      Begin VB.CommandButton cmdSave1 
         Height          =   375
         Left            =   2940
         Picture         =   "PalMaker.frx":3D8C
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Save"
         Top             =   570
         Width           =   900
      End
   End
   Begin VB.CheckBox chkHideWebStyle 
      Caption         =   "Hide WebStyles"
      Height          =   195
      Left            =   1740
      TabIndex        =   23
      Top             =   5055
      Width           =   1770
   End
   Begin VB.PictureBox PicWebStyle1 
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   225
      ScaleHeight     =   1875
      ScaleWidth      =   3945
      TabIndex        =   45
      Top             =   5850
      Width           =   3945
      Begin VB.CommandButton cmdSave2 
         Height          =   375
         Left            =   2940
         Picture         =   "PalMaker.frx":3E7F
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Save"
         Top             =   570
         Width           =   900
      End
      Begin VB.CommandButton cmdOpen1 
         Height          =   375
         Left            =   2940
         Picture         =   "PalMaker.frx":3F72
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Open"
         Top             =   60
         Width           =   900
      End
      Begin VB.PictureBox PicWebLinks 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1860
         Left            =   0
         ScaleHeight     =   1830
         ScaleWidth      =   2835
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   15
         Width           =   2865
         Begin VB.Label lblLinkObj 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Welcome"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   840
            TabIndex        =   49
            Top             =   90
            Width           =   1125
         End
         Begin VB.Label lblLinkObj 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ActiveLink = #000000"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   510
            TabIndex        =   48
            Top             =   645
            Width           =   1890
         End
         Begin VB.Label lblLinkObj 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VistedLink = #990033"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   195
            Index           =   2
            Left            =   510
            TabIndex        =   47
            Top             =   975
            Width           =   1890
         End
      End
   End
   Begin VB.CommandButton cmdPageStyle 
      Caption         =   "Page Layout"
      Height          =   315
      Left            =   1425
      TabIndex        =   44
      Top             =   5385
      Width           =   1215
   End
   Begin VB.CommandButton cmdWebLinks 
      Caption         =   "Web Links"
      Height          =   315
      Left            =   210
      TabIndex        =   43
      Top             =   5385
      Width           =   1110
   End
   Begin Project1.Flat2 Flat23 
      Height          =   495
      Left            =   60
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2280
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   873
   End
   Begin Project1.Flat2 Flat22 
      Height          =   2130
      Left            =   60
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2805
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   3757
   End
   Begin VB.CheckBox chkWebSafe 
      Height          =   375
      Left            =   3285
      Picture         =   "PalMaker.frx":40EB
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Show websafe Colors"
      Top             =   2340
      Width           =   900
   End
   Begin VB.PictureBox p2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   4065
      ScaleHeight     =   315
      ScaleWidth      =   360
      TabIndex        =   36
      Top             =   2790
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox PicPallet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   2040
      MousePointer    =   99  'Custom
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   34
      Top             =   165
      Width           =   3840
   End
   Begin VB.PictureBox PicCol 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4545
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   33
      Top             =   2745
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdDialog 
      Height          =   375
      Left            =   135
      Picture         =   "PalMaker.frx":4364
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Color from Dialog"
      Top             =   2340
      Width           =   900
   End
   Begin VB.PictureBox PicGrab 
      Height          =   1860
      Left            =   5340
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   270
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2985
      Width           =   4110
      Begin VB.PictureBox Picture1 
         Height          =   1830
         Left            =   3015
         ScaleHeight     =   1770
         ScaleWidth      =   990
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   -30
         Width           =   1050
         Begin VB.CheckBox cmdWebSafe 
            Height          =   375
            Left            =   60
            Picture         =   "PalMaker.frx":460E
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Show websafe Colors"
            Top             =   1365
            Width           =   900
         End
         Begin VB.CommandButton cmdzoom 
            Height          =   375
            Left            =   60
            Picture         =   "PalMaker.frx":4887
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Zoom"
            Top             =   930
            Width           =   900
         End
         Begin VB.CommandButton cmdOpen2 
            Height          =   375
            Left            =   60
            Picture         =   "PalMaker.frx":49F1
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Open Picture"
            Top             =   30
            Width           =   900
         End
         Begin VB.CommandButton cmdSave 
            Height          =   375
            Left            =   60
            Picture         =   "PalMaker.frx":4B6A
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Save Picture"
            Top             =   480
            Width           =   900
         End
      End
      Begin VB.PictureBox p1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1800
         Left            =   0
         MousePointer    =   2  'Cross
         ScaleHeight     =   120
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   0
         Width           =   3000
      End
   End
   Begin VB.Timer TmrGrab 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5085
      Top             =   2775
   End
   Begin VB.TextBox txtrgb 
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
      Height          =   360
      Index           =   1
      Left            =   2760
      TabIndex        =   11
      Top             =   3855
      Width           =   480
   End
   Begin VB.TextBox txtrgb 
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
      Height          =   360
      Index           =   2
      Left            =   2760
      TabIndex        =   12
      Top             =   4470
      Width           =   480
   End
   Begin VB.TextBox txtrgb 
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
      Height          =   360
      Index           =   0
      Left            =   2760
      TabIndex        =   10
      Top             =   3150
      Width           =   480
   End
   Begin VB.TextBox txtHex 
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
      Height          =   360
      Index           =   0
      Left            =   3435
      MaxLength       =   2
      TabIndex        =   13
      Top             =   3150
      Width           =   480
   End
   Begin VB.TextBox txtHex 
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
      Height          =   360
      Index           =   1
      Left            =   3435
      MaxLength       =   2
      TabIndex        =   14
      Top             =   3855
      Width           =   480
   End
   Begin VB.TextBox txtHex 
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
      Height          =   360
      Index           =   2
      Left            =   3435
      MaxLength       =   2
      TabIndex        =   15
      Top             =   4470
      Width           =   480
   End
   Begin VB.CommandButton cmdinvert 
      Height          =   375
      Left            =   2235
      Picture         =   "PalMaker.frx":4C5D
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Invert"
      Top             =   2340
      Width           =   900
   End
   Begin VB.CommandButton RndBlue 
      Caption         =   "Random"
      Height          =   360
      Left            =   4170
      TabIndex        =   18
      Top             =   4470
      Width           =   960
   End
   Begin VB.CommandButton RndGreen 
      Caption         =   "Random"
      Height          =   360
      Left            =   4170
      TabIndex        =   17
      Top             =   3855
      Width           =   960
   End
   Begin VB.CommandButton RndRed 
      Caption         =   "Random"
      Height          =   360
      Left            =   4170
      TabIndex        =   16
      Top             =   3135
      Width           =   960
   End
   Begin VB.CommandButton cmdcopy 
      Caption         =   "Copy"
      Height          =   345
      Left            =   7320
      TabIndex        =   2
      Top             =   1740
      Width           =   2145
   End
   Begin VB.TextBox txtShCol 
      Height          =   300
      Left            =   7305
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1380
      Width           =   2145
   End
   Begin VB.ListBox lstOp 
      Height          =   1050
      IntegralHeight  =   0   'False
      Left            =   7320
      TabIndex        =   0
      Top             =   210
      Width           =   2145
   End
   Begin VB.CommandButton cmdScreen 
      Height          =   375
      Left            =   1185
      Picture         =   "PalMaker.frx":4D42
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Color from Screen"
      Top             =   2340
      Width           =   900
   End
   Begin VB.PictureBox PicGrad 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   5895
      MousePointer    =   99  'Custom
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   27
      Top             =   165
      Width           =   1380
   End
   Begin Project1.Line3D Line3D3 
      Height          =   90
      Left            =   0
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   0
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   159
   End
   Begin VB.HScrollBar hsbgreen 
      Height          =   360
      Left            =   450
      Max             =   255
      TabIndex        =   8
      Top             =   3855
      Width           =   2205
   End
   Begin VB.HScrollBar hsbred 
      Height          =   360
      Left            =   450
      Max             =   255
      TabIndex        =   7
      Top             =   3150
      Width           =   2205
   End
   Begin VB.HScrollBar hsbblue 
      Height          =   360
      Left            =   450
      Max             =   255
      TabIndex        =   9
      Top             =   4455
      Width           =   2205
   End
   Begin Project1.Flat2 Flat21 
      Height          =   2010
      Left            =   75
      TabIndex        =   25
      Top             =   135
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   3545
   End
   Begin VB.PictureBox PicView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   120
      ScaleHeight     =   1920
      ScaleWidth      =   1860
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   165
      Width           =   1860
      Begin VB.Label lblPreview 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   390
         TabIndex        =   32
         Top             =   750
         Width           =   930
      End
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DM Color Picker Version 3.0"
      Height          =   195
      Left            =   7485
      TabIndex        =   52
      Top             =   5025
      Width           =   1995
   End
   Begin VB.Label bllweb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web page styles:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   42
      Top             =   5055
      Width           =   1260
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Index           =   0
      Left            =   195
      TabIndex        =   39
      Top             =   3195
      Width           =   105
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Index           =   1
      Left            =   195
      TabIndex        =   38
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   2
      Left            =   195
      TabIndex        =   37
      Top             =   4560
      Width           =   105
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      Index           =   1
      X1              =   5910
      X2              =   5910
      Y1              =   180
      Y2              =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      Index           =   0
      X1              =   2010
      X2              =   2010
      Y1              =   180
      Y2              =   2100
   End
   Begin VB.Label lblHex 
      BackStyle       =   0  'Transparent
      Caption         =   "Hex"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3525
      TabIndex        =   29
      Top             =   2880
      Width           =   300
   End
   Begin VB.Label lblDec 
      BackStyle       =   0  'Transparent
      Caption         =   "Dec"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2850
      TabIndex        =   28
      Top             =   2880
      Width           =   300
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open Pallet"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnunewpal 
         Caption         =   "&New Pallet"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnublank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "&About DM Color Picker..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstTime As Boolean
Dim isMouseDown As Boolean, ColorFromDC As Boolean, CanMove As Boolean
Dim DskTop_dc As Long, bCanUpdate As Boolean

Dim OldX As Integer, OldY As Integer
Dim TempR As Integer, TempG As Integer, TempB As Integer
Dim Selected_Object As Object

Private Function BuildCSS() As String
Dim StrA As String

    StrA = "<STYLE TYPE=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf
    StrA = StrA & "<!--" & vbCrLf
    StrA = StrA & "BODY{" & vbCrLf
    StrA = StrA & "scrollbar-arrow-color:" & Dec2Web(PicsBar(0).BackColor) & ";" & vbCrLf
    StrA = StrA & "scrollbar-highlight-color:" & Dec2Web(PicsBar(2).BackColor) & ";" & vbCrLf
    StrA = StrA & "scrollbar-face-color:" & Dec2Web(PicsBar(1).BackColor) & ";" & vbCrLf
    StrA = StrA & "scrollbar-3dlight-color:" & Dec2Web(PicsBar(3).BackColor) & ";" & vbCrLf
    StrA = StrA & "scrollbar-track-color:" & Dec2Web(PicsBar(6).BackColor) & ";" & vbCrLf
    StrA = StrA & "scrollbar-darkshadow-color:" & Dec2Web(PicsBar(5).BackColor) & ";" & vbCrLf
    StrA = StrA & "scrollbar-shadow-color:" & Dec2Web(PicsBar(4).BackColor) & ";" & vbCrLf
    StrA = StrA & "}" & vbCrLf
    StrA = StrA & "-->" & vbCrLf
    
    StrA = StrA & "</STYLE>" & vbCrLf & vbCrLf
    
    BuildCSS = StrA

End Function

Sub UpdateScrollBar()
Dim X As Integer, Y As Integer, Col As Long
    'Draw the Scrollbar
    
    For X = 0 To PicHolder.Width - 1
        For Y = 0 To PicHolder.Height - 1
        Col = GetPixel(PicHolder.hdc, X, Y)
        
        Select Case Col
            Case vbBlack
                Col = PicsBar(0).BackColor
            Case vbYellow
                Col = PicsBar(1).BackColor
            Case vbWhite
                Col = PicsBar(2).BackColor
            Case vbBlue
                Col = PicsBar(3).BackColor
            Case vbCyan
                Col = PicsBar(4).BackColor
            Case vbGreen
                Col = PicsBar(5).BackColor
            Case vbRed
                Col = PicsBar(6).BackColor
        End Select
        
        SetPixel PicBar.hdc, X, Y, Col
    Next
    Next
    
    PicBar.Refresh
End Sub

Sub UpdateObject()
On Error Resume Next
    If bCanUpdate Then Exit Sub
    
    If PicWebStyle1.Visible Or PicPayLayout.Visible Or PicScrollbar.Visible Then
        If Selected_Object.Name = "lblLinkObj" Then
            Selected_Object.ForeColor = RGB(hsbred, hsbgreen, hsbblue)
        ElseIf Selected_Object.Name = "PicWebLinks" Then
            Selected_Object.BackColor = RGB(hsbred, hsbgreen, hsbblue)
        ElseIf Selected_Object.Name = "PicWebStyle" Then
            Selected_Object.BackColor = RGB(hsbred, hsbgreen, hsbblue)
        ElseIf Selected_Object.Name = "lblWebStyle" Then
            Selected_Object.ForeColor = RGB(hsbred, hsbgreen, hsbblue)
        ElseIf Selected_Object.Name = "PicsBar" Then
            Selected_Object.BackColor = RGB(hsbred, hsbgreen, hsbblue)
        End If
    End If
      
      If PicScrollbar.Visible And chkHideWebStyle Then Call UpdateScrollBar
      
End Sub

Function OpenPallet(lpFile As String) As Integer
Dim nFile As Long
    nFile = FreeFile
    
    OpenPallet = 0
    
    Erase m_Pallet.Pallet
    m_Pallet.Sig = ""
    
    Open lpFile For Binary As #nFile
        Get #nFile, , m_Pallet
    Close #nFile
    
    If m_Pallet.Sig <> "Pal" Then
        Exit Function
    Else
        ShowPallet PicPallet, PicCol
        OpenPallet = 1
    End If
    
End Function

Sub GetCoords(X As Single, Y As Single)
    OldX = (X \ TileSize) * TileSize
    OldY = (Y \ TileSize) * TileSize
    
    If OldX < 0 Then OldX = 0
    If OldY < 0 Then OldY = Y
    If OldX > (TileSize * 8) Then OldX = (TileSize * 8) - TileSize
    If OldY > (TileSize * 4) - TileSize Then OldY = (TileSize * 4) - TileSize
    
End Sub

Sub ColorFromPoint(PicBox As PictureBox, X As Single, Y As Single)
    
    Selected_LngCol = PicBox.Point(X, Y)
    LongToRgb
    
    hsbred.Value = T_RGB.Red
    hsbgreen.Value = T_RGB.Green
    hsbblue.Value = T_RGB.Blue
    
    ReleaseCapture
    GetColour lstOp.ListIndex
    
End Sub

Sub GetColour(Index As Integer)
Dim iColor As Long
    Select Case Index
        Case 0
            txtShCol.Text = DoHex(PicView.BackColor, mCPlus)
        Case 1
            txtShCol.Text = DoHex(PicView.BackColor, mDelphi)
        Case 2
            txtShCol.Text = DoHex(PicView.BackColor, mVB)
        Case 3
            txtShCol.Text = PicView.BackColor
        Case 4
            txtShCol.Text = txtrgb(0) & "," & txtrgb(1) & "," & txtrgb(2)
        Case 5
            iColor = PicView.BackColor
            txtShCol.Text = Dec2Web(iColor)
        Case 6
            txtShCol.Text = DoHex(PicView.BackColor, mJava)
        Case 7
            txtShCol.Text = DoHex(PicView.BackColor, mPhotoShop)
    End Select
    UpdateObject
End Sub

Sub Update()
    On Error Resume Next
    
    LongToRgb
    GetColour lstOp.ListIndex
    txtrgb(0).Text = hsbred
    txtrgb(1).Text = hsbgreen
    txtrgb(2).Text = hsbblue
    
    txtHex(0).Text = Hex(txtrgb(0).Text)
    txtHex(1).Text = Hex(txtrgb(1).Text)
    txtHex(2).Text = Hex(txtrgb(2).Text)
    
    PicView.BackColor = RGB(hsbred, hsbgreen, hsbblue)
    
    If isMouseDown Then Exit Sub
    Gradient PicGrad, hsbred, hsbgreen, hsbblue, True

End Sub

Private Sub chkHideWebStyle_Click()
    If chkHideWebStyle.Caption = "Hide WebStyles" Then
        chkHideWebStyle.Caption = "Show WebStyles"
    Else
        chkHideWebStyle.Caption = "Hide WebStyles"
    End If
    
    PicWebStyle1.Visible = chkHideWebStyle
    cmdWebLinks.Visible = chkHideWebStyle
    cmdPageStyle.Visible = chkHideWebStyle
   
    If chkHideWebStyle Then
        Form1.Height = 8475
        lblVersion.Top = 7560
    Else
        Form1.Height = 6015
        lblVersion.Top = 5070
    End If
    
End Sub

Private Sub chkWebSafe_Click()

    Use_WebSafe = chkWebSafe

    If chkWebSafe Then
        TempR = hsbred.Value
        TempG = hsbgreen.Value
        TempB = hsbblue.Value
        
        hsbred.Value = WebSafe(TempR)
        hsbgreen.Value = WebSafe(TempG)
        hsbblue.Value = WebSafe(TempB)
    Else
        hsbred.Value = TempR
        hsbgreen.Value = TempG
        hsbblue.Value = TempB
    End If

    GetColour lstOp.ListIndex
    Gradient PicGrad, hsbred, hsbgreen, hsbblue, True
    
End Sub



Private Sub cmdcopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtShCol.Text
End Sub

Private Sub cmdCopy1_Click()
    Clipboard.Clear
    Clipboard.SetText BuildCSS
    MsgBox "CSS Style-Sheetcode code has been copiyed to the clipboard.", vbInformation, Form1.Caption
    
End Sub

Private Sub cmdDialog_Click()
    TDialog.flags = 2
    TDialog.ShowColor
    If TDialog.CancelError Then Exit Sub
    
    Selected_LngCol = TDialog.Color

    LongToRgb
    
    hsbred.Value = T_RGB.Red
    hsbgreen.Value = T_RGB.Green
    hsbblue.Value = T_RGB.Blue
    
    GetColour lstOp.ListIndex
    Gradient PicGrad, hsbred, hsbgreen, hsbblue, True
    
End Sub

Private Sub cmdinvert_Click()
    PicView.BackColor = InvertColour(PicView.BackColor)
    Selected_LngCol = PicView.BackColor
    
    'Update the RGB scrollbars
    hsbred_Change
    hsbred.Value = T_RGB.Red
    hsbgreen.Value = T_RGB.Green
    hsbblue.Value = T_RGB.Blue
    
    
    Gradient PicGrad, Val(txtrgb(0)), Val(txtrgb(1)), Val(txtrgb(2)), True

End Sub

Private Sub cmdOpen1_Click()
On Error GoTo OpenErr:
    InitDialog frmnew.Hwnd
    TDialog.DialogTitle = "Open Linkpage Layout"
    TDialog.Filter = "Link Page Layout Files(*.lpl*)" + Chr$(0) + "*.lpl*" + Chr$(0)
    TDialog.InitialDir = FixPath(App.Path) & "layouts\Text and Links\"
    TDialog.ShowOpen
    If Not TDialog.CancelError Then Exit Sub
    Open TDialog.FileName For Binary As #1
        Get #1, , TemplateData
    Close #1
    
    PicWebLinks.BackColor = TemplateData(0)
    lblLinkObj(0).ForeColor = TemplateData(1)
    lblLinkObj(1).ForeColor = TemplateData(2)
    lblLinkObj(2).ForeColor = TemplateData(3)

    Erase TemplateData
    Exit Sub
    
OpenErr:
    If Err Then MsgBox "Error while opening file.", vbCritical, "Invaild Picture Format"
    
End Sub

Private Sub cmdOpen_Click()
On Error GoTo OpenErr:

    InitDialog frmnew.Hwnd
    TDialog.DialogTitle = "Open Web Template"
    TDialog.Filter = "Web Template Files(*.wtp*)" + Chr$(0) + "*.wtp*" + Chr$(0)
    TDialog.InitialDir = FixPath(App.Path) & "layouts\Webpage Templates\"
    TDialog.ShowOpen
    If Not TDialog.CancelError Then Exit Sub
    Open TDialog.FileName For Binary As #1
        Get #1, , TemplateData
    Close #1
    
    PicWebStyle(0).BackColor = TemplateData(0)
    PicWebStyle(1).BackColor = TemplateData(1)
    PicWebStyle(2).BackColor = TemplateData(2)
    PicWebStyle(3).BackColor = TemplateData(3)
    '
    lblWebStyle(0).ForeColor = TemplateData(4)
    lblWebStyle(1).ForeColor = TemplateData(5)
    lblWebStyle(2).ForeColor = TemplateData(6)
    lblWebStyle(3).ForeColor = TemplateData(7)
    lblWebStyle(4).ForeColor = TemplateData(8)
    Exit Sub
OpenErr:
    If Err Then MsgBox "Error while opening file.", vbCritical, "Invaild Picture Format"
    
End Sub

Private Sub cmdOpen2_Click()
On Error GoTo OpenErr:
    InitDialog frmnew.Hwnd
    TDialog.DialogTitle = "Open Picture"
    TDialog.Filter = "Picture Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    TDialog.InitialDir = FixPath(App.Path) & "pallets\"
    TDialog.ShowOpen
    If Not TDialog.CancelError Then Exit Sub
    p1.Picture = LoadPicture(TDialog.FileName)
    Exit Sub
OpenErr:
    If Err Then MsgBox "Error opening picture.", vbCritical, "Invaild Picture Format"
    
End Sub

Private Sub cmdOpen3_Click()
On Error GoTo OpenErr:

    InitDialog frmnew.Hwnd
    TDialog.DialogTitle = "Open Scrollbar Project"
    TDialog.Filter = "Scrollbar Project Files(*.sbp*)" + Chr$(0) + "*.sbp*" + Chr$(0)
    TDialog.InitialDir = FixPath(App.Path) & "layouts\WebBars\"
    TDialog.ShowOpen
    If Not TDialog.CancelError Then Exit Sub
    Open TDialog.FileName For Binary As #1
        Get #1, , TemplateData
    Close #1
    
    PicsBar(0).BackColor = TemplateData(0)
    PicsBar(1).BackColor = TemplateData(1)
    PicsBar(2).BackColor = TemplateData(2)
    PicsBar(3).BackColor = TemplateData(3)
    PicsBar(4).BackColor = TemplateData(4)
    PicsBar(5).BackColor = TemplateData(5)
    PicsBar(6).BackColor = TemplateData(6)
    UpdateScrollBar
    Erase TemplateData
    Exit Sub
OpenErr:
    If Err Then MsgBox "Error while opening file.", vbCritical, "Invaild Picture Format"
    
End Sub

Private Sub cmdPageStyle_Click()
    PicPayLayout.Visible = True
    PicWebStyle1.Visible = False
    PicScrollbar.Visible = False
End Sub

Private Sub cmdsave_Click()
Dim sFile As String
    InitDialog frmnew.Hwnd
    TDialog.DialogTitle = "Save Picture"
    TDialog.Filter = "Picture Files (*.bmp)" + Chr$(0) + "*.bmp" + Chr$(0)
    TDialog.InitialDir = FixPath(App.Path) & "pallets\"
    TDialog.ShowSave
    If Not TDialog.CancelError Then Exit Sub
    sFile = TDialog.FileName
    If Not GetFileExt(sFile) = "BMP" Then sFile = sFile & ".bmp"
    SavePicture p1.Image, sFile
    Exit Sub
OpenErr:
    If Err Then MsgBox "Error opening picture.", vbCritical, "Invaild Picture Format"
    
End Sub

Private Sub cmdSave1_Click()
Dim sFile As String, sBuff As String
    
    nFile = FreeFile
    InitDialog frmnew.Hwnd
    TDialog.DialogTitle = "Save"
    TDialog.Filter = "Web Template Files(*.wtp*)" + Chr$(0) + "*.wtp*" + Chr$(0)
    TDialog.InitialDir = FixPath(App.Path) & "layouts\Webpage Templates\"
    TDialog.ShowSave
    
    If Not TDialog.CancelError Then Exit Sub
    sFile = TDialog.FileName
    If Not GetFileExt(sFile) = "WTP" Then sFile = sFile & ".wtp"
    
    'Template code
    TemplateData(0) = PicWebStyle(0).BackColor
    TemplateData(1) = PicWebStyle(1).BackColor
    TemplateData(2) = PicWebStyle(2).BackColor
    TemplateData(3) = PicWebStyle(3).BackColor
    '
    TemplateData(4) = lblWebStyle(0).ForeColor
    TemplateData(5) = lblWebStyle(1).ForeColor
    TemplateData(7) = lblWebStyle(2).ForeColor
    TemplateData(8) = lblWebStyle(3).ForeColor
    TemplateData(9) = lblWebStyle(3).ForeColor
    
    SaveFile sFile, TemplateData
    Erase TemplateData
    sFile = ""
End Sub

Private Sub cmdSave2_Click()
Dim sFile As String
    InitDialog frmnew.Hwnd
    TDialog.DialogTitle = "Save"
    TDialog.Filter = "Link Page Layout Files(*.lpl*)" + Chr$(0) + "*.lpl*" + Chr$(0)
    TDialog.InitialDir = FixPath(App.Path) & "layouts\Text and Links\"
    TDialog.ShowSave
    
    If Not TDialog.CancelError Then Exit Sub
    sFile = TDialog.FileName
    If Not GetFileExt(sFile) = "LPL" Then sFile = sFile & ".lpl"
    
    'Template code
    TemplateData(0) = PicWebLinks.BackColor
    TemplateData(1) = lblLinkObj(0).ForeColor
    TemplateData(2) = lblLinkObj(1).ForeColor
    TemplateData(3) = lblLinkObj(2).ForeColor
    
    SaveFile sFile, TemplateData
    Erase TemplateData
    sFile = ""
End Sub

Private Sub cmdSave3_Click()
Dim sFile As String, sBuff As String
    
    nFile = FreeFile
    InitDialog frmnew.Hwnd
    TDialog.DialogTitle = "Save Scrollbar Project"
    TDialog.Filter = "Scrollbar Project Files(*.sbp*)" + Chr$(0) + "*.sbp*" + Chr$(0)
    TDialog.InitialDir = FixPath(App.Path) & "layouts\WebBars\"
    TDialog.ShowSave
    
    If Not TDialog.CancelError Then Exit Sub
    sFile = TDialog.FileName
    If Not GetFileExt(sFile) = "SBP" Then sFile = sFile & ".sbp"
    
    'Template code
    TemplateData(0) = PicsBar(0).BackColor
    TemplateData(1) = PicsBar(1).BackColor
    TemplateData(2) = PicsBar(2).BackColor
    TemplateData(3) = PicsBar(3).BackColor
    TemplateData(4) = PicsBar(4).BackColor
    TemplateData(5) = PicsBar(5).BackColor
    TemplateData(6) = PicsBar(6).BackColor
    
    SaveFile sFile, TemplateData
    Erase TemplateData
    sFile = ""
End Sub

Private Sub cmdScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    TmrGrab = True
End Sub

Private Sub cmdScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TmrGrab = False
    cmdScreen.MousePointer = vbDefault
    If cmdWebSafe.Value = 1 Then cmdWebSafe.Value = 0
End Sub

Private Sub cmdWebBar_Click()
    PicPayLayout.Visible = False
    PicWebStyle1.Visible = False
    PicScrollbar.Visible = True
End Sub

Private Sub cmdWebLinks_Click()
    PicWebStyle1.Visible = True
    PicPayLayout.Visible = False
    PicScrollbar.Visible = False
End Sub

Private Sub cmdWebSafe_Click()
Dim X As Integer, Y As Integer

    p1.SetFocus
    
    If cmdWebSafe.Value = 1 Then
        p2.Picture = p1.Image
        
        For X = 0 To p1.ScaleWidth - 1
            For Y = 0 To p1.ScaleHeight - 1
                LongToRgb
                Selected_LngCol = p1.Point(X, Y)
                p1.PSet (X, Y), RGB(WebSafe(T_RGB.Red), WebSafe(T_RGB.Green), WebSafe(T_RGB.Blue))
            Next
        Next
        X = 0: Y = 0: Selected_LngCol = 0
        Exit Sub
    Else
        p1.Picture = p2.Image
        Set p2.Picture = Nothing
    End If

End Sub

Private Sub cmdzoom_Click()
    PopupMenu frmMenu.mnuzoom
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
    ZoomFactor = 100
    
    If OpenPallet(FixPath(App.Path) & "pallets\Windows XP.pal") <> 1 Then
        MsgBox "The Pallet can not be loaded.", vbExclamation, Form1.Caption
    End If
    
    PicGrad.MouseIcon = cmdScreen.Picture
    'Set the mouse pointer for out color pallet
    PicPallet.MouseIcon = cmdScreen.Picture
    
    lstOp.AddItem "C++ Hex"
    lstOp.AddItem "Delphi Hex"
    lstOp.AddItem "Visual Basic Hex"
    lstOp.AddItem "VB Long Colour"
    lstOp.AddItem "RGB Colour"
    lstOp.AddItem "HTML"
    lstOp.AddItem "Java"
    lstOp.AddItem "PhotoShop"
    lstOp.ListIndex = 0
    
    DskTop_dc = GetDC(GetDesktopWindow)
    PicPallet_MouseDown vbLeftButton, 0, 0, 0
    chkHideWebStyle_Click
    cmdWebLinks_Click
    UpdateScrollBar
End Sub

Private Sub Form_Paint()
    Line3D3.Width = Form1.Width
    If Not FirstTime Then Exit Sub
    FirstTime = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
    Set frmnew = Nothing
    Set frmabout = Nothing
    End
    
End Sub

Private Sub hsbblue_Change()
    hsbred_Change
End Sub

Private Sub hsbblue_Scroll()
    hsbblue_Change
End Sub

Private Sub hsbgreen_Change()
On Error Resume Next
    hsbred_Change
End Sub

Private Sub hsbgreen_Scroll()
    hsbgreen_Change
End Sub

Private Sub hsbred_Change()
    Update
End Sub

Private Sub hsbred_Scroll()
    hsbred_Change
End Sub

Private Sub Label1_Click(Index As Integer)
    Set Selected_Object = PicsBar(Index)
End Sub

Private Sub lblLinkObj_Click(Index As Integer)
    Set Selected_Object = lblLinkObj(Index)
End Sub

Private Sub lblLinkObj_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    bCanUpdate = True
    Selected_LngCol = lblLinkObj(Index).ForeColor 'Get color from under the cursor
    'Update the RGB scrollbars
    hsbred_Change
    hsbred.Value = T_RGB.Red
    hsbgreen.Value = T_RGB.Green
    hsbblue.Value = T_RGB.Blue
    GetColour lstOp.ListIndex
End Sub

Private Sub lblLinkObj_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
bCanUpdate = False
End Sub

Private Sub lblWebStyle_Click(Index As Integer)
    Set Selected_Object = lblWebStyle(Index)
End Sub

Private Sub lblWebStyle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    bCanUpdate = True
    Selected_LngCol = lblWebStyle(Index).ForeColor 'Get color from under the cursor
    'Update the RGB scrollbars
    hsbred_Change
    hsbred.Value = T_RGB.Red
    hsbgreen.Value = T_RGB.Green
    hsbblue.Value = T_RGB.Blue
    GetColour lstOp.ListIndex
End Sub

Private Sub lblWebStyle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    bCanUpdate = False
End Sub

Private Sub lstOp_Click()
    GetColour lstOp.ListIndex
End Sub

Private Sub mnuabout_Click()
    frmabout.Show vbModal, Form1
End Sub

Private Sub mnuexit_Click()
    Unload Form1
End Sub

Private Sub mnunewpal_Click()
    frmnew.Show vbModal, Form1
End Sub

Private Sub mnuopen_Click()
    
    InitDialog frmnew.Hwnd
    TDialog.DialogTitle = "Open Pallet"
    TDialog.Filter = "Pallet Files (*.pal)" + Chr$(0) + "*.pal" + Chr$(0)
    TDialog.InitialDir = FixPath(App.Path) & "pallets\"
    TDialog.ShowOpen
    If Not TDialog.CancelError Then Exit Sub
    
    If GetFileExt(TDialog.FileName) <> "PAL" Then
        MsgBox "This is not a vaild Pallet filename.", vbCritical, "Invaild File Format"
        Exit Sub
    End If
    
    If OpenPallet(TDialog.FileName) <> 1 Then
        MsgBox "The Pallet can not be loaded.", vbExclamation, Form1.Caption
        Exit Sub
    Else
        PicPallet_MouseDown vbLeftButton, 0, 0, 0 'Select first color
    End If
    
End Sub

Private Sub p1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    
    If Button <> vbLeftButton Then
        Exit Sub
    Else
        isMouseDown = True
        ColorFromPoint p1, X, Y
        Gradient PicGrad, hsbred, hsbgreen, hsbblue, True
    End If
End Sub


Private Sub p1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isMouseDown = False
End Sub

Private Sub PicGrad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    
    If Button <> vbLeftButton Then
        Exit Sub
    Else
        isMouseDown = True
        ColorFromPoint PicGrad, X, Y
    End If

End Sub

Private Sub PicGrad_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isMouseDown = False
End Sub

Private Sub PicPallet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    
    Selected_LngCol = PicPallet.Point(X, Y) 'Get color from under the cursor
    
    'Update the RGB scrollbars
    hsbred_Change
    hsbred.Value = T_RGB.Red
    hsbgreen.Value = T_RGB.Green
    hsbblue.Value = T_RGB.Blue
    
    GetColour lstOp.ListIndex

End Sub

Private Sub PicPallet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetCoords X, Y
End Sub

Private Sub PicsBar_Click(Index As Integer)
    Label1_Click Index
End Sub

Private Sub PicWebLinks_Click()
    Set Selected_Object = PicWebLinks
End Sub

Private Sub PicWebLinks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    bCanUpdate = True
    Selected_LngCol = PicWebLinks.Point(X, Y) 'Get color from under the cursor
    'Update the RGB scrollbars
    hsbred_Change
    hsbred.Value = T_RGB.Red
    hsbgreen.Value = T_RGB.Green
    hsbblue.Value = T_RGB.Blue
    GetColour lstOp.ListIndex
End Sub

Private Sub PicWebLinks_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bCanUpdate = False
End Sub

Private Sub PicWebStyle_Click(Index As Integer)
    Set Selected_Object = PicWebStyle(Index)
End Sub

Private Sub PicWebStyle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    bCanUpdate = True
    Selected_LngCol = PicWebStyle(Index).Point(X, Y) 'Get color from under the cursor
    'Update the RGB scrollbars
    hsbred_Change
    hsbred.Value = T_RGB.Red
    hsbgreen.Value = T_RGB.Green
    hsbblue.Value = T_RGB.Blue
    GetColour lstOp.ListIndex
End Sub

Private Sub PicWebStyle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    bCanUpdate = False
End Sub

Private Sub RndBlue_Click()
    hsbblue.Value = Int(Rnd * 256) + 1
End Sub

Private Sub RndGreen_Click()
    hsbgreen.Value = Int(Rnd * 256) + 1
End Sub

Private Sub RndRed_Click()
    hsbred.Value = Int(Rnd * 256) + 1
End Sub
Private Sub TmrGrab_Timer()
Dim aWidth As Long, aHeight As Long, ZoomWidth As Long, ZoomHeight As Long, X As Long, Y As Long
Dim mouse As POINTAPI
Dim k As Long

    cmdScreen.MousePointer = vbCrosshair 'Set cursor
    GetCursorPos mouse
        
    If WindowFromPoint(mouse.X, mouse.Y) <> p1.Hwnd Then
        p1.Cls
        aWidth = p1.ScaleWidth
        aHeight = p1.ScaleHeight
        ZoomWidth = aWidth * (1 / (120 / ZoomFactor))
        ZoomHeight = aHeight * (1 / (200 / ZoomFactor))
        X = (mouse.X - ZoomWidth / 2)
        Y = (mouse.Y - ZoomHeight / 2)
        StretchBlt p1.hdc, 0, 0, aWidth, aHeight, DskTop_dc, X, Y, ZoomWidth, ZoomHeight, vbSrcCopy
    End If
End Sub

Private Sub txtHex_Change(Index As Integer)
   txtrgb(Index).Text = Val("&H" & txtHex(Index))
End Sub

Private Sub txtrgb_Change(Index As Integer)
Dim Col_Val As Integer
    Col_Val = Val(txtrgb(Index).Text)
    If (Col_Val > 255) Then Col_Val = 255: txtrgb(Index).Text = Col_Val
    
    hsbred.Value = Val(txtrgb(0).Text)
    hsbgreen.Value = Val(txtrgb(1).Text)
    hsbblue.Value = Val(txtrgb(2).Text)
End Sub

