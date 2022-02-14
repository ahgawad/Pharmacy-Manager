VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmDrugs 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·√’‰«›"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "frmDrugs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdMain 
      Caption         =   "«·ﬁ«∆„… «·—∆Ì”Ì…"
      Height          =   375
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   6120
      Width           =   3000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   -2147483640
      TabCaption(0)   =   "„⁄·Ê„« "
      TabPicture(0)   =   "frmDrugs.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chrSales"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraSource0"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "≈÷«›…"
      TabPicture(1)   =   "frmDrugs.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " Õ–›"
      TabPicture(2)   =   "frmDrugs.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraSure"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "fraSource2"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   " ⁄œÌ·"
      TabPicture(3)   =   "frmDrugs.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).Control(1)=   "Frame8"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "„⁄œ· «·»Ì⁄ ·› —… „⁄Ì‰…"
      TabPicture(4)   =   "frmDrugs.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(1)=   "Frame10"
      Tab(4).Control(2)=   "Frame11"
      Tab(4).ControlCount=   3
      Begin VB.Frame Frame11 
         Height          =   1335
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   2280
         Width           =   9015
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ﬁÿ⁄…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   480
            Width           =   450
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "⁄œœ «·ﬁÿ⁄ Œ·«· «·› —… «·„Õœœ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   480
            Width           =   2700
         End
      End
      Begin VB.Frame Frame10 
         Height          =   855
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1320
         Width           =   9015
         Begin VB.CommandButton Command4 
            Caption         =   "„Ê«›ﬁ"
            Height          =   375
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   6480
            TabIndex        =   40
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyy\MMM\d"
            Format          =   24510467
            CurrentDate     =   36412
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   3120
            TabIndex        =   41
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyy\MMM\d"
            Format          =   24510467
            CurrentDate     =   36412
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "≈·Ï"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "„‰"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8400
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   240
            Width           =   195
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   480
         Width           =   9015
         Begin VB.ListBox lstItemrec 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy\\M\\d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            ItemData        =   "frmDrugs.frx":04CE
            Left            =   120
            List            =   "frmDrugs.frx":04D0
            RightToLeft     =   -1  'True
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.ListBox cmpItemcode 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            ItemData        =   "frmDrugs.frx":04D2
            Left            =   6720
            List            =   "frmDrugs.frx":04D4
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   1005
         End
         Begin VB.ListBox txtItemname 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            ItemData        =   "frmDrugs.frx":04D6
            Left            =   3000
            List            =   "frmDrugs.frx":04D8
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "’‰›"
            Height          =   195
            Index           =   5
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   240
            Width           =   330
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "ﬂÊœ"
            Height          =   195
            Index           =   15
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   240
            Width           =   270
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame9 
         Height          =   735
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   480
         Width           =   9015
         Begin VB.ListBox txtItemname 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            ItemData        =   "frmDrugs.frx":04DA
            Left            =   3000
            List            =   "frmDrugs.frx":04DC
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   2895
         End
         Begin VB.ListBox cmpItemcode 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            ItemData        =   "frmDrugs.frx":04DE
            Left            =   6720
            List            =   "frmDrugs.frx":04E0
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   1005
         End
         Begin VB.ListBox lstItemrec 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy\\M\\d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            ItemData        =   "frmDrugs.frx":04E2
            Left            =   120
            List            =   "frmDrugs.frx":04E4
            RightToLeft     =   -1  'True
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "ﬂÊœ"
            Height          =   195
            Index           =   14
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   240
            Width           =   270
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "’‰›"
            Height          =   195
            Index           =   4
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   240
            Width           =   330
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1935
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1320
         Width           =   9015
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
            Caption         =   "„Ê«›ﬁ"
            Height          =   360
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "‰Ê⁄ «·’‰›"
            Height          =   195
            Index           =   13
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   1440
            Width           =   1020
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·Õœ «·√œ‰Ï"
            Height          =   195
            Index           =   12
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   840
            Width           =   1020
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "À„‰ «·ÊÕœ…"
            Height          =   195
            Index           =   11
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   240
            Width           =   1050
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraSure 
         Enabled         =   0   'False
         Height          =   2295
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   3360
         Width           =   9015
         Begin VB.CommandButton cmdNo 
            Caption         =   "·«"
            Height          =   360
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   1320
            Width           =   3135
         End
         Begin VB.CommandButton Command2 
            Caption         =   "‰⁄„"
            Height          =   360
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Â· √‰  „ √ﬂœ ø"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   960
            Width           =   1800
         End
      End
      Begin VB.Frame Frame6 
         Height          =   4335
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   9015
         Begin VB.CommandButton cmdCancel 
            Caption         =   "≈·€«¡"
            Height          =   480
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   3000
            Width           =   2415
         End
         Begin VB.CommandButton Command1 
            Caption         =   "„Ê«›ﬁ"
            Height          =   480
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   3000
            Width           =   2415
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   1920
            Width           =   2535
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "‰Ê⁄ «·’‰›"
            Height          =   195
            Index           =   7
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   1920
            Width           =   1020
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·Õœ «·√œ‰Ï"
            Height          =   195
            Index           =   6
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1320
            Width           =   1020
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "À„‰ «·ÊÕœ…"
            Height          =   195
            Index           =   5
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   720
            Width           =   1050
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1935
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1320
         Width           =   9015
         Begin VB.CommandButton cmdAgreesource2 
            Caption         =   "„Ê«›ﬁ"
            Height          =   360
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblKind 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label lblLimit 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblPrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "À„‰ «·ÊÕœ…"
            Height          =   195
            Index           =   10
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   240
            Width           =   1050
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·Õœ «·√œ‰Ï"
            Height          =   195
            Index           =   9
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   840
            Width           =   1020
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "‰Ê⁄ «·’‰›"
            Height          =   195
            Index           =   8
            Left            =   7320
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   1440
            Width           =   1020
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraSource2 
         Height          =   735
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   480
         Width           =   9015
         Begin VB.ListBox lstItemrec 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy\\M\\d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            ItemData        =   "frmDrugs.frx":04E6
            Left            =   120
            List            =   "frmDrugs.frx":04E8
            RightToLeft     =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.ListBox cmpItemcode 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            ItemData        =   "frmDrugs.frx":04EA
            Left            =   6720
            List            =   "frmDrugs.frx":04EC
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   240
            Width           =   1005
         End
         Begin VB.ListBox txtItemname 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            ItemData        =   "frmDrugs.frx":04EE
            Left            =   3000
            List            =   "frmDrugs.frx":04F0
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "’‰›"
            Height          =   195
            Index           =   2
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   240
            Width           =   330
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "ﬂÊœ"
            Height          =   195
            Index           =   2
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   240
            Width           =   270
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraSource0 
         Height          =   735
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   9015
         Begin VB.ListBox lstItemrec 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy\\M\\d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            ItemData        =   "frmDrugs.frx":04F2
            Left            =   120
            List            =   "frmDrugs.frx":04F4
            RightToLeft     =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.ListBox cmpItemcode 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            ItemData        =   "frmDrugs.frx":04F6
            Left            =   6720
            List            =   "frmDrugs.frx":04F8
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   1005
         End
         Begin VB.ListBox txtItemname 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            ItemData        =   "frmDrugs.frx":04FA
            Left            =   3000
            List            =   "frmDrugs.frx":04FC
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton cmdAgreesource0 
            Caption         =   "„Ê«›ﬁ"
            Height          =   360
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "’‰›"
            Height          =   195
            Index           =   0
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   240
            Width           =   330
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "ﬂÊœ"
            Height          =   195
            Index           =   0
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   270
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   9015
         Begin VB.TextBox txtMinlimit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   5
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3480
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   480
            Width           =   855
         End
         Begin VB.Frame Frame7 
            Height          =   855
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   840
            Width           =   8055
            Begin VB.ListBox lstKind 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy\\M\\d"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   0
               EndProperty
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "frmDrugs.frx":04FE
               Left            =   240
               List            =   "frmDrugs.frx":0500
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   360
               Width           =   1695
            End
            Begin VB.ListBox lstRecno 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy\\M\\d"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   0
               EndProperty
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "frmDrugs.frx":0502
               Left            =   120
               List            =   "frmDrugs.frx":0504
               RightToLeft     =   -1  'True
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   360
               Visible         =   0   'False
               Width           =   180
            End
            Begin VB.ListBox lstSource 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy\\M\\d"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   0
               EndProperty
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "frmDrugs.frx":0506
               Left            =   2160
               List            =   "frmDrugs.frx":0508
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   360
               Width           =   2535
            End
            Begin VB.ListBox cmpExpirydate 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy\\M\\d"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   0
               EndProperty
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "frmDrugs.frx":050A
               Left            =   4920
               List            =   "frmDrugs.frx":050C
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   360
               Width           =   1695
            End
            Begin VB.ListBox txtStock 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "yyyy\\M\\d"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   0
               EndProperty
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "frmDrugs.frx":050E
               Left            =   6840
               List            =   "frmDrugs.frx":0510
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "‰Ê⁄ «·„’œ—"
               Height          =   195
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   120
               Width           =   1230
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "«”„ «·„’œ—"
               Height          =   195
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   120
               Width           =   1230
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   " «—ÌŒ «·’·«ÕÌ…"
               Height          =   195
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   120
               Width           =   1230
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "«·—’Ìœ"
               Height          =   195
               Left            =   7080
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   120
               Width           =   465
               WordWrap        =   -1  'True
            End
         End
         Begin VB.TextBox txtAll 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Height          =   360
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   480
            Width           =   1005
         End
         Begin VB.TextBox txtPceprc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Height          =   360
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   480
            Width           =   1005
         End
         Begin VB.TextBox txtAllstock 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Height          =   360
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   480
            Width           =   885
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "«·Õœ «·√œ‰Ï"
            Height          =   195
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   240
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·À„‰ «·≈Ã„«·Ì"
            Height          =   195
            Index           =   4
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·”⁄— ··Ã„ÂÊ—"
            Height          =   195
            Index           =   3
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·ﬂ„Ì… «·ﬂ·Ì…"
            Height          =   195
            Index           =   3
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   480
         Width           =   9015
         Begin VB.ComboBox cmpItemname 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmDrugs.frx":0512
            Left            =   3000
            List            =   "frmDrugs.frx":0514
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label txtItemcode 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "’‰›"
            Height          =   195
            Index           =   1
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   240
            Width           =   330
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "ﬂÊœ"
            Height          =   195
            Index           =   1
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   240
            Width           =   270
            WordWrap        =   -1  'True
         End
      End
      Begin MSChart20Lib.MSChart chrSales 
         Height          =   2535
         Left            =   240
         OleObjectBlob   =   "frmDrugs.frx":0516
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   3120
         Width           =   9015
      End
   End
End
Attribute VB_Name = "frmDrugs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAgreesource0_Click()
On Error Resume Next
Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    Set rsStock = db.OpenRecordset("stock", dbOpenTable)
    Set rsSales = db.OpenRecordset("sales", dbOpenTable)
    
    Dim rec_no As Integer
    rec_no = txtItemname(0).ListIndex

    If cmpItemcode(0).ListCount > 0 Then
        cmpItemcode(0).Selected(rec_no) = True
        lstItemrec(0).Selected(rec_no) = True

        rsItem_data.MoveFirst
        rsItem_data.Move Val(lstItemrec(0).Text) - 1
        txtPceprc.Text = rsItem_data.Fields(4)
        txtMinlimit.Text = rsItem_data.Fields(3)
        txtAllstock.Text = rsItem_data.Fields(6)
        If Val(txtAllstock.Text) <= Val(txtMinlimit.Text) Then txtAllstock.BackColor = RGB(255, 0, 0) Else txtAllstock.BackColor = &H80000004
        txtAll.Text = Val(txtAllstock.Text) * Val(txtPceprc.Text)
        
        cmpExpirydate.Clear
        txtStock.Clear
        lstRecno.Clear
        lstSource.Clear
        lstKind.Clear
                
        Dim expDate As String
        Dim stkQnt As Long
        Dim rec As Integer
        Dim Rec1 As Integer
        Dim no As Integer
        Dim srcName As String
        Dim srcKind As String
        
        rsStock.MoveLast
        rec = rsStock.RecordCount
        rsStock.MoveFirst
        
        For Rec1 = 1 To rec
            If rsStock.Fields(1) = Val(cmpItemcode(0).Text) Then
                expDate = Format(rsStock.Fields(3), "d\\MMM\\yyyy")
                stkQnt = rsStock.Fields(2)
                srcName = rsStock.Fields(4)
                srcKind = rsStock.Fields(5)
                cmpExpirydate.AddItem expDate
                txtStock.AddItem stkQnt
                lstRecno.AddItem Rec1
                lstSource.AddItem srcName
                lstKind.AddItem srcKind
            End If
            rsStock.Move 1
        Next Rec1
        If lstRecno.ListCount > 0 Then
            cmpExpirydate.Selected(0) = True
            txtStock.Selected(0) = True
            lstRecno.Selected(0) = True
            lstSource.Selected(0) = True
            lstKind.Selected(0) = True
        End If
    End If
    Dim T As Integer
    Dim S As Integer
    Dim DD(1 To 30) As Long
    Dim XX As Integer
    rsSales.MoveLast
    T = rsSales.RecordCount
    rsSales.MoveFirst
For XX = 1 To 30
    rsSales.MoveFirst
    For S = 1 To T
        If rsSales.Fields(1) = cmpItemcode(0).Text Then
            If rsSales.Fields(4) = Date + 1 - XX Then
                DD(31 - XX) = DD(31 - XX) + rsSales.Fields(2)
            End If
        End If
    rsSales.Move 1
    Next S
Next XX
    

    For XX = 1 To 30
        chrSales.Column = XX
        chrSales.Data = DD(XX)
    Next XX

End Sub

Private Sub cmdAgreesource2_Click()
On Error Resume Next
    fraSure.Enabled = True
    cmdNo.SetFocus

End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
cmpItemname.Text = ""
Text2.Text = ""
Text1.Text = ""
Text3.Text = ""
cmpItemname.SetFocus
End Sub

Private Sub cmdMain_Click()
On Error Resume Next
    frmDrugs.Visible = False
    frmMain.Visible = True
    Unload frmDrugs

End Sub

Private Sub cmdNo_Click()
On Error Resume Next
    fraSure.Enabled = False
    cmpItemcode(1).SetFocus

End Sub

Private Sub cmpExpirydate_Scroll()
On Error Resume Next
    Dim rec_no As Integer
    rec_no = cmpExpirydate.ListIndex
    If lstRecno.ListCount > 0 Then
        txtStock.Selected(rec_no) = True
        lstRecno.Selected(rec_no) = True
        lstSource.Selected(rec_no) = True
        lstKind.Selected(rec_no) = True
    End If

End Sub


Private Sub cmpItemcode_LostFocus(Index As Integer)
On Error Resume Next
    Call cmpItemcode_Scroll(Index)
End Sub

Private Sub cmpItemcode_Scroll(Index As Integer)
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    
    Dim rec_no0 As Integer
    rec_no0 = cmpItemcode(0).ListIndex
    Dim rec_no As Integer
    rec_no = cmpItemcode(1).ListIndex
    Dim rec_no1 As Integer
    rec_no1 = cmpItemcode(2).ListIndex
    Dim rec_no2 As Integer
    rec_no2 = cmpItemcode(3).ListIndex


    If cmpItemcode(0).ListCount > 0 Then
        txtItemname(0).Selected(rec_no0) = True
        lstItemrec(0).Selected(rec_no0) = True
    End If


    If cmpItemcode(1).ListCount > 0 Then
        txtItemname(1).Selected(rec_no) = True
        lstItemrec(1).Selected(rec_no) = True

        rsItem_data.MoveFirst
        rsItem_data.Move Val(lstItemrec(1).Text) - 1
        lblPrice.Caption = rsItem_data.Fields(4)
        lblLimit.Caption = rsItem_data.Fields(3)
        lblKind.Caption = rsItem_data.Fields(5)
        
    End If
    
    If cmpItemcode(2).ListCount > 0 Then
        txtItemname(2).Selected(rec_no1) = True
        lstItemrec(2).Selected(rec_no1) = True

        rsItem_data.MoveFirst
        rsItem_data.Move Val(lstItemrec(2).Text) - 1
        Text4.Text = rsItem_data.Fields(4)
        Text5.Text = rsItem_data.Fields(3)
        Text6.Text = rsItem_data.Fields(5)
        
    End If

    If cmpItemcode(3).ListCount > 0 Then
        txtItemname(3).Selected(rec_no2) = True
        lstItemrec(3).Selected(rec_no2) = True
    End If

End Sub
Private Sub Command1_Click()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    If cmpItemname.Text = "" Then
        MsgBox "√œŒ· «”„ «·’‰›", vbExclamation, "Œÿ√"
        cmpItemname.SetFocus
    ElseIf Not IsNumeric(Text1.Text) Or Val(Text1.Text) < 0 Then
        MsgBox "√œŒ· À„‰ «·ﬁÿ⁄…", vbExclamation, "Œÿ√"
        Text1.Text = ""
        Text1.SetFocus
    ElseIf Not IsNumeric(Text2.Text) Or Val(Text2.Text) < 0 Then
        MsgBox "√œŒ· «·Õœ «·√œ‰Ï", vbExclamation, "Œÿ√"
        Text2.Text = ""
        Text2.SetFocus
    ElseIf Text3.Text = "" Then
        MsgBox "√œŒ· ‰Ê⁄ «·’‰›", vbExclamation, "Œÿ√"
        Text3.SetFocus
    Else
        rsItem_data.MoveLast
        rsItem_data.Edit
        rsItem_data.AddNew
        rsItem_data.Fields(1) = txtItemcode.Caption
        rsItem_data.Fields(2) = cmpItemname.Text
        rsItem_data.Fields(3) = Text2.Text
        rsItem_data.Fields(4) = Text1.Text
        rsItem_data.Fields(5) = Text3.Text
        rsItem_data.Fields(6) = 0
        rsItem_data.Update
        rsItem_data.MoveLast
        txtItemcode.Caption = rsItem_data.Fields(0) + 1
        cmpItemname.AddItem cmpItemname.Text
        Dim X As Integer
        rsItem_data.MoveLast
        For X = 0 To 3
            cmpItemcode(X).AddItem rsItem_data.Fields(1)
            txtItemname(X).AddItem rsItem_data.Fields(2)
            lstItemrec(X).AddItem rsItem_data.RecordCount
        Next X
        
        cmpItemname.Text = ""
        Text2.Text = ""
        Text1.Text = ""
        Text3.Text = ""
        cmpItemname.SetFocus
    End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
Rem
    Dim Rc As Integer
    Dim Rn As Integer
    Dim Inv As Integer
    Dim Y As Integer
    rsItem_data.MoveFirst
    Rn = Val(lstItemrec(1).Text) - 1
    rsItem_data.Move Rn
    If rsItem_data.Fields(6) = 0 Then
        rsItem_data.Edit
        rsItem_data.Delete
        Inv = cmpItemcode(1).ListIndex
        For Y = 0 To 2
            cmpItemcode(Y).RemoveItem (Inv)
            txtItemname(Y).RemoveItem (Inv)
        Next Y
    Else: MsgBox "·« Ì„ﬂ‰ Õ–› Â–« «·’‰›° ·ÊÃÊœ „Œ“Ê‰ „‰Â", vbExclamation, "Œÿ√"
    End If
cmpItemcode(2).SetFocus
     
fraSure.Enabled = False

cmpItemcode(1).SetFocus
cmpItemcode(1).Selected(0) = True
txtItemname(1).Selected(0) = True
End Sub

Private Sub Command4_Click()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsSales = db.OpenRecordset("sales", dbOpenTable)
    
    Dim rec_no As Integer
    rec_no = txtItemname(0).ListIndex

    If cmpItemcode(0).ListCount > 0 Then
        cmpItemcode(0).Selected(rec_no) = True
        lstItemrec(0).Selected(rec_no) = True
    End If
    
    Dim T As Integer
    Dim S As Integer
    Dim XX2, XX1 As Integer
    rsSales.MoveLast
    T = rsSales.RecordCount
    rsSales.MoveFirst
    For S = 1 To T
        If rsSales.Fields(1) = cmpItemcode(3).Text Then
            If rsSales.Fields(4) >= DTPicker1.Value And rsSales.Fields(4) <= DTPicker2.Value Then
                XX1 = rsSales.Fields(2)
                XX2 = XX2 + XX1
            End If
        End If
        rsSales.Move 1
    Next S
    Label12.Caption = XX2
End Sub
Private Sub Command5_Click()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    
    Dim rec_no As Integer
    rec_no = txtItemname(1).ListIndex
    Dim rec_no1 As Integer
    rec_no1 = txtItemname(2).ListIndex
    
    rsItem_data.MoveFirst
    rsItem_data.Move Val(lstItemrec(2).Text) - 1
    rsItem_data.Edit
    rsItem_data.Fields(4) = Text4.Text
    rsItem_data.Fields(3) = Text5.Text
    rsItem_data.Fields(5) = Text6.Text
    rsItem_data.Update
    cmpItemcode(2).SetFocus
       
End Sub
Private Sub Form_Load()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)

    Dim l_rec_no1 As Integer
    
    rsItem_data.MoveLast
    l_rec_no1 = rsItem_data.RecordCount
    Dim zoro As Integer
    zoro = rsItem_data.Fields(0)
    txtItemcode.Caption = 1 + zoro
    
    Dim J1 As Integer
    Dim X As Integer
For X = 0 To 3
    rsItem_data.MoveFirst
    For J1 = 1 To l_rec_no1
        cmpItemcode(X).AddItem rsItem_data.Fields(1)
        txtItemname(X).AddItem rsItem_data.Fields(2)
        lstItemrec(X).AddItem J1
        rsItem_data.Move 1
    Next J1
    
    If cmpItemcode(X).ListCount > 0 Then
        cmpItemcode(X).Selected(0) = True
        txtItemname(X).Selected(0) = True
        lstItemrec(X).Selected(0) = True
    End If
Next X

rsItem_data.MoveFirst
For J1 = 1 To l_rec_no1
    cmpItemname.AddItem rsItem_data.Fields(2)
    rsItem_data.Move 1
Next J1

chrSales.ColumnCount = 30

Dim XX As Integer
For XX = 1 To 30
    chrSales.Column = XX
    chrSales.Data = 0
    chrSales.ColumnLabel = XX
Next XX

    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker1.Year = Year(Date) - 1
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    frmMain.Visible = True
End Sub
Private Sub lstKind_Scroll()
On Error Resume Next
    Dim rec_no As Integer
    rec_no = lstKind.ListIndex
    If lstRecno.ListCount > 0 Then
        cmpExpirydate.Selected(rec_no) = True
        lstRecno.Selected(rec_no) = True
        lstSource.Selected(rec_no) = True
        txtStock.Selected(rec_no) = True
    
    End If

End Sub
Private Sub lstSource_Scroll()
On Error Resume Next
    Dim rec_no As Integer
    rec_no = lstSource.ListIndex
    If lstRecno.ListCount > 0 Then
        cmpExpirydate.Selected(rec_no) = True
        lstRecno.Selected(rec_no) = True
        txtStock.Selected(rec_no) = True
        lstKind.Selected(rec_no) = True
    
    End If

End Sub


Private Sub txtItemname_LostFocus(Index As Integer)
On Error Resume Next
    Call txtItemname_Scroll(Index)
End Sub
Private Sub txtItemname_Scroll(Index As Integer)
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    
    
    Dim rec_no0 As Integer
    rec_no0 = txtItemname(0).ListIndex
    Dim rec_no As Integer
    rec_no = txtItemname(1).ListIndex
    Dim rec_no1 As Integer
    rec_no1 = txtItemname(2).ListIndex
    Dim rec_no2 As Integer
    rec_no2 = txtItemname(3).ListIndex


    If cmpItemcode(0).ListCount > 0 Then
        cmpItemcode(0).Selected(rec_no0) = True
        lstItemrec(0).Selected(rec_no0) = True
    End If

    If cmpItemcode(1).ListCount > 0 Then
        cmpItemcode(1).Selected(rec_no) = True
        lstItemrec(1).Selected(rec_no) = True

        rsItem_data.MoveFirst
        rsItem_data.Move Val(lstItemrec(1).Text) - 1
        lblPrice.Caption = rsItem_data.Fields(4)
        lblLimit.Caption = rsItem_data.Fields(3)
        lblKind.Caption = rsItem_data.Fields(5)
    
    End If
    
    If cmpItemcode(2).ListCount > 0 Then
        cmpItemcode(2).Selected(rec_no1) = True
        lstItemrec(2).Selected(rec_no1) = True

        rsItem_data.MoveFirst
        rsItem_data.Move Val(lstItemrec(2).Text) - 1
        Text4.Text = rsItem_data.Fields(4)
        Text5.Text = rsItem_data.Fields(3)
        Text6.Text = rsItem_data.Fields(5)
    End If
    If cmpItemcode(3).ListCount > 0 Then
        cmpItemcode(3).Selected(rec_no2) = True
        lstItemrec(3).Selected(rec_no2) = True
    End If
    
End Sub

Private Sub txtStock_Scroll()
On Error Resume Next
    Dim rec_no As Integer
    rec_no = txtStock.ListIndex
    If lstRecno.ListCount > 0 Then
        cmpExpirydate.Selected(rec_no) = True
        lstRecno.Selected(rec_no) = True
        lstSource.Selected(rec_no) = True
        lstKind.Selected(rec_no) = True
    End If

End Sub


