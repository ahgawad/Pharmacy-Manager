VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·Õ”«»« "
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdMain 
      Caption         =   "«·ﬁ«∆„… «·—∆Ì”Ì…"
      Height          =   375
      Left            =   3120
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
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   -2147483640
      TabCaption(0)   =   "„⁄·Ê„« "
      TabPicture(0)   =   "frmAccount.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraSource0"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "≈÷«›…"
      TabPicture(1)   =   "frmAccount.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " Õ–›"
      TabPicture(2)   =   "frmAccount.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraSure"
      Tab(2).Control(1)=   "fraAdd"
      Tab(2).Control(2)=   "fraSource2"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   " ⁄œÌ·"
      TabPicture(3)   =   "frmAccount.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "Frame6"
      Tab(3).ControlCount=   2
      Begin VB.Frame fraSource2 
         Height          =   735
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   480
         Width           =   9015
         Begin VB.ListBox cmpSource2 
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
            ItemData        =   "frmAccount.frx":0070
            Left            =   4080
            List            =   "frmAccount.frx":0072
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   240
            Width           =   3045
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«”„ «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Frame fraSource0 
         Height          =   735
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   480
         Width           =   9015
         Begin VB.CommandButton cmdAgreesource0 
            Caption         =   "„Ê«›ﬁ"
            Height          =   360
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   240
            Width           =   3135
         End
         Begin VB.ListBox cmpSource0 
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
            ItemData        =   "frmAccount.frx":0074
            Left            =   4080
            List            =   "frmAccount.frx":0076
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   240
            Width           =   3045
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«”„ «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2295
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1320
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid grdStk 
            Height          =   1935
            Left            =   480
            TabIndex        =   37
            Top             =   240
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   3413
            _Version        =   393216
            Rows            =   1
            Cols            =   5
            FixedCols       =   0
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   480
         Width           =   9015
         Begin VB.ComboBox txtSourcename 
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
            ItemData        =   "frmAccount.frx":0078
            Left            =   4080
            List            =   "frmAccount.frx":007A
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   3045
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«”„ «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1935
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1320
         Width           =   9015
         Begin VB.TextBox txtAddresssave 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   480
            Width           =   6645
         End
         Begin VB.TextBox txtTelsave 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   1200
            Width           =   1635
         End
         Begin VB.CommandButton Command1 
            Caption         =   "„Ê«›ﬁ"
            Height          =   360
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "⁄‰Ê«‰ «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Â« › «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   1200
            Width           =   930
         End
      End
      Begin VB.Frame Frame4 
         Enabled         =   0   'False
         Height          =   1935
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   3720
         Width           =   9015
         Begin VB.TextBox txtTel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Height          =   360
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1200
            Width           =   1845
         End
         Begin VB.TextBox txtAddress 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Height          =   360
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   480
            Width           =   6765
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Â« › «·„‰‘√…"
            Height          =   195
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "⁄‰Ê«‰ «·„‰‘√…"
            Height          =   195
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame fraAdd 
         Height          =   1935
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1320
         Width           =   9015
         Begin VB.CommandButton cmdOK 
            Caption         =   "„Ê«›ﬁ"
            Height          =   360
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "⁄‰Ê«‰ «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Â« › «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label txtAdddisplay 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   480
            Width           =   6735
         End
         Begin VB.Label txtTeldisplay 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   1200
            Width           =   1815
         End
      End
      Begin VB.Frame fraSure 
         Enabled         =   0   'False
         Height          =   2295
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   3360
         Width           =   9015
         Begin VB.CommandButton cmdAgreesource2 
            Caption         =   "‰⁄„"
            Height          =   360
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   600
            Width           =   3135
         End
         Begin VB.CommandButton cmdNo 
            Caption         =   "·«"
            Height          =   360
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label Label10 
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
            TabIndex        =   15
            Top             =   840
            Width           =   1800
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1935
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   9015
         Begin VB.CommandButton Command2 
            Caption         =   "„Ê«›ﬁ"
            Height          =   360
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   480
            Width           =   6735
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Â« › «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "⁄‰Ê«‰ «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   9015
         Begin VB.ListBox List1 
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
            ItemData        =   "frmAccount.frx":007C
            Left            =   4080
            List            =   "frmAccount.frx":007E
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   3045
         End
         Begin VB.ListBox List2 
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
            ItemData        =   "frmAccount.frx":0080
            Left            =   360
            List            =   "frmAccount.frx":0082
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   240
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«”„ «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMain_Click()
On Error Resume Next
    frmAccount.Visible = False
    frmMain.Visible = True
    Unload frmAccount

End Sub

Private Sub Form_Load()
On Error Resume Next
'    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
'    Set rsStock = db.OpenRecordset("stock", dbOpenTable)
'    Set rsBank = db.OpenRecordset("bank", dbOpenTable)
'    Set rsChecks = db.OpenRecordset("checks", dbOpenTable)
'    Set rsCompany_name = db.OpenRecordset("company name", dbOpenTable)
'    Set rsFirm_name = db.OpenRecordset("firm name", dbOpenTable)
'    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
'    Set rsOffice_name = db.OpenRecordset("office name", dbOpenTable)
'    Set rsPerson_name = db.OpenRecordset("person name", dbOpenTable)
'    Set rsPharmacy_name = db.OpenRecordset("pharmacy name", dbOpenTable)
'    Set rsPurchases = db.OpenRecordset("purchases", dbOpenTable)
'    Set rsSales = db.OpenRecordset("sales", dbOpenTable)
'    Set rsSuper_market_name = db.OpenRecordset("super market name", dbOpenTable)
'    Set rsReturns= db.OpenRecordset("returns", dbOpenTable)

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    frmMain.Visible = True
End Sub
