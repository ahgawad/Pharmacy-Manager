VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmDates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«· Ê«—ÌŒ"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "frmDates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "’·«ÕÌ… √’‰«›"
      TabPicture(0)   =   "frmDates.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "«” Õﬁ«ﬁ ‘Ìﬂ« "
      TabPicture(1)   =   "frmDates.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   -74640
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   540
         Width           =   8895
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "ﬁ«∆„… «·„œÌ‰"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "„Ê«›ﬁ"
            Height          =   375
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   6480
            TabIndex        =   8
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyy\MMM\d"
            Format          =   24510467
            CurrentDate     =   36412
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   3960
            TabIndex        =   9
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyy\MMM\d"
            Format          =   24510467
            CurrentDate     =   36412
         End
         Begin VB.Label Label4 
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
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label3 
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
            TabIndex        =   17
            Top             =   240
            Width           =   195
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3495
         Left            =   -74640
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1500
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid grdChecks 
            Height          =   3135
            Left            =   720
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   5530
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FixedCols       =   0
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            RightToLeft     =   -1  'True
            FillStyle       =   1
            GridLines       =   3
            ScrollBars      =   2
            FormatString    =   "<—ﬁ„ «·‘Ìﬂ|<«·«”„|<ﬁÌ„… «·‘Ìﬂ|< «—ÌŒ «·≈” Õﬁ«ﬁ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   8895
         Begin VB.CommandButton Command1 
            Caption         =   "„Ê«›ﬁ"
            Height          =   375
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   6480
            TabIndex        =   3
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
            TabIndex        =   4
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyy\MMM\d"
            Format          =   24510467
            CurrentDate     =   36412
         End
         Begin VB.Label Label1 
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
            TabIndex        =   15
            Top             =   240
            Width           =   195
         End
         Begin VB.Label Label2 
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
            TabIndex        =   14
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3495
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1500
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid grdBill 
            Height          =   3135
            Left            =   960
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5530
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FixedCols       =   0
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            RightToLeft     =   -1  'True
            FillStyle       =   1
            GridLines       =   3
            ScrollBars      =   2
            FormatString    =   "<ﬂÊœ|<’‰›|< «—ÌŒ «·’·«ÕÌ…|<«·ﬂ„Ì…"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "«·ﬁ«∆„… «·—∆Ì”Ì…"
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   6120
      Width           =   3000
   End
End
Attribute VB_Name = "frmDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMain_Click()
  On Error Resume Next
  frmDates.Visible = False
    frmMain.Visible = True
    Unload frmDates

End Sub
Private Sub Command1_Click()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    Set rsStock = db.OpenRecordset("stock", dbOpenTable)
    Dim Rn As Integer
    Dim I As Integer
    Dim X, Y As Integer

    grdBill.Rows = 1
    grdBill.FormatString = "<ﬂÊœ|<’‰›|< «—ÌŒ «·’·«ÕÌ…|<«·ﬂ„Ì…"
    grdBill.ColWidth(0) = 160 * 6
    grdBill.ColWidth(1) = 160 * 15
    grdBill.ColWidth(2) = 160 * 15
    grdBill.ColWidth(3) = 160 * 6

    Dim ED As String

    rsItem_data.MoveLast
    X = rsItem_data.RecordCount

    rsStock.MoveLast
    Rn = rsStock.RecordCount
    
    rsStock.MoveFirst
    
    For I = 1 To Rn
        If rsStock.Fields(3) >= DTPicker1.Value And rsStock.Fields(3) <= DTPicker2.Value Then
            rsItem_data.MoveFirst
            For Y = 1 To X
                If rsItem_data.Fields(1) = rsStock.Fields(1) Then
                    Rem grid bill
                    ED = Format(rsStock.Fields(3), "d\\MMM\\yyyy")
                    grdBill.AddItem (rsStock.Fields(1) & vbTab & rsItem_data.Fields(2) & vbTab & ED & vbTab & rsStock.Fields(2))
                    Rem
                End If
            rsItem_data.Move 1
            Next Y
            End If
        rsStock.Move 1
    Next I

       
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsChecks = db.OpenRecordset("checks", dbOpenTable)

    Dim I As Integer
    Dim X As Integer
    Dim B As Boolean
    
    grdChecks.Rows = 1
    grdChecks.FormatString = "<—ﬁ„ «·‘Ìﬂ|<«·«”„|<ﬁÌ„… «·‘Ìﬂ|< «—ÌŒ «·≈” Õﬁ«ﬁ"
    grdChecks.ColWidth(0) = 160 * 9
    grdChecks.ColWidth(1) = 160 * 15
    grdChecks.ColWidth(2) = 160 * 6
    grdChecks.ColWidth(3) = 160 * 15

    Dim ED As String

    rsChecks.MoveLast
    X = rsChecks.RecordCount
    
    rsChecks.MoveFirst
    If Check1.Value = 1 Then B = True Else B = False
    
    For I = 1 To X
        If rsChecks.Fields(3) >= DTPicker3.Value And rsChecks.Fields(3) <= DTPicker4.Value And rsChecks.Fields(6) = B Then
                Rem grid bill
                ED = Format(rsChecks.Fields(3), "d\\MMM\\yyyy")
                grdChecks.AddItem (rsChecks.Fields(1) & vbTab & rsChecks.Fields(4) & vbTab & rsChecks.Fields(2) & vbTab & ED)
                Rem
        End If
        rsChecks.Move 1
    Next I
End Sub
Private Sub Form_Load()
On Error Resume Next
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker3.Value = Date
    DTPicker4.Value = Date
    DTPicker2.Year = Year(Date) + 2
    DTPicker4.Month = Month(Date) + 2
    
    grdBill.ColWidth(0) = 160 * 6
    grdBill.ColWidth(1) = 160 * 15
    grdBill.ColWidth(2) = 160 * 15
    grdBill.ColWidth(3) = 160 * 6


    grdChecks.ColWidth(0) = 160 * 9
    grdChecks.ColWidth(1) = 160 * 15
    grdChecks.ColWidth(2) = 160 * 6
    grdChecks.ColWidth(3) = 160 * 15


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    frmMain.Visible = True
End Sub

