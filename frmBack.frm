VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBack 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " «·„— Ã⁄« "
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "frmBack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGrid 
      Enabled         =   0   'False
      Height          =   2415
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3360
      Width           =   9255
      Begin VB.Frame fraCheck 
         Height          =   2175
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   120
         Width           =   2775
         Begin VB.TextBox txtCheckno 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   360
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton optCash 
            Alignment       =   1  'Right Justify
            Caption         =   "‘Ìﬂ"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optCash 
            Alignment       =   1  'Right Justify
            Caption         =   "ﬂ«‘"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.CommandButton cmdClosebill 
            Caption         =   "≈€·«ﬁ «·≈— Ã«⁄"
            Enabled         =   0   'False
            Height          =   375
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1680
            Width           =   1800
         End
         Begin MSComCtl2.DTPicker dtpDuedate 
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyy\MMM\d"
            Format          =   24444931
            CurrentDate     =   36412
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "—ﬁ„ «·‘Ìﬂ"
            Height          =   240
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   720
            Width           =   705
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   " «—ÌŒ «·”œ«œ"
            Height          =   195
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   1200
            Width           =   840
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "√÷› ≈·Ï «·≈— Ã«⁄"
         Height          =   375
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1920
         Width           =   1800
      End
      Begin VB.TextBox txtAlltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   345
         Left            =   3360
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1920
         Width           =   960
      End
      Begin MSFlexGridLib.MSFlexGrid grdBill 
         Height          =   1695
         Left            =   3000
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         RightToLeft     =   -1  'True
         FillStyle       =   1
         GridLines       =   3
         ScrollBars      =   2
         FormatString    =   "<ﬂÊœ|<’‰›|<À„‰ «·ÊÕœ…|<«·ﬂ„Ì…|< «—ÌŒ «·’·«ÕÌ…|<«·≈Ã„«·Ì"
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
   Begin VB.Frame fraBill 
      Enabled         =   0   'False
      Height          =   2535
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   780
      Width           =   9255
      Begin VB.CommandButton Command1 
         Caption         =   " €ÌÌ— «·„’œ—"
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
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2040
         Width           =   2160
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
         ItemData        =   "frmBack.frx":0442
         Left            =   960
         List            =   "frmBack.frx":0444
         RightToLeft     =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.ListBox txtPceprc 
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
         ItemData        =   "frmBack.frx":0446
         Left            =   3120
         List            =   "frmBack.frx":0448
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   765
      End
      Begin VB.CommandButton cmdCompute 
         Caption         =   "«Õ”»"
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
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   " ‰›Ì– ⁄„·Ì… «·≈— Ã«⁄"
         Enabled         =   0   'False
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2040
         Width           =   2160
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   1
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
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
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
         Left            =   2040
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Ã.„.˛"" #,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   2
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
         Left            =   840
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1440
         Width           =   975
      End
      Begin VB.ListBox cmpExpirydate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MM.yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   3
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
         ItemData        =   "frmBack.frx":044A
         Left            =   5280
         List            =   "frmBack.frx":044C
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
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
         ItemData        =   "frmBack.frx":044E
         Left            =   7440
         List            =   "frmBack.frx":0450
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   600
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
         ItemData        =   "frmBack.frx":0452
         Left            =   4200
         List            =   "frmBack.frx":0454
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   2895
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
         ItemData        =   "frmBack.frx":0456
         Left            =   7560
         List            =   "frmBack.frx":0458
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtAllstock 
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
         Left            =   840
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
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
         ItemData        =   "frmBack.frx":045A
         Left            =   7440
         List            =   "frmBack.frx":045C
         RightToLeft     =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "”⁄— ‘—«¡ «·ÊÕœ…"
         Height          =   390
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   120
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·≈Ã„«·Ì"
         Height          =   195
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1080
         Width           =   570
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·’·«ÕÌ…"
         Height          =   435
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   960
         Width           =   750
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ"
         Height          =   195
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1080
         Width           =   435
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·Õœ «·√œ‰Ï"
         Height          =   435
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   120
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂ„Ì… «·„— Ã⁄…"
         Height          =   390
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   960
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "’‰›"
         Height          =   195
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   330
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ"
         Height          =   195
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   270
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ﬂ· «·—’Ìœ"
         Height          =   435
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   120
         Width           =   615
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "«·ﬁ«∆„… «·—∆Ì”Ì…"
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
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   6120
      Width           =   3000
   End
   Begin VB.Frame fraSource0 
      Height          =   735
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdAgreesource0 
         Caption         =   "„Ê«›ﬁ"
         Height          =   360
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
      Begin VB.ListBox cmpSource 
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
         ItemData        =   "frmBack.frx":045E
         Left            =   4320
         List            =   "frmBack.frx":0460
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3045
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«”„ «·„‰‘√…"
         Height          =   195
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error Resume Next
fraGrid.Enabled = False
fraBill.Enabled = True
fraSource0.Enabled = False
Call cmdAgreesource0_Click
End Sub
Private Sub cmdAgreesource0_Click()
On Error Resume Next
If cmpSource.Text = "" Then
    MsgBox "√œŒ· «”„ «·„’œ—", vbExclamation, "Œÿ√"
    cmpSource.SetFocus
Else
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsStock = db.OpenRecordset("stock", dbOpenTable)
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    Set rsCheckno = db.OpenRecordset("Checkno", dbOpenTable)

    Dim Rc As Integer
    Dim Rn As Integer
    Dim Rc1 As Integer
    Dim Rn1 As Integer
    Dim Name As String
    Dim Rc2 As Integer
    Dim Rn2 As Integer

    fraSource0.Enabled = False
    fraBill.Enabled = True
    fraGrid.Enabled = False
    
    cmpItemcode.Clear
    txtItemname.Clear
    txtPceprc.Clear
    txtStock.Clear
    cmpExpirydate.Clear
    lstItemrec.Clear
    lstRecno.Clear

    rsItem_data.MoveFirst
    'txtMinlimit = rsItem_data.Fields(3)
    'txtAllstock = rsItem_data.Fields(6)

    rsStock.MoveLast
    Rc = rsStock.RecordCount

rsStock.MoveFirst
For Rn = 1 To Rc
    If rsStock.Fields(5) = SourceKind And rsStock.Fields(4) = cmpSource.Text Then
        cmpItemcode.AddItem rsStock.Fields(1)
        txtPceprc.AddItem rsStock.Fields(6)
        txtStock.AddItem rsStock.Fields(2)
        cmpExpirydate.AddItem Format(rsStock.Fields(3), "d\\MMM\\yyyy")
        lstRecno.AddItem Rn
    End If
    rsStock.Move 1
Next Rn

rsItem_data.MoveLast
Rc2 = rsItem_data.RecordCount

Rc1 = cmpItemcode.ListCount
rsItem_data.MoveFirst
For Rn1 = 1 To Rc1
    rsItem_data.MoveFirst
    For Rn2 = 1 To Rc2
    If cmpItemcode.List(Rn1 - 1) = rsItem_data.Fields(1) Then
        txtItemname.AddItem rsItem_data.Fields(2)
        lstItemrec.AddItem Rn2
    End If
    rsItem_data.Move 1
    Next Rn2
    
Next Rn1

        cmpItemcode.Selected(0) = True
        txtItemname.Selected(0) = True
        txtPceprc.Selected(0) = True
        txtStock.Selected(0) = True
        cmpExpirydate.Selected(0) = True
        lstItemrec.Selected(0) = True
        lstRecno.Selected(0) = True
        cmpItemcode.SetFocus

    txtTotal = ""
    txtQuantity = ""
    Call cmpItemcode_Scroll
End If
End Sub
Private Sub cmdClosebill_Click()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Select Case SourceKind
        Case "‘—ﬂ…"
            Set rsFound_name = db.OpenRecordset("company name", dbOpenTable)
        Case "„ﬂ »"
            Set rsFound_name = db.OpenRecordset("office name", dbOpenTable)
        Case "‘Œ’"
            Set rsFound_name = db.OpenRecordset("person name", dbOpenTable)
        Case "’Ìœ·Ì…"
            Set rsFound_name = db.OpenRecordset("pharmacy name", dbOpenTable)
        Case "”Ê»— „«—ﬂ "
            Set rsFound_name = db.OpenRecordset("super market name", dbOpenTable)
        Case "„’‰⁄"
            Set rsFound_name = db.OpenRecordset("firm name", dbOpenTable)
    End Select
Rem rem rem
If optCash(1) And (Not IsNumeric(txtCheckno.Text) Or Val(txtCheckno.Text) < 0) Then
        MsgBox "√œŒ· —ﬁ„ «·‘Ìﬂ", vbExclamation, "Œÿ√"
        txtCheckno.Text = ""
        txtCheckno.SetFocus
Else
        If optCash(1) Then
            Set rsChecks = db.OpenRecordset("checks", dbOpenTable)
            Set rsCheckno = db.OpenRecordset("Checkno", dbOpenTable)
            Set rsReturns = db.OpenRecordset("returns", dbOpenTable)

            rsChecks.AddNew
            rsChecks.Fields(1) = Val(txtCheckno.Text)
            rsChecks.Fields(2) = Val(txtAlltotal.Text)
            rsChecks.Fields(3) = dtpDuedate.Value
            rsChecks.Fields(4) = cmpSource.Text
            rsChecks.Fields(6) = True
            rsCheckno.MoveLast
            rsChecks.Fields(7) = rsCheckno.Fields(0)
            rsChecks.Update
            Rem
            Dim X, Y, XY As Integer
            rsFound_name.MoveLast
            X = rsFound_name.RecordCount
            rsFound_name.MoveFirst
            For Y = 1 To X
                If rsFound_name.Fields(1) = cmpSource.Text Then
                    XY = Val(rsFound_name.Fields(4))
                    rsFound_name.Edit
                    rsFound_name.Fields(4) = XY - Val(txtAlltotal.Text)
                    rsFound_name.Update
                End If
            rsFound_name.Move 1
            Next Y
            Rem
        
        
        End If
Rem
    rsCheckno.Edit
    rsCheckno.AddNew
    rsCheckno.Fields(1) = "any"
    rsCheckno.Update
Rem
    grdBill.Clear
    grdBill.Rows = 1
    grdBill.FormatString = "<ﬂÊœ|<’‰›|<À„‰ «·ÊÕœ…|<«·ﬂ„Ì…|< «—ÌŒ «·’·«ÕÌ…|<«·≈Ã„«·Ì"
    grdBill.ColWidth(0) = 160 * 4
    grdBill.ColWidth(1) = 160 * 7
    grdBill.ColWidth(2) = 160 * 6
    grdBill.ColWidth(3) = 160 * 5
    grdBill.ColWidth(4) = 160 * 9
    grdBill.ColWidth(5) = 160 * 5
    'txtBill.Text = ""
    'txtItemname.Text = ""
    'txtStock.Text = ""
    'txtPceprc.Text = ""
    'txtQuantity.Text = ""
    'txtDiscount.Text = ""
    'txtBonus.Text = ""
    'txtTotal.Text = ""
    txtAlltotal.Text = ""
    txtCheckno.Text = ""
    dtpDuedate.Month = Month(Date) + 2

    cmdSave.Enabled = False
    cmdClosebill.Enabled = False
    cmdAdd.Enabled = False
        
    optCash(0).Enabled = False
    optCash(0).Value = True
    optCash(1).Enabled = False

    fraSource0.Enabled = True
    fraBill.Enabled = False
    fraGrid.Enabled = False
'    Call cmdAgreesource0_Click

End If
End Sub
Private Sub cmdCompute_Click()
On Error Resume Next
    If cmpItemcode.Text = "" Then
        cmpItemcode.SetFocus
    ElseIf cmpExpirydate.Text = "" Then
        MsgBox "«Œ —  «—ÌŒ «·’·«ÕÌ…", vbExclamation, "Œÿ√"
        cmpExpirydate.SetFocus
    ElseIf Not IsNumeric(txtQuantity.Text) Or Val(txtQuantity.Text) <= 0 Then
        MsgBox "√œŒ· «·ﬂ„Ì…", vbExclamation, "Œÿ√"
        txtQuantity.Text = ""
        txtQuantity.SetFocus
    ElseIf Val(txtQuantity.Text) > txtStock.Text Then
        MsgBox "«·ﬂ„Ì… ÌÃ» √·«   Ã«Ê“ «·„Œ“Ê‰", vbExclamation, "Œÿ√"
        txtQuantity.Text = ""
        txtQuantity.SetFocus
    Else
        txtTotal.Text = Val(txtPceprc.Text) * Val(txtQuantity.Text)
        cmdSave.Enabled = True
    End If
End Sub
Private Sub cmdMain_Click()
On Error Resume Next
    frmBack.Visible = False
    frmMain.Visible = True
    Unload frmBack
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsStock = db.OpenRecordset("stock", dbOpenTable)
'    Set rsBank = db.OpenRecordset("bank", dbOpenTable)
    Set rsCheckno = db.OpenRecordset("checkno", dbOpenTable)
'    Set rsCompany_name = db.OpenRecordset("company name", dbOpenTable)
'    Set rsFirm_name = db.OpenRecordset("firm name", dbOpenTable)
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
'    Set rsOffice_name = db.OpenRecordset("office name", dbOpenTable)
'    Set rsPerson_name = db.OpenRecordset("person name", dbOpenTable)
'    Set rsPharmacy_name = db.OpenRecordset("pharmacy name", dbOpenTable)
'    Set rsPurchases = db.OpenRecordset("purchases", dbOpenTable)
'    Set rsSales = db.OpenRecordset("sales", dbOpenTable)
'    Set rsSuper_market_name = db.OpenRecordset("super market name", dbOpenTable)
    Set rsReturns = db.OpenRecordset("returns", dbOpenTable)

    
Rem rem
    Dim nn1 As Long
    rsItem_data.MoveFirst
    rsItem_data.Move Val(lstItemrec.Text) - 1
    nn1 = rsItem_data.Fields(6)
    rsItem_data.Edit
    rsItem_data.Fields(6) = nn1 - Val(txtQuantity.Text)
    rsItem_data.Update

    Dim Nn11 As Long
    Dim Nn111 As Long
    rsStock.MoveFirst
    rsStock.Move Val(lstRecno.Text) - 1
    Nn11 = rsStock.Fields(2)
'    Nn111 = rsStock.Fields(7)
    rsStock.Edit
    rsStock.Fields(2) = Nn11 - Val(txtQuantity.Text)
    rsStock.Update
    If rsStock.Fields(2) <= 0 Then rsStock.Delete

Rem *****************************
    If Not (rsReturns.BOF And rsReturns.EOF) Then
        rsReturns.MoveLast
    End If
Rem *****************************
    
    rsReturns.AddNew
    rsReturns.Fields(1) = Val(cmpItemcode.Text)
    rsReturns.Fields(2) = Val(txtQuantity.Text)
    rsReturns.Fields(3) = Val(txtPceprc.Text)
    rsReturns.Fields(4) = Date
    rsReturns.Fields(5) = Format(Val(cmpExpirydate.Text), "Short Date")
    rsReturns.Fields(6) = cmpSource.Text
    rsReturns.Fields(7) = SourceKind
    'rsReturns.Fields(8) = 'check yes\no
    
    rsCheckno.MoveLast
    rsReturns.Fields(9) = rsCheckno.Fields(0)
    rsReturns.Update

    fraGrid.Enabled = True

Rem grid bill
    grdBill.AddItem (cmpItemcode.Text & vbTab & txtItemname.Text & vbTab & txtQuantity.Text & vbTab & txtStock.Text & vbTab & Format(cmpExpirydate.Text, "Long Date") & vbTab & txtTotal.Text)
Rem
    txtAlltotal.Text = Val(txtAlltotal.Text) + Val(txtTotal.Text)
'    txtItemname.Text = ""
'    txtStock.Text = ""
'    txtPceprc.Text = ""
'    txtQuantity.Text = ""
'    txtDiscount.Text = ""
'    txtBonus.Text = ""
'    txtTotal.Text = ""
    cmdSave.Enabled = False
    cmdClosebill.Enabled = True
    cmdAdd.Enabled = True
    optCash(0).Enabled = True
    optCash(1).Enabled = True
    fraGrid.Enabled = True
Rem
    fraBill.Enabled = False
Rem rem

End Sub

Private Sub cmpExpirydate_LostFocus()
On Error Resume Next
    Call cmpExpirydate_Scroll
End Sub
Private Sub cmpExpirydate_Scroll()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    
    Dim rec_no As Integer
    rec_no = cmpExpirydate.ListIndex
    If cmpExpirydate.ListCount > 0 Then
        cmpItemcode.Selected(rec_no) = True
        txtItemname.Selected(rec_no) = True
        txtPceprc.Selected(rec_no) = True
        txtStock.Selected(rec_no) = True
        lstItemrec.Selected(rec_no) = True
        lstRecno.Selected(rec_no) = True
    End If
    rsItem_data.MoveFirst
    rsItem_data.Move Val(lstItemrec.Text) - 1
        
    txtMinlimit.Text = rsItem_data.Fields(3)
    txtAllstock.Text = rsItem_data.Fields(6)
If Val(txtAllstock.Text) <= Val(txtMinlimit.Text) Then txtAllstock.BackColor = RGB(255, 0, 0) Else txtAllstock.BackColor = &H80000004
End Sub
Private Sub cmpItemcode_LostFocus()
On Error Resume Next
    Call cmpItemcode_Scroll
End Sub
Private Sub cmpItemcode_Scroll()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)

    Dim rec_no As Integer
    rec_no = cmpItemcode.ListIndex
    If cmpItemcode.ListCount > 0 Then
        txtItemname.Selected(rec_no) = True
        txtPceprc.Selected(rec_no) = True
        txtStock.Selected(rec_no) = True
        cmpExpirydate.Selected(rec_no) = True
        lstItemrec.Selected(rec_no) = True
        lstRecno.Selected(rec_no) = True
    End If
    
    rsItem_data.MoveFirst
    rsItem_data.Move Val(lstItemrec.Text) - 1
        
    txtMinlimit.Text = rsItem_data.Fields(3)
    txtAllstock.Text = rsItem_data.Fields(6)
    If Val(txtAllstock.Text) <= Val(txtMinlimit.Text) Then txtAllstock.BackColor = RGB(255, 0, 0) Else txtAllstock.BackColor = &H80000004
End Sub

Private Sub Command1_Click()
On Error Resume Next
    fraSource0.Enabled = True
    fraBill.Enabled = False
    fraGrid.Enabled = False
    grdBill.Clear
    grdBill.Rows = 1
    grdBill.FormatString = "<ﬂÊœ|<’‰›|<À„‰ «·ÊÕœ…|<«·ﬂ„Ì…|< «—ÌŒ «·’·«ÕÌ…|<«·≈Ã„«·Ì"
    grdBill.ColWidth(0) = 160 * 4
    grdBill.ColWidth(1) = 160 * 7
    grdBill.ColWidth(2) = 160 * 6
    grdBill.ColWidth(3) = 160 * 5
    grdBill.ColWidth(4) = 160 * 9
    grdBill.ColWidth(5) = 160 * 5
    txtAlltotal.Text = ""
    txtCheckno.Text = ""
    dtpDuedate.Month = Month(Date) + 2

End Sub
Private Sub Form_Load()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsCheckno = db.OpenRecordset("Checkno", dbOpenTable)
Rem
    rsCheckno.Edit
    rsCheckno.AddNew
    rsCheckno.Fields(1) = "any"
    rsCheckno.Update
Rem
    
    grdBill.ColWidth(0) = 160 * 4
    grdBill.ColWidth(1) = 160 * 7
    grdBill.ColWidth(2) = 160 * 6
    grdBill.ColWidth(3) = 160 * 5
    grdBill.ColWidth(4) = 160 * 9
    grdBill.ColWidth(5) = 160 * 5

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    frmMain.Visible = True
End Sub
Private Sub optCash_Click(Index As Integer)
On Error Resume Next
    If Index = 1 Then
    dtpDuedate.Enabled = True
    txtCheckno.Enabled = True
    Else
    dtpDuedate.Enabled = False
    txtCheckno.Enabled = False
    End If
End Sub
Private Sub txtItemname_LostFocus()
On Error Resume Next
    Call txtItemname_Scroll
End Sub
Private Sub txtItemname_Scroll()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    
    Dim rec_no As Integer
    rec_no = txtItemname.ListIndex
    If txtItemname.ListCount > 0 Then
        cmpItemcode.Selected(rec_no) = True
        txtPceprc.Selected(rec_no) = True
        txtStock.Selected(rec_no) = True
        cmpExpirydate.Selected(rec_no) = True
        lstItemrec.Selected(rec_no) = True
        lstRecno.Selected(rec_no) = True
    End If
    rsItem_data.MoveFirst
    rsItem_data.Move Val(lstItemrec.Text) - 1
        
    txtMinlimit.Text = rsItem_data.Fields(3)
    txtAllstock.Text = rsItem_data.Fields(6)
    If Val(txtAllstock.Text) <= Val(txtMinlimit.Text) Then txtAllstock.BackColor = RGB(255, 0, 0) Else txtAllstock.BackColor = &H80000004
End Sub
Private Sub txtPceprc_LostFocus()
On Error Resume Next
    Call txtPceprc_Scroll
End Sub
Private Sub txtPceprc_Scroll()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    rsItem_data.MoveFirst
    rsItem_data.Move Val(lstItemrec.Text) - 1
        
    txtMinlimit.Text = rsItem_data.Fields(3)
    txtAllstock.Text = rsItem_data.Fields(6)
    
    Dim rec_no As Integer
    rec_no = txtPceprc.ListIndex

    If txtPceprc.ListCount > 0 Then
        cmpItemcode.Selected(rec_no) = True
        txtItemname.Selected(rec_no) = True
        txtStock.Selected(rec_no) = True
        cmpExpirydate.Selected(rec_no) = True
        lstItemrec.Selected(rec_no) = True
        lstRecno.Selected(rec_no) = True
    If Val(txtAllstock.Text) <= Val(txtMinlimit.Text) Then txtAllstock.BackColor = RGB(255, 0, 0) Else txtAllstock.BackColor = &H80000004
    End If
End Sub
Private Sub txtStock_LostFocus()
On Error Resume Next
    Call txtStock_Scroll
End Sub
Private Sub txtStock_Scroll()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    
    Dim rec_no As Integer
    rec_no = txtStock.ListIndex
    If txtStock.ListCount > 0 Then
        cmpItemcode.Selected(rec_no) = True
        txtItemname.Selected(rec_no) = True
        txtPceprc.Selected(rec_no) = True
        cmpExpirydate.Selected(rec_no) = True
        lstItemrec.Selected(rec_no) = True
        lstRecno.Selected(rec_no) = True
    End If
    rsItem_data.MoveFirst
    rsItem_data.Move Val(lstItemrec.Text) - 1
        
    txtMinlimit.Text = rsItem_data.Fields(3)
    txtAllstock.Text = rsItem_data.Fields(6)
If Val(txtAllstock.Text) <= Val(txtMinlimit.Text) Then txtAllstock.BackColor = RGB(255, 0, 0) Else txtAllstock.BackColor = &H80000004
End Sub
