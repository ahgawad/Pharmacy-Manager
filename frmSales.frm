VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·„»Ì⁄« "
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "frmSales.frx":0000
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
   Begin VB.Frame fraBill 
      Height          =   2775
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   9255
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
         ItemData        =   "frmSales.frx":0442
         Left            =   6360
         List            =   "frmSales.frx":0444
         RightToLeft     =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   900
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
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
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
         ItemData        =   "frmSales.frx":0446
         Left            =   7800
         List            =   "frmSales.frx":0448
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1560
         Width           =   975
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
         ItemData        =   "frmSales.frx":044A
         Left            =   4200
         List            =   "frmSales.frx":044C
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   600
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
         ItemData        =   "frmSales.frx":044E
         Left            =   7440
         List            =   "frmSales.frx":0450
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   1005
      End
      Begin VB.ListBox cmpExpirydate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "yyyy\\MMM\\d"
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
         ItemData        =   "frmSales.frx":0452
         Left            =   5640
         List            =   "frmSales.frx":0454
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtPceprc 
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
         Left            =   3120
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   22
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
         Left            =   480
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1560
         Width           =   975
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   " ‰›Ì– ⁄„·Ì… «·»Ì⁄"
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2280
         Width           =   3000
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
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1560
         Width           =   2175
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
         ItemData        =   "frmSales.frx":0456
         Left            =   2280
         List            =   "frmSales.frx":0458
         RightToLeft     =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ﬂ· «·—’Ìœ"
         Height          =   435
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   120
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ"
         Height          =   195
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   270
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂ„Ì… «·„»«⁄…"
         Height          =   390
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1080
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·Õœ «·√œ‰Ï"
         Height          =   435
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   120
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ"
         Height          =   195
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1200
         Width           =   435
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·’·«ÕÌ…"
         Height          =   435
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1080
         Width           =   750
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·≈Ã„«·Ì"
         Height          =   195
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1200
         Width           =   570
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "”⁄— «·ÊÕœ…"
         Height          =   435
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   120
         Width           =   435
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "«·ﬁ«∆„… «·—∆Ì”Ì…"
      Height          =   375
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   6120
      Width           =   3000
   End
   Begin VB.Frame fraGrid 
      Enabled         =   0   'False
      Height          =   3015
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2880
      Width           =   9255
      Begin VB.CommandButton cmdAdd 
         Caption         =   "√÷› ≈·Ï «·›« Ê—…"
         Height          =   375
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Frame Frame1 
         Height          =   2775
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   2775
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   360
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton cmdClosebill 
            Caption         =   "≈€·«ﬁ «·›« Ê—…"
            Enabled         =   0   'False
            Height          =   375
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   2160
            Width           =   1800
         End
         Begin VB.OptionButton optCash 
            Alignment       =   1  'Right Justify
            Caption         =   "ﬂ«‘"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optCash 
            Alignment       =   1  'Right Justify
            Caption         =   "‘Ìﬂ"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtCheckno 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   360
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   720
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtpDuedate 
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   1680
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyy\MMM\d"
            Format          =   24510467
            CurrentDate     =   36412
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "„” Õﬁ „‰"
            Height          =   195
            Left            =   1785
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   " «—ÌŒ «·”œ«œ"
            Height          =   195
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   1680
            Width           =   840
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "—ﬁ„ «·‘Ìﬂ"
            Height          =   240
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   720
            Width           =   705
         End
      End
      Begin VB.TextBox txtAlltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   345
         Left            =   3360
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2520
         Width           =   960
      End
      Begin MSFlexGridLib.MSFlexGrid grdBill 
         Height          =   2295
         Left            =   3000
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         RightToLeft     =   -1  'True
         FillStyle       =   1
         GridLines       =   3
         ScrollBars      =   2
         FormatString    =   "<’‰›|< «—ÌŒ «·’·«ÕÌ…|<À„‰ «·ÊÕœ…|<«·ﬂ„Ì…|<«·≈Ã„«·Ì"
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
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
On Error Resume Next
        txtMinlimit.Text = ""
        txtAllstock.Text = ""
        txtPceprc.Text = ""
        txtQuantity.Text = ""
        txtTotal.Text = ""
        txtAlltotal.Text = ""
        txtCheckno.Text = ""
        txtStock.Clear
        cmpExpirydate.Clear
        dtpDuedate.Month = Month(Date) + 2
    
    fraGrid.Enabled = False
    fraBill.Enabled = True
    cmpItemcode.SetFocus


End Sub

Private Sub cmdClosebill_Click()
On Error Resume Next
    If optCash(1) And (Not IsNumeric(txtCheckno.Text) Or Val(txtCheckno.Text) < 0) Then
        MsgBox "√œŒ· —ﬁ„ «·‘Ìﬂ", vbExclamation, "Œÿ√"
        txtCheckno.Text = ""
        txtCheckno.SetFocus
    ElseIf optCash(1) And Text1.Text = "" Then
        MsgBox "√œŒ· «·„” Õﬁ „‰Â «·‘Ìﬂ", vbExclamation, "Œÿ√"
        Text1.Text = ""
        Text1.SetFocus
    Else
        If optCash(1) Then
            Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
            Set rsChecks = db.OpenRecordset("checks", dbOpenTable)
            Set rsCheckno = db.OpenRecordset("Checkno", dbOpenTable)
            rsChecks.AddNew
            rsChecks.Fields(1) = Val(txtCheckno.Text)
            rsChecks.Fields(2) = Val(txtAlltotal.Text)
            rsChecks.Fields(3) = dtpDuedate.Value
            rsChecks.Fields(4) = Text1.Text
            rsChecks.Fields(6) = True
            rsCheckno.MoveLast
            rsChecks.Fields(7) = rsCheckno.Fields(0)
            rsChecks.Update

        End If
Rem
    rsCheckno.Edit
    rsCheckno.AddNew
    rsCheckno.Fields(1) = "any"
    rsCheckno.Update
Rem
    
        grdBill.Clear
        grdBill.Rows = 1


        txtMinlimit.Text = ""
        txtAllstock.Text = ""
        txtPceprc.Text = ""
        txtQuantity.Text = ""
        txtTotal.Text = ""
        txtAlltotal.Text = ""
        txtCheckno.Text = ""
        txtStock.Clear
        cmpExpirydate.Clear
        dtpDuedate.Month = Month(Date) + 2
    
        cmdSave.Enabled = False
        cmdClosebill.Enabled = False
        cmdAdd.Enabled = False
    
        optCash(0).Enabled = False
        optCash(0).Value = True
        optCash(1).Enabled = False
    
        fraBill.Enabled = True
        fraGrid.Enabled = False
    Dim Z1 As Integer
    Z1 = cmpItemcode.ListIndex
    cmpItemcode.Selected(Z1) = True
    txtItemname.Selected(Z1) = True
'    cmpItemcode.SetFocus
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
    frmSales.Visible = False
    frmMain.Visible = True
    Unload frmSales
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsSales = db.OpenRecordset("sales", dbOpenTable)
    Set rsStock = db.OpenRecordset("stock", dbOpenTable)
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    Set rsCheckno = db.OpenRecordset("Checkno", dbOpenTable)

    rsSales.AddNew
    rsSales.Fields(1) = Val(cmpItemcode.Text)
    rsSales.Fields(2) = Val(txtQuantity.Text)
    rsSales.Fields(3) = Val(txtTotal.Text)
    rsSales.Fields(4) = Date
    rsSales.Fields(5) = Format(Val(cmpExpirydate.Text), "Short Date")
    rsCheckno.MoveLast
    rsSales.Fields(7) = rsCheckno.Fields(0)
    rsSales.Update
    rsSales.MoveLast

    Dim Qnt As Long
    Dim Qnt2 As Long
    
    rsStock.MoveFirst
    rsStock.Move Val(lstRecno.Text) - 1
    Qnt = rsStock.Fields(2)
    rsStock.Edit
    rsStock.Fields(2) = Qnt - Val(txtQuantity.Text)
    rsStock.Update
    If rsStock.Fields(2) = 0 Then
        rsStock.Edit
        rsStock.Delete

    End If

    rsItem_data.MoveFirst
    rsItem_data.Move Val(lstItemrec.Text) - 1
    Qnt2 = rsItem_data.Fields(6)
    rsItem_data.Edit
    rsItem_data.Fields(6) = Qnt2 - Val(txtQuantity.Text)
    rsItem_data.Update

    Dim Toto As Long
    Toto = Val(txtAlltotal.Text)
    txtAlltotal.Text = Toto + Val(txtTotal.Text)

    grdBill.AddItem (txtItemname.Text & vbTab & cmpExpirydate.Text & vbTab & txtPceprc.Text & vbTab & txtQuantity.Text & vbTab & txtTotal.Text)
    
    cmpItemcode.SetFocus
    cmdSave.Enabled = False

    fraGrid.Enabled = True

    cmdSave.Enabled = False
    cmdClosebill.Enabled = True
    cmdAdd.Enabled = True
    optCash(0).Enabled = True
    optCash(1).Enabled = True
    fraBill.Enabled = False

End Sub
Private Sub cmpExpirydate_LostFocus()
On Error Resume Next
    Call cmpExpirydate_Scroll
    
End Sub


Private Sub cmpExpirydate_Scroll()
On Error Resume Next
    Dim rec_no As Integer
    rec_no = cmpExpirydate.ListIndex
    If lstRecno.ListCount > 0 Then
        txtStock.Selected(rec_no) = True
        lstRecno.Selected(rec_no) = True
    End If
End Sub


Private Sub cmpItemcode_LostFocus()
On Error Resume Next
    Call cmpItemcode_Scroll
End Sub
Private Sub cmpItemcode_Scroll()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    Set rsStock = db.OpenRecordset("stock", dbOpenTable)
    
    Dim rec_no As Integer
    rec_no = cmpItemcode.ListIndex

    If cmpItemcode.ListCount > 0 Then
        txtItemname.Selected(rec_no) = True
        lstItemrec.Selected(rec_no) = True

        rsItem_data.MoveFirst
        rsItem_data.Move Val(lstItemrec.Text) - 1
    
        txtPceprc.Text = rsItem_data.Fields(4)
        txtMinlimit.Text = rsItem_data.Fields(3)
        txtAllstock.Text = rsItem_data.Fields(6)
        If Val(txtAllstock.Text) <= Val(txtMinlimit.Text) Then txtAllstock.BackColor = RGB(255, 0, 0) Else txtAllstock.BackColor = &H80000004

        cmpExpirydate.Clear
        txtStock.Clear
        lstRecno.Clear
    
        Dim expDate As String
        Dim stkQnt As Long
        Dim rec As Integer
        Dim Rec1 As Integer
        Dim no As Integer

        rsStock.MoveLast
        rec = rsStock.RecordCount
        rsStock.MoveFirst
  
        For Rec1 = 1 To rec
            If rsStock.Fields(1) = Val(cmpItemcode.Text) Then
                expDate = Format(rsStock.Fields(3), "d\\MMM\\yyyy")
                stkQnt = rsStock.Fields(2)
                cmpExpirydate.AddItem expDate
                txtStock.AddItem stkQnt
                lstRecno.AddItem Rec1
            End If
            rsStock.Move 1
        Next Rec1
    
        If lstRecno.ListCount > 0 Then
            cmpExpirydate.Selected(0) = True
            txtStock.Selected(0) = True
            lstRecno.Selected(0) = True
        End If
    End If
End Sub
Private Sub Form_Load()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    Set rsStock = db.OpenRecordset("stock", dbOpenTable)
    dtpDuedate.Value = Date
    dtpDuedate.Month = Month(Date) + 2
Rem
    Set rsCheckno = db.OpenRecordset("Checkno", dbOpenTable)
    rsCheckno.Edit
    rsCheckno.AddNew
    rsCheckno.Fields(1) = "any"
    rsCheckno.Update
Rem
    
    
    Dim l_rec_no1 As Integer
    
    rsItem_data.MoveLast
    l_rec_no1 = rsItem_data.RecordCount
    rsItem_data.MoveFirst
    
    Dim J1 As Integer
    For J1 = 1 To l_rec_no1
        If rsItem_data.Fields(6) > 0 Then
            cmpItemcode.AddItem rsItem_data.Fields(1)
            txtItemname.AddItem rsItem_data.Fields(2)
            lstItemrec.AddItem J1
        End If
        rsItem_data.Move 1
    Next J1
    
    If cmpItemcode.ListCount > 0 Then
        cmpItemcode.Selected(0) = True
        txtItemname.Selected(0) = True
        lstItemrec.Selected(0) = True
    End If
    
    grdBill.ColWidth(0) = 160 * 8
    grdBill.ColWidth(1) = 160 * 10
    grdBill.ColWidth(2) = 160 * 6
    grdBill.ColWidth(3) = 160 * 6
    grdBill.ColWidth(4) = 160 * 6

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
        Text1.Enabled = True
    Else
        dtpDuedate.Enabled = False
        txtCheckno.Enabled = False
        txtCheckno.Text = ""
        Text1.Text = ""
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
    Set rsStock = db.OpenRecordset("stock", dbOpenTable)
    
    Dim rec_no As Integer
    rec_no = txtItemname.ListIndex

    If cmpItemcode.ListCount > 0 Then
        cmpItemcode.Selected(rec_no) = True
        lstItemrec.Selected(rec_no) = True

        rsItem_data.MoveFirst
        rsItem_data.Move Val(lstItemrec.Text) - 1
        txtPceprc.Text = rsItem_data.Fields(4)
        txtMinlimit.Text = rsItem_data.Fields(3)
        txtAllstock.Text = rsItem_data.Fields(6)
        If Val(txtAllstock.Text) <= Val(txtMinlimit.Text) Then txtAllstock.BackColor = RGB(255, 0, 0) Else txtAllstock.BackColor = &H80000004
        
        cmpExpirydate.Clear
        txtStock.Clear
        lstRecno.Clear
        
        Dim expDate As String
        Dim stkQnt As Long
        Dim rec As Integer
        Dim Rec1 As Integer
        Dim no As Integer
    
        rsStock.MoveLast
        rec = rsStock.RecordCount
        rsStock.MoveFirst
        
        For Rec1 = 1 To rec
            If rsStock.Fields(1) = Val(cmpItemcode.Text) Then
                expDate = Format(rsStock.Fields(3), "d\\MMM\\yyyy")
                stkQnt = rsStock.Fields(2)
                cmpExpirydate.AddItem expDate
                txtStock.AddItem stkQnt
                lstRecno.AddItem Rec1
            End If
            rsStock.Move 1
        Next Rec1
        If lstRecno.ListCount > 0 Then
            cmpExpirydate.Selected(0) = True
            txtStock.Selected(0) = True
            lstRecno.Selected(0) = True
        End If
    End If
End Sub
Private Sub txtStock_LostFocus()
On Error Resume Next
    Call txtStock_Scroll

End Sub
Private Sub txtStock_Scroll()
On Error Resume Next
    Dim rec_no As Integer
    rec_no = txtStock.ListIndex
    If lstRecno.ListCount > 0 Then
        cmpExpirydate.Selected(rec_no) = True
        lstRecno.Selected(rec_no) = True
    End If
End Sub
