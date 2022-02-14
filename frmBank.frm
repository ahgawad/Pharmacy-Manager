VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmBank 
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   10398
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "‘Ìﬂ« "
      TabPicture(0)   =   "frmBank.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "≈Ìœ«⁄"
      TabPicture(1)   =   "frmBank.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "”Õ»"
      TabPicture(2)   =   "frmBank.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame8 
         Height          =   2175
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid grdOut 
            Height          =   1815
            Left            =   120
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   3201
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2175
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   600
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid grdIn 
            Height          =   1815
            Left            =   120
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   3201
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.Frame Frame6 
         Height          =   855
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   9015
         Begin VB.CommandButton Command3 
            Caption         =   "„Ê«›ﬁ"
            Height          =   375
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   2175
         End
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
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2775
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1320
         Width           =   9015
         Begin VB.Frame Frame5 
            Caption         =   "Frame5"
            Height          =   15
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   0
            Width           =   1455
         End
         Begin MSFlexGridLib.MSFlexGrid grdChecks 
            Height          =   2415
            Left            =   120
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   4260
            _Version        =   393216
            Rows            =   1
            Cols            =   5
            FixedCols       =   0
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2760
         Width           =   9015
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "”Õ»"
            Height          =   375
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "«·—’Ìœ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   7080
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   795
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   1215
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "«·ﬁÌ„…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   1215
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2760
         Width           =   9015
         Begin VB.CommandButton Command1 
            Caption         =   "≈Ìœ«⁄"
            Height          =   375
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "«·—’Ìœ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   7080
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   795
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   240
            Width           =   1215
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "«·ﬁÌ„…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   1215
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   975
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   4080
         Width           =   9015
         Begin VB.CommandButton Command4 
            Caption         =   "’—›"
            Height          =   375
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   360
            Width           =   2895
         End
         Begin VB.ListBox List1 
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
            ItemData        =   "frmBank.frx":0054
            Left            =   4440
            List            =   "frmBank.frx":0056
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   360
            Width           =   1215
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "«·—’Ìœ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   360
            Width           =   795
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "„”·”·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5595
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   360
            Width           =   615
            WordWrap        =   -1  'True
         End
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
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   6120
      Width           =   3000
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMain_Click()
On Error Resume Next
    frmBank.Visible = False
    frmMain.Visible = True
    Unload frmBank

End Sub

Private Sub Command1_Click()
On Error Resume Next
    If Not IsNumeric(Text1.Text) Or Val(Text1.Text) <= 0 Then
        MsgBox "√œŒ· ﬁÌ„… «·≈Ìœ«⁄", vbExclamation, "Œÿ√"
        Text1.Text = ""
        Text1.SetFocus
    Else
        Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
        Set rsBank = db.OpenRecordset("bank", dbOpenTable)
        Dim X As Currency
        rsBank.MoveLast
        X = rsBank.Fields(4)
        rsBank.AddNew
        rsBank.Fields(1) = Date
        rsBank.Fields(2) = True
        rsBank.Fields(3) = Val(Text1.Text)
        rsBank.Fields(4) = X + Val(Text1.Text)
        rsBank.Update
        rsBank.MoveLast
        grdIn.AddItem (Format(rsBank.Fields(1), "long date") & vbTab & rsBank.Fields(3))
        Label4(0).Caption = X + Val(Text1.Text)
        Label4(1).Caption = X + Val(Text1.Text)
        Label4(2).Caption = X + Val(Text1.Text)
        Text1.Text = ""
        Text1.SetFocus
    End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
    If Not IsNumeric(Text2.Text) Or Val(Text2.Text) <= 0 Then
        MsgBox "√œŒ· ﬁÌ„… «·”Õ»", vbExclamation, "Œÿ√"
        Text2.Text = ""
        Text2.SetFocus
    ElseIf Val(Text2.Text) > Val(Label4(2).Caption) Then
        MsgBox "ﬁÌ„… «·”Õ» ·«   Ã«Ê“ «·—’Ìœ", vbExclamation, "Œÿ√"
        Text2.Text = ""
        Text2.SetFocus
    Else
        Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
        Set rsBank = db.OpenRecordset("bank", dbOpenTable)
        Dim X As Currency
        rsBank.MoveLast
        X = rsBank.Fields(4)
        rsBank.AddNew
        rsBank.Fields(1) = Date
        rsBank.Fields(2) = False
        rsBank.Fields(3) = Val(Text2.Text)
        rsBank.Fields(4) = X - Val(Text2.Text)
        rsBank.Update
        rsBank.MoveLast
        grdOut.AddItem (Format(rsBank.Fields(1), "long date") & vbTab & rsBank.Fields(3))
        Label4(0).Caption = X - Val(Text2.Text)
        Label4(1).Caption = X - Val(Text2.Text)
        Label4(2).Caption = X - Val(Text2.Text)
        Text2.Text = ""
        Text2.SetFocus
    End If
End Sub


Private Sub Command3_Click()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsChecks = db.OpenRecordset("checks", dbOpenTable)
    
    List1.Clear

    Dim i As Integer
    Dim X As Integer
    Dim B As Boolean
    
    grdChecks.Rows = 1
    grdChecks.FormatString = "<„”·”·|<—ﬁ„ «·‘Ìﬂ|<«·«”„|<ﬁÌ„… «·‘Ìﬂ|< «—ÌŒ «·≈” Õﬁ«ﬁ"
    grdChecks.ColWidth(0) = 160 * 4
    grdChecks.ColWidth(1) = 160 * 9
    grdChecks.ColWidth(2) = 160 * 12
    grdChecks.ColWidth(3) = 160 * 6
    grdChecks.ColWidth(4) = 160 * 12

    Dim ED As String

    rsChecks.MoveLast
    X = rsChecks.RecordCount
    
    rsChecks.MoveFirst
    If Check1.Value = 1 Then B = True Else B = False
    
    For i = 1 To X
        If rsChecks.Fields(3) <= Date And rsChecks.Fields(5) = False And rsChecks.Fields(6) = B Then
                Rem grid bill
                ED = Format(rsChecks.Fields(3), "d\\MMM\\yyyy")
                grdChecks.AddItem (rsChecks.Fields(0) & vbTab & rsChecks.Fields(1) & vbTab & rsChecks.Fields(4) & vbTab & rsChecks.Fields(2) & vbTab & ED)
                List1.AddItem rsChecks.Fields(0)
                Rem
        End If
        rsChecks.Move 1
    Next i
    Frame1.Enabled = True
If List1.ListCount = 0 Then Frame1.Enabled = False Else List1.Selected(0) = True
End Sub
Private Sub Command4_Click()
On Error Resume Next
If List1.ListCount = 0 Then
    Frame1.Enabled = False
Else
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsChecks = db.OpenRecordset("checks", dbOpenTable)
    Set rsBank = db.OpenRecordset("bank", dbOpenTable)

    Dim i As Integer
    Dim X As Integer
    Dim B As Boolean
    Dim ACD As Currency

    If Check1.Value = 1 Then B = True Else B = False

    rsChecks.MoveLast
    X = rsChecks.RecordCount
    rsChecks.MoveFirst

    For i = 1 To X
        If rsChecks.Fields(0) = List1.Text Then
            ACD = rsChecks.Fields(2)
        End If
        rsChecks.Move 1
    Next i

    Dim ACC As Currency
    rsBank.MoveLast
    ACC = rsBank.Fields(4)

If (B = False And ACD <= ACC) Or (B = True) Then
    rsChecks.MoveFirst
    For i = 1 To X
        If rsChecks.Fields(0) = List1.Text Then
                ACD = rsChecks.Fields(2)

                rsChecks.Edit
                rsChecks.Fields(5) = True
                rsChecks.Update
        End If
        rsChecks.Move 1
    Next i
    rsBank.AddNew
    rsBank.Fields(1) = Date
    rsBank.Fields(3) = ACD
    rsBank.Fields(2) = B
    If B Then rsBank.Fields(4) = ACC + ACD Else rsBank.Fields(4) = ACC - ACD
    rsBank.Update
    Call Command3_Click

    Dim ASS As Currency
    rsBank.MoveLast
    ASS = rsBank.Fields(4)
    Label4(0).Caption = ASS
    Label4(1).Caption = ASS
    Label4(2).Caption = ASS
    
    Else
        MsgBox "«·—’Ìœ ·« Ì”„Õ", vbExclamation, "Œÿ√"
    End If
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
'    Set rsStock = db.OpenRecordset("stock", dbOpenTable)
    Set rsBank = db.OpenRecordset("bank", dbOpenTable)
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
    
    grdChecks.Rows = 1
    grdChecks.FormatString = "<„”·”·|<—ﬁ„ «·‘Ìﬂ|<«·«”„|<ﬁÌ„… «·‘Ìﬂ|< «—ÌŒ «·≈” Õﬁ«ﬁ"
    grdChecks.ColWidth(0) = 160 * 4
    grdChecks.ColWidth(1) = 160 * 9
    grdChecks.ColWidth(2) = 160 * 12
    grdChecks.ColWidth(3) = 160 * 6
    grdChecks.ColWidth(4) = 160 * 12

    grdIn.Rows = 1
    grdIn.FormatString = "< «—ÌŒ «·≈Ìœ«⁄|<ﬁÌ„… «·≈Ìœ«⁄"
    grdIn.ColWidth(0) = 160 * 15
    grdIn.ColWidth(1) = 160 * 7

    grdOut.Rows = 1
    grdOut.FormatString = "< «—ÌŒ «·”Õ»|<ﬁÌ„… «·”Õ»"
    grdOut.ColWidth(0) = 160 * 15
    grdOut.ColWidth(1) = 160 * 7

Dim ASS As Currency
    rsBank.MoveLast
    ASS = rsBank.Fields(4)
    Label4(0).Caption = ASS
    Label4(1).Caption = ASS
    Label4(2).Caption = ASS

Dim X, i As Integer
rsBank.MoveLast
X = rsBank.RecordCount
rsBank.MoveFirst
For i = 1 To X
    If rsBank.Fields(1) >= Date - Month(6) And rsBank.Fields(2) <= Date Then
        If rsBank.Fields(2) = True Then grdIn.AddItem (Format(rsBank.Fields(1), "long date") & vbTab & rsBank.Fields(3)) Else grdOut.AddItem (Format(rsBank.Fields(1), "long date") & vbTab & rsBank.Fields(3))
    End If
    rsBank.Move 1
Next i

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    frmMain.Visible = True
End Sub
