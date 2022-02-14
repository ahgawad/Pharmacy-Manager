VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPurchases 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·„‘ —Ì« "
   ClientHeight    =   6825
   ClientLeft      =   1680
   ClientTop       =   1545
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPurchases.frx":0000
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
   Begin VB.Frame fraGrid 
      Enabled         =   0   'False
      Height          =   2535
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3360
      Width           =   9255
      Begin VB.TextBox txtCheckno 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   360
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2040
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpDuedate 
         Height          =   375
         Left            =   2880
         TabIndex        =   21
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyy\M\d"
         Format          =   24707075
         CurrentDate     =   36390
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "√÷› ≈·Ï «·›« Ê—…"
         Height          =   375
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   1800
      End
      Begin VB.OptionButton optCash 
         Alignment       =   1  'Right Justify
         Caption         =   "‘Ìﬂ"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton optCash 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ«‘"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1680
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdClosebill 
         Caption         =   "≈€·«ﬁ «·›« Ê—…"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2040
         Width           =   1800
      End
      Begin VB.TextBox txtAlltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   345
         Left            =   480
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1680
         Width           =   960
      End
      Begin MSFlexGridLib.MSFlexGrid grdBill 
         Height          =   1455
         Left            =   120
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         RightToLeft     =   -1  'True
         FillStyle       =   1
         GridLines       =   3
         ScrollBars      =   2
         FormatString    =   "<ﬂÊœ|<’‰›|<«·ﬂ„Ì…|<«·»Ê‰’|<«·Œ’„|<«·—’Ìœ|< «—ÌŒ «·’·«ÕÌ…|<«·≈Ã„«·Ì"
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
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·‘Ìﬂ"
         Height          =   240
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   2040
         Width           =   705
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
   Begin VB.Frame fraSource 
      Height          =   735
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   9255
      Begin VB.ComboBox cmpSource 
         Height          =   360
         ItemData        =   "frmPurchases.frx":0442
         Left            =   6240
         List            =   "frmPurchases.frx":0444
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdAgreesource 
         Caption         =   "„Ê«›ﬁ"
         Height          =   360
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   360
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyy\MMM\d"
         Format          =   24576003
         CurrentDate     =   36389
      End
      Begin VB.TextBox txtBill 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«· «—ÌŒ"
         Height          =   240
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·›« Ê—…"
         Height          =   240
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„’œ—"
         Height          =   240
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraBill 
      Enabled         =   0   'False
      Height          =   2645
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   720
      Width           =   9255
      Begin VB.ComboBox txtItemname 
         Height          =   360
         ItemData        =   "frmPurchases.frx":0446
         Left            =   3360
         List            =   "frmPurchases.frx":0448
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   3375
      End
      Begin VB.ComboBox cmpItemcode 
         Height          =   360
         ItemData        =   "frmPurchases.frx":044A
         Left            =   6840
         List            =   "frmPurchases.frx":044C
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtMinlimit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
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
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdCompute 
         Caption         =   "«Õ”»"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   " ‰›Ì– ⁄„·Ì… «·‘—«¡"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   2160
         Width           =   3000
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
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
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
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtBonus 
         Alignment       =   1  'Right Justify
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
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtStock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
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
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
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
            Size            =   8.25
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtPceprc 
         Alignment       =   1  'Right Justify
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
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpExpiry 
         Height          =   360
         Left            =   3720
         TabIndex        =   11
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy\MMM\d"
         Format          =   24576003
         CurrentDate     =   36526
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·Õœ «·√œ‰Ï"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "”⁄— «·ÊÕœ…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1320
         Width           =   885
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·≈Ã„«·Ì"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·’·«ÕÌ…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·Œ’„"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·»Ê‰’"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂ„Ì…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "’‰›"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   270
      End
   End
End
Attribute VB_Name = "frmPurchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
   On Error Resume Next
 fraGrid.Enabled = False
    fraBill.Enabled = True
    cmpItemcode.SetFocus
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
            Set rsPurchases = db.OpenRecordset("purchases", dbOpenTable)
            
            rsChecks.AddNew
            rsChecks.Fields(1) = Val(txtCheckno.Text)
            rsChecks.Fields(2) = Val(txtAlltotal.Text)
            rsChecks.Fields(3) = dtpDuedate.Value
            rsChecks.Fields(4) = cmpSource.Text
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
                    rsFound_name.Fields(4) = XY + Val(txtAlltotal.Text)
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
        grdBill.FormatString = "<ﬂÊœ|<’‰›|<«·ﬂ„Ì…|<«·»Ê‰’|<«·Œ’„|<«·—’Ìœ|< «—ÌŒ «·’·«ÕÌ…|<«·≈Ã„«·Ì"
        grdBill.ColWidth(0) = 160 * 6
        grdBill.ColWidth(1) = 160 * 8
        grdBill.ColWidth(2) = 160 * 6
        grdBill.ColWidth(3) = 160 * 6
        grdBill.ColWidth(4) = 160 * 6
        grdBill.ColWidth(5) = 160 * 6
        grdBill.ColWidth(6) = 160 * 10
        grdBill.ColWidth(7) = 160 * 6
        txtBill.Text = ""
        txtItemname.Text = ""
        txtStock.Text = ""
        txtPceprc.Text = ""
        txtQuantity.Text = ""
        txtDiscount.Text = ""
        txtBonus.Text = ""
        txtTotal.Text = ""
        txtAlltotal.Text = ""
        txtCheckno.Text = ""
        dtpDuedate.Month = Month(Date) + 2
    
        cmdSave.Enabled = False
        cmdClosebill.Enabled = False
        cmdAdd.Enabled = False
    
        optCash(0).Enabled = False
        optCash(0).Value = True
        optCash(1).Enabled = False
    
        fraSource.Enabled = True
        fraBill.Enabled = False
        fraGrid.Enabled = False
        
    End If
End Sub

Private Sub cmdMain_Click()
On Error Resume Next
    frmPurchases.Visible = False
    frmMain.Visible = True
    Unload frmPurchases
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsBank = db.OpenRecordset("bank", dbOpenTable)
    Set rsChecks = db.OpenRecordset("checks", dbOpenTable)
    Set rsPurchases = db.OpenRecordset("purchases", dbOpenTable)
    Set rsStock = db.OpenRecordset("stock", dbOpenTable)
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    Set rsCheckno = db.OpenRecordset("Checkno", dbOpenTable)

    Dim Recc As Integer
    Dim Recn As Integer
    Dim nn1 As Long
    rsItem_data.MoveLast
    Recc = rsItem_data.RecordCount
    rsItem_data.MoveFirst
    For Recn = 1 To Recc
    nn1 = rsItem_data.Fields(6)
    If rsItem_data.Fields(1) = Val(cmpItemcode.Text) Then
        rsItem_data.Edit
        rsItem_data.Fields(6) = nn1 + Val(txtQuantity.Text) + Val(txtBonus.Text)
        rsItem_data.Update
    End If
    rsItem_data.Move 1
    Next Recn

    rsPurchases.MoveLast
    rsStock.MoveLast

    rsPurchases.AddNew
    rsStock.AddNew

    rsPurchases.Fields(1) = SourceKind
    rsStock.Fields(5) = SourceKind
    
    rsPurchases.Fields(2) = cmpSource.Text
    rsStock.Fields(4) = cmpSource.Text
    
    rsPurchases.Fields(3) = Val(cmpItemcode.Text)
    rsStock.Fields(1) = Val(cmpItemcode.Text)
    
    rsPurchases.Fields(4) = Val(txtQuantity.Text) + Val(txtBonus.Text)
    rsStock.Fields(2) = Val(txtQuantity.Text) + Val(txtBonus.Text)
    
    rsPurchases.Fields(5) = Val(txtPceprc.Text)
    rsStock.Fields(6) = Val(txtPceprc.Text)
    
    rsPurchases.Fields(6) = Val(txtDiscount.Text)
    rsPurchases.Fields(7) = dtpDate.Value
    rsPurchases.Fields(8) = dtpExpiry.Value
    rsStock.Fields(3) = dtpExpiry.Value
    
    rsPurchases.Fields(9) = Val(txtBill.Text)
    
    rsCheckno.MoveLast
    rsPurchases.Fields(11) = rsCheckno.Fields(0)
    rsStock.Fields(7) = rsCheckno.Fields(0)
    
    rsPurchases.Update
    rsStock.Update
    
    rsPurchases.MoveLast
    rsStock.MoveLast

    fraGrid.Enabled = True

Rem grid bill
    grdBill.AddItem (cmpItemcode & vbTab & txtItemname.Text & vbTab & txtQuantity.Text & vbTab & txtBonus.Text & vbTab & txtDiscount.Text & vbTab & txtStock.Text & vbTab & Format(dtpExpiry.Value, "long date") & vbTab & txtTotal.Text)
Rem
    txtAlltotal.Text = Val(txtAlltotal.Text) + Val(txtTotal.Text)
    txtItemname.Text = ""
    txtStock.Text = ""
    txtPceprc.Text = ""
    txtQuantity.Text = ""
    txtDiscount.Text = ""
    txtBonus.Text = ""
    txtTotal.Text = ""
    cmdSave.Enabled = False
    cmdClosebill.Enabled = True
    cmdAdd.Enabled = True
    optCash(0).Enabled = True
    optCash(1).Enabled = True
    fraBill.Enabled = False

End Sub
Private Sub cmdCompute_Click()
On Error Resume Next
    If cmpItemcode.Text = "" Then
        cmpItemcode.SetFocus
    ElseIf Not IsNumeric(txtPceprc.Text) Or Val(txtPceprc.Text) <= 0 Then
        MsgBox "√œŒ· ”⁄— «·ÊÕœ…", vbExclamation, "Œÿ√"
        txtPceprc.Text = ""
        txtPceprc.SetFocus
    ElseIf Not IsNumeric(txtQuantity.Text) Or Val(txtQuantity.Text) <= 0 Then
        MsgBox "√œŒ· «·ﬂ„Ì…", vbExclamation, "Œÿ√"
        txtQuantity.Text = ""
        txtQuantity.SetFocus
    ElseIf (txtDiscount.Text <> "" And Not IsNumeric(txtDiscount.Text)) Or Val(txtDiscount.Text) < 0 Or Val(txtDiscount.Text) >= 100 Then
        MsgBox "√œŒ· ‰”»… Œ’„ ’ÕÌÕ…", vbExclamation, "Œÿ√"
        txtDiscount.Text = ""
        txtDiscount.SetFocus
    Else
        cmdSave.Enabled = True
        txtTotal.Text = Val(txtPceprc.Text) * Val(txtQuantity.Text)(100 - Val(txtDiscount.Text)) / 100
    End If
End Sub

Private Sub cmpItemcode_LostFocus()
On Error Resume Next
    Call cmpItemcode_Scroll
End Sub
Private Sub cmdAgreeSource_Click()
On Error Resume Next
    If cmpSource.Text = "" Then
        MsgBox "√œŒ· «”„ «·„’œ—", vbExclamation, "Œÿ√"
        cmpSource.SetFocus
    ElseIf Not IsNumeric(txtBill.Text) Or Val(txtBill.Text) < 0 Then
        MsgBox "√œŒ· —ﬁ„ «·›« Ê—…", vbExclamation, "Œÿ√"
        txtBill.Text = ""
        txtBill.SetFocus
    Else
        fraBill.Enabled = True
        cmpItemcode.SetFocus
        fraSource.Enabled = False
        
    End If
End Sub
Private Sub cmpItemcode_Scroll()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    
    If cmpItemcode.ListCount > 0 Then
        Dim rec_no As Integer
        rec_no = cmpItemcode.ListIndex
    End If
        rsItem_data.MoveFirst
        rsItem_data.Move rec_no
        txtStock.Text = rsItem_data.Fields(6)
        txtMinlimit.Text = rsItem_data.Fields(3)
        If Val(txtStock.Text) <= Val(txtMinlimit.Text) Then txtStock.BackColor = RGB(255, 0, 0) Else txtStock.BackColor = &H80000004
End Sub



Private Sub Form_Load()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    Set rsCheckno = db.OpenRecordset("Checkno", dbOpenTable)

    rsCheckno.Edit
    rsCheckno.AddNew
    rsCheckno.Fields(1) = "any"
    rsCheckno.Update

    Dim l_rec_no1 As Integer
    rsItem_data.MoveLast
    l_rec_no1 = rsItem_data.RecordCount
    rsItem_data.MoveFirst
    
    Dim J1 As Integer
    For J1 = 1 To l_rec_no1
        cmpItemcode.AddItem rsItem_data.Fields(1)
        txtItemname.AddItem rsItem_data.Fields(2)
        rsItem_data.Move 1
    Next J1

    dtpDate.Value = Date
    dtpExpiry.Value = Date
    dtpDuedate.Value = Date
    dtpExpiry.Year = Year(Date) + 2
    dtpDuedate.Month = Month(Date) + 2
        
    grdBill.ColWidth(0) = 160 * 6
    grdBill.ColWidth(1) = 160 * 8
    grdBill.ColWidth(2) = 160 * 6
    grdBill.ColWidth(3) = 160 * 6
    grdBill.ColWidth(4) = 160 * 6
    grdBill.ColWidth(5) = 160 * 6
    grdBill.ColWidth(6) = 160 * 10
    grdBill.ColWidth(7) = 160 * 6

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
        txtCheckno.Text = ""
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
    
    If cmpItemcode.ListCount > 0 Then
        Dim rec_no As Integer
        rec_no = txtItemname.ListIndex
    End If
        rsItem_data.MoveFirst
        rsItem_data.Move rec_no
        txtStock.Text = rsItem_data.Fields(6)
        txtMinlimit.Text = rsItem_data.Fields(3)
        If Val(txtStock.Text) <= Val(txtMinlimit.Text) Then txtStock.BackColor = RGB(255, 0, 0) Else txtStock.BackColor = &H80000004
End Sub
