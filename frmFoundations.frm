VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFoundations 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·„‰‘¬ "
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "frmFoundations.frx":0000
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   26
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
      TabPicture(0)   =   "frmFoundations.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSource0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "≈÷«›…"
      TabPicture(1)   =   "frmFoundations.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " Õ–›"
      TabPicture(2)   =   "frmFoundations.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraSure"
      Tab(2).Control(1)=   "fraAdd"
      Tab(2).Control(2)=   "fraSource2"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   " ⁄œÌ·"
      TabPicture(3)   =   "frmFoundations.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   9015
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
            ItemData        =   "frmFoundations.frx":04B2
            Left            =   360
            List            =   "frmFoundations.frx":04B4
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   240
            Visible         =   0   'False
            Width           =   765
         End
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
            ItemData        =   "frmFoundations.frx":04B6
            Left            =   4080
            List            =   "frmFoundations.frx":04B8
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   3045
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«”„ «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1935
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1320
         Width           =   9015
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   480
            Width           =   6735
         End
         Begin VB.CommandButton Command2 
            Caption         =   "„Ê«›ﬁ"
            Height          =   360
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "⁄‰Ê«‰ «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Â« › «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   1200
            Width           =   930
         End
      End
      Begin VB.Frame fraSure 
         Enabled         =   0   'False
         Height          =   2295
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   3360
         Width           =   9015
         Begin VB.CommandButton cmdNo 
            Caption         =   "·«"
            Height          =   360
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   1200
            Width           =   3135
         End
         Begin VB.CommandButton cmdAgreesource2 
            Caption         =   "‰⁄„"
            Height          =   360
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   600
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
            TabIndex        =   38
            Top             =   840
            Width           =   1800
         End
      End
      Begin VB.Frame fraAdd 
         Height          =   1935
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   9015
         Begin VB.CommandButton cmdOK 
            Caption         =   "„Ê«›ﬁ"
            Height          =   360
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label txtTeldisplay 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label txtAdddisplay 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   480
            Width           =   6735
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Â« › «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   1200
            Width           =   930
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "⁄‰Ê«‰ «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Enabled         =   0   'False
         Height          =   1935
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   3720
         Width           =   9015
         Begin VB.TextBox txtAddress 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Height          =   360
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   480
            Width           =   6765
         End
         Begin VB.TextBox txtTel 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Height          =   360
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   1200
            Width           =   1845
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "⁄‰Ê«‰ «·„‰‘√…"
            Height          =   195
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Â« › «·„‰‘√…"
            Height          =   195
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   1200
            Width           =   930
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1935
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1320
         Width           =   9015
         Begin VB.CommandButton Command1 
            Caption         =   "„Ê«›ﬁ"
            Height          =   360
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox txtTelsave 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   1200
            Width           =   1635
         End
         Begin VB.TextBox txtAddresssave 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   480
            Width           =   6645
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
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "⁄‰Ê«‰ «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   6
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
            ItemData        =   "frmFoundations.frx":04BA
            Left            =   4080
            List            =   "frmFoundations.frx":04BC
            RightToLeft     =   -1  'True
            TabIndex        =   7
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
            TabIndex        =   29
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2295
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   9015
         Begin MSFlexGridLib.MSFlexGrid grdStk 
            Height          =   1935
            Left            =   480
            TabIndex        =   25
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
      Begin VB.Frame fraSource0 
         Height          =   735
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   9015
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
            ItemData        =   "frmFoundations.frx":04BE
            Left            =   4080
            List            =   "frmFoundations.frx":04C0
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   240
            Width           =   3045
         End
         Begin VB.CommandButton cmdAgreesource0 
            Caption         =   "„Ê«›ﬁ"
            Height          =   360
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«”„ «·„‰‘√…"
            Height          =   195
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Frame fraSource2 
         Height          =   735
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   12
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
            ItemData        =   "frmFoundations.frx":04C2
            Left            =   4080
            List            =   "frmFoundations.frx":04C4
            RightToLeft     =   -1  'True
            TabIndex        =   13
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
            TabIndex        =   27
            Top             =   240
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frmFoundations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAgreesource0_Click()
On Error Resume Next
    Dim Recc As Integer
    Dim Recn As Integer
    Dim Recc1 As Integer
    Dim Recn1 As Integer
    Dim Recc2 As Integer
    Dim Recn2 As Integer
    Dim Name As String
    Dim Min As Integer

grdStk.Clear
grdStk.Rows = 1
grdStk.FormatString = "<ﬂÊœ|<’‰›|<«·—’Ìœ|<«·Õœ «·√œ‰Ï|< «—ÌŒ «·’·«ÕÌ…"
grdStk.ColWidth(0) = 160 * 6
grdStk.ColWidth(1) = 160 * 20
grdStk.ColWidth(2) = 160 * 6
grdStk.ColWidth(3) = 160 * 6
grdStk.ColWidth(4) = 160 * 10
    
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsPurchases = db.OpenRecordset("purchases", dbOpenTable)
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    
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

Rem
    rsFound_name.MoveLast
    Recc = rsFound_name.RecordCount
    rsFound_name.MoveFirst
For Recn = 1 To Recc
    If rsFound_name.Fields(1) = cmpSource0.Text Then
        txtAddress.Text = rsFound_name.Fields(2)
        txtTel.Text = rsFound_name.Fields(3)
    End If
    rsFound_name.Move 1
Next Recn

    rsPurchases.MoveLast
    Recc1 = rsPurchases.RecordCount
    rsPurchases.MoveFirst

    rsItem_data.MoveLast
    Recc2 = rsItem_data.RecordCount
    rsItem_data.MoveFirst

Rem


For Recn1 = 1 To Recc1
    If rsPurchases.Fields(1) = SourceKind And rsPurchases.Fields(2) = cmpSource0 Then
        rsItem_data.MoveFirst
        For Recn2 = 2 To Recc2
            If rsItem_data.Fields(1) = rsPurchases.Fields(3) Then
                Name = rsItem_data.Fields(2)
                Min = rsItem_data.Fields(3)
            End If
            rsItem_data.Move 1
        Next Recn2
Rem grid
    grdStk.AddItem (rsPurchases.Fields(3) & vbTab & Name & vbTab & rsPurchases.Fields(4) & vbTab & Min & vbTab & Format(rsPurchases.Fields(8), "d\\MMM\\yyyy"))
Rem

    End If
    rsPurchases.Move 1
Next Recn1
cmpSource0.SetFocus
End Sub
Private Sub cmdAgreesource2_Click()
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
Rem
    Dim Rc As Integer
    Dim Rn As Integer
    Dim Inv As Integer
    rsFound_name.MoveLast
    Rc = rsFound_name.RecordCount
    rsFound_name.MoveFirst
For Rn = 1 To Rc
    If rsFound_name.Fields(1) = cmpSource2.Text Then
        rsFound_name.Edit
        rsFound_name.Delete
        fraSure.Enabled = False
        Inv = cmpSource2.ListIndex
        cmpSource2.RemoveItem (Inv)
        cmpSource0.RemoveItem (Inv)
        txtSourcename.RemoveItem (Inv)
        cmpSource0.Selected(0) = True
        cmpSource2.Selected(0) = True
        Call cmpSource2_Scroll
        Exit For
    End If
    rsFound_name.Move 1
Next Rn

End Sub
Private Sub cmdMain_Click()
On Error Resume Next
    frmFoundations.Visible = False
    frmMain.Visible = True
    Unload frmFoundations
End Sub
Private Sub cmdNo_Click()
On Error Resume Next
    fraSure.Enabled = False
    cmpSource2.SetFocus
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
    fraSure.Enabled = True
    cmdNo.SetFocus
End Sub
Private Sub cmpSource2_GotFocus()
On Error Resume Next
    Call cmpSource2_Scroll
End Sub

Private Sub cmpSource2_LostFocus()
On Error Resume Next
    Call cmpSource2_Scroll
End Sub

Private Sub cmpSource2_Scroll()
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

Rem
    Dim Rc As Integer
    Dim Rn As Integer
    rsFound_name.MoveLast
    Rc = rsFound_name.RecordCount
    rsFound_name.MoveFirst
For Rn = 1 To Rc
    If rsFound_name.Fields(1) = cmpSource2.Text Then
        txtAdddisplay.Caption = rsFound_name.Fields(2)
        txtTeldisplay.Caption = rsFound_name.Fields(3)
        Exit For
    End If
    rsFound_name.Move 1
Next Rn

End Sub

Private Sub Command1_Click()
On Error Resume Next
    If txtSourcename.Text = "" Then
        MsgBox "√œŒ· «”„ «·„‰‘√…", vbExclamation, "Œÿ√"
        txtSourcename.SetFocus
    ElseIf txtAddresssave = "" Then
        MsgBox "√œŒ· «·⁄‰Ê«‰", vbExclamation, "Œÿ√"
        txtAddresssave.Text = ""
        txtAddresssave.SetFocus
    ElseIf txtTelsave = "" Or Not IsNumeric(txtTelsave.Text) Then
        MsgBox "√œŒ· «·Â« ›", vbExclamation, "Œÿ√"
        txtTelsave.Text = ""
        txtTelsave.SetFocus
    Else
    
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsPurchases = db.OpenRecordset("purchases", dbOpenTable)
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
    
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

rsFound_name.MoveLast
rsFound_name.Edit
rsFound_name.AddNew
    rsFound_name.Fields(1) = txtSourcename.Text
    rsFound_name.Fields(2) = txtAddresssave.Text
    rsFound_name.Fields(3) = txtTelsave.Text
    rsFound_name.Fields(4) = 0
rsFound_name.Update
cmpSource0.AddItem txtSourcename.Text
cmpSource2.AddItem txtSourcename.Text
txtSourcename.AddItem txtSourcename.Text
List1.AddItem txtSourcename.Text
List2.AddItem rsFound_name.RecordCount

txtSourcename.Text = ""
txtAddresssave.Text = ""
txtTelsave.Text = ""

End If
End Sub
Private Sub Command2_Click()
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

Rem
    Dim Rc As Integer
    Dim Rn As Integer
    rsFound_name.MoveLast
    Rc = rsFound_name.RecordCount
    rsFound_name.MoveFirst
For Rn = 1 To Rc
    If rsFound_name.Fields(1) = List1.Text Then
        rsFound_name.Edit
        rsFound_name.Fields(2) = Text1.Text
        rsFound_name.Fields(3) = Text2.Text
        rsFound_name.Update
        Exit For
    End If
    rsFound_name.Move 1
Next Rn
List1.SetFocus

End Sub

Private Sub Form_Load()
On Error Resume Next
    grdStk.FormatString = "<ﬂÊœ|<’‰›|<«·—’Ìœ|<«·Õœ «·√œ‰Ï|< «—ÌŒ «·’·«ÕÌ…"
    grdStk.ColWidth(0) = 160 * 6
    grdStk.ColWidth(1) = 160 * 20
    grdStk.ColWidth(2) = 160 * 6
    grdStk.ColWidth(3) = 160 * 6
    grdStk.ColWidth(4) = 160 * 10

'    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
'    Set rsCompany_name = db.OpenRecordset("company name", dbOpenTable)
'    Set rsFirm_name = db.OpenRecordset("firm name", dbOpenTable)
'    Set rsOffice_name = db.OpenRecordset("office name", dbOpenTable)
'    Set rsPerson_name = db.OpenRecordset("person name", dbOpenTable)
'    Set rsPharmacy_name = db.OpenRecordset("pharmacy name", dbOpenTable)
'    Set rsSuper_market_name = db.OpenRecordset("super market name", dbOpenTable)
    
'DataSource.DatabaseName = db.name

'Select Case SourceKind
'Case "‘—ﬂ…"
'    DataSource.RecordSource = rsCompany_name.name
'Case "„ﬂ »"
'    DataSource.RecordSource = rsOffice_name.name
'Case "‘Œ’"
'    DataSource.RecordSource = rsPerson_name.name
'Case "’Ìœ·Ì…"
'    DataSource.RecordSource = rsPharmacy_name.name
'Case "”Ê»— „«—ﬂ "
'    DataSource.RecordSource = rsSuper_market_name.name
'Case "„’‰⁄"
'    DataSource.RecordSource = rsFirm_name.name
'End Select

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    frmMain.Visible = True
End Sub
Private Sub List1_GotFocus()
On Error Resume Next
    Call List1_Scroll
End Sub
Private Sub List1_LostFocus()
On Error Resume Next
    Call List1_Scroll
End Sub
Private Sub List1_Scroll()
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

Rem
    Dim Rc As Integer
    Dim Rn As Integer
    rsFound_name.MoveLast
    Rc = rsFound_name.RecordCount
    
    rsFound_name.MoveFirst
For Rn = 1 To Rc
    If rsFound_name.Fields(1) = List1.Text Then
        Text1.Text = rsFound_name.Fields(2)
        Text2.Text = rsFound_name.Fields(3)
        Exit For
    End If
    rsFound_name.Move 1
Next Rn
        Dim rec_no1 As Integer
        rec_no1 = List1.ListIndex
        List2.Selected(rec_no1) = True

End Sub
