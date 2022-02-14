VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMinlmt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·Õœ «·√œ‰Ï"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "frmMinlmt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      Begin MSFlexGridLib.MSFlexGrid grdBill 
         Height          =   5535
         Left            =   1440
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   9763
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
         FormatString    =   "<ﬂÊœ|<’‰›|<«·Õœ «·√œ‰Ï|<«·ﬂ„Ì…"
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
   Begin VB.CommandButton cmdMain 
      Caption         =   "«·ﬁ«∆„… «·—∆Ì”Ì…"
      Height          =   375
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   6120
      Width           =   3000
   End
End
Attribute VB_Name = "frmMinlmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMain_Click()
On Error Resume Next
    frmMinlmt.Visible = False
    frmMain.Visible = True
    Unload frmMinlmt

End Sub

Private Sub Form_Load()
On Error Resume Next
    grdBill.ColWidth(0) = 160 * 7
    grdBill.ColWidth(1) = 160 * 15
    grdBill.ColWidth(2) = 160 * 7
    grdBill.ColWidth(3) = 160 * 7

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    frmMain.Visible = True
End Sub
