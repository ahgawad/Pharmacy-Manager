VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "ÕíÏáíÉ 2000"
   ClientHeight    =   5220
   ClientLeft      =   405
   ClientTop       =   1095
   ClientWidth     =   7875
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   240
      Top             =   4560
   End
   Begin VB.Menu mnuPurchases 
      Caption         =   "ÇáãÔÊÑíÇÊ"
      Begin VB.Menu mnuCompany 
         Caption         =   "ÔÑßÉ"
      End
      Begin VB.Menu mnuOffice 
         Caption         =   "ãßÊÈ"
      End
      Begin VB.Menu mnuPerson 
         Caption         =   "ÔÎÕ"
      End
      Begin VB.Menu mnuPharmacy 
         Caption         =   "ÕíÏáíÉ"
      End
      Begin VB.Menu mnuSuperMarket 
         Caption         =   "ÓæÈÑ ãÇÑßÊ"
      End
      Begin VB.Menu mnuFirm 
         Caption         =   "ãÕäÚ"
      End
   End
   Begin VB.Menu mnuBack 
      Caption         =   "ÇáãÑÊÌÚÇÊ"
      Begin VB.Menu mnuCompany2 
         Caption         =   "ÔÑßÉ"
      End
      Begin VB.Menu mnuOffice2 
         Caption         =   "ãßÊÈ"
      End
      Begin VB.Menu mnuPerson2 
         Caption         =   "ÔÎÕ"
      End
      Begin VB.Menu mnuPharmacy2 
         Caption         =   "ÕíÏáíÉ"
      End
      Begin VB.Menu mnuSuperMarket2 
         Caption         =   "ÓæÈÑ ãÇÑßÊ"
      End
      Begin VB.Menu mnuFirm2 
         Caption         =   "ãÕäÚ"
      End
   End
   Begin VB.Menu mnuSales 
      Caption         =   "ÇáãÈíÚÇÊ"
   End
   Begin VB.Menu mnuFoundations 
      Caption         =   "ÇáãäÔÂÊ"
      Begin VB.Menu mnuCompany1 
         Caption         =   "ÔÑßÉ"
      End
      Begin VB.Menu mnuOffice1 
         Caption         =   "ãßÊÈ"
      End
      Begin VB.Menu mnuPerson1 
         Caption         =   "ÔÎÕ"
      End
      Begin VB.Menu mnuPharmacy1 
         Caption         =   "ÕíÏáíÉ"
      End
      Begin VB.Menu mnuSuperMarket1 
         Caption         =   "ÓæÈÑ ãÇÑßÊ"
      End
      Begin VB.Menu mnuFirm1 
         Caption         =   "ãÕäÚ"
      End
   End
   Begin VB.Menu mnuDrugs 
      Caption         =   "ÇáÃÕäÇÝ"
   End
   Begin VB.Menu mnuMinlmt 
      Caption         =   "ÇáÍÏ ÇáÃÏäì"
   End
   Begin VB.Menu mnuDates 
      Caption         =   "ÇáÊæÇÑíÎ"
   End
   Begin VB.Menu mnuBank 
      Caption         =   "ÇáÍÓÇÈÇÊ"
   End
   Begin VB.Menu mnutools 
      Caption         =   "ÃÏæÇÊ"
      Begin VB.Menu mnuBackup 
         Caption         =   "Úãá äÓÎÉ ÇÍÊíÇØíÉ"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "ÎÑæÌ"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsDays = db.OpenRecordset("days", dbOpenTable)
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

Rem *********************************
'Timer1.Interval = 60000
'Dim DD As Integer
'rsDays.MoveFirst
'If rsDays(1) >= 31 Then
'    MsgBox "ÞÏ ÇäÊåÊ ÇáÝÊÑÉ ÇáããäæÍÉ áÊÌÑíÈ ÇáÈÑäÇãÌ¡ ãä ÝÖáß ÇÊÕá ÈÇáãÈÑãÌ", vbCritical
'    End
'Else
'    DD = rsDays.Fields(1)
'    rsDays.Edit
'    rsDays.Fields(1) = DD + 1
'    rsDays.Update
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill (App.Path & "\Data\backup.mdb")
ChDir (App.Path & "\Data\")
Shell App.Path & "\Data\backup.bat", vbHide
ChDir (App.Path)
End Sub

Private Sub mnuBackup_Click()
On Error Resume Next
Kill (App.Path & "\Data\backup.mdb")
ChDir (App.Path & "\Data\")
Shell App.Path & "\Data\backup.bat", vbHide
ChDir (App.Path)
End Sub

Private Sub mnuBank_Click()
On Error Resume Next
    frmMain.Visible = False
    frmBank.Show

End Sub

Private Sub mnuCompany_Click()
On Error Resume Next
    SourceKind = "ÔÑßÉ"

    frmMain.Visible = False
    frmPurchases.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsCompany_name = db.OpenRecordset("company name", dbOpenTable)

    rsCompany_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsCompany_name.RecordCount
    rsCompany_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsCompany_name.Fields(1)
        frmPurchases.cmpSource.AddItem Name
        rsCompany_name.Move 1
    Next j

End Sub
Private Sub mnuCompany1_Click()
On Error Resume Next
    SourceKind = "ÔÑßÉ"

    frmMain.Visible = False
    frmFoundations.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsCompany_name = db.OpenRecordset("company name", dbOpenTable)

    rsCompany_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsCompany_name.RecordCount
    rsCompany_name.MoveFirst
    
    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsCompany_name.Fields(1)
        frmFoundations.cmpSource0.AddItem Name
        frmFoundations.cmpSource2.AddItem Name
        frmFoundations.txtSourcename.AddItem Name
        frmFoundations.List1.AddItem Name
        frmFoundations.List2.AddItem j
        rsCompany_name.Move 1
    Next j
    frmFoundations.cmpSource0.Selected(0) = True
    frmFoundations.cmpSource2.Selected(0) = True
    frmFoundations.List1.Selected(0) = True
    frmFoundations.List2.Selected(0) = True
End Sub
Private Sub mnuCompany2_Click()
On Error Resume Next
    SourceKind = "ÔÑßÉ"

    frmMain.Visible = False
    frmBack.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsCompany_name = db.OpenRecordset("company name", dbOpenTable)

    rsCompany_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsCompany_name.RecordCount
    rsCompany_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsCompany_name.Fields(1)
        frmBack.cmpSource.AddItem Name
        rsCompany_name.Move 1
    Next j
    frmBack.cmpSource.Selected(0) = True
End Sub
Private Sub mnuDates_Click()
On Error Resume Next
    frmMain.Visible = False
    frmDates.Show
End Sub
Private Sub mnuDrugs_Click()
On Error Resume Next
    frmMain.Visible = False
    frmDrugs.Show
End Sub
Private Sub mnuExit_Click()
On Error Resume Next
    End
End Sub
Private Sub mnuFirm_Click()
On Error Resume Next
    SourceKind = "ãÕäÚ"

    frmMain.Visible = False
    frmPurchases.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsFirm_name = db.OpenRecordset("Firm name", dbOpenTable)

    rsFirm_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsFirm_name.RecordCount
    rsFirm_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsFirm_name.Fields(1)
        frmPurchases.cmpSource.AddItem Name
        rsFirm_name.Move 1
    Next j
End Sub
Private Sub mnuFirm1_Click()
On Error Resume Next
    SourceKind = "ãÕäÚ"

    frmMain.Visible = False
    frmFoundations.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsFirm_name = db.OpenRecordset("Firm name", dbOpenTable)

    rsFirm_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsFirm_name.RecordCount
    rsFirm_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsFirm_name.Fields(1)
        frmFoundations.cmpSource0.AddItem Name
        frmFoundations.cmpSource2.AddItem Name
        frmFoundations.txtSourcename.AddItem Name
        frmFoundations.List1.AddItem Name
        frmFoundations.List2.AddItem j
        rsFirm_name.Move 1
    Next j
    frmFoundations.cmpSource0.Selected(0) = True
    frmFoundations.cmpSource2.Selected(0) = True
    frmFoundations.List1.Selected(0) = True
    frmFoundations.List2.Selected(0) = True
End Sub
Private Sub mnuFirm2_Click()
On Error Resume Next
    SourceKind = "ãÕäÚ"

    frmMain.Visible = False
    frmBack.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsFirm_name = db.OpenRecordset("Firm name", dbOpenTable)

    rsFirm_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsFirm_name.RecordCount
    rsFirm_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsFirm_name.Fields(1)
        frmBack.cmpSource.AddItem Name
        rsFirm_name.Move 1
    Next j
    frmBack.cmpSource.Selected(0) = True
End Sub
Private Sub mnuMinlmt_Click()
On Error Resume Next
    frmMain.Visible = False
    frmMinlmt.Show
    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsItem_data = db.OpenRecordset("item data", dbOpenTable)
Dim X, Y As Integer
rsItem_data.MoveLast
X = rsItem_data.RecordCount
rsItem_data.MoveFirst
For Y = 1 To X
rsItem_data.Index = Trim(rsItem_data.Fields(6).Name)
Dim ZXZX As String
ZXZX = Trim(rsItem_data.Fields(3))
rsItem_data.Seek ("=" + ZXZX)
frmMinlmt.grdBill.AddItem (rsItem_data.Fields(1) & vbTab & rsItem_data.Fields(2) & vbTab & rsItem_data.Fields(3) & vbTab & rsItem_data.Fields(6))
    End If
    rsItem_data.Move 1
Next Y
End Sub
Private Sub mnuOffice_Click()
On Error Resume Next
    SourceKind = "ãßÊÈ"

    frmMain.Visible = False
    frmPurchases.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsOffice_name = db.OpenRecordset("office name", dbOpenTable)

    rsOffice_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsOffice_name.RecordCount
    rsOffice_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsOffice_name.Fields(1)
        frmPurchases.cmpSource.AddItem Name
        rsOffice_name.Move 1
    Next j
End Sub
Private Sub mnuOffice1_Click()
On Error Resume Next
    SourceKind = "ãßÊÈ"

    frmMain.Visible = False
    frmFoundations.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsOffice_name = db.OpenRecordset("office name", dbOpenTable)

    rsOffice_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsOffice_name.RecordCount
    rsOffice_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsOffice_name.Fields(1)
        frmFoundations.cmpSource0.AddItem Name
        frmFoundations.cmpSource2.AddItem Name
        frmFoundations.txtSourcename.AddItem Name
        frmFoundations.List1.AddItem Name
        frmFoundations.List2.AddItem j
        rsOffice_name.Move 1
    Next j
    frmFoundations.cmpSource0.Selected(0) = True
    frmFoundations.cmpSource2.Selected(0) = True
    frmFoundations.List1.Selected(0) = True
    frmFoundations.List2.Selected(0) = True
End Sub
Private Sub mnuOffice2_Click()
On Error Resume Next
    SourceKind = "ãßÊÈ"

    frmMain.Visible = False
    frmBack.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsOffice_name = db.OpenRecordset("office name", dbOpenTable)

    rsOffice_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsOffice_name.RecordCount
    rsOffice_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsOffice_name.Fields(1)
        frmBack.cmpSource.AddItem Name
        rsOffice_name.Move 1
    Next j
    frmBack.cmpSource.Selected(0) = True
End Sub
Private Sub mnuPerson_Click()
On Error Resume Next
    SourceKind = "ÔÎÕ"

    frmMain.Visible = False
    frmPurchases.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsPerson_name = db.OpenRecordset("Person name", dbOpenTable)

    rsPerson_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsPerson_name.RecordCount
    rsPerson_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsPerson_name.Fields(1)
        frmPurchases.cmpSource.AddItem Name
        rsPerson_name.Move 1
    Next j
End Sub
Private Sub mnuPerson1_Click()
On Error Resume Next
    SourceKind = "ÔÎÕ"

    frmMain.Visible = False
    frmFoundations.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsPerson_name = db.OpenRecordset("Person name", dbOpenTable)

    rsPerson_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsPerson_name.RecordCount
    rsPerson_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsPerson_name.Fields(1)
        frmFoundations.cmpSource0.AddItem Name
        frmFoundations.cmpSource2.AddItem Name
        frmFoundations.txtSourcename.AddItem Name
        frmFoundations.List1.AddItem Name
        frmFoundations.List2.AddItem j
        rsPerson_name.Move 1
    Next j
    frmFoundations.cmpSource0.Selected(0) = True
    frmFoundations.cmpSource2.Selected(0) = True
    frmFoundations.List1.Selected(0) = True
    frmFoundations.List2.Selected(0) = True
End Sub
Private Sub mnuPerson2_Click()
On Error Resume Next
    SourceKind = "ÔÎÕ"

    frmMain.Visible = False
    frmBack.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsPerson_name = db.OpenRecordset("Person name", dbOpenTable)

    rsPerson_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsPerson_name.RecordCount
    rsPerson_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsPerson_name.Fields(1)
        frmBack.cmpSource.AddItem Name
        rsPerson_name.Move 1
    Next j
    frmBack.cmpSource.Selected(0) = True
End Sub
Private Sub mnuPharmacy_Click()
On Error Resume Next
    SourceKind = "ÕíÏáíÉ"

    frmMain.Visible = False
    frmPurchases.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsPharmacy_name = db.OpenRecordset("Pharmacy name", dbOpenTable)

    rsPharmacy_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsPharmacy_name.RecordCount
    rsPharmacy_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsPharmacy_name.Fields(1)
        frmPurchases.cmpSource.AddItem Name
        rsPharmacy_name.Move 1
    Next j
End Sub
Private Sub mnuPharmacy1_Click()
On Error Resume Next
    SourceKind = "ÕíÏáíÉ"

    frmMain.Visible = False
    frmFoundations.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsPharmacy_name = db.OpenRecordset("Pharmacy name", dbOpenTable)

    rsPharmacy_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsPharmacy_name.RecordCount
    rsPharmacy_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsPharmacy_name.Fields(1)
        frmFoundations.cmpSource0.AddItem Name
        frmFoundations.cmpSource2.AddItem Name
        frmFoundations.txtSourcename.AddItem Name
        frmFoundations.List1.AddItem Name
        frmFoundations.List2.AddItem j
        rsPharmacy_name.Move 1
    Next j
    frmFoundations.cmpSource0.Selected(0) = True
    frmFoundations.cmpSource2.Selected(0) = True
    frmFoundations.List1.Selected(0) = True
    frmFoundations.List2.Selected(0) = True
End Sub
Private Sub mnuPharmacy2_Click()
On Error Resume Next
    SourceKind = "ÕíÏáíÉ"

    frmMain.Visible = False
    frmBack.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsPharmacy_name = db.OpenRecordset("Pharmacy name", dbOpenTable)

    rsPharmacy_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsPharmacy_name.RecordCount
    rsPharmacy_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsPharmacy_name.Fields(1)
        frmBack.cmpSource.AddItem Name
        rsPharmacy_name.Move 1
    Next j
    frmBack.cmpSource.Selected(0) = True
End Sub

Private Sub mnurestore_Click()
On Error Resume Next
Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
db.Close
Kill (App.Path & "\Data\data.mdb")
ChDir (App.Path & "\Data\")
Shell App.Path & "\Data\Restore.bat", vbHide
ChDir (App.Path)
End Sub

Private Sub mnuSales_Click()
On Error Resume Next
    frmMain.Visible = False
    frmSales.Show
End Sub
Private Sub mnuSuperMarket_Click()
On Error Resume Next
    SourceKind = "ÓæÈÑ ãÇÑßÊ"

    frmMain.Visible = False
    frmPurchases.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsSuperMarket_name = db.OpenRecordset("Super Market name", dbOpenTable)

    rsSuperMarket_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsSuperMarket_name.RecordCount
    rsSuperMarket_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsSuperMarket_name.Fields(1)
        frmPurchases.cmpSource.AddItem Name
        rsSuperMarket_name.Move 1
    Next j
End Sub
Private Sub mnuSuperMarket1_Click()
On Error Resume Next
    SourceKind = "ÓæÈÑ ãÇÑßÊ"

    frmMain.Visible = False
    frmFoundations.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsSuperMarket_name = db.OpenRecordset("Super Market name", dbOpenTable)

    rsSuperMarket_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsSuperMarket_name.RecordCount
    rsSuperMarket_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsSuperMarket_name.Fields(1)
        frmFoundations.cmpSource0.AddItem Name
        frmFoundations.cmpSource2.AddItem Name
        frmFoundations.txtSourcename.AddItem Name
        frmFoundations.List1.AddItem Name
        frmFoundations.List2.AddItem j
        rsSuperMarket_name.Move 1
    Next j
    frmFoundations.cmpSource0.Selected(0) = True
    frmFoundations.cmpSource2.Selected(0) = True
    frmFoundations.List1.Selected(0) = True
    frmFoundations.List2.Selected(0) = True
End Sub
Private Sub mnuSuperMarket2_Click()

On Error Resume Next
    SourceKind = "ÓæÈÑ ãÇÑßÊ"

    frmMain.Visible = False
    frmBack.Show

    Set db = DBEngine.OpenDatabase(App.Path + "\DATA\DATA.MDB")
    Set rsSuperMarket_name = db.OpenRecordset("Super Market name", dbOpenTable)

    rsSuperMarket_name.MoveLast
    Dim l_rec_no As Integer
    l_rec_no = rsSuperMarket_name.RecordCount
    rsSuperMarket_name.MoveFirst

    Dim j As Integer
    Dim Name As String
    For j = 1 To l_rec_no
        Name = rsSuperMarket_name.Fields(1)
        frmBack.cmpSource.AddItem Name
        rsSuperMarket_name.Move 1
    Next j
    frmBack.cmpSource.Selected(0) = True
End Sub
Private Sub Timer1_Timer()
'Dim DD As Integer
'rsDays.MoveFirst
'If rsDays(1) >= 31 Then
'    MsgBox "ÞÏ ÇäÊåÊ ÇáÝÊÑÉ ÇáããäæÍÉ áÊÌÑíÈ ÇáÈÑäÇãÌ¡ ãä ÝÖáß ÇÊÕá ÈÇáãÈÑãÌ", vbCritical
'    End
'Else
'    If ASKTimer >= 240 Then
'        DD = rsDays.Fields(1)
'        rsDays.Edit
'        rsDays.Fields(1) = DD + 1
'        rsDays.Update
'    Else:
'        ASKTimer = ASKTimer + 1
'    End If
'End If
End Sub
