VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProcessBilling 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PROCESS BILLING STATEMENT"
   ClientHeight    =   2985
   ClientLeft      =   3030
   ClientTop       =   3135
   ClientWidth     =   5925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdButton 
      BackColor       =   &H00808000&
      Caption         =   "..."
      Height          =   375
      Index           =   3
      Left            =   5280
      TabIndex        =   15
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox FileName1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox Filename 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CommandButton CmdButton 
      BackColor       =   &H00808000&
      Caption         =   "..."
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   5160
      Top             =   360
   End
   Begin ComctlLib.ProgressBar Progress 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton CmdButton 
      Cancel          =   -1  'True
      Caption         =   "Cl&ose"
      Height          =   372
      Index           =   1
      Left            =   3000
      TabIndex        =   6
      Top             =   2520
      Width           =   972
   End
   Begin VB.CommandButton CmdButton 
      Caption         =   "&Process"
      Height          =   372
      Index           =   0
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   972
   End
   Begin MSComDlg.CommonDialog Cmd 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   1092
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox DateField 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   4200
         TabIndex        =   7
         Top             =   600
         Width           =   1155
      End
      Begin VB.TextBox CompanyField 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   120
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox CompanyField 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   1056
         MaxLength       =   50
         TabIndex        =   3
         Top             =   240
         Width           =   2916
      End
      Begin VB.TextBox CompanyField 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox CompanyField 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   1056
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   2916
      End
      Begin VB.Label LabelText 
         Caption         =   "Active Period"
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label LabelText 
      Caption         =   "EV Charging"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label LabelText 
      Caption         =   "EV Charging"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1260
      Width           =   1215
   End
End
Attribute VB_Name = "frmProcessBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Refresh_On_Activate As Boolean
Dim TermDue As Integer, PhoneDir As String, Penalty As Double, MinPEC As Double, xBillDate As Integer
Dim HardCopyCharges As Double, SheetName As String, SheetName1 As String
Dim Period_AR As Date, Period_CAR As Date
Private Const PPN As Double = (11 / 100)

Private Sub cmdButton_Click(Index As Integer)
    Dim errMsg As String, time_start As Date, menit As Integer, detik As Integer, how_long As Long
    If Index = 0 Then
        If MsgBox("Are you sure want to continue this process ?", vbYesNo) = vbYes Then
            If Not IsValidProcess(errMsg) Then
                MsgBox errMsg
                Exit Sub
            End If
            Me.MousePointer = vbHourglass
            Progress.Visible = True
            time_start = Now
            If Not ProsesBilling Then
                MsgBox "Process Failed!", vbCritical
                Progress.Visible = False
                Exit Sub
            End If
            how_long = DateDiff("s", time_start, Now)
            menit = how_long \ 60
            detik = how_long Mod 60
            Progress.Visible = False
            Me.MousePointer = vbNormal
            MsgBox "Process success in " & Format(menit, "#,#0") & " minutes " & Format(detik, "#,#0") & " seconds !"
        End If
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        Cmd.DialogTitle = "Import Excel File"
        Cmd.DefaultExt = "xls"
        Cmd.Filter = "Microsoft Excel File (*.xls)|*.xls"
        Cmd.Action = 2
        Filename = Cmd.Filename
    Else
        Cmd.DialogTitle = "Import Excel File"
        Cmd.DefaultExt = "xls"
        Cmd.Filter = "Microsoft Excel File (*.xls)|*.xls"
        Cmd.Action = 2
        FileName1 = Cmd.Filename
    End If
End Sub

Private Sub Form_Activate()
    If Refresh_On_Activate Then
        CompanyField(0) = companyId
        CompanyField(1) = companyName
        CompanyField(2) = LocationId
        CompanyField(3) = LocationName
        Progress.Visible = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Load()
    Dim rs As RemoteSet
    Dim mtgl As Date
    Refresh_On_Activate = True
    Set rs = db.OpenRemoteSet("SELECT * FROM M_Active_Period WHERE AUTONUM = 'KAR" & companyId & LocationId & "' OR AUTONUM = 'KCAR" & companyId & LocationId & "' ORDER BY AUTONUM")
    If TypeName(rs) = "Nothing" Then
        MsgBox "RS Nothing M_Active_Period", vbCritical
        CmdButton(0).Enabled = False
        Exit Sub
    Else
        rs.MoveFirst
        Period_AR = todate(rs("ActivePeriod"))
        rs.MoveLast
        Period_CAR = todate(rs("ActivePeriod"))
    End If
    Set rs = db.OpenRemoteSet("select * from kta_m_penalty where companyid='" & companyId & "' and locationid='" & LocationId & "'")
    If rs Is Nothing Then
        MsgBox "KTA_M_Penalty Nothing !"
        CmdButton(0).Enabled = False
        Exit Sub
    ElseIf rs.EOF Then
        MsgBox "KTA_M_Penalty EOF !"
        CmdButton(0).Enabled = False
        Exit Sub
    End If
    TermDue = rs("Term")
    PhoneDir = rs("PhoneDir")
    Penalty = rs("Penalty")
    MinPEC = rs("MinPEC")
    HardCopyCharges = rs("HardCopyCharges")
    DateField(0) = Format(Period_CAR, "dd-mmm-yyyy")
    xBillDate = rs("BillDate")
    rs.CloseRemoteSet
End Sub

Private Function ProsesBilling() As Boolean
    Dim rs As RemoteSet, rsTemp As RemoteSet, rsTmp1 As RemoteSet
    Dim rs_H_Billing As RemoteSet
    Dim rs_H_Outstanding As RemoteSet, rs_D_Outstanding As RemoteSet
    Dim rs_H_SJ As RemoteSet, rs_D_SJ As RemoteSet, rsRate As RemoteSet, mRate As Double
    Dim mKeyHD As Long, mPeriode As Date, mPrev As Double, mPayment As Double, mTotal As Double
    Dim mSeqNo As Integer, strDBF As String
    Dim conFox As New ADODB.Connection, rsTmp As ADODB.Recordset, mTotalAmount As Double, i As Integer
    Dim mBillingNo As String, mTotalSJ As Double, mTotalOuts As Double
    Dim mBillDate As Date, mDueDate As Date, lInsertOuts As Boolean, mIncomeCode As String, c_Amount As Double
    Dim ARUpTo As Date, ARFrom As Date, mPeriode1 As Date
    Dim rs_D_BillPhone As RemoteSet
    Dim dbLocal As DAO.Database, rsLocal As DAO.Recordset
    Dim dbLocal_1 As DAO.Database, rsLocal_1 As DAO.Recordset
    Dim dbLocal_2 As DAO.Database, rsLocal_2 As DAO.Recordset, lProcessKelebihan As Boolean
    Dim dbLocal_3 As DAO.Database
    
    On Error GoTo errMsg
    strDBF = PhoneDir & "BILL-TLP.DBF"
'    If Dir(strDBF) = "" Then GoTo errMsg
'    conFox.ConnectionString = "Driver={Microsoft Visual FoxPro Driver};UID=;SourceDB=" & PhoneDir & _
'                               ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;" & _
'                               "Null=Yes;Deleted=Yes;"
'    conFox.CursorLocation = adUseClient
'    conFox.Open
    
    Set rsRate = db.OpenRemoteSet("select top 1 * from m_rate where convert(varchar(6),date,112)=" & Format(Period_CAR, "yyyymm"))
    If rsRate Is Nothing Then
        MsgBox "M_Rate Nothing !"
        GoTo ErrMessage
    ElseIf rsRate.EOF Then
        MsgBox "M_Rate EOF !"
        GoTo ErrMessage
    Else
        mRate = rsRate("AcctRate")
    End If
    ' cek apakah tagihan telepon sudah siap utk periode KAR-1 !!
    Set rsTmp1 = db.OpenRemoteSet("SELECT TOP 1 * FROM KTA_T_BillingPhone where convert(varchar(8),billdate,112)='" + Format(DateAdd("d", -1, DateAdd("m", -1, DateAdd("d", 1, Period_AR))), "yyyymmdd") + "'")
'    Set rsTmp = conFox.Execute("select * from bill-tlp where period='" & Format(DateAdd("m", -1, Period_AR), "yyyymm") & "'")
'    Set rsTmp = conFox.Execute("select * from bill-tlp where period='" & Format(Period_AR, "yyyymm") & "'")
    If rsTmp1.EOF Then
        MsgBox "Billing for telephone doesn't exists for period " & Format(DateAdd("m", -1, Period_AR), "mmm-yyyy") & " !"
'        MsgBox "Billing for telephone doesn't exists for period " & Format(Period_AR, "mmm-yyyy") & " !"
        GoTo ErrMessage
    End If
    'If Dir(PhoneDir & "Billing_Member.mdb") = "" Then
    If Dir(PhoneDir & "Billing_Member.csv") = "" Then
        MsgBox "Billing parking file " & PhoneDir & "Billing_Member.csv doesn't exists !"
        GoTo ErrMessage
    End If
    'Set dbLocal = OpenDatabase(PhoneDir & "Billing_Member.mdb")
    Set dbLocal = OpenDatabase(PhoneDir, False, False, "Text;Database=" & PhoneDir & ";table=Billing_Member.csv")
    'Set rsLocal = dbLocal.OpenRecordset("select top 1 * from billing_parking where format(date_trans,'yyyymm')='" & Format(Period_AR, "yyyymm") & "'", dbOpenForwardOnly)
    Set rsLocal = dbLocal.OpenRecordset("select top 1 * from Billing_Member.csv where format(date_trans,'yyyymm')='" & Format(Period_AR, "yyyymm") & "'", dbOpenForwardOnly)
    If rsLocal.EOF Then
        MsgBox "Billing for parking doesn't exists for period " & Format(Period_AR, "mmm-yyyy") & " !"
        GoTo ErrMessage
    End If
    rsLocal.Close
    If Dir(PhoneDir & "Billing_KelebihanTagih.csv") <> "" Then
        Set dbLocal_2 = OpenDatabase(PhoneDir, False, False, "Text;Database=" & PhoneDir & ";table=Billing_KelebihanTagih.csv")
        lProcessKelebihan = True
    Else
        lProcessKelebihan = False
    End If
    
    If Filename.Text <> "" Then
        Set dbLocal_1 = OpenDatabase(Filename.Text, False, False, "Excel 8.0;")
    End If
    If FileName1.Text <> "" Then
        Set dbLocal_3 = OpenDatabase(FileName1.Text, False, False, "Excel 8.0;")
    End If
    
    On Error GoTo ErrMessage
    
    Set rs_H_Outstanding = db.OpenRemoteSet("select * from kta_t_h_outstanding where 1<>1")
    Set rs_D_Outstanding = db.OpenRemoteSet("select * from kta_t_d_outstanding where 1<>1")
    Set rs_H_SJ = db.OpenRemoteSet("select * from ar_t_h_salesjournal where 1<>1")
    Set rs_D_SJ = db.OpenRemoteSet("select * from ar_t_d_salesjournal where 1<>1")
    Set rs_H_Billing = db.OpenRemoteSet("select * from kta_t_h_billing where 1<>1")
    Set rs_D_BillPhone = db.OpenRemoteSet("select * from kta_t_d_billing_phone where 1<>1")
    
    mPeriode = DateAdd("m", -1, Period_AR + 1) - 1  ' periode bulan lalu
    mPeriode1 = IIf(Month(todate(DateField(0))) = 3, DateSerial(Year(DateField(0)), Month(DateField(0)), 2), mPeriode)  ' batas pengecekan duedate khusus utk BS feb pada proses LPC..
    ARUpTo = DateSerial(Year(Period_AR), Month(Period_AR), xBillDate) - 1
    ARFrom = DateSerial(Year(mPeriode), Month(mPeriode), xBillDate)
'    Set rs = db.OpenRemoteSet("Select a.unitno,a.ownerid,startdate,badate,handoverdate,b.*,d.payment,newbalance from kta_m_unit a inner join kta_m_roomtype b on a.companyid=b.companyid and a.locationid=b.locationid and a.roomtype=b.roomtype left join kta_t_h_billing c" & _
'                        " on a.companyid=c.comp1000036.00anyid and a.locationid=c.locationid and a.unitno=c.unitno left join (select companyid, locationid, unitno, payment=sum(d.amount) from kta_t_d_outstanding d inner join kta_t_h_outstanding h on d.keyhd=h.keyhd and d.refno<>''" & _
'                        " and txtype in ('W','B') and locationno='' and recordtype in (3,4,5) where convert(varchar(6),refdate,112)=" & Format(mPeriode, "yyyymm") & " or (convert(varchar(6),refdate,112)>" & Format(mPeriode, "yyyymm") & " and refdate<='" & Format(ARUpTo, "dd-mmm-yyyy") & _
'                        "' and (txtype='B' or (txtype='W' and convert(varchar(6),documentdate,112)<=" & Format(mPeriode, "yyyymm") & ")))" & _
'                        " group by companyid, locationid, unitno) d on a.companyid=d.companyid and a.locationid=d.locationid and a.unitno=d.unitno" & _
'                        " where a.companyid='" & companyId & "' and a.locationid='" & LocationId & "' and a.tower<>'' and a.ownerid<>'' and convert(varchar(6),c.billdate,112)=" & Format(mPeriode, "yyyymm") & " order by a.unitno")
'    Set rs = db.OpenRemoteSet("Select a.unitno,a.ownerid,startdate,badate,handoverdate,b.*,d.payment,newbalance,a.sendhardcopy from kta_m_unit a inner join kta_m_roomtype b on a.companyid=b.companyid and a.locationid=b.locationid and a.roomtype=b.roomtype left join kta_t_h_billing c" & _
'                        " on a.companyid=c.companyid and a.locationid=c.locationid and a.unitno=c.unitno left join (select companyid, locationid, unitno, payment=sum(d.amount) from kta_t_d_outstanding d inner join kta_t_h_outstanding h on d.keyhd=h.keyhd where d.refno<>''" & _
'                        " and txtype in ('W','B') and locationno='' and recordtype in (3,4,5) and refdate <= '" & Format(ARUpTo, "dd-mmm-yyyy") & "' and ((txtype='w' and convert(varchar(6),documentdate,112)=" & Format(mPeriode, "yyyymm") & _
'                        " and convert(varchar(6),refdate,112)>=" & Format(mPeriode, "yyyymm") & ") or (txtype='w' and convert(varchar(6),documentdate,112)<" & Format(mPeriode, "yyyymm") & " and refdate>='" & Format(ARFrom, "dd-mmm-yyyy") & _
'                        "') or (txtype='b' and refdate>='" & Format(ARFrom, "dd-mmm-yyyy") & "'))" & _
'                        " group by companyid, locationid, unitno) d on a.companyid=d.companyid and a.locationid=d.locationid and a.unitno=d.unitno" & _
'                        " where a.companyid='" & companyId & "' and a.locationid='" & LocationId & "' and a.tower<>'' and a.ownerid<>'' and convert(varchar(6),c.billdate,112)=" & Format(mPeriode, "yyyymm") & " order by a.unitno")
    Set rs = db.OpenRemoteSet("Select a.unitno,a.ownerid,startdate,badate,handoverdate,b.*,d.payment,newbalance,a.sendhardcopy from kta_m_unit a inner join kta_m_roomtype b on a.companyid=b.companyid and a.locationid=b.locationid and a.roomtype=b.roomtype left join kta_t_h_billing c" & _
                        " on a.companyid=c.companyid and a.locationid=c.locationid and a.unitno=c.unitno left join (select companyid, locationid, unitno, payment=sum(CASE WHEN d.RefNo='-' THEN 0 ELSE d.amount END) from kta_t_d_outstanding d inner join vw_kta_t_h_outstanding h on d.keyhd=h.keyhd where d.refno<>''" & _
                        " and txtype in ('W','B') and locationno='' and recordtype in (3,4,5) and refdate <= h.enddate and ((txtype='w' and convert(varchar(6),documentdate,112)=convert(varchar(6),prevperiod,112)" & _
                        " and convert(varchar(6),refdate,112)>=convert(varchar(6),prevperiod,112)) or (txtype='w' and convert(varchar(6),documentdate,112)<convert(varchar(6),prevperiod,112) and refdate>=h.startdate" & _
                        ") or (txtype='b' and refdate>=h.startdate))" & _
                        " group by companyid, locationid, unitno) d on a.companyid=d.companyid and a.locationid=d.locationid and a.unitno=d.unitno" & _
                        " where a.companyid='" & companyId & "' and a.locationid='" & LocationId & "' and a.tower<>'' and a.ownerid<>'' and convert(varchar(6),c.billdate,112)=" & Format(mPeriode, "yyyymm") & " order by a.unitno")
    Progress.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
    i = 1
    db.Begintrans
    Do While Not rs.EOF
        Progress.Value = i
        mTotalAmount = 0
'        Set rsTemp = db.OpenRemoteSet("select * from kta_t_h_billing where companyid='" & companyId & "' and locationid='" & LocationId & "' and unitno='" & rs("UnitNo") & _
'                            "' and convert(varchar(6),billdate,112)=" & Format(mPeriode, "yyyymm"))
'        If rsTemp.EOF Then mPrev = 0 Else mPrev = rsTemp("NewBalance")
'        Set rsTemp = db.OpenRemoteSet("select payment=sum(amount) from kta_t_d_outstanding d, kta_t_h_outstanding h where d.keyhd=h.keyhd and companyid='" & companyId & _
'                            "' and locationid='" & LocationId & "' and txtype in ('W','B') and locationno='' and recordtype in (3,4) and unitno='" & _
'                            rs("UnitNo") & "' and convert(varchar(6),refdate,112)=" & Format(mPeriode, "yyyymm"))
        mPayment = toval(rs("Payment"))
        mPrev = toval(rs("NewBalance"))
        mBillDate = DateSerial(Year(Period_AR), Month(Period_AR), xBillDate)
        With rs_H_Billing
            .AddNew
            !companyId = companyId
            !LocationId = LocationId
            !billingno = GetNextDocNo("KBI" & companyId & LocationId, Period_CAR)
            !unitno = rs("UnitNo")
            '!BillDate = DateAdd("d", -Day(Period_CAR), Period_CAR) + 1
            !BillDate = mBillDate
            !duedate = DateAdd("d", TermDue, !BillDate) - 1
            !CurrencyID = rs("CurrencyID")
            !Previous = mPrev
            !payment = mPayment
            !NewBilling = 0     ' di-update pada saat break unitno
            !NewBalance = 0     ' di-update pada saat break unitno
            !PrintCode = 0
            !RePrint = 0
            !UpdateUser = tostr(db_user)
            !UpdateDate = todate(db.ServerDate + db.ServerTime)
            If Not .Update Then
            GoTo ErrMessage
            End If
            mKeyHD = !KeyHD
        End With
        mSeqNo = 0
        ' maintenance fee & sinking fund
        ' -----------------------------------------------------------
'        If Format(rs("badate"), "yyyymm") < Format(Period_CAR, "yyyymm") Then
            mSeqNo = mSeqNo + 1
            If Not InsertDetail(mKeyHD, mSeqNo, "MFE", rs("Area") * (rs("MaintenanceFee") + rs("SinkingFund")), "", todate("1/1/1900"), "", mTotalAmount, "") Then
                GoTo ErrMessage
            End If
'        End If
'        If Not InsertDetail(mKeyHD, mSeqNo, "MFE", rs("Area") * rs("MaintenanceFee"), "", todate("1/1/1900"), "", mTotalAmount) Then
'            GoTo errMessage
'        End If
        If rs("SendHardCopy") = "Y" Then
            mSeqNo = mSeqNo + 1
            If Not InsertDetail(mKeyHD, mSeqNo, "PHC", HardCopyCharges, "", todate("1/1/1900"), "", mTotalAmount, "") Then
                GoTo ErrMessage
            End If
        End If
        ' work order
        ' -----------------------------------------------------------
        Set rsTemp = db.OpenRemoteSet("select * from kta_t_h_outstanding h, kta_t_d_outstanding d where d.keyhd=h.keyhd and companyid='" & companyId & _
                            "' and locationid='" & LocationId & "' and recordtype=1 and txtype='W' and locationno='' and convert(varchar(6),documentdate,112)=" & _
                            Format(mPeriode, "yyyymm") & " and unitno='" & rs("UnitNo") & "'")
        Do While Not rsTemp.EOF
            mSeqNo = mSeqNo + 1
            If Not InsertDetail(mKeyHD, mSeqNo, rsTemp("IncomeCode"), rsTemp("Amount"), rsTemp("RefNo"), rsTemp("RefDate"), "", mTotalAmount, "") Then
                GoTo ErrMessage
            End If
            rsTemp.MoveNext
        Loop
        ' collect call
        ' -----------------------------------------------------------
        Set rsTemp = db.OpenRemoteSet("select * from kta_t_collectcall where companyid='" & companyId & "' and locationid='" & LocationId & "' and unitno='" & rs("unitno") & "' and convert(varchar(6),period,112)=" & Format(mPeriode, "yyyymm"))
        If Not rsTemp.EOF Then
            mSeqNo = mSeqNo + 1
'            If Not InsertDetail(mKeyHD, mSeqNo, "TEC", rsTemp("Amount") * (100 + rsTemp("VAT")) / 100, "", todate("1/1/1900"), "", mTotalAmount) Then
            If Not InsertDetail(mKeyHD, mSeqNo, rsTemp("IncomeCode"), rsTemp("Amount") + rsTemp("VAT"), "", todate("1/1/1900"), "", mTotalAmount, rsTemp("Description")) Then
                GoTo ErrMessage
            End If
        End If
        ' telepon
        ' -----------------------------------------------------------
'        Set rsTmp = conFox.Execute("select phoneno, int((monthlyfee+local+sljj+sli+airfee+japati)*1.1) as amount from bill-tlp where rtnmtc='" & rs("UnitNo") & "' and period='" & Format(DateAdd("m", -1, mPeriode), "yyyymm") & "'")
        'Set rsTmp = conFox.Execute("select phoneno, int(((monthlyfee+local+sljj+sli+airfee+japati)*10/9)*1.1) as amount from bill-tlp where rtnmtc='" & rs("UnitNo") & "' and period='" & Format(mPeriode, "yyyymm") & "'")
'        Set rsTmp = conFox.Execute("select phoneno, int((monthlyfee+local+sljj+sli+airfee+japati)*1.1) as amount from bill-tlp where rtnmtc='" & Rs("UnitNo") & "' and period='" & Format(mPeriode, "yyyymm") & "'")
        'Set rsTmp = conFox.Execute("select phoneno, monthlyfee,local,sljj,sli,airfee,japati,localpulse,sljjpulse from bill-tlp where rtnmtc='" & rs("UnitNo") & "' and period='" & Format(mPeriode, "yyyymm") & "'")
        'Set rsTmp = conFox.Execute("select period, phoneno, monthlyfee,local,sljj,sli,airfee,japati,localpulse,sljjpulse, other from bill-tlp where rtnmtc='" & rs("UnitNo") & "' and period='" & Format(mPeriode, "yyyymm") & "' order by phoneno")
        Set rsTmp1 = db.OpenRemoteSet("SELECT CONVERT(VARCHAR(6),BillDate,112) period, phoneno, fee monthlyfee,[local],sljj,international sli,airtime airfee,japati,pulsa_local localpulse,pulsa_sljj sljjpulse, 0 other " + _
        "from KTA_T_BillingPhone where unitno='" + rs("UnitNo") + "' and convert(varchar(8),billdate,112)='" + Format(DateAdd("d", -1, DateAdd("m", -1, DateAdd("d", 1, Period_AR))), "yyyymmdd") + "' order by phoneno")
        
        Do While Not rsTmp1.EOF
            mSeqNo = mSeqNo + 1
'            mTotal = Int((rsTmp("MonthlyFee") + rsTmp("Local") + rsTmp("SLJJ") + rsTmp("SLI") + rsTmp("AirFee") + rsTmp("Japati")) * 1.1) '+ rsTmp("SLJJPulse") + rsTmp("LocalPulse")
            With rs_D_BillPhone
                .AddNew
                .Fields("keyhd") = mKeyHD
                .Fields("PhoneNo") = rsTmp1("PhoneNo")
                .Fields("Fee") = rsTmp1("monthlyfee")
                .Fields("Local") = rsTmp1("local")
                .Fields("SLJJ") = rsTmp1("SLJJ")
                .Fields("International") = rsTmp1("sli")
                .Fields("Airtime") = rsTmp1("airfee")
                .Fields("japati") = rsTmp1("japati")
                .Fields("SubTotal") = rsTmp1("monthlyfee") + rsTmp1("local") + rsTmp1("SLJJ") + rsTmp1("sli") + rsTmp1("airfee") + rsTmp1("japati")
                .Fields("Pulsa_Local") = rsTmp1("localpulse")
                .Fields("Pulsa_SLJJ") = rsTmp1("sljjpulse")
                .Fields("Other") = rsTmp1("other")
                .Fields("BeaAdmin") = Round(.Fields("SubTotal") / 9, 0)
                .Fields("PPN") = Round((.Fields("SubTotal") + .Fields("BeaAdmin")) * PPN, 0)
                .Fields("Total") = .Fields("SubTotal") + .Fields("BeaAdmin") + .Fields("PPN") + .Fields("other")
                mTotal = .Fields("Total")
                If Not .Update Then
                GoTo ErrMessage
                End If
            End With
            If Not InsertDetail(mKeyHD, mSeqNo, "TEC", mTotal, "", todate("1/1/1900"), rsTmp1("PhoneNo"), mTotalAmount, "") Then
                GoTo ErrMessage
            End If
            rsTmp1.MoveNext
        Loop

        ' parking fee
        ' -----------------------------------------------------------
        'Set rsLocal = dbLocal.OpenRecordset("select * from billing_parking where format(date_trans,'yyyymm')='" & Format(Period_AR, "yyyymm") & "' and unitno='" & rs("UnitNo") & "' order by description", dbOpenForwardOnly)
        'Set rsLocal = dbLocal.OpenRecordset("select * from Billing_Member.csv where format(date_trans,'yyyymm')='" & Format(Period_AR, "yyyymm") & "' and unitno='" & rs("UnitNo") & "' order by description", dbOpenForwardOnly)
        Set rsLocal = dbLocal.OpenRecordset("select * from Billing_Member.csv where format(date_trans,'yyyymm')='" & Format(Period_AR, "yyyymm") & "' and unitno='" & rs("UnitNo") & "' order by date_trans, description", dbOpenForwardOnly)
        Do While Not rsLocal.EOF
            mSeqNo = mSeqNo + 1
            If Not InsertDetail(mKeyHD, mSeqNo, "PCR", rsLocal("Amount").Value, "", todate("1/1/1900"), "", mTotalAmount, Left(rsLocal("Description").Value, 100)) Then
                GoTo ErrMessage
            End If
            rsLocal.MoveNext
        Loop
        rsLocal.Close
        
        If lProcessKelebihan Then
            Set rsLocal_2 = dbLocal_2.OpenRecordset("select * from Billing_KelebihanTagih.csv where format(date_trans,'yyyymm')='" & Format(Period_AR, "yyyymm") & "' and unitno='" & rs("UnitNo") & "' order by date_trans, description", dbOpenForwardOnly)
            Do While Not rsLocal_2.EOF
                mSeqNo = mSeqNo + 1
                If Not InsertDetail(mKeyHD, mSeqNo, rsLocal_2("income").Value, rsLocal_2("Amount").Value, "", todate("1/1/1900"), "", mTotalAmount, Left(rsLocal_2("Description").Value, 100)) Then
                    GoTo ErrMessage
                End If
                rsLocal_2.MoveNext
            Loop
            rsLocal_2.Close
            dbLocal_2.Close
        End If
        
        
        If Filename.Text <> "" Then
        ' ev charging
        ' -----------------------------------------------------------
        Set rsLocal_1 = dbLocal_1.OpenRecordset("select * from " + SheetName + " where left(period,10)='" & Format(DateAdd("m", -2, DateAdd("d", 1, Period_AR)), "yyyy-mm-dd") & "' and [DRIVER ACCOUNT]='" & rs("UnitNo") & "'", dbOpenForwardOnly)
        If Not rsLocal_1.EOF Then
            mSeqNo = mSeqNo + 1
            If Not InsertDetail(mKeyHD, mSeqNo, "ELV", Round(rsLocal_1("SUBTOTAL AMOUNT").Value, 0), "", todate("1/1/1900"), "", mTotalAmount, "Electricity Reimbursement for EV Charging EV-1, Usage Period " + Format(DateAdd("m", -2, DateAdd("d", 1, Period_AR)), "MMMM, YYYY")) Then
                GoTo ErrMessage
            End If
        End If
        rsLocal_1.Close
        'dbLocal_1.Close
        End If
        
        If FileName1.Text <> "" Then
        ' ev charging
        ' -----------------------------------------------------------
        Set rsLocal_1 = dbLocal_3.OpenRecordset("select * from " + SheetName1 + " where [CHARGING CARD NAME]='" & rs("UnitNo") & "'", dbOpenForwardOnly)
        While Not rsLocal_1.EOF
            mSeqNo = mSeqNo + 1
            If Not InsertDetail(mKeyHD, mSeqNo, "ELV", Round(rsLocal_1("TOTAL POWER COST").Value, 0), "", todate("1/1/1900"), "", mTotalAmount, "Electricity Reimbursement for EV Charging EV-2, Usage Period " + Format(CDate(Left(rsLocal_1("STATISTICAL PERIOD").Value, 10)), "MMMM, YYYY")) Then
                GoTo ErrMessage
            End If
            rsLocal_1.MoveNext
        Wend
        rsLocal_1.Close
        'dbLocal_3.Close
        End If
        
        ' air & listrik
        ' -----------------------------------------------------------
        If rs("handoverdate") <> todate("1/1/1900") And Format(rs("handoverdate"), "yyyymm") < Format(Period_CAR, "yyyymm") Then
'            Set rsTemp = db.OpenRemoteSet("select d.amount from ar_t_h_billing h, ar_t_d_billing d where d.keyhd=h.keyhd and companyid='" & companyId & "' and locationid='" & LocationId & _
'                                    "' and convert(varchar(6),billdate,112)=" & Format(mPeriode, "yyyymm") & " and tenantid='" & rs("UnitNo") & "' and incomecode='elc' and cancel='n' and flag='' and sendno<>''")
            Set rsTemp = db.OpenRemoteSet("select d.amount,[description]='Electricity Reimbursement, usage period '+DATENAME(MONTH,billdate)+', '+CAST(YEAR(billdate) AS VARCHAR) from ar_t_h_billing h, ar_t_d_billing d where d.keyhd=h.keyhd and companyid='" & companyId & "' and locationid='" & LocationId & _
                                    "' and convert(varchar(6),billdate,112)=" & Format(mPeriode, "yyyymm") & " and tenantid='" & rs("UnitNo") & "' and incomecode='elc' and cancel='n' and flag='' and sendno<>''")
            Do While Not rsTemp.EOF
                mSeqNo = mSeqNo + 1
                If Not InsertDetail(mKeyHD, mSeqNo, "ELC", rsTemp("Amount"), "", todate("1/1/1900"), "", mTotalAmount, rsTemp("description")) Then
                    GoTo ErrMessage
                End If
                rsTemp.MoveNext
            Loop
'            Set rsTemp = db.OpenRemoteSet("select d.amount from ar_t_h_billing h, ar_t_d_billing d where d.keyhd=h.keyhd and companyid='" & companyId & "' and locationid='" & LocationId & _
'                                    "' and convert(varchar(6),billdate,112)=" & Format(mPeriode, "yyyymm") & " and tenantid='" & rs("UnitNo") & "' and incomecode='air' and cancel='n' and flag='' and sendno<>''")
            Set rsTemp = db.OpenRemoteSet("select d.amount,[description]='Water Reimbursement, usage period '+DATENAME(MONTH,billdate)+', '+CAST(YEAR(billdate) AS VARCHAR) from ar_t_h_billing h, ar_t_d_billing d where d.keyhd=h.keyhd and companyid='" & companyId & "' and locationid='" & LocationId & _
                                    "' and convert(varchar(6),billdate,112)=" & Format(mPeriode, "yyyymm") & " and tenantid='" & rs("UnitNo") & "' and incomecode='air' and cancel='n' and flag='' and sendno<>''")
            Do While Not rsTemp.EOF
                mSeqNo = mSeqNo + 1
                If Not InsertDetail(mKeyHD, mSeqNo, "AIR", rsTemp("Amount"), "", todate("1/1/1900"), "", mTotalAmount, rsTemp("description")) Then
                    GoTo ErrMessage
                End If
                rsTemp.MoveNext
            Loop
        End If
        ' late payment charges
        ' -----------------------------------------------------------
        '
        Dim c_Break As Integer, c_KeyHD As Long, c_DueDate As Date
        Dim c_OutsAmount As Double, c_Month As Integer, c_TotPenalty As Double, c_DocumentNo As String, c_StartDate As Date, c_EndDate As Date
        Dim c_ChargesAmount As Double, c_SeqNo As Integer, c_BillAmount As Double, c_CalAmount As Double, c_OffsetAmount As Double
        Dim c_BillDate As Date, c_OwnerID As String, c_UnitNo As String, c_IncomeCode As String, c_CurrencyID As String
        Dim c_OutsAmount1 As Double, c_IncomeCode1 As String, c_RKDate1 As Date

'        Set rsTemp = db.OpenRemoteSet("select h.*,d.* from kta_t_h_outstanding h, kta_t_d_outstanding d, kta_m_owner a where d.keyhd=h.keyhd and txtype='B' and paiddate='1-Jan-1900' and refno<>'' and h.unitno='" & _
'                                    rs("UnitNo") & "' and h.ownerid='" & rs("OwnerID") & "' and h.companyid='" & companyId & "' and h.locationid='" & LocationId & _
'                                    "' and a.companyid+a.locationid+a.ownerid=h.companyid+h.locationid+h.ownerid and a.penalty='y' order by d.keyhd,case when recordtype=1 then 0 else 1 end,rkdate")
'        Set rsTemp = db.OpenRemoteSet("select d.*,h.*,billdate from kta_t_h_outstanding h inner join kta_t_d_outstanding d on d.keyhd=h.keyhd inner join (kta_m_owner x inner join kta_m_unit y on x.companyid=y.companyid and x.locationid=y.locationid and x.ownerid=y.ownerid) " & _
'                                "on y.companyid=h.companyid and y.locationid=h.locationid and y.unitno=h.unitno left join (SELECT h.CompanyID, h.LocationID, h.UnitNo, h.BillDate, d.InvoiceNo, d.IncomeCode " & _
'                                "FROM kta_t_d_billing d, kta_t_h_billing h WHERE d.KeyHD=h.KeyHD AND d.IncomeCode IN ('WOL','WOM')) z on z.companyid=h.companyid and z.locationid=h.locationid and z.unitno=h.unitno and z.incomecode=h.incomecode and z.invoiceno=h.documentno " & _
'                                "where x.penalty='y' and exists (select * from kta_t_h_outstanding a inner join kta_t_d_outstanding b on a.keyhd=b.keyhd where b.refno<>''" & _
'                                " and a.companyid=h.companyid and a.locationid=h.locationid and a.ownerid=h.ownerid" & _
'                                " and a.companyid=h.companyid and a.locationid=h.locationid and a.documentno=h.documentno and a.txtype=h.txtype and rkdate<=duedate and convert(varchar(8),duedate,112)<='" & Format(ARUpTo, "yyyymmdd") & "' group by a.companyid,a.locationid,a.documentno,a.txtype" & _
'                                " having sum(case when b.recordtype=1 then b.amount else -b.amount end)>0 and convert(varchar(8),max(b.rkdate),112)<='" & Format(ARUpTo, "yyyymmdd") & _
'                                "') and h.unitno='" & rs("UnitNo") & "' and (h.paiddate='1/1/1900' or convert(varchar(6),h.paiddate,112)=" & Format(mPeriode, "yyyymm") & ") and h.companyid='" & companyId & "' and h.locationid='" & LocationId & "' and d.refno<>'' order by d.keyhd, case when d.recordtype=1 then 1 when d.recordtype=3 then 2 else 3 end, d.rkdate")
        
'        Set rsTemp = db.OpenRemoteSet("select h.*,d.* from kta_t_h_outstanding h, kta_t_d_outstanding d, kta_m_owner a where d.keyhd=h.keyhd and txtype='B' AND convert(varchar(6),refdate,112)<=" & _
'                                    Format(Period_AR, "yyyymm") & " and paiddate='1-Jan-1900' and refno<>'' and h.unitno='" & rs("UnitNo") & "' and h.ownerid='" & rs("OwnerID") & _
'                                    "' and h.companyid='" & companyId & "' and h.locationid='" & LocationId & _
'                                    "' and a.companyid+a.locationid+a.ownerid=h.companyid+h.locationid+h.ownerid and a.penalty='y' order by d.keyhd,case when recordtype=1 then 0 else 1 end,rkdate")
        Set rsTemp = db.OpenRemoteSet("sp_KTA_selectLPC '" & companyId & "','" & LocationId & "','" & rs("UnitNo") & "','" & Format(DateField(0), "yyyymm") & "'")
        c_Break = 9
        c_Month = Day(mPeriode)
        c_OutsAmount1 = 0
        c_IncomeCode1 = ""
        Do While Not rsTemp.EOF
            If c_Break > 0 Then
                c_CurrencyID = rsTemp("CurrencyID")
                c_OutsAmount = rsTemp("Amount")
                c_KeyHD = rsTemp("KeyHD")
                c_SeqNo = 0
                c_DueDate = rsTemp("DueDate")
                c_DocumentNo = rsTemp("DocumentNo")
                c_StartDate = rsTemp("StartDate")
                c_EndDate = rsTemp("EndDate")
                c_BillAmount = rsTemp("Amount")
                c_OffsetAmount = rsTemp("OffsetAmount")
                c_OwnerID = rsTemp("OwnerID")
                c_UnitNo = rsTemp("UnitNo")
                c_IncomeCode = rsTemp("IncomeCode")
                c_BillDate = rsTemp("DocumentDate")
            If c_Break > 1 Then
                c_TotPenalty = 0
            End If
            End If
            'If rsTemp("RecordType") = 4 And rsTemp("RKDate") >= ARFrom And rsTemp("RKDate") <= ARUpTo And rsTemp("RKDate") > c_DueDate Then
            If rsTemp("RecordType") = 4 And rsTemp("RKDate") >= c_StartDate And rsTemp("RKDate") <= c_EndDate And rsTemp("RKDate") > c_DueDate Then
                c_CalAmount = IIf(c_OutsAmount > rsTemp("Amount"), rsTemp("Amount"), c_OutsAmount)
                'c_ChargesAmount = toval((((rsTemp("RKDate") - ARFrom + 1) / c_Month) * c_CalAmount * (Penalty / 100)), 0)
                c_ChargesAmount = toval((((rsTemp("RKDate") - c_StartDate + 1) / c_Month) * c_CalAmount * (Penalty / 100)), 0)
                c_TotPenalty = c_TotPenalty + c_ChargesAmount
                c_SeqNo = c_SeqNo + 1
'                If Not InsertLPC(c_UnitNo, c_DocumentNo, c_CurrencyID, c_IncomeCode, c_SeqNo, c_BillDate, c_DueDate, _
'                                c_OwnerID, c_CalAmount, rsTemp("RKDate") - ARFrom + 1, c_Month, c_CalAmount, Penalty, c_ChargesAmount) Then
                If Not InsertLPC(c_UnitNo, c_DocumentNo, c_CurrencyID, c_IncomeCode, c_SeqNo, c_BillDate, c_DueDate, _
                                c_OwnerID, c_CalAmount, rsTemp("RKDate") - c_StartDate + 1, c_Month, c_CalAmount, Penalty, c_ChargesAmount) Then
                    GoTo ErrMessage
                End If
                c_RKDate1 = rsTemp("RKDate")
            End If
            'If rsTemp("RecordType") <> "1" Then
            If rsTemp("RecordType") <> "1" And rsTemp("RKDate") <= rsTemp("EndDate") Then
                c_OutsAmount = c_OutsAmount - rsTemp("Amount")
            End If
            rsTemp.MoveNext
            If rsTemp.EOF Then
                c_Break = 9
            ElseIf rsTemp("KeyHD") <> c_KeyHD Then
                c_Break = 1
            Else
                c_Break = 0
            End If
            If c_Break = 1 And c_OutsAmount < 0 And c_OutsAmount1 = 0 And c_IncomeCode = "AIR" And Format(c_BillDate, "yyyymm") = "202511" Then
                c_OutsAmount1 = c_OutsAmount
                c_IncomeCode1 = c_IncomeCode
            End If
            If c_Break > 0 Then
                If c_OutsAmount > 0 And Format(c_DueDate, "yyyymm") <= Format(mPeriode1, "yyyymm") Then
                    'c_ChargesAmount = toval(c_OutsAmount * (c_Month / c_Month) * (Penalty / 100), 0)
                    c_ChargesAmount = toval(c_OutsAmount * ((c_EndDate - c_StartDate + 1) / c_Month) * (Penalty / 100), 0)
                    c_TotPenalty = c_TotPenalty + c_ChargesAmount
                    c_SeqNo = c_SeqNo + 1
'                    If Not InsertLPC(c_UnitNo, c_DocumentNo, c_CurrencyID, c_IncomeCode, c_SeqNo, c_BillDate, c_DueDate, _
'                                    c_OwnerID, c_OutsAmount, c_Month, c_Month, c_OutsAmount, Penalty, c_ChargesAmount) Then
                    If Not InsertLPC(c_UnitNo, c_DocumentNo, c_CurrencyID, c_IncomeCode, c_SeqNo, c_BillDate, c_DueDate, _
                                    c_OwnerID, c_OutsAmount, c_EndDate - c_StartDate + 1, c_Month, c_OutsAmount, Penalty, c_ChargesAmount) Then
                        GoTo ErrMessage
                    End If
                End If
            If c_Break > 1 Then
                If c_OutsAmount1 < 0 And c_ChargesAmount <> 0 Then
                    c_ChargesAmount = toval(c_OutsAmount1 * ((c_RKDate1 - c_StartDate + 1) / c_Month) * (Penalty / 100), 0)
                    c_TotPenalty = c_TotPenalty + c_ChargesAmount
                    c_SeqNo = c_SeqNo + 1
                    If Not InsertLPC(c_UnitNo, c_DocumentNo, c_CurrencyID, c_IncomeCode1, c_SeqNo, c_BillDate, c_DueDate, _
                                    c_OwnerID, c_OutsAmount1, c_RKDate1 - c_StartDate + 1, c_Month, c_OutsAmount1, Penalty, c_ChargesAmount) Then
                        GoTo ErrMessage
                    End If
                End If
                c_OutsAmount1 = 0
                c_IncomeCode1 = ""
                If c_TotPenalty > MinPEC Then
                    mSeqNo = mSeqNo + 1
                    If Not InsertDetail(mKeyHD, mSeqNo, "PEC", c_TotPenalty, "", todate("1/1/1900"), "", mTotalAmount, "") Then
                        GoTo ErrMessage
                    End If
                End If
            End If
            End If
        Loop
        ' -----------------------------------------------------------
        rs.MoveNext
        ' update new billing & new balance header
        With rs_H_Billing
            .Edit
            !NewBilling = mTotalAmount
            !NewBalance = !Previous + !NewBilling - !payment
            If Not .Update Then
            GoTo ErrMessage
            End If
        End With
        i = i + 1
    Loop
    ' insert outstanding & sales journal
    ' -----------------------------------------------
    Set rs = db.OpenRemoteSet("select d.*,h.*,ownerid,craccount,deptid,y.description,maintenancefee,sinkingfund,area from kta_t_h_billing h, kta_t_d_billing d, kta_m_unit x, kta_m_income y, kta_m_roomtype z" & _
                        " where d.keyhd=h.keyhd and h.companyid+h.locationid=z.companyid+z.locationid and z.roomtype=x.roomtype and h.companyid='" & companyId & _
                        "' and h.locationid='" & LocationId & "' and convert(varchar(6),billdate,112)=" & Format(Period_CAR, "yyyymm") & _
                        " and printcode=0 and d.invoiceno='' and x.companyid+x.locationid+x.unitno=h.companyid+h.locationid+h.unitno" & _
                        " and y.companyid+y.locationid+y.incomecode=h.companyid+h.locationid+d.incomecode order by billingno,d.incomecode")
    c_Break = 9
    lInsertOuts = True
    Progress.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
    i = 1
    Set rsTemp = db.OpenRemoteSet("select * from kta_m_income where companyid='" & companyId & "' and locationid='" & LocationId & "' and incomecode='SKF'")
    Do While Not rs.EOF
        mIncomeCode = rs("IncomeCode")
        Progress.Value = i
        If lInsertOuts Then
            lInsertOuts = False
            mTotalOuts = 0
            With rs_H_Outstanding
                .AddNew
                !companyId = companyId
                !LocationId = LocationId
                !TxType = "B"
                !DocumentNo = rs("BillingNo")
                !IncomeCode = rs("incomecode")
                !DocumentDate = rs("BillDate")
                !duedate = rs("DueDate")
                !OwnerID = rs("OwnerID")
                !unitno = rs("UnitNo")
                !LocationNo = ""
                !CurrencyID = rs("CurrencyID")
                !PaidDate = "1-Jan-1900"
                !Status = 0
                !UpdateUser = tostr(db_user)
                !UpdateDate = todate(db.ServerDate + db.ServerTime)
                !CancelDate = todate("1-Jan-1900")
                If Not .Update Then
                GoTo ErrMessage
                End If
            End With
        End If
        If c_Break > 0 Then
            mBillingNo = rs("BillingNo")
            mBillDate = rs("BillDate")
            mDueDate = rs("DueDate")
            With rs_H_SJ
                .AddNew
                !companyId = companyId
                !LocationId = LocationId
                !DocumentNo = rs("BillingNo")
                !RecordType = 1
                !PaymentFirst = 0
                !refno = ""
                !Date = rs("BillDate")
                !duedate = rs("DueDate")
                !CurrencyID = rs("CurrencyID")
                !OriginalAmount = 0
                !TenantID = rs("UnitNo")
                !SuiteNo = ""
                !PrintCode = 1
                !Status = 0
                !UpdateUser = tostr(db_user)
                !UpdateDate = todate(db.ServerDate + db.ServerTime)
                !Posted_to_GL = 0
                !Category = ""
                !ARType = "B"
                !VAT = 0
                !PPH23 = 0
                If Not .Update Then
                GoTo ErrMessage
                End If
            End With
            mTotalSJ = 0
            mSeqNo = 0
        End If
        mTotalOuts = mTotalOuts + rs("Amount")
        mTotalSJ = mTotalSJ + rs("Amount")
        mSeqNo = mSeqNo + 1
        If rs("IncomeCode") = "MFE" Then
            c_Amount = rs("Area") * rs("MaintenanceFee")
            If Not InsertDetailSJ(rs_D_SJ, rs_H_SJ!KeyHD, mSeqNo, rs("IncomeCode"), rs("CrAccount"), rs("DeptID"), rs("Description"), "", rs("CurrencyID"), c_Amount, mRate) Then
                GoTo ErrMessage
            End If
            mSeqNo = mSeqNo + 1
            c_Amount = rs("Area") * rs("SinkingFund")
            If Not InsertDetailSJ(rs_D_SJ, rs_H_SJ!KeyHD, mSeqNo, rsTemp("IncomeCode"), rsTemp("CrAccount"), rsTemp("DeptID"), rsTemp("Description"), "", rs("CurrencyID"), c_Amount, mRate) Then
                GoTo ErrMessage
            End If
        Else
            If Not InsertDetailSJ(rs_D_SJ, rs_H_SJ!KeyHD, mSeqNo, rs("IncomeCode"), rs("CrAccount"), rs("DeptID"), rs("Description"), IIf(mIncomeCode = "TEC", "Phone No. " & rs("PhoneNo"), ""), rs("CurrencyID"), rs("Amount"), mRate) Then
                GoTo ErrMessage
            End If
        End If
        rs.MoveNext
        If rs.EOF Then
            c_Break = 9
            lInsertOuts = True
        ElseIf mBillingNo <> rs("BillingNo") Then
            c_Break = 1
            lInsertOuts = True
        Else
            c_Break = 0
            lInsertOuts = mIncomeCode <> rs("IncomeCode")
        End If
        If lInsertOuts Then
            With rs_D_Outstanding
                .AddNew
                !KeyHD = rs_H_Outstanding!KeyHD
                !RecordType = 1
                !refno = mBillingNo
                !refdate = mBillDate
                !RKDate = mDueDate
                !PaymentNo = ""
                !Amount = mTotalOuts
                !UpdateUser = tostr(db_user)
                !UpdateDate = todate(db.ServerDate + db.ServerTime)
                If Not .Update Then
                GoTo ErrMessage
                End If
            End With
        End If
        If c_Break > 0 Then
            rs_H_SJ.Edit
            rs_H_SJ!OriginalAmount = mTotalSJ
            If Not rs_H_SJ.Update Then
            GoTo ErrMessage
            End If
        End If
        i = i + 1
    Loop
    ' -----------------------------------------------------------------
    Set rs = db.OpenRemoteSet("select * from kta_t_processbilling where 1<>1")
    With rs
        .AddNew
        !companyId = companyId
        !LocationId = LocationId
        !Period = Format(Period_CAR, "yyyymm")
        !Status = 1
        !UpdateUser = tostr(db_user)
        !UpdateDate = todate(db.ServerDate + db.ServerTime)
        If Not .Update Then
        GoTo ErrMessage
        End If
    End With
    db.Execute ("DELETE KTA_T_BillingPhone where convert(varchar(8),billdate,112)='" + Format(DateAdd("d", -1, DateAdd("m", -1, DateAdd("d", 1, Period_AR))), "yyyymmdd") + "'")
    db.Committrans
    If lProcessKelebihan Then
        FileCopy PhoneDir & "Billing_KelebihanTagih.csv", PhoneDir & "Billing_KelebihanTagih_" & Format(Period_AR, "yyyymm") + ".csv"
        Kill PhoneDir & "Billing_KelebihanTagih.csv"
    End If
    ProsesBilling = True
    Set conFox = Nothing
    Exit Function

errMsg:
    ProsesBilling = False
    'MsgBox "Error opening " & strDBF & " !"
    MsgBox err.Description
    Exit Function
ErrMessage:
    db.RollBack
    Set conFox = Nothing
    ProsesBilling = False
End Function

Private Function InsertDetailSJ(rs As RemoteSet, mKeyHD As Long, mSeqNo As Integer, mIncomeID As String, mAccountCr As String, mProfitCenterNo As String, mDescription1 As String, mDescription2 As String, mCurrencyID As String, mOriginalAmount As Double, mRate As Double) As Boolean
    On Error GoTo errMsg
    With rs
        .AddNew
        !KeyHD = mKeyHD
        !SeqNo = mSeqNo
        !IncomeID = mIncomeID
        !AccountCr = mAccountCr
        !ProfitCenterNo = mProfitCenterNo
        !Description1 = mDescription1
        !Description2 = mDescription2
        !Description3 = ""
        !CurrencyID = mCurrencyID
        !OriginalAmount = mOriginalAmount
        !RpAmount = toval(IIf(mCurrencyID = "IDR", mOriginalAmount, mOriginalAmount * mRate), 0)
        !USAmount = toval(IIf(mCurrencyID = "IDR", mOriginalAmount / mRate, mOriginalAmount), 2)
        If Not .Update Then GoTo errMsg
    End With
    InsertDetailSJ = True
    Exit Function
errMsg:
    InsertDetailSJ = False
End Function

Private Function IsValidProcess(ByRef errMsg As String) As Boolean
    Dim rs As RemoteSet
    On Error GoTo ErrMessage
'    Set rs = db.OpenRemoteSet("select * from kta_t_workorder where companyid='" & companyId & "' and locationid='" & LocationId & "' and status=2 and invoiceno=''")
'    If rs Is Nothing Then
'        errmsg = "KTA_T_WorkOrder Nothing !"
'        GoTo ErrMessage
'    ElseIf Not rs.EOF Then
'        errmsg = "There're some completed WO invoice not yet issued !"
'        GoTo ErrMessage
'    End If
    Set rs = db.OpenRemoteSet("select * from kta_t_h_outstanding h where companyid='" & companyId & "' and locationid='" & LocationId & "' and txtype='W'" & _
                       " and unitno<>'' and locationno='' and paiddate='1-jan-1900'")
    If rs Is Nothing Then
        errMsg = "KTA_T_Outstanding Nothing !"
        GoTo ErrMessage
    ElseIf Not rs.EOF Then
'        Set rs = db.OpenRemoteSet("select * from kta_t_h_outstanding h where companyid='" & companyId & "' and locationid='" & LocationId & "' and txtype='W'" & _
'                           " and unitno<>'' and locationno='' and paiddate='1-jan-1900' and exists (select * from kta_t_d_outstanding d where d.keyhd=h.keyhd and recordtype=4)")
'        If rs.EOF Then
'            errmsg = "There're no collections for WO invoice !"
'            GoTo errMessage
'        End If
    End If
'    Set rs = db.OpenRemoteSet("select * from kta_t_processbilling where companyid='" & companyId & "' and locationid='" & LocationId & "' and period=" & Format(Period_CAR, "yyyymm"))
'    If rs Is Nothing Then
'        errMsg = "KTA_T_ProcessBilling Nothing !"
'        GoTo errMessage
'    ElseIf Not rs.EOF Then
'        errMsg = "Billing Statement for " & Format(Period_CAR, "mmm, yyyy") & " already processed !"
'        GoTo errMessage
'    End If
    If todate(Period_AR) <> todate(Period_CAR) Then
        errMsg = "Process can be done after EOM..."
        GoTo ErrMessage
    End If
    Dim oXls As Object
    Dim oDoc As Object
    Dim oSheet As Object
    Dim dbLocal As DAO.Database, rsLocal As DAO.Recordset
    If errMsg = "" And Filename.Text <> "" Then
        'check excel file
        Set oXls = GetObject("", "Excel.Application")
        Set oDoc = oXls.Workbooks.Open(Filename.Text)
        oXls.Visible = False
        Set oSheet = oDoc.Sheets(1)
        SheetName = "[" + oSheet.name + "$]"
        If Not (UCase(oSheet.cells(1, 1)) = "DRIVER ACCOUNT" And UCase(oSheet.cells(1, 2)) = "EMAIL" And UCase(oSheet.cells(1, 3)) = "RFID" _
        And UCase(oSheet.cells(1, 4)) = "CURRENCY" And UCase(oSheet.cells(1, 5)) = "SUBTOTAL AMOUNT" And UCase(oSheet.cells(1, 6)) = "SUBTOTAL POWER USAGE (KWH)" _
        And UCase(oSheet.cells(1, 7)) = "PERIOD") Then
            errMsg = "Invalid columns specification for EV Charging excel file (Format Lama)..."
        End If
        oDoc.Close
        Set oXls = Nothing
        If errMsg = "" Then
            Set dbLocal = OpenDatabase(Filename.Text, False, False, "Excel 8.0;")
            Set rsLocal = dbLocal.OpenRecordset("select top 1 * from " + SheetName + " where left(period,10)='" & Format(DateAdd("m", -2, DateAdd("d", 1, Period_AR)), "yyyy-mm-dd") & "'", dbOpenForwardOnly)
            'Set rsLocal = dbLocal.OpenRecordset("select top 1 * from " + SheetName + " ", dbOpenForwardOnly)
            If rsLocal.EOF Then
                errMsg = "Billing for EV Charging doesn't exists for period " & Format(Period_AR, "mmm-yyyy") & " (Format Lama) !"
            End If
            rsLocal.Close
            dbLocal.Close
        End If
        If errMsg <> "" Then GoTo ErrMessage
    End If
    If errMsg = "" And FileName1.Text <> "" Then
        'check excel file
        Set oXls = GetObject("", "Excel.Application")
        Set oDoc = oXls.Workbooks.Open(FileName1.Text)
        oXls.Visible = False
        Set oSheet = oDoc.Sheets(1)
        SheetName1 = "[" + oSheet.name + "$]"
        If Not (UCase(oSheet.cells(1, 1)) = "CHARGING CARD NAME" And UCase(oSheet.cells(1, 2)) = "CHARGING CARD NO." And UCase(oSheet.cells(1, 3)) = "EMAIL" _
        And UCase(oSheet.cells(1, 4)) = "CURRENCY" And UCase(oSheet.cells(1, 5)) = "TOTAL POWER COST" And UCase(oSheet.cells(1, 6)) = "TOTAL POWER CONSUMPTION (KWH)" _
        And UCase(oSheet.cells(1, 7)) = "STATISTICAL PERIOD") Then
            errMsg = "Invalid columns specification for EV Charging excel file (Format Baru)..."
        End If
        oDoc.Close
        Set oXls = Nothing
        If errMsg = "" Then
            Set dbLocal = OpenDatabase(FileName1.Text, False, False, "Excel 8.0;")
            Set rsLocal = dbLocal.OpenRecordset("select top 1 * from " + SheetName1 + " where left([STATISTICAL PERIOD],10)='" & Format(DateAdd("m", -2, DateAdd("d", 1, Period_AR)), "yyyy-mm-dd") & "'", dbOpenForwardOnly)
            If rsLocal.EOF Then
                errMsg = "Billing for EV Charging doesn't exists for period " & Format(Period_AR, "mmm-yyyy") & " (Format Baru) !"
            End If
            rsLocal.Close
            dbLocal.Close
        End If
        If errMsg <> "" Then GoTo ErrMessage
    End If
    errMsg = ""
    IsValidProcess = True
    Exit Function
ErrMessage:
    If errMsg = "" Then errMsg = err.Description
    IsValidProcess = False
End Function

Private Function InsertDetail(mKeyHD As Long, mSeqNo As Integer, mIncomeCode As String, mAmount As Double, mInvoiceNo As String, mInvoiceDate As Date, mPhoneNo As String, ByRef mTotalAmount As Double, mDescription As String) As Boolean
    Dim rsD As RemoteSet
    On Error GoTo ErrMessage
    
    If mAmount <> 0 Or mIncomeCode = "PHC" Then
        Set rsD = db.OpenRemoteSet("select * from kta_t_d_billing where 1<>1")
        With rsD
            .AddNew
            !KeyHD = mKeyHD
            !SeqNo = mSeqNo
            !IncomeCode = mIncomeCode
            !Amount = mAmount
            !InvoiceNo = mInvoiceNo
            !InvoiceDate = mInvoiceDate
            !PhoneNo = mPhoneNo
            !Description = mDescription
            mTotalAmount = mTotalAmount + mAmount
            If Not .Update Then GoTo ErrMessage
        End With
    End If
    InsertDetail = True
    Exit Function
ErrMessage:
    InsertDetail = False
End Function

Private Function InsertLPC(c_UnitNo As String, c_BillingNo As String, c_CurrencyID As String, c_IncomeCode As String, c_SeqNo As Integer, c_BillDate As Date, c_DueDate As Date, c_OwnerID As String, c_BillAmount As Double, c_CalDay As Integer, c_CalMonth As Integer, c_CalAmount As Double, c_CalPersen As Double, c_Charges As Double) As Boolean
    Dim rs As RemoteSet
    On Error GoTo ErrMessage
    
    Set rs = db.OpenRemoteSet("select * from kta_t_processbilling_calculation where 1<>1")
    With rs
        .AddNew
        !companyId = companyId
        !LocationId = LocationId
        !Period = Format(Period_CAR, "yyyymm")
        !unitno = c_UnitNo
        !billingno = c_BillingNo
        !IncomeCode = c_IncomeCode
        !SeqNo = c_SeqNo
        !BillDate = c_BillDate
        !CurrencyID = c_CurrencyID
        !duedate = c_DueDate
        !OwnerID = c_OwnerID
        !BillAmount = c_BillAmount
        !CalDay = c_CalDay
        !CalMonth = c_CalMonth
        !CalAmount = c_CalAmount
        !CalPersen = c_CalPersen
        !Charges = c_Charges
        !UpdateUser = tostr(db_user)
        !UpdateDate = todate(db.ServerDate + db.ServerTime)
        If Not .Update Then GoTo ErrMessage
    End With
    InsertLPC = True
    Exit Function
ErrMessage:
    InsertLPC = False
End Function

