VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExpenseReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Expenses"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrintExpenses 
      Caption         =   "Print &Expenses"
      Height          =   405
      Left            =   3195
      TabIndex        =   4
      Top             =   885
      Visible         =   0   'False
      Width           =   1950
   End
   Begin MSComCtl2.DTPicker DTStart 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   480
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dddd, dd-MMM-yyyy"
      Format          =   109445123
      CurrentDate     =   38957
   End
   Begin MSComCtl2.DTPicker DTEnd 
      Height          =   315
      Left            =   2655
      TabIndex        =   3
      Top             =   480
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dddd, dd-MMM-yyyy"
      Format          =   109445123
      CurrentDate     =   38957
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Finishing Date :"
      Height          =   195
      Index           =   1
      Left            =   2700
      TabIndex        =   2
      Top             =   255
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Starting Date :"
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   255
      Width           =   1065
   End
End
Attribute VB_Name = "frmExpenseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim myDTEnd As Date
Dim ForVenders As Boolean
Dim VendersDone As Boolean
Dim iPageNo As Integer
Dim i As Integer
Dim bExpense As Boolean

Private Sub chkMonthly_Click()
    'Call DTStart_Change
End Sub


Public Sub ShowMe(Optional p_bExpense As Boolean = False)
    bExpense = p_bExpense
    If bExpense Then
        cmdPrintExpenses.Visible = True
    End If
    Me.Show , MainForm
End Sub

Private Sub ArrangeData(DTStart As Date, DTEnd As Date)

    On Error GoTo err
    
    Dim DTPrevStart As Date, DTPrevEnd As Date
    DTEnd = DateAdd("d", -1, DTEnd)
    myDTEnd = DTEnd
    DTPrevStart = DateAdd("yyyy", -1, DTStart)
    DTPrevEnd = DateAdd("yyyy", -1, DTEnd)
    
    Dim myDT As Date
    Dim myList As String
    Dim myList1 As String
    con.Execute "Delete From rptSaleHeaders"
    con.Execute "Delete From rptSalePeriods"
    con.Execute "Delete From rptSalePeriodsPrevious"
    
    '293501
    
    myList = "'"
    myDT = DTStart
    Do While myDT < DTEnd
        myList1 = Format(myDT, "mmm-yy")
        myList = myList & myList1 & "','"
        con.Execute "Insert Into rptSalePeriods(Period) Values('" & myList1 & "')"
        myDT = DateAdd("m", 1, myDT)
    Loop
    
    myList = Left(myList, Len(myList) - 2)
    
    con.Execute "Insert Into rptSaleHeaders(Period1,Period2,Period3,Period4" & _
     ",Period5,Period6,Period7,Period8,Period9,Period10,Period11,Period12) Values(" & _
     myList & ")"
     
    
    myDT = DTPrevStart
    Do While myDT < DTPrevEnd
        myList1 = Format(myDT, "mmm-yy")
        con.Execute "Insert Into rptSalePeriodsPrevious(Period) Values('" & myList1 & "')"
        myDT = DateAdd("m", 1, myDT)
    Loop
    
    Dim lRefID As Long
    lRefID = Val(con.Execute("Select Max(EntryID) From rptSaleHeaders").Fields(0).Value & "")
    
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    myList = ""
    With rs
        .Open "SELECT * FROM dbo.F_SalesAcc ('" & DTStart & "' ,'" & DTEnd & "') Order By AccNo,SortID", con, adOpenForwardOnly, adLockReadOnly
        rs1.Open "SELECT AccNo FROM Accounts Where Left(AccNo,7)='15-003-'", con, adOpenForwardOnly, adLockReadOnly
        Do Until rs1.EOF
            .Filter = "AccNo='" & rs1!AccNo & "'"
            myList = ""
            Do Until .EOF
                myList = myList & Val(!Amount & "") & ","
                .MoveNEXT
            Loop
            myList = Left(myList, Len(myList) - 1)
            con.Execute "INSERT INTO rptSales(RefID,Customer,Period1,Period2,Period3,Period4,Period5,Period6,Period7,Period8,Period9,Period10,Period11,Period12) Values(" & lRefID & ",'" & rs1!AccNo & "'," & myList & ")"
            rs1.MoveNEXT
        Loop
        
        .Close
        rs1.Close
        '/////Previous///////
        .Open "Select * From dbo.F_SalesAccPrevious ('" & DTPrevStart & "' ,'" & DTPrevEnd & "') Order By AccNo,SortID", con, adOpenForwardOnly, adLockReadOnly
        rs1.Open "Select AccNo From Accounts Where Left(AccNo,7)='15-003-'", con, adOpenForwardOnly, adLockReadOnly
        Do Until rs1.EOF
            .Filter = "AccNo='" & rs1!AccNo & "'"
            myList = ""
            Do Until .EOF
                myList = myList & Val(!Amount & "") & ","
                .MoveNEXT
            Loop
            myList = Left(myList, Len(myList) - 1)
            con.Execute "Insert Into rptSalesPrevious(RefID,Customer,Period1,Period2,Period3,Period4,Period5,Period6,Period7,Period8,Period9,Period10,Period11,Period12) Values(" & lRefID & ",'" & rs1!AccNo & "'," & myList & ")"
            rs1.MoveNEXT
        Loop
        .Close
        rs1.Close
        '////////////////////
    End With
    
    Set rs = Nothing
    Set rs1 = Nothing
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdPrintExpenses_Click()
            
    If DateDiff("m", DTStart, DTEnd) <> 6 Then
        MsgBox "Invalid Date Selection.", vbInformation
        Exit Sub
    End If

    Me.MousePointer = vbHourglass
    Call ArrangeDataForExpenses(DTStart, DTEnd)
    Dim rpt As CRAXDDRT.Report, frm As New frmPrevRpt
    Set rpt = rptApp.OpenReport(RptPath & "\ExpenseReport6Months.rpt")
    rpt.FormulaFields(5).Text = "'" & Format(DTStart, "dd-MMM-yyyy") & " to " & Format(myDTEnd, "dd-MMM-yyyy") & "'"
    frm.ShowReport "ToNumber(LEFT({VChildAccounts.AccNo},2))>=31", rpt
    Me.MousePointer = vbNormal

    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub


Private Sub DTStart_Change()

   DTEnd = DateAdd("m", 6, DTStart)

End Sub


Private Sub ArrangeDataForExpenses(DTStart As Date, DTEnd As Date)

    On Error GoTo err
    
    Dim DTPrevStart As Date, DTPrevEnd As Date
    DTEnd = DateAdd("d", -1, DTEnd)
    myDTEnd = DTEnd
    myDTEnd = DateAdd("m", 12, DTStart)
    myDTEnd = DateAdd("d", -1, myDTEnd)
    
    DTPrevStart = DateAdd("yyyy", -1, DTStart)
    DTPrevEnd = DateAdd("yyyy", -1, myDTEnd)
    
    Dim myDT As Date
    Dim myList As String
    Dim myList1 As String
    con.Execute "DELETE FROM rptSaleHeaders"
    con.Execute "DELETE FROM rptSalePeriods"
    con.Execute "DELETE FROM rptSalePeriodsPrevious"
    
    '293501
    
    myList = "'"
    myDT = DTStart
    Do While myDT < myDTEnd
        myList1 = Format(myDT, "mmm-yy")
        myList = myList & myList1 & "','"
        con.Execute "INSERT INTO rptSalePeriods(Period) Values('" & myList1 & "')"
        myDT = DateAdd("m", 1, myDT)
    Loop
    
    myList = Left(myList, Len(myList) - 2)
    
    con.Execute "INSERT INTO rptSaleHeaders(Period1,Period2,Period3,Period4" & _
     ",Period5,Period6,Period7,Period8,Period9,Period10,Period11,Period12) Values(" & _
     myList & ")"

    myDT = DTPrevStart
    Do While myDT < DTPrevEnd
        myList1 = Format(myDT, "mmm-yy")
        con.Execute "INSERT INTO rptSalePeriodsPrevious(Period) Values('" & myList1 & "')"
        myDT = DateAdd("m", 1, myDT)
    Loop
    
    Dim lRefID As Long
    lRefID = Val(con.Execute("SELECT MAX(EntryID) FROM rptSaleHeaders").Fields(0).Value & "")
    
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    myList = ""
    With rs
        .Open "SELECT * FROM dbo.F_Expenses ('" & DTStart & "' ,'" & DTEnd & "') Order By AccNo,SortID", con, adOpenForwardOnly, adLockReadOnly
        rs1.Open "SELECT AccNo From Accounts Where CAST(Left(AccNo,2) AS INT)>=31", con, adOpenForwardOnly, adLockReadOnly
        Do Until rs1.EOF
            .Filter = "AccNo='" & rs1!AccNo & "'"
            myList = ""
            Do Until .EOF
                myList = myList & Val(!Amount & "") & ","
                .MoveNEXT
            Loop
            myList = Left(myList, Len(myList) - 1)
            con.Execute "INSERT INTO rptSales(RefID,Customer,Period1,Period2,Period3,Period4,Period5,Period6,Period7,Period8,Period9,Period10,Period11,Period12) VALUES(" & lRefID & ",'" & rs1!AccNo & "'," & myList & ")"
            rs1.MoveNEXT
        Loop
        
        .Close
        rs1.Close
        '/////Previous///////
        .Open "SELECT * FROM dbo.F_ExpensesPrevious ('" & DTPrevStart & "' ,'" & DTPrevEnd & "') Order By AccNo,SortID", con, adOpenForwardOnly, adLockReadOnly
        rs1.Open "SELECT AccNo FROM Accounts WHERE CAST(Left(AccNo,2) AS INT)>=31", con, adOpenForwardOnly, adLockReadOnly
        Do Until rs1.EOF
            .Filter = "AccNo='" & rs1!AccNo & "'"
            
            myList = ""
            Do Until .EOF
                myList = myList & Val(!Amount & "") & ","
                .MoveNEXT
            Loop
            
            myList = Left(myList, Len(myList) - 1)
            con.Execute "INSERT INTO rptSalesPrevious(RefID,Customer,Period1,Period2,Period3,Period4,Period5,Period6,Period7,Period8,Period9,Period10,Period11,Period12) VALUES(" & lRefID & ",'" & rs1!AccNo & "'," & myList & ")"
            rs1.MoveNEXT
        Loop
        .Close
        rs1.Close
        '////////////////////
    End With
    
    Set rs = Nothing
    Set rs1 = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub ShowForVenders()

    ForVenders = True
    Me.Caption = "Weekly Issuence Report of Stitchers"
    cmdPrintExpenses.Visible = False
    
    
    
    
    Me.Show
    
End Sub


Private Function ArrangeDataForStitchers(DTStart As Date, DTEnd As Date) As Boolean

    On Error GoTo err
    
    Dim myDT As Date
    Dim myList As String
    Dim myList1 As String
    
    Call StartTrans(con)
    con.Execute "Delete From rptStitcherNames"
    con.Execute "Delete From rptStitcherDetails"
    
    myList = "'"
    
    Dim rs As New ADODB.Recordset
    Dim strServices As String
    Dim iEntries As Integer
    iEntries = 0
    myDT = DTStart
    
    With rs
        .Open "Select Top 15 MakerID,MakerName,WeeklyCapacity,FootballA,FootballB,Handball From Makers Where Active=1 AND MakerID In(Select StitcherID From rptStitcherDoneList) Order By WeeklyCapacity Desc,MakerID Desc", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            VendersDone = True
            ArrangeDataForStitchers = False
            Exit Function
        End If
        iEntries = 0
        Do Until .EOF
            strServices = ""
            If !FootballA Then
                strServices = "Football A" & Chr(13)
            End If
            If !FootballB Then
                strServices = strServices & "Football B" & Chr(13)
            End If
            If !Handball Then
                strServices = strServices & "Handball" & Chr(13)
            End If
            If strServices <> "" Then
                strServices = Left(strServices, Len(strServices) - 1)
            End If
            myList = myList & !MakerName & "','" & strServices & "'," & Val(!WeeklyCapacity & "") & ",'"
            iEntries = iEntries + 1
            .MoveNEXT
        Loop
        .Close
    End With
    For i = iEntries + 1 To 15
        myList = myList & "-','-',0,'"
    Next
    
    myList = Left(myList, Len(myList) - 2)
    con.Execute "Insert Into rptStitcherNames(Stitcher1Name,Stitcher1Services,Stitcher1Capacity" & _
    ",Stitcher2Name,Stitcher2Services,Stitcher2Capacity,Stitcher3Name,Stitcher3Services,Stitcher3Capacity" & _
    ",Stitcher4Name,Stitcher4Services,Stitcher4Capacity,Stitcher5Name,Stitcher5Services,Stitcher5Capacity" & _
    ",Stitcher6Name,Stitcher6Services,Stitcher6Capacity,Stitcher7Name,Stitcher7Services,Stitcher7Capacity" & _
    ",Stitcher8Name,Stitcher8Services,Stitcher8Capacity,Stitcher9Name,Stitcher9Services,Stitcher9Capacity" & _
    ",Stitcher10Name,Stitcher10Services,Stitcher10Capacity,Stitcher11Name,Stitcher11Services,Stitcher11Capacity" & _
    ",Stitcher12Name,Stitcher12Services,Stitcher12Capacity,Stitcher13Name,Stitcher13Services,Stitcher13Capacity" & _
    ",Stitcher14Name,Stitcher14Services,Stitcher14Capacity,Stitcher15Name,Stitcher15Services,Stitcher15Capacity" & _
    ") Values(" & myList & ")"
        
    Dim lRefID As Long
    lRefID = Val(con.Execute("Select Max(EntryID) From rptStitcherNames").Fields(0).Value & "")
    
    myList = ""
    
    myDT = DTStart
    Do While myDT < DTEnd
    
        With rs
            .Open "Select Top 15 Makers.MakerID,DT,Qty From Makers LEFT OUTER JOIN " & _
            "(Select MakerID,DT,Sum(IssQty) As Qty From VMakerIssItems Where DT='" & myDT & _
            "' Group By MakerID,DT) Tab1 On Makers.MakerID=Tab1.MakerID " & _
            "Where  Makers.MakerID In(Select StitcherID From rptStitcherDoneList) Order By WeeklyCapacity Desc,Makers.MakerID Desc", con, adOpenForwardOnly, adLockReadOnly
        
            myList = ""
            iEntries = 0
            Do Until .EOF
                iEntries = iEntries + 1
                myList = myList & Val(!Qty & "") & ","
                .MoveNEXT
            Loop
            .Close
            For i = iEntries + 1 To 15
                myList = myList & 0 & ","
            Next
            myList = Left(myList, Len(myList) - 1)
            con.Execute "Insert Into rptStitcherDetails(RefID,DT,Stitcher1,Stitcher2,Stitcher3" & _
            ",Stitcher4,Stitcher5,Stitcher6,Stitcher7,Stitcher8,Stitcher9,Stitcher10" & _
            ",Stitcher11,Stitcher12,Stitcher13,Stitcher14,Stitcher15) Values(" & _
            lRefID & ",'" & Format(myDT, "dd-MMM-yyyy") & "'," & myList & ")"
            myDT = DateAdd("d", 1, myDT)
            
         
        End With
    Loop
    
    Set rs = Nothing
    con.Execute "Delete From rptStitcherDoneList Where StitcherID In(Select Top 15 MakerID From Makers " & _
     "Where MakerID In(Select StitcherID From rptStitcherDoneList) Order By WeeklyCapacity Desc,MakerID Desc)"
             
    Call con.CommitTrans
    ArrangeDataForStitchers = True
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub Form_Load()
    ForVenders = False
    bExpense = False
    '-----------------------
    Dim myDT As Date
    myDT = Date
    If Month(myDT) >= 7 Then
        myDT = DateSerial(year(myDT), 7, 1)
    Else
        myDT = DateSerial(year(myDT) - 1, 7, 1)
    End If
    DTStart = myDT
    Call DTStart_Change
End Sub
