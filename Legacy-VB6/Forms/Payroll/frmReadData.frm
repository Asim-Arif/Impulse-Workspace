VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmReadData 
   Caption         =   "RFID"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   9315
   WindowState     =   2  'Maximized
   Begin VB.Frame fraMain 
      Height          =   6915
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   9240
      Begin MSWinsockLib.Winsock tcpClient 
         Left            =   360
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7515
         TabIndex        =   17
         Top             =   2655
         Width           =   1605
      End
      Begin VB.Frame fTitle 
         Height          =   570
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   9240
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Barcode Machine Options"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   390
            Index           =   2
            Left            =   60
            TabIndex        =   16
            Top             =   120
            Width           =   9090
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Barcode Machine Options"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Index           =   3
            Left            =   60
            TabIndex        =   15
            Top             =   135
            Width           =   9090
         End
      End
      Begin VB.CommandButton cmdRead 
         Caption         =   "Read From RFID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7515
         TabIndex        =   13
         Top             =   765
         Width           =   1605
      End
      Begin VB.Frame Frame1 
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1140
         Left            =   7470
         TabIndex        =   6
         Top             =   3330
         Width           =   1710
         Begin VB.Label Label1 
            Caption         =   "Total Records :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   10
            Top             =   330
            Width           =   1110
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   1155
            TabIndex        =   9
            Top             =   330
            Width           =   465
         End
         Begin VB.Label Label1 
            Caption         =   "Records Read :"
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   1
            Left            =   30
            TabIndex        =   8
            Top             =   780
            Width           =   1305
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   1155
            TabIndex        =   7
            Top             =   795
            Width           =   465
         End
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "Transfer To Database"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7515
         TabIndex        =   5
         Top             =   1395
         Width           =   1605
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear RFID Buffer"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7515
         TabIndex        =   4
         Top             =   2040
         Width           =   1605
      End
      Begin VB.Frame Frame2 
         Caption         =   "Machine Settings"
         Height          =   690
         Left            =   90
         TabIndex        =   1
         Top             =   6150
         Width           =   3780
         Begin VB.CommandButton cmdSetDate 
            Caption         =   "Set Date/Time"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2250
            TabIndex        =   2
            Top             =   240
            Width           =   1440
         End
         Begin MSComCtl2.DTPicker DT 
            Height          =   330
            Left            =   90
            TabIndex        =   3
            Top             =   240
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   8388608
            CalendarTitleForeColor=   8388608
            CustomFormat    =   "dd-MMM-yyyy hh:mm:ss"
            Format          =   16646147
            CurrentDate     =   38089
         End
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   5835
         Visible         =   0   'False
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   5940
         Top             =   2055
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   2
         DTREnable       =   -1  'True
         NullDiscard     =   -1  'True
         RThreshold      =   1
         RTSEnable       =   -1  'True
         BaudRate        =   19200
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5370
         Left            =   75
         TabIndex        =   12
         Top             =   765
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   9472
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SNo"
            Object.Width           =   2249
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp. ID"
            Object.Width           =   2805
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "In/Out"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   3440
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Time"
            Object.Width           =   3916
         EndProperty
      End
   End
End
Attribute VB_Name = "frmReadData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tot As Long
Dim i As Integer
Dim Arr() As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRead_Click()
    
    Screen.MousePointer = vbHourglass
    Call ReadFromMachine
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdClear_Click()

    Dim InputValue As String
    
    If MsgBox("After Clearing The Buffer You Cant Recover The Entries." & vbNewLine & "Are You Sure You Want To Clear The Machine Buffer", vbYesNo + vbInformation) = vbYes Then
        MSComm1.Output = Chr$(&H2&) & "E" & Chr$(&HD&)
        Do
            InputValue = InputValue & MSComm1.Input
        Loop Until InStr(InputValue, Chr(13))
            
        If InStr(Ans, "A") Then
            MsgBox "Machine Buffer Cleared Successfully !", vbInformation
        'Else
        '    MsgBox "Error Clearing Machine Buffer !", vbInformation
        End If
        LV.ListItems.Clear
    End If
    
End Sub

Private Sub cmdSetDate_Click_OLD()
    Dim strTime As String, Ans As String
    
    If MsgBox("This Will Reset The Date/Time Of Machine." & vbNewLine & "Do You Want To Continue.", vbQuestion + vbYesNo) = vbYes Then
        strTime = Format(DT, "YYYYMMDDhhnnss")
        
        MSComm1.Output = Chr$(&H2&) & "D00" & Chr$(&HD&)
        
        i = 0
        
        Do
            
            Ans = Ans & MSComm1.Input
            i = i + 1
            If i >= 1000 Then Exit Do
        Loop Until InStr(Ans, Chr(13))
        
        
        
        MSComm1.Output = Chr$(&H2) & "S" & strTime & Chr$(&HD&)
    
        Do
            i = i + 1
            Ans = Ans & MSComm1.Input
            If i >= 1000 Then Exit Do
            
        Loop Until InStr(Ans, Chr(13))
        
        If InStr(Ans, "A") Then
            MsgBox "Date/Time Has Been Changed", vbInformation
        Else
            MsgBox "Error Setting Date/Time !", vbInformation
        End If
        
    End If
    
End Sub

Private Sub Command1_Click()
    Dim count As Integer, a As String
    
    MSComm1.Output = Chr$(&H2&) & "D00" & Chr$(&HD&)
    Do
        a = a & MSComm1.Input
        count = count + 1
        If count > 60 Then Exit Do
    Loop Until InStr(a, Chr(13))
    
    If a = "" Then
        MsgBox "Could Not Connect To Machine !", vbInformation
        Unload Me
    End If

End Sub

Private Sub cmdSetDate_Click()
     
'     Screen.MousePointer = vbHourglass
'
'        Dim Rec As New ADODB.Recordset
'        With Rec
'            .Open "SELECT CompName FROM Computers", con, adOpenForwardOnly, adLockReadOnly
'
'            For i = 1 To .RecordCount
'                If tcpClient.State <> sckClosed Then tcpClient.Close
'                tcpClient.RemoteHost = !CompName & ""
'                tcpClient.RemotePort = 1012
'                tcpClient.Connect
'            Next
'
'        End With
'
'    Screen.MousePointer = vbDefault
    Call SetMachineTime(DT.value)
End Sub



Private Sub cmdTransfer_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    Dim rs As New ADODB.Recordset
    Call StartTrans(con)
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            rs.Open "SELECT EmpId FROM Employees WHERE EmployeeId=" & Val(.SubItems(1)) & "", con, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount > 0 Then
                'Call TakeAtt( rs.Fields(0), CDate(.ListSubItems(3).Tag), CDate(.SubItems(4)))
                'Call TakeAttendance(rs!EmpId & "", CDate(.ListSubItems(3).Tag), CDate(.SubItems(4)), True )
                Call TakeAtt(rs!EmpId & "", CDate(.ListSubItems(3).Tag), CDate(.SubItems(4)))
            End If
            rs.Close
        End With
    Next
    
    MSComm1.Output = Chr$(&H2&) & "E" & Chr$(&HD&)
    Do
        InputValue = inpurtvalue & MSComm1.Input
    Loop Until InStr(InputValue, Chr(13))
    LV.ListItems.Clear
    con.CommitTrans
    
    Exit Sub

err:
    MsgBox err.Description, vbInformation
    con.RollbackTrans
    
End Sub

Private Function TakeAtt(EmployeeID As String, AttDate As Date, AttTime As Date) As Boolean

    Dim rs As New ADODB.Recordset
    Dim strEmpID As String, strDeptID As String, bExemptSettings As Boolean
    Dim bLunchTimeScanning As Boolean
    Dim strLunchInTime As String, strLunchOutTime As String, iEmpAttCatID As Integer
    Dim iTotalEntries As Integer, bActive As Boolean
    Dim bOverTimePaid As Boolean
    
    With rs
        .Open "Select DeptID,EmpID,ExemptSettings,Active From Employees Where EmpID='" & EmployeeID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            strEmpID = !EmpId & ""
            strDeptID = !DeptID & ""
            bExemptSettings = !ExemptSettings
            bActive = !Active
            .Close
        End If
        
       
        If bExemptSettings Then
            .Open "Select LunchInTime,LunchOutTime,EmpAttCat_RefID From EmpSettings Where EmpID='" & strEmpID & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select LunchInTime,LunchOutTime,EmpAttCat_RefID From GeneralSettings Where DeptID='" & strDeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        
        strLunchInTime = !LunchInTime & ""
        strLunchOutTime = !LunchOutTime & ""
        iEmpAttCatID = Val(!EmpAttCat_RefID & "")
        
        .Close
        
        .Open "Select LunchTimeScanning,OverTimePaid From EmployeeAttendanceCatagories Where EntryID=" & iEmpAttCatID, con, adOpenForwardOnly, adLockReadOnly
        bLunchTimeScanning = !LunchTimeScanning
        bOverTimePaid = !OverTimePaid
        .Close
        
        .Open "Select TotalEntries From VEmpDailyEntries Where EmpID='" & strEmpID & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            iTotalEntries = 0
        Else
            iTotalEntries = .Fields(0) & ""
        End If
        .Close
        
    End With
    Set rs = Nothing
    
    If bLunchTimeScanning = False Then
        If iTotalEntries = 1 Then
            If Format(AttDate, "ddd") = "Fri" Then
                Call TakeAttendance(strEmpID, AttDate, #1:00:00 PM#, True)
                DoEvents
                Call TakeAttendance(strEmpID, AttDate, #2:00:00 PM#, True)
                DoEvents
'                Call TakeAttendance(strEmpID, AttDate, CDate(strLunchInTime), True)
'                DoEvents
'                Call TakeAttendance(strEmpID, AttDate, CDate(strLunchOutTime), True)
'                DoEvents
            Else
                Call TakeAttendance(strEmpID, AttDate, CDate(strLunchInTime), True)
                DoEvents
                Call TakeAttendance(strEmpID, AttDate, CDate(strLunchOutTime), True)
                DoEvents
            End If
        End If
    End If
    TakeAtt = TakeAttendance(strEmpID, AttDate, AttTime, True)
    DoEvents
    
End Function

Private Sub Form_Load()

    
    Dim rs As New ADODB.Recordset
    Dim iPortNo As Integer
    With rs
        .Open "SELECT PortNo FROM MachinePortNo", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            iPortNo = Val(!PortNo & "")
        Else
            iPortNo = 3
        End If
        .Close
    End With
    Set rs = Nothing
    MSComm1.CommPort = iPortNo
    
    MSComm1.PortOpen = True
    i = 0

    DT = Now


End Sub

Public Sub ReadFromMachine()
On Error GoTo err
    LV.ListItems.Clear
    Dim a As String, count As Integer
    
    MSComm1.Output = Chr$(&H2&) & "D00" & Chr$(&HD&)
    Sleep (50)
    Do
        a = a & MSComm1.Input
        count = count + 1
        If count > 1000 Then Exit Do
    Loop Until InStr(a, Chr(13))
    
    If a = "" Then
        MsgBox "Could Not Connect To Machine !", vbInformation
        Exit Sub
    End If
    
    count = 0
    MSComm1.Output = Chr$(&H2&) & "N" & Chr$(&HD&)
    Sleep (50)
    a = ""
    Do
        a = a & MSComm1.Input
        count = count + 1
        If count > 60 Then
            Exit Do
        End If
    Loop Until InStr(a, Chr(13))
    
    Tot = Val("&H" & Mid(a, 3, 4) & "&")
    
    If Tot = 0 Then
        MsgBox "No Record Found In The Machine Database !", vbInformation
        Exit Sub
    End If
    
    PB.Max = Tot
    lblTotal(0).Caption = Tot
    PB.Visible = True
    lblTotal(0).Refresh
    For i = 0 To Tot - 1
        
        MSComm1.Output = Chr$(&H2&) & "G" & String(4 - Len(Hex(i)), "0") & Hex(i) & Chr$(&HD&)
         'MSComm1.Output = Chr$(&H2&) & "G0001" & Chr$(&HD&)
        
        a = ""
        Do
            a = a & MSComm1.Input
                
        Loop Until InStr(a, Chr(13))
        
        Call FillList(a)
        PB.value = PB.value + 1
        lblTotal(1) = PB.value
        lblTotal(1).Refresh
    Next i
    PB.value = 0
    PB.Visible = False
    'MSComm1.Output = Chr$(&H2&) & "G0001" & Chr$(&HD&)
Exit Sub

err:

MsgBox "Cannot Read from Machine due to following error : & vbnewline & err.Description"
End Sub


'Public Sub ReadFromMachine()
'
'    LV.ListItems.Clear
'    Dim a As String, count As Integer
'
'    MSComm1.Output = Chr$(&H2&) & "D00" & Chr$(&HD&)
'
'    Do
'        a = a & MSComm1.Input
'        count = count + 1
'        If count > 1000 Then Exit Do
'    Loop Until InStr(a, Chr(13))
'
'    If a = "" Then
'        MsgBox "Could Not Connect To Machine !", vbInformation
'        Exit Sub
'    End If
'
'    count = 0
'    MSComm1.Output = Chr$(&H2&) & "N" & Chr$(&HD&)
'
'    a = ""
'    Do
'        a = a & MSComm1.Input
'        count = count + 1
'        If count > 60 Then Exit Do
'    Loop Until InStr(a, Chr(13))
'
'    Tot = Val("&H" & Mid(a, 3, 4) & "&")
'
'    If Tot = 0 Then
'        MsgBox "No Record Found In The Machine Database !", vbInformation
'        Exit Sub
'    End If
'
'    PB.Max = Tot
'    lblTotal(0).Caption = Tot
'    PB.Visible = True
'    lblTotal(0).Refresh
'    For i = 0 To Tot - 1
'
'        MSComm1.Output = Chr$(&H2&) & "G" & String(4 - Len(Hex(i)), "0") & Hex(i) & Chr$(&HD&)
'         'MSComm1.Output = Chr$(&H2&) & "G0001" & Chr$(&HD&)
'
'        a = ""
'        Do
'            a = a & MSComm1.Input
'
'        Loop Until InStr(a, Chr(13))
'
'        Call FillList(a)
'        PB.value = PB.value + 1
'        lblTotal(1) = PB.value
'        lblTotal(1).Refresh
'    Next i
'    PB.value = 0
'    PB.Visible = False
'    'MSComm1.Output = Chr$(&H2&) & "G0001" & Chr$(&HD&)
'
'
'
'
'End Sub


Private Sub FillList(str As String)

    Arr = Split(str, ",")
    Dim ITM As ListItem
    Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
    
    ITM.SubItems(1) = Replace(Arr(1), ":", "-")
    ITM.SubItems(2) = Arr(2)
    ITM.SubItems(3) = Format(CDate(Arr(3)), "ddd dd-MM-yy")
    ITM.ListSubItems(3).Tag = Format(CDate(Arr(3)), "dd-MMM-yy")
    ITM.SubItems(4) = Left(Arr(4), Len(Arr(4)) - 3)
    
End Sub

Private Sub Form_Resize()
    'fraMain.Move (Me.Width - fraMain.Width) / 2, (Me.Height - fraMain.Height - frmMain.Toolbar1.Height) / 2
End Sub

Private Sub TakeAttOld(EmpId As String, AttDate As Date, AttTime As Date)

'''''''''''   This Will Is The Employee Exists Or Not    '''''''''''
If Val(con.Execute("SELECT COUNT(*) FROM Employees WHERE EmpID='" & EmpId & "'").Fields(0)) = 0 Then
    MsgBox "The Following Employee " & EmpId & " Does Not Exists !", vbInformation
    Exit Sub
End If



Dim rec As New ADODB.Recordset, CurrTime As String, Hrs As Double, Sql As String, Cond As String
Dim TimeDiff As Integer
Dim LateHours As Double
Dim DeptID As String

Cond = "WHERE EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
CurrTime = Format(AttTime, "HH:NN")

TimeDiff = Val(con.Execute("SELECT TimeDiff FROM TimeSettings").Fields(0) & "")

With rec
    .Open "SELECT * FROM EmpTimes WHERE EntryID=(SELECT MAX(EntryID) FROM EmpTimes " & Cond & ")", con, adOpenStatic, adLockReadOnly
    
    Sql = "INSERT INTO EmpTimes(EmpID,DT,InTime) VALUES('" & EmpId _
        & "','" & AttDate & "','" & CurrTime & "')"
    If .RecordCount > 0 Then
        If !OutTime & "" = "" Then
            Hrs = DateDiff("n", CDate(!InTime), CDate(CurrTime))
            
            If Hrs < TimeDiff Then Exit Sub
            Hrs = Round(Hrs / 60, 2)
            Sql = "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & " WHERE EntryID=" & Val(!EntryID & "") & ""
        Else
            Hrs = DateDiff("n", CDate(!OutTime), CDate(CurrTime))
            If Hrs < TimeDiff Then Exit Sub
        End If
    Else
        .Close
        
        .Open "select * from  Departments where deptid='" & Mid(EmpId, 1, 5) & "'", con, adOpenForwardOnly, adLockReadOnly
        LateHours = DateDiff("n", CDate(!InTime), CDate(CurrTime))
        LateHours = Round(LateHours / 60, 2)
        If LateHours < 0.25 Then
            LateHours = 0
        End If
        con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance,LateHours) VALUES('" & EmpId & "','" & AttDate & "',1," & LateHours & ")"
        
    End If

    con.Execute Sql
        
End With

End Sub

Private Sub tcpClient_Connect()

    frmWait.Show
    tcpClient.SendData "SetTime" & DT.value

End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
    
    Dim str As String
    tcpClient.GetData str
    If str = "Completed" Then
        MsgBox "Machine Date and Time Set Successfully."
    Else
        MsgBox "Error In Setting Machine Date and Time."
    End If
    
End Sub

Private Function SetMachineTime(DateToSet As Date) As Boolean

    On Error GoTo err
    
    strTime = Format(DateToSet, "YYYYMMDDhhnnss")
    MSComm1.Output = Chr$(&H2&) & "D00" & Chr$(&HD&)

    i = 0
    
    Do
        Ans = Ans & MSComm1.Input
        i = i + 1
        If i >= 1000 Then Exit Do
    Loop Until InStr(Ans, Chr(13))
    
    MSComm1.Output = Chr$(&H2) & "S" & strTime & Chr$(&HD&)

    Do
        i = i + 1
        Ans = Ans & MSComm1.Input
        If i >= 1000 Then
            Exit Do
        End If
    Loop Until InStr(Ans, Chr(13))
            
    If InStr(Ans, "N04") Then
        SetMachineTime = False
    Else
        SetMachineTime = True
    End If
    
    Exit Function
    
err:
    SetMachineTime = False
    
End Function
