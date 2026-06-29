Attribute VB_Name = "mainModule"
Option Explicit
'-----------------------------------------------------
Public iReadyFinishProcessID As Integer
Public iSemiFinishProcessID As Integer, iFinishedPackagingProcessID As Integer, iFilingProcessID As Integer, iFiling2ProcessID As Integer
Public iMakerPOProess_I As Integer, iMakerPOProess_II As Integer, iMakerPOProess_III As Integer, iMakerPOProess_IV As Integer, iMakerPOProess_V As Integer
Public iMakerPOProess_VI As Integer, iMakerPOProess_VII As Integer, iMakerPOProess_VIII As Integer

Public strEmpShortTermAccNo As String, strEmpLongTermAccNo As String
Public iFactoryMakerID As Integer, iInsulationValuesProcessID As Integer, strFactoryMakers As String

Public iTemper_Inspection_Process_ID As Integer, iFirst_Inspection_Process_ID As Integer

Public bPrintWhenSaved As Boolean, bVoucherViewOnlyWhenSaved As Boolean
Public bAutoProforma As Boolean

Public rptApp As New CRAXDDRT.Application
Public strConnectionStringrpt As String

Public Export_Package As Boolean

Public RptPRPath As String

Public FormsMaxWidth As Long

Public OurOwnForm As Boolean
Public AwayTime As Date

Public ServerDate As Date
Public ServerTime As Date
Public con As New ADODB.Connection, con_ServerSide As New ADODB.Connection
Private TestCon As New ADODB.Connection
Public sString As String
Public sName As String
Global DatabasePath As String
Public strExpNo As String

Public FormProcessID As Long    'For Save and New Of Issue and Receive...

Public CurrentUserName As String
Public strComputerName As String
Public bUrduSupported As Boolean

Public UserID As Long
Public HicoVisible As Boolean
'Public UserReqAuth As Integer 'Whether Or Not User Required Authorization ON Issue and Receive.


Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const WS_DLGFRAME = &H400000
Public Const WS_THICKFRAME = &H40000
Public Const WS_BORDER = &H800000

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function CreateMenu Lib "User32" () As Long
Public Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenu Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemInfo Lib "User32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal BOOL As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const flags = SWP_NOSIZE Or SWP_NOMOVE

Public Const Err_Trans As String = "Cannot start more transactions on this session."
Public BackPic As New StdPicture

Public Const GRID_BK_COLOR_ODD As Long = &HD7DDD5       '&HF8F3FE       '&HE3ECF4         ' &HE9F9FA   '&HE3F4E3
Public Const GRID_BK_COLOR_EVEN As Long = vbWhite '&HE3F4E3
Public VerHigherThanWin2000 As Boolean

Public rsUserSettings As ADODB.Recordset

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function SetCursorPos Lib "User32" (ByVal x As Long, ByVal y As Long) As Long

'Public ttp As New cToolTip

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

Public Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Const lForeColor_RepairLot = 11414324 '13158400

Public iProductionType As Integer

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_HIDE As Long = 0
Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWMINIMIZED As Long = 2

Sub Main()
      
    'Very Important... Set Company Name
'    strCompany = "PAKSMITH GROUP"
    'strCompany = "Weldon Instruments"
'    strCompany = "Weldon Industries"
'   strCompany = "All Seas"
    'strCompany = "Banzai"
    strCompany = "IAA"
    'strCompany = "Kami"
    'strCompany = "Scard"
    
    If strCompany = "IAA" Or strCompany = "Kami" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "Towne" Then
        iProductionType = 0
    ElseIf strCompany = "QEL" Then
        iProductionType = 1
    ElseIf strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
        iProductionType = 2
    End If
    
'    If strCompany = "Tecno" Then
'        bAutoProforma = True
'    Else
'        bAutoProforma = False
'    End If
    bAutoProforma = True
    ' we need to call InitCommonControls before we
    ' can use XP visual styles.  Here I'm using
    ' InitCommonControlsEx, which is the extended
    ' version provided in v4.72 upwards (you need
    ' v6.00 or higher to get XP styles)
    On Error Resume Next
    ' this will fail if Comctl not available
    '  - unlikely now though!
    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
   
    ' now start the application
    On Error GoTo 0
        
    Dim retValue As Integer
    retValue = TryToConnect
    
    If retValue = 1 Then
        'Call getAccNames
        RptPRPath = Left(App.Path, Len(App.Path) - 12) & "Reports\PayRoll"
        'bPrintWhenSaved = True
        If Command$ = "FromVBIDE" Then
            RptPath = Left(App.Path, Len(App.Path) - 13)
            If strCompany = "IAA" Then
                RptPath = RptPath & "\German Health Care Reports\Reports"
            ElseIf strCompany = "Weldon Instruments" Then
                RptPath = RptPath & "\Weldon Instruments Reports\Reports"
            ElseIf strCompany = "Weldon Industries" Then
                RptPath = RptPath & "\Weldon Industries Reports\Reports"
            ElseIf strCompany = "All Seas" Then
                RptPath = RptPath & "\All Seas Reports\Reports"
            ElseIf strCompany = "Sunlike" Then
                RptPath = RptPath & "\Sunlike Reports\Reports"
                bPrintWhenSaved = False
            ElseIf strCompany = "QEL" Then
                RptPath = RptPath & "\QEL Reports\Reports"
            ElseIf strCompany = "Tecno" Then
                RptPath = RptPath & "\Tecno Reports\Reports"
            ElseIf strCompany = "Instrumed" Then
                RptPath = RptPath & "\Instrumed Reports\Reports"
            ElseIf strCompany = "Dr-Frgz" Then
                If strSubCompany = "Electro" Then
                    RptPath = RptPath & "\Dr-Frgz Reports\Reports"
                ElseIf strSubCompany = "Intl" Then
                    RptPath = RptPath & "\Frgz-Intl Reports\Reports"
                End If
            ElseIf strCompany = "Banzai" Then
                RptPath = RptPath & "\Banzai Reports\Reports"
                bPrintWhenSaved = False
            ElseIf strCompany = "IAA" Then
                RptPath = RptPath & "\Wrangler Reports\Reports"
                bPrintWhenSaved = False
            ElseIf strCompany = "Scard" Then
                RptPath = RptPath & "\Scard Reports\Reports"
                bPrintWhenSaved = False
            ElseIf strCompany = "Towne" Then
                RptPath = RptPath & "\Towne Reports\Reports"
                bPrintWhenSaved = False
            ElseIf strCompany = "Kami" Then
                RptPath = RptPath & "\Kami Reports\Reports"
                bPrintWhenSaved = False
            End If
        Else
            RptPath = App.Path & "\Reports"
        End If
        
        bVoucherViewOnlyWhenSaved = True
        
        rptPayrollPath = RptPath & "\Payroll"
        
        iReadyFinishProcessID = Val(getGeneralDataValue("ReadyFinishProcessID"))
        iSemiFinishProcessID = Val(getGeneralDataValue("SemiFinishProcessID"))
        iFilingProcessID = Val(getGeneralDataValue("FilingProcessID"))
        iFiling2ProcessID = Val(getGeneralDataValue("Filing2ProcessID"))
        
        iMakerPOProess_I = Val(getGeneralDataValue("Maker_PO_Process_I"))
        iMakerPOProess_II = Val(getGeneralDataValue("Maker_PO_Process_II"))
        iMakerPOProess_III = Val(getGeneralDataValue("Maker_PO_Process_III"))
        iMakerPOProess_IV = Val(getGeneralDataValue("Maker_PO_Process_IV"))
        iMakerPOProess_V = Val(getGeneralDataValue("Maker_PO_Process_V"))
        
        iMakerPOProess_VI = Val(getGeneralDataValue("Maker_PO_Process_VI"))
        iMakerPOProess_VII = Val(getGeneralDataValue("Maker_PO_Process_VII"))
        iMakerPOProess_VIII = Val(getGeneralDataValue("Maker_PO_Process_VIII"))
        
        iTemper_Inspection_Process_ID = Val(getGeneralDataValue("Temper_Inspection_ProcessID"))
        iFirst_Inspection_Process_ID = Val(getGeneralDataValue("First_Inspection_ProcessID"))
        
        
        CashInHand = getGeneralDataValue("CashInHandParent")
        If CashInHand = "" Then
            CashInHand = "15-001"
        End If
        '------------------------------------------------------------------------
        iFactoryMakerID = Val(getGeneralDataValue("FactoryMaker"))
        
        Call getFactoryMakers
        
        iInsulationValuesProcessID = Val(getGeneralDataValue("ProcessIDForInsulationLeakage"))
        
        strEmpShortTermAccNo = getGeneralDataValue("EmpShortTermAccNo")
        strEmpLongTermAccNo = getGeneralDataValue("EmpLongTermAccNo")
        
        
        
        iFinishedPackagingProcessID = 84
        
        
'        MainForm.Show
'
'        FrmLogin.Show 1
        frmSplash.Show
        'MainForm.Show
        
    ElseIf retValue = 2 Then
        'Do Nothing
    Else
        End
    End If
    
End Sub

Private Sub getFactoryMakers()

    Dim rs As New ADODB.Recordset
    strFactoryMakers = "0,"
    With rs
        .Open "SELECT * FROM FactoryMakers", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            strFactoryMakers = strFactoryMakers & Val(!MakerID & "") & ","
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    strFactoryMakers = Left(strFactoryMakers, Len(strFactoryMakers) - 1)
    
End Sub

Private Function TryToConnect() As Integer

    On Error GoTo err
    'con.CursorLocation = adUseServer
    con.CursorLocation = adUseClient
    con.CommandTimeout = 120
    
    con_ServerSide.CursorLocation = adUseServer
    con_ServerSide.CommandTimeout = 120
    If strCompany = "PAKSMITH GROUP" Then
        con.Open "DSN=DSN_PakSmith_AWM"
        con_ServerSide.Open "DSN=DSN_PakSmith_AWM"
        strConnectionStringrpt = "DSN=DSN_PakSmith_AWM"
    ElseIf strCompany = "Weldon Instruments" Then
        con.Open "DSN=DSN_WeldonInstr_AWM"
        con_ServerSide.Open "DSN=DSN_WeldonInstr_AWM"
        strConnectionStringrpt = "DSN=DSN_WeldonInstr_AWM"
    ElseIf strCompany = "Weldon Industries" Then
        con.Open "DSN=DSN_WeldonInd_AWM"
        con_ServerSide.Open "DSN=DSN_WeldonInd_AWM"
        strConnectionStringrpt = "DSN=DSN_WeldonInd_AWM"
    'ElseIf strCompany = "All Seas" Then
    Else
        If Command$ = "FromVBIDE" Then
            If strCompany = "Dr-Frgz" Then
                con.Open "DSN=DSN_SMBI_AWM;;UID=sa;;PWD=Abc@123"
                con_ServerSide.Open "DSN=DSN_SMBI_AWM;;UID=sa;;PWD=Abc@123"
                strConnectionStringrpt = "DSN=DSN_SMBI_AWM;;UID=sa;;PWD=Abc@123"
            Else
                con.Open "DSN=DSN_SMBI_AWM;;UID=sa;;PWD=awm"
                con_ServerSide.Open "DSN=DSN_SMBI_AWM;;UID=sa;;PWD=awm"
                strConnectionStringrpt = "DSN=DSN_SMBI_AWM;;UID=sa;;PWD=awm"
            End If
        Else
            If strCompany = "Dr-Frgz" Then
                If strSubCompany = "Electro" Then
'                    strConnectionStringrpt = "DSN=DSN_SMBI_AWM;;Server=Server"
'                    con.Open "DSN=DSN_SMBI_AWM;Server=Server"
'                    con_ServerSide.Open "DSN=DSN_SMBI_AWM;Server=Server"
                    con.Open "DSN=DSN_SMBI_AWM"
                    con_ServerSide.Open "DSN=DSN_SMBI_AWM"
                    strConnectionStringrpt = "DSN=DSN_SMBI_AWM"
                ElseIf strSubCompany = "Intl" Then
'                    strConnectionStringrpt = "DSN=DSN_SMBI_AWM;;Server=DC2"
'                    con.Open "DSN=DSN_SMBI_AWM;Server=DC2"
'                    con_ServerSide.Open "DSN=DSN_SMBI_AWM;Server=DC2"
                    con.Open "DSN=DSN_SMBI_AWM"
                    con_ServerSide.Open "DSN=DSN_SMBI_AWM"
                    strConnectionStringrpt = "DSN=DSN_SMBI_AWM"
                End If
            Else
                If strCompany = "IAA" Then
                    con.Open "DSN=DSN_SMBI_AWM;;UID=sa;;PWD=awm"
                    con_ServerSide.Open "DSN=DSN_SMBI_AWM;;UID=sa;;PWD=awm"
                    strConnectionStringrpt = "DSN=DSN_SMBI_AWM;;UID=sa;;PWD=awm"
                Else
                    strConnectionStringrpt = "DSN=DSN_SMBI_AWM"
                    con.Open "DSN=DSN_SMBI_AWM"
                    con_ServerSide.Open "DSN=DSN_SMBI_AWM"
                End If
            End If
        End If
    End If
    TryToConnect = 1
'    con.Execute "set Dateformat 'dmy'"
    
    
    Exit Function
err:
    If err.Number = -2147467259 Then
        Call CheckErrorStatus(err.Number, err.Description)
        TryToConnect = 2
    Else
        MsgBox err.Description
    End If
    
End Function

Sub CheckErrorStatus(errNumber As Long, errDesc As String)
    
'    If Command$ <> "FromVBIDE" Then
'        If strSubCompany = "Intl" Then
'            sName = "Server2"
'        ElseIf strSubCompany = "Electro" Then
'            sName = "Server"
'        Else
'            sName = InputBox("Please Enter Server Name : ")
'        End If
'    Else
'            sName = InputBox("Please Enter Server Name : ")
'    End If

    sName = InputBox("Please Enter Server Name : ")
    If sName = "" Then
        MsgBox "Invalid Server Name." & vbNewLine & "Quiting!", vbCritical
        End
    End If
            
    If Right$(err.Description, 14) = "access denied." Then
        Call ConfigDSN(sName)
    Else
        Call CreateDSN(sName)
    End If
    
    On Error GoTo 0
    Call Main
        
    
    MainForm.Show
    
End Sub

Public Sub setDateRange(ByRef DT As DTPicker)

    If DateDiff("d", Format(DT, "dd/MMM/yyyy"), Format(YearStartDate, "dd/MMM/yyyy")) > 0 Then
        DT = YearStartDate
    ElseIf DateDiff("d", Format(DT, "dd/MMM/yyyy"), Format(YearLastDate, "dd/MMM/yyyy")) < 0 Then
        DT = YearLastDate
    End If
End Sub

Public Function SaveQtyTotal(strOrdNo As String) As String
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Tot From VrptOrders Where OrderNo='" & strOrdNo & "'", con, adOpenForwardOnly, adLockReadOnly
    End With
    Set rs = Nothing
End Function

Public Sub GetWinVersionInfo()

    On Error GoTo err
    
    VerHigherThanWin2000 = False
    
    Dim OSV As OSVERSIONINFO
    OSV.dwOSVersionInfoSize = Len(OSV)
    
    Call GetVersionEx(OSV)
    
    If OSV.dwMajorVersion >= 5 Then
        If OSV.dwMinorVersion >= 1 Then
            VerHigherThanWin2000 = True
        End If
    End If
        
    
    Exit Sub
err:
    Exit Sub
End Sub

Public Sub SetToolTip(lHWnd As Long, strToolTip As String, Optional BackColor As Long = vbRed, Optional ForeColor As Long = vbWhite, Optional DelayTime As Long = 300, Optional VisibleTime As Long = 1000, Optional TipWidth As Long = 20, Optional Font As StdFont)

    On Error GoTo err
    
    'Set tooltip properties
'    With ttp
'        .BkColor = BackColor
'        .TxtColor = ForeColor
'        .DelayTime = DelayTime
'        .VisibleTime = VisibleTime
'        .TipWidth = TipWidth
'        .Style = ttStyleBalloon
'        'Set .Font = oFont
'        'Set tooltip for control using controls hwnd
'        .SetToolTipObj lHwnd, strToolTip
'
'    End With
    
    'Set ttp = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub PrintReport(strSel As String, ArgSourceRpt As CRAXDDRT.Report, Optional strQueryString As String, Optional bDiscardSaved As Boolean = True)

    Dim SrcReport As CRAXDDRT.Report
    Set SrcReport = ArgSourceRpt
    
    On Error Resume Next
    SrcReport.Database.Tables(1).ConnectBufferString = strConnectionStringrpt
    
'    On Error Resume Next
'    SrcReport.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=vision;;Password=visionasim"
'    On Error GoTo 0
    strSel = strSel & " "
    Screen.MousePointer = vbHourglass
    
    If bDiscardSaved Then
        SrcReport.DiscardSavedData
    End If
    If strSel <> "" Then SrcReport.RecordSelectionFormula = Left(strSel, Len(strSel) - 1)
    SrcReport.RecordSelectionFormula = strSel '& "'"
    If strQueryString <> "" Then
        SrcReport.SQLQueryString = strQueryString
    End If
    
    SrcReport.PrintOut False
    Screen.MousePointer = vbDefault
    
    
End Sub

Public Function GetSingleStringValue(strFieldName As String, strTableName As String, Optional strCondition As String) As String

    On Error GoTo err
    
    Dim strValue As String
    Dim rs As New ADODB.Recordset
    strValue = ""
    With rs
        .Open "SELECT " & strFieldName & " FROM " & strTableName & " " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            strValue = .Fields(0) & ""
        End If
        .Close
    End With
    Set rs = Nothing
    GetSingleStringValue = strValue
    
    Exit Function
err:
    MsgBox err.Description
    
End Function


Public Function GetSingleLongValue(strFieldName As String, strTableName As String, Optional strCondition As String) As Long

    On Error GoTo err
    
    Dim lValue As Long
    Dim rs As New ADODB.Recordset
    lValue = 0
    With rs
        .Open "SELECT " & strFieldName & " FROM " & strTableName & " " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            lValue = Val(.Fields(0) & "")
        End If
        .Close
    End With
    Set rs = Nothing
    GetSingleLongValue = lValue
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Public Function GetSingleDoubleValue(strFieldName As String, strTableName As String, Optional strCondition As String) As Double

    On Error GoTo err
    
    Dim dValue As Double
    Dim rs As New ADODB.Recordset
    dValue = 0
    With rs
        .Open "SELECT " & strFieldName & " FROM " & strTableName & " " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            dValue = Val(.Fields(0) & "")
        End If
        .Close
    End With
    Set rs = Nothing
    GetSingleDoubleValue = dValue
    
    Exit Function
err:
    MsgBox err.Description
End Function

Public Function GetSingleBooleanValue(strFieldName As String, strTableName As String, Optional strCondition As String) As Boolean

    On Error GoTo err
    
    Dim bValue As Boolean
    Dim rs As New ADODB.Recordset
    bValue = False
    With rs
        .Open "SELECT " & strFieldName & " FROM " & strTableName & " " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            If IsNull(.Fields(0)) Then
                bValue = False
            Else
                bValue = .Fields(0)
            End If
        End If
        .Close
    End With
    Set rs = Nothing
    GetSingleBooleanValue = bValue
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Public Function GetSingleDateValue(strFieldName As String, strTableName As String, Optional strCondition As String) As Date

    On Error GoTo err
    
    Dim myDT As Date
    Dim rs As New ADODB.Recordset
    myDT = #1/1/1900#
    With rs
        .Open "SELECT " & strFieldName & " FROM " & strTableName & " " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False And IsNull(.Fields(0)) = False Then
            myDT = .Fields(0) & ""
        End If
        .Close
    End With
    Set rs = Nothing
    GetSingleDateValue = myDT
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Public Sub RollBackTransactions()
    On Error Resume Next
    con.RollbackTrans
End Sub

Public Sub ImportCompanyCatalog()
     
    On Error Resume Next
    Dim strFileName As String
    Dim strFileName_Tag As String
    MainForm.CDExcel.ShowSave
    
    If Not (err.Number = 32755) Then
        strFileName = MainForm.CDExcel.FileName
        strFileName_Tag = strFileName
    Else
        Exit Sub
    End If
    On Error GoTo err
    
    Dim Con_Excel As ADODB.Connection
    Set Con_Excel = New ADODB.Connection
    Con_Excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strFileName & ";" & "Extended Properties=""Excel 8.0;HDR=YES;"""
    Dim sTableName As String
    sTableName = "[sheet1$]"
    
    Dim strItemID As String, strItemName As String, strUnit As String, strItemGroup As String, strCategory As String, strItemGroupMain As String
    Dim lGroupID As Long, lCatID As Long, lGroupIDMain As Long
    
    Dim lCount As Long
    Call StartTrans(con)
    con.Execute "DELETE FROM Import_Excel_CustomerOrder"
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM " & sTableName, Con_Excel, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
        
            strItemID = .Fields(0).Value & ""
            strItemName = .Fields(1).Value & ""
            strUnit = .Fields(2).Value & ""
            strItemGroup = .Fields(3).Value & ""
            strCategory = .Fields(4).Value & ""
            strItemGroupMain = .Fields(5).Value & ""
            
            lGroupID = GetSingleLongValue("ID", "ItemGroups", " WHERE Description='" & strItemGroup & "'")
            lCatID = GetSingleLongValue("CatID", "ItemCatagories", " WHERE Description='" & strCategory & "'")
            lGroupIDMain = GetSingleLongValue("MainGroupID", "ItemGroupsMain", " WHERE MainGroupName='" & strItemGroupMain & "'")
            
            
            'Now Validate...
            If strItemID <> "" And lGroupID > 0 And lCatID > 0 Then
                lCount = GetSingleLongValue("COUNT(*)", "Items", " WHERE ItemID='" & strItemID & "'")
                If lCount = 0 Then
                    con.Execute "INSERT INTO Items(ItemID,ItemName,Unit,CatID,GroupID,MainGroupID,Type,ItemSize,SizeUnit,ItemType) VALUES('" & strItemID & "','" & _
                     strItemName & "','" & strUnit & "'," & lCatID & "," & lGroupID & "," & lGroupIDMain & ",'','','',0)"
                End If
            End If
            
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    Set Con_Excel = Nothing
    con.CommitTrans
    MsgBox "Done.", vbInformation
    'Call ShowImportedData
    
    Exit Sub
err:
    MsgBox err.Description

End Sub
