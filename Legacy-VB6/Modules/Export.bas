Attribute VB_Name = "Export"
Public Function GetProformaNo() As Long

    On Error GoTo err
    
    Dim ProformaNo As Long
    Dim stNo As String
'    Dim rs As New ADODB.Recordset
'    With rs
'        .Open "Select Max(Cast(SubString(PInvoice,5,CharIndex('-',PInvoice,5)-5) as Int)) From FPInvoice Where CompanyRefID=" & iCompanyID, con, adOpenForwardOnly, adLockReadOnly
'        ProformaNo = Val(.Fields(0).value & "")
'        .Close
'        .Open "Select InvoicePrefix Where ComopanyID=" & iCompanyID, con, adOpenForwardOnly, adLockReadOnly
'        .Close
'    End With
'    Set rs = Nothing
    
    'ProformaNo = Val(con.Execute("Select Max(Cast(SubString(PInvoice,6,CharIndex('-',PInvoice,6)-6) as Int)) From FPInvoice ").Fields(0).value & "")
    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Instrumed" Or strCompany = "Towne" Then
        ProformaNo = Val(con.Execute("SELECT MAX(CAST(SUBSTRING(PInvoice,CHARINDEX('-',PInvoice)+1,50) AS INT)) FROM FPInvoice WHERE ISNUMERIC(SUBSTRING(PInvoice,CHARINDEX('-',PInvoice)+1,50))=1").Fields(0).Value & "")
    Else
        ProformaNo = Val(con.Execute("Select Max(Cast(LEFT(PInvoice,5) as Int)) From FPInvoice ").Fields(0).Value & "")
    End If
    If ProformaNo = 0 Then
        'StNo = Val(InputBox("Enter Starting Serial #"))
        ProformaNo = 0
    Else
        ProformaNo = ProformaNo
    End If
    ProformaNo = ProformaNo + 1
    GetProformaNo = ProformaNo
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Public Function GetBatchNo() As Long

    On Error GoTo err
    
    
    Dim ProformaNo As Long
    Dim stNo As String
    
    ProformaNo = Val(con.Execute("Select Max(BatchNo) From FPInvoice").Fields(0).Value & "")
    
    If ProformaNo = 0 Then
        'StNo = Val(InputBox("Enter Starting Batch #"))
        ProformaNo = 1
    Else
        ProformaNo = ProformaNo + 1
    End If
    
    GetBatchNo = ProformaNo
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Public Function CheckHoliday(DTCheck As Date) As Date

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim DT1 As Date
    Dim InHoliday As Boolean
    DT1 = DTCheck
CheckAgain:
    With rs
        .Open "Select DT From Holidays Where DT='" & DT1 & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            InHoliday = True
        Else
            InHoliday = False
        End If
        .Close
    End With
    Set rs = Nothing
    
    If InHoliday Then
        DT1 = DateAdd("d", 1, DT1)
    End If
    
    If Format(DT1, "dddd") = "Sunday" Then
        DT1 = DateAdd("d", 1, DT1)
        GoTo CheckAgain
    Else
        CheckHoliday = DT1
    End If
    
    
    
    Exit Function
err:
    MsgBox err.Description
End Function

Public Function CheckForNullDT(TempDT As Date) As String
    If TempDT = "1/1/1900" Then
        CheckForNullDT = ""
    Else
        CheckForNullDT = TempDT
    End If
End Function


Public Sub PrintInvoiceForBank(rpt As CRAXDDRT.Report)

    
'    With FPrintInv
'
'
'        rpt.Sections(4).ReportObjects(3).Suppress = True    'rpt.FakeAddress.Suppress = True
'
'        rpt.Sections(4).ReportObjects(2).Suppress = False    'rpt.ActualAddress.Suppress = False
'        rpt.Sections(14).Suppress = True                    'rpt.ReportFooterSection3.Suppress = True
'
'        rpt.Sections(14).Suppress = True 'rpt.ReportFooterSection4.Suppress = True
'        rpt.Sections(16).Suppress = True 'rpt.ReportFooterSection2.Suppress = True
'
'        rpt.FormulaFields(10).Text = False  'FormulaName=ForBank
'
'    End With
    
    Call PrintInvoice(rpt)
    
End Sub

Public Sub PrintInvoice(rpt As CRAXDDRT.Report)

    Dim frmPrint As New frmPrint
    
    With frmPrint
        Set .rpt = rpt
        .cmbCust.ClearVals
        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
        .cmbCust.AddItem "<All Customers>", "0"
        .cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'Hico'"
        .SelFor = "{VrptProformas.CustomInvoice}='"
        .JobCondField = " CustCode "
        .JobField = " CustomInvoice "
        .JobIDField = " CustomInvoice "
        .JobTable = " CustomInvoice "
        .TheEnd = "'"
        .Show , MainForm
    End With
    
End Sub


Public Function UpdateRequistion(PInvoice As String) As Boolean

     
    On Error GoTo err
    
    Dim PchrOrdNo As String
    Dim iNextNo As String
    Dim rs As New ADODB.Recordset
    Dim ColLv As New Collection
    Dim ColObj As UpdateSteelReq
    
    Dim ExtraPer As Double
    
    With rs
                        
        .Open "Select * From VFOrderSteelWeightEP Where PInvoice='" & PInvoice & "'", con, adOpenForwardOnly, adLockReadOnly
        'LV1.ListItems.Clear
        For i = 1 To ColLv.count
            ColLv.Remove 1
        Next i
        
        Do Until .EOF
            Set ColObj = New UpdateSteelReq
            'Set itm = LV1.ListItems.Add(, ![SteelUsed] & "'" & .AbsolutePosition, ![SteelType] & "")
            'itm.SubItems(1) = ![TotalWeight] & ""
            
            ExtraPer = (Val(![TotalWeightEP] & "") - Val(![TotalWeight] & "")) / Val(![TotalWeight] & "")
            ExtraPer = ExtraPer * 100
            
            ColObj.SteelID = ![SteelUsed] & ""
            ColObj.SteelWeight = ![TotalWeightEP] & ""
            ColObj.ExtraPer = ExtraPer
            ColLv.add ColObj
            Set ColObj = Nothing
            .MoveNEXT
        Loop
        
        .Close
        
        
        con.Execute "Delete From FSteelPurchDetail2 Where PInvoice='" & PInvoice & "'"
        con.Execute "Delete From FPurchOrderComps1 Where RefID In(Select EntryID From FSteelPurchDetail1 Where PInvoice='" & PInvoice & "')"

        
        .Open "Select * From VFSteelPurchOrdGageWise Where PInvoice='" & PInvoice & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            
            dW = ![TotalWeightEP]
            
            con.Execute "Insert Into FSteelPurchDetail2(PInvoice,SteelID,GageID,RealWeight,SaleInvWeight) Values('" & PInvoice & "'," & _
             ![SteelUsed] & ",'" & ![Gage] & "'," & dW & "," & dW & ")"
             
            .MoveNEXT
        Loop
        .Close
        
    End With
    '-----------------------------------------------------------------
    
    '{[(<<<<<<<<<<<<Start From Here>>>>>>>>>>>)]}
    Randomize
    
    con.Execute "Delete from FSteelPurchDetail Where PInvoice='" & PInvoice & "'"
    con.Execute "Update FSteelPurchaseOrder Set ExtraPer=" & ExtraPer & _
     " Where PInvoice='" & PInvoice & "'"
     
    '--------------------------------------
    'Call DoForSaleInvoice(PInvoice, ExtraPer)
    '--------------------------------------
    
    
    Dim iChNo As Integer
    
    
    Dim dWeight As Double, dWeight1 As Double
    Dim dWUsed As Double
    Dim dBal As Double
    Dim iRnd As Double
    
    Dim dSW As Double
    
    Dim dRndVal As Double
    
    
    Dim iRefID As Double
    
    dWUsed = 0
    'Get Charge No.
    Randomize
    iChNo = Int(Rnd * (8000 - (2000 - 1)) + (2000))
    
    For i = 1 To ColLv.count
        
        
    'rs.Open "Select * From VSteelTypeComps Where SteelID=" & Val(LV1.ListItems(i).Key), con, adOpenForwardOnly, adLockReadOnly
        rs.Open "Select * From VSteelTypeComps Where SteelID=" & Val(ColLv(i).SteelID), con, adOpenForwardOnly, adLockReadOnly
        
        dWeight = Val(ColLv(i).SteelWeight) 'Total Steel
        
        
        'dWeight = dWeight + (dWeight * (Val(TBox(0)) / 100)) 'Steel With %
        
        
        dWUsed = 0 'Reset To 0.
        
        Do While dWUsed < dWeight
            If rs.RecordCount > 0 Then rs.MoveFirst
            dBal = dWeight - dWUsed
            
            Randomize
            'iRnd = Int(Rnd * (Val(rs![StripeMaxW]) - (Val(rs![StripeMinW]) - 1)) + (Val(rs![StripeMinW])))
            iRnd = Int(Rnd * (Val(400) - (Val(385) - 1)) + (Val(385)))
            
            'The Following Line is commented out To keep the weight of Last Stripe Consistent with other stripes
            'If iRnd > dBal Then iRnd = dBal 'If Random Value is More Than Bal
            dSW = iRnd
            
            con.Execute "Insert Into FSteelPurchDetail(PInvoice,SteelUsed" & _
             ",Guage,StripeWeight,ChargeNo)" & _
             " Values('" & PInvoice & "','" & Val(ColLv(i).SteelID) & _
             "',''," & dSW & "," & _
             iChNo & ")"
            
            iRefID = con.Execute("Select Max(EntryID) From FSteelPurchDetail").Fields(0).Value
            
            Do Until rs.EOF
                dRndVal = Format((Rnd * (Val(rs![UVal] & "") - (Val(rs![lVal] & ""))) + (Val(rs![lVal] & ""))), "#.00")
                
                con.Execute "Insert Into FPurchOrderComps(RefID,CompID" & _
                 ",CompVal) Values(" & iRefID & "," & rs![CompID] & "," & _
                 dRndVal & ")"
    
                rs.MoveNEXT
            Loop
            
            'Increment Charge No. By One.
            iChNo = iChNo + 1
            
            dWUsed = dWUsed + dSW
        Loop
        
        rs.Close
        
    Next i
    
    
    ''''''''''''''
    
    Dim TestRefNo As String
    Dim TestDT As Date
    Dim MaterialTestDT As Date
    Dim strBNo As String
    Dim iRndRow As Integer
    
    
    'TestDT = DateAdd("d", 1, DTRec)
    
    'TestDT = CheckHoliday(TestDT)
    
    'MaterialTestDT = DateAdd("d", 2, TestDT)
    'MaterialTestDT = CheckHoliday(MaterialTestDT)
    
    
    'TestRefNo = Int(Rnd * (999 - (100 - 1)) + (100)) & "/" & Format(TestDT, "yyyy")
    
    'strBNo = Format(TestDT, "ddMMyy")
    
    'con.Execute "Insert Into FSteelPurchDetail1(PInvoice,OurRefNo,DT,BatchNo,MaterialTestDT) Values('" & _
     PInvoice & "','" & TestRefNo & "','" & TestDT & "','" & strBNo & "','" & MaterialTestDT & "')"
     
    Randomize
    'iRndRow = Int(Rnd * (LV1.ListItems.Count - (1 - 1)) + (1))
    iRndRow = Int(Rnd * (ColLv.count - (1 - 1)) + (1))
    
    rs.Open "Select * From VSteelTypeComps Where SteelID=" & Val(ColLv(iRndRow).SteelID), con, adOpenForwardOnly, adLockReadOnly
    
    iRefID = con.Execute("Select EntryID From FSteelPurchDetail1 Where PInvoice='" & PInvoice & "'").Fields(0).Value
    
    Do Until rs.EOF
        Randomize
        dRndVal = Format((Rnd * (Val(rs![UVal] & "") - (Val(rs![lVal] & ""))) + (Val(rs![lVal] & ""))), "#.00")
                
        con.Execute "Insert Into FPurchOrderComps1(RefID,CompID" & _
         ",CompVal) Values(" & iRefID & "," & rs![CompID] & "," & _
         dRndVal & ")"

        rs.MoveNEXT
    Loop
    
    rs.Close
    
    
    Set rs = Nothing
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    strOrderNo = ""
    UpdateRequistion = True
    
    Exit Function
    
err:
    
    MsgBox err.Description
    
End Function
Private Sub DoForSaleInvoice(PInvoice As String, ExtraPer As Double)

    Dim rs As New ADODB.Recordset
    Dim rndVal As Integer
    Dim dW As Double
    
    With rs
        .Open "Select * From VFSteelPurchOrdGageWise Where PInvoice='" & PInvoice & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
        
            'rndVal = Int(Rnd * (Val(2) - (Val(1) - 1)) + (Val(1)))
            'If rndVal = 1 Then
            'Else
            'End If
            dW = ![TotalWeight] + (![TotalWeight] * (ExtraPer / 100))
            
            con.Execute "Insert Into FSteelPurchDetail2(PInvoice,SteelID,GageID,RealWeight,SaleInvWeight) Values('" & cmbOrders.Text & "'," & _
             ![SteelUsed] & ",'" & ![Gage] & "'," & dW & "," & dW + 1 & ")"
             
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub


Public Sub ShowCustListrpt()
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report 'New rptCustList
    Set rpt = rptApp.OpenReport(RptPath & "\rptCustList.rpt") '
    Dim f As New frmPrevRpt
    f.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Public Function GetInvoiceNo(iCompanyID As Integer, Optional bCourierInvoice As Boolean = False, Optional myDT As Date) As String

    On Error GoTo err
    
    Dim InvoiceNo As Long, lYearlyInvoiceNo As Long
    Dim stNo As String
    
    Dim rs As New ADODB.Recordset
    Dim strPrefix As String
    With rs
        .Open "Select InvoicePrefix From Companies Where EntryID=" & iCompanyID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            strPrefix = ""
        Else
            strPrefix = !InvoicePrefix & ""
        End If
        .Close
    End With
    Set rs = Nothing
    If strPrefix = "" Then
        strPrefix = "ABC"
    End If
    
    If bCourierInvoice Then
        InvoiceNo = Val(con.Execute("Select Max(Cast(Right(CustomInvoice,Len(CustomInvoice)-5) As int)) From CustomInvoice Where Left(CustomInvoice,4)='" & strCourierInvoicePrefix & "'").Fields(0).Value & "")
    ElseIf strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        'InvoiceNo = Val(con.Execute("SELECT MAX(CAST(SUBSTRING(CustomInvoice,CHARINDEX('-',CustomInvoice)+1,CHARINDEX('/',CustomInvoice)-CHARINDEX('-',CustomInvoice)-1) As Int)) From CustomInvoice WHERE LEFT(CustomInvoice," & Len(strPrefix) & ")='" & strPrefix & "'").Fields(0).Value & "")
        'InvoiceNo = Val(con.Execute("SELECT MAX(CAST(LEFT(CustomInvoice,CHARINDEX('/',CustomInvoice,1)-1) AS INT)) FROM CustomInvoice WHERE CHARINDEX('/',CustomInvoice,1)>1 AND CHARINDEX('MSL',CustomInvoice,1)>1").Fields(0).Value & "")
        InvoiceNo = Val(con.Execute("SELECT MAX(CAST(SUBSTRING(CustomInvoice,5,LEN(CustomInvoice)-4) AS INT)) FROM CustomInvoice WHERE LEFT(CustomInvoice,2)='E-' AND LEN(CustomInvoice)>=7 AND LEN(REPLACE(CustomInvoice,'-',''))=(LEN(CustomInvoice)-1)").Fields(0).Value & "")
    ElseIf strCompany = "Banzai" Then
        InvoiceNo = Val(con.Execute("SELECT MAX(CAST(SUBSTRING(CustomInvoice,CHARINDEX('-',CustomInvoice)+1,4) As Int)) From CustomInvoice WHERE LEFT(CustomInvoice," & Len(strPrefix) & ")='" & strPrefix & "'").Fields(0).Value & "")
        lYearlyInvoiceNo = Val(con.Execute("SELECT MAX(CAST(SUBSTRING(CustomInvoice,LEN(CustomInvoice)-CHARINDEX('-',REVERSE(CustomInvoice))+2,CHARINDEX('-',REVERSE(CustomInvoice))) As Int)) From CustomInvoice WHERE LEFT(CustomInvoice," & Len(strPrefix) & ")='" & strPrefix & "' AND YEAR(DT)=" & year(myDT)).Fields(0).Value & "")
    ElseIf strCompany = "Instrumed" Then
        InvoiceNo = Val(con.Execute("SELECT MAX(CAST(RIGHT(CustomInvoice,4) As Int)) FROM CustomInvoice WHERE LEFT(CustomInvoice,8)='IL/" & Format(myDT, "yyyy") & "/'").Fields(0).Value & "")
    Else
        'InvoiceNo = Val(con.Execute("Select Max(Cast(Right(CustomInvoice,Len(CustomInvoice)-4) As int)) From CustomInvoice Where Left(CustomInvoice,3)='VTL'").Fields(0).value & "")
        InvoiceNo = Val(con.Execute("Select Max(Cast(SUBSTRING(CustomInvoice,4,4) As Int)) From CustomInvoice WHERE LEFT(CustomInvoice,3)='" & strPrefix & "'").Fields(0).Value & "")
    End If
    
    If strCompany <> "Instrumed" Then
        If InvoiceNo = 0 Then
            stNo = Val(InputBox("Enter Starting Serial #"))
            InvoiceNo = stNo
            lInvoiceNo = InvoiceNo
            If strCompany = "Banzai" Then
                stNo = Val(InputBox("Enter Yearly Starting Serial #"))
                lYearlyInvoiceNo = stNo
            End If
        Else
            lInvoiceNo = InvoiceNo + 1
        End If
    Else
        lInvoiceNo = 1
    End If
    If strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        Dim lYear As Long
        If Month(myDT) <= 6 Then
            lYear = year(myDT) - 1
        Else
            lYear = year(myDT)
        End If
        GetInvoiceNo = "E-" & Format(myDT, "YY") & lInvoiceNo
        'GetInvoiceNo = lInvoiceNo & "/MSL/" & Right(lYear, 2)
        'GetInvoiceNo = strPrefix & InvoiceNo
    ElseIf strCompany = "Banzai" Then
        GetInvoiceNo = strPrefix & "-" & InvoiceNo & "-" & Format(myDT, "yyyy") & "-" & lYearlyInvoiceNo
    ElseIf strCompany = "Instrumed" Then
        GetInvoiceNo = "IL/" & Format(myDT, "yyyy") & "/" & Format(InvoiceNo, "0000")
    Else
        GetInvoiceNo = strPrefix & Format(InvoiceNo, "0000/")
    End If
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Public Function GetSampleInvoiceNo(iCompanyID As Integer, Optional bCourierInvoice As Boolean = False) As String

    On Error GoTo err
    
    Dim InvoiceNo As Long
    Dim stNo As String
    
    Dim rs As New ADODB.Recordset
    Dim strPrefix As String
    With rs
        .Open "SELECT InvoicePrefix FROM Companies WHERE EntryID=" & iCompanyID, con, adOpenForwardOnly, adLockReadOnly
        strPrefix = !InvoicePrefix & ""
        .Close
    End With
    Set rs = Nothing
    If strPrefix = "" Then
        strPrefix = "ABC"
    End If
    
    
    InvoiceNo = Val(con.Execute("Select Max(Cast(SUBSTRING(CustomInvoice,4,4) As Int)) From SampleInvoice WHERE LEFT(CustomInvoice,3)='" & strPrefix & "'").Fields(0).Value & "")
   
    
    If InvoiceNo = 0 Then
        stNo = 0 'Val(InputBox("Enter Starting Serial #"))
        InvoiceNo = 0
    Else
        InvoiceNo = InvoiceNo
    End If
    
    
    InvoiceNo = InvoiceNo + 1
    GetSampleInvoiceNo = strPrefix & Format(InvoiceNo, "0000/")
    
    Exit Function
err:
    MsgBox err.Description
    
End Function


Public Sub PrintReceivableAgingReport()
    'PrintShippedUnShippedSummary
    Dim rpt As CRAXDDRT.Report
    Dim rptSub As CRAXDDRT.Report
   
    Set rpt = rptApp.OpenReport(RptPath & "\ReceivablesInvoicesAging.rpt")
     
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "{SP_PaymentStatus;1.EntryID}=0 AND {SP_PaymentStatus;1.AWBNoDT}<>#12/25/2001#", rpt, , False
    
End Sub

Public Function GetNewTecnoInvoiceNo(myDT As Date, Optional bSampleInvoice As Boolean = False, Optional bFreeSamples As Boolean = False) As String

    Dim lInvNo As Long, strInvNo As String
    
    If bSampleInvoice Then
        If bFreeSamples Then
            lInvNo = GetSingleLongValue("MAX(CAST(RIGHT(CustomInvoice,CHARINDEX('/',REVERSE(CustomInvoice))-1) AS INT))", "CustomInvoice", " WHERE LEFT(CustomInvoice,3)='TCI/S/' AND SampleInvoice=1 AND FreeSamples=1")
        Else
            lInvNo = GetSingleLongValue("MAX(CAST(RIGHT(CustomInvoice,CHARINDEX('/',REVERSE(CustomInvoice))-1) AS INT))", "CustomInvoice", " WHERE LEFT(CustomInvoice,3)='TCI/S/' AND SampleInvoice=1 AND FreeSamples=0")
        End If
    Else
        lInvNo = GetSingleLongValue("MAX(CAST(RIGHT(CustomInvoice,CHARINDEX('/',REVERSE(CustomInvoice))-1) AS INT))", "CustomInvoice", " WHERE LEFT(CustomInvoice,3)='TCI'")
    End If
    
    lInvNo = lInvNo + 1
    
    'strInvNo = "EXP/" & Format(myDT, "yyyy") & "/" & Format(lInvNo, "0000")
    If bSampleInvoice Then
        If bFreeSamples Then
            strInvNo = "TCI/S/" & Format(lInvNo, "0000")
        Else
            strInvNo = "TCI/S/" & Format(lInvNo, "000")
        End If
    Else
        strInvNo = "TCI/" & Format(lInvNo, "0000")
    End If
    
    GetNewTecnoInvoiceNo = strInvNo
    
End Function


Public Function GetNewQELInvoiceNo(myDT As Date) As String

    Dim lInvNo As Long, strInvNo As String
    lInvNo = GetSingleLongValue("MAX(CAST(RIGHT(CustomInvoice,CHARINDEX('/',REVERSE(CustomInvoice))-1) AS INT))", "CustomInvoice", " WHERE LEFT(CustomInvoice,3)='EXP'")
    lInvNo = lInvNo + 1
    
    strInvNo = "EXP/" & Format(myDT, "yyyy") & "/" & Format(lInvNo, "0000")
    GetNewQELInvoiceNo = strInvNo
    
End Function

Public Sub PrintShippedUnShippedSummary()
    'PrintDeliveryAnalysis
    Dim myDTFrom As Date, myDTTo As Date
    Dim f As New frmGetDateDouble
    If f.getDate(myDTFrom, myDTTo) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\OrdersReceivedShippedUnShippedSummary.rpt")
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
    
End Sub

Public Sub PrintDeliveryAnalysis(Optional strOrderNo As String)
    '
    Dim myDTFrom As Date, myDTTo As Date
'    Dim f As New frmGetDateDouble
'    If f.getDate(myDTFrom, myDTTo) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\ItemDeliveryAnalysis.rpt")
    'rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue myDTFrom
    'rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue myDTTo
    'rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue Null
    rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue strOrderNo
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
    
End Sub


Public Sub PrintPackingList(strInvoiceNo As String, Optional bCustom As Boolean = True, Optional iReportType As Integer = 0)

    On Error GoTo err
        
    Dim rpt As CRAXDDRT.Report
    If bCustom Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptCustomPList.rpt")
    Else
        If iReportType = 2 Then
            Set rpt = rptApp.OpenReport(RptPath & "\Bank_Packing_List.rpt")
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\rptCommercialPList.rpt")
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VCustomPList.CustomInvoice}='" & strInvoiceNo & "' ", rpt
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    MsgBox err.Description
End Sub
