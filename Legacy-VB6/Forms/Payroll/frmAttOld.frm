VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmAttOld 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8685
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   12840
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   12840
   Begin MSComCtl2.DTPicker DTEdit 
      Height          =   315
      Left            =   7275
      TabIndex        =   9
      Top             =   1335
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "hh:mm  tt"
      Format          =   64749571
      UpDown          =   -1  'True
      CurrentDate     =   37384
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   4845
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttOld.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttOld.frx":0454
            Key             =   "ORD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttOld.frx":08AC
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttOld.frx":0D00
            Key             =   "ITMIN"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttOld.frx":1154
            Key             =   "CAT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttOld.frx":15A8
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttOld.frx":1A00
            Key             =   "TYPE"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   7485
      Left            =   60
      TabIndex        =   8
      Top             =   720
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   13203
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   6495
      TabIndex        =   1
      Top             =   390
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "ddd dd-MMM yy"
      Format          =   64749571
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTFrom 
      Height          =   330
      Left            =   8175
      TabIndex        =   5
      Top             =   390
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "hh:mm  tt"
      Format          =   64749571
      UpDown          =   -1  'True
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTTo 
      Height          =   330
      Left            =   9615
      TabIndex        =   6
      Top             =   390
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "hh:mm  tt"
      Format          =   64749571
      UpDown          =   -1  'True
      CurrentDate     =   37384
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7485
      Left            =   6450
      TabIndex        =   7
      Top             =   720
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   13203
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ITM"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "In Time"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Out Time"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Net Time"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   435
      Left            =   4065
      TabIndex        =   10
      Top             =   8235
      Visible         =   0   'False
      Width           =   1965
      Caption         =   "Feed All Attendance"
      PicturePosition =   327683
      Size            =   "3466;767"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAttOld.frx":1E54
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
      Height          =   225
      Index           =   4
      Left            =   60
      TabIndex        =   4
      Top             =   150
      Width           =   10995
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   435
      Left            =   11100
      TabIndex        =   2
      Top             =   255
      Width           =   1590
      Caption         =   "Save Time"
      PicturePosition =   327683
      Size            =   "2805;767"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   360
      Left            =   11085
      TabIndex        =   3
      Top             =   8265
      Width           =   1605
      Caption         =   "Close            "
      PicturePosition =   327683
      Size            =   "2831;635"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmAttOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public edit As Boolean
Dim EditOutTime As Boolean
Dim bEmployees As Boolean
Dim bOverTime As Boolean

Private Sub AddHeads()

    Dim Rec As New ADODB.Recordset, RecDept As New ADODB.Recordset
    
    Tv.ImageList = ImageList1
    Tv.Nodes.Clear
    
    If bEmployees Then
        RecDept.Open "SELECT DeptID,Name FROM Departments Where TempDept=0", con, adOpenForwardOnly, adLockReadOnly
    Else
        RecDept.Open "SELECT DeptID,Name FROM Departments Where TempDept=1", con, adOpenForwardOnly, adLockReadOnly
    End If
    
    For i = 1 To RecDept.RecordCount
        Tv.Nodes.add , , RecDept![DeptID] & "'", RecDept![Name] & "", "TYPE"
        With Rec
            .Open "Select EmpID,Name from Employees WHERE DeptID='" & RecDept!DeptID & "' and Active=1 ORDER BY EmpID", con, adOpenForwardOnly, adLockReadOnly
            For j = 1 To .RecordCount
                Tv.Nodes.add RecDept![DeptID] & "'", tvwChild, ![EmpId] & "'", ![Name] & "", "CAT"
                .MoveNext
            Next
            .Close
        End With
        RecDept.MoveNext
    Next

End Sub

Private Sub TakeRegularAtt()
        'Call MakeAllPresentForJuly
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    Dim i As Integer
    rs1.Open "Select DeptID,LunchInTime,LunchOutTime From GeneralSettings", con, adOpenKeyset, adLockReadOnly
    Dim iRet As Integer
    Dim AttTime As String
    Call StartTrans(con)
    With rs
        '.Open "Select EmpID From EmpTimes Where DT='" & DTPicker1 & "' AND OutTime Is Null", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EmpID From Employees Where Active=1 AND DeptID<>'ASTT' AND EmpID Not In(Select EmpID From AttendanceSheet Where DT='" & DTPicker1.value & "') Order By EmpID", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select DeptID,EmpID From Employees Where EmpID In('VBPE-0007','VHRM-0005','VHRM-0006','VPDT-0063','VPDT-0064','VPDT-0065','VPDT-0066','VPDT-0067','VPDT-0068','VQAS-0011','VQTC-0057','VQTC-0058','VQTC-0059','VQTC-0060','VQTC-0061','VQTC-0062','VQTC-0063','VQTC-0064','VQTC-0065','VQTC-0066','VRND-0015','VRND-0016','VRND-0017','VRND-0018','VSIP-0004','VSIP-0005','VSRT-0033','VSRT-0034','VSRT-0035','VSRT-0036','VSRT-0037','VSRT-0038','VSRT-0039','VSRT-0040','VSRT-0041','VSRTA-0001','VSRTA-0002','VSVS-0086','VSVS-0087','VWSH-0060','VWSH-0061','VWSH-0062','VWSH-0063','VWSH-0064','VWSH-0065','VWSH-0066','VWSH-0067','VWSH-0068','VWSH-0069','VWSH-0070','VWSH-0071','VWSH-0072','VWSH-0073','VWSH-0074','VWSH-0075') Order By EmpID", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            rs1.Find "DeptID='" & !DeptID & "'"
            For i = 1 To 29
                DTPicker1.value = DateSerial(2008, 2, i)
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                AttTime = "08:" & Format(iRet, "00")
                Call SaveNow(!EmpId, DTPicker1.value, Format(AttTime, "HH:NN"))
                
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                'AttTime = "12:" & Format(30 + iRet, "00")
                AttTime = DateAdd("n", iRet, rs1!LunchInTime)
                Call SaveNow(!EmpId, DTPicker1.value, Format(AttTime, "HH:NN"))
                
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                'AttTime = "13:" & Format(25 + iRet, "00")
                AttTime = DateAdd("n", -iRet, rs1!LunchOutTime)
                Call SaveNow(!EmpId, DTPicker1.value, Format(AttTime, "HH:NN"))
                
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                AttTime = "17:" & Format(iRet, "00")
                Call SaveNow(!EmpId, DTPicker1.value, Format(AttTime, "HH:NN"))
            Next
            Debug.Print !EmpId
            .MoveNext
        Loop
        .Close
    End With
    con.CommitTrans
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub CommandButton1_Click()
    Call TakeRegularAtt
    MsgBox "OK"
    'Call TakeRamzanBreakAtt
End Sub

Private Sub TakeRamzanBreakAtt()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim iRet As Integer
    Dim AttTime As String
    Dim AttDT As Date
    Dim strTime As String
    Dim lEntryID As Long
    Dim dHrs As Double
    
    Call StartTrans(con)
    With rs
        'It was taken from 14-Sep-07 to 20-Sep-07
        .Open "Select Min(EntryID) as EntryID,EmpID,DT,Min(InTime) as InTime,Min(OutTime) as OutTime From EmpTimes " & _
        "Where DT Between '10/11/2007' AND '10/12/2007' GROUP BY EmpID,DT Having Count(EmpId) = 1", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EntryID,EmpID,DT,InTime,OutTime From EmpTimes Where DT Between '9/14/2007' AND '9/20/2007' AND PayableHrs<5 ", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EntryID,EmpID,DT,InTime,OutTime From EmpTimes Where month(DT)=10 and year(DT)=2006 And OutTime Is Null AND EmpID In(Select EmpID From VEmp Where TempDept=0) AND InTime<'10:00' ORDER By EmpID,DT", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EntryID,EmpID,DT,InTime,OutTime From EmpTimes Where DT In('10/13/2006','10/14/2006','10/18/2006') AND EmpID='VSVS-0071'", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EntryID,EmpID,DT,InTime,OutTime From EmpTimes Where DT Between '10/01/2006' AND '10/21/2006' AND EmpTimes.EmpID In(Select EmpID From VEmp Where TempDept=0 AND Designation<>'Security Guard') Group By EmpTimes.EmpID,DT Having Count(EmpTimes.EmpID)=1", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EntryID,EmpID,DT,InTime,OutTime From EmpTimes Where DT ='10/20/2006' AND PayableHrs<5 AND EmpID='VMIS-0001'", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EmpID,DT From EmpTimes Where DT Between '10/1/2006' AND '10/21/2006' AND EmpID Not In('VSVS-0065') AND EmpID In(Select EmpID From VEmp Where TempDept=0 AND Designation<>'Security Guard') Group By EmpID,DT Having Count(EmpID)=1 Order By EmpID,DT", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            
            lEntryID = Val(!EntryID & "")
            
            iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
            AttTime = "13:" & Format(iRet, "00")
            If !InTime < "08:11" Then
                dHrs = DateDiff("n", "08:00", AttTime) / 60
            Else
                dHrs = DateDiff("n", !InTime, AttTime) / 60
            End If
            
            con.Execute "Update EmpTimes Set OutTime='" & AttTime & "',Hrs=" & dHrs & ",PayableHrs=" & dHrs & " Where EntryID=" & lEntryID
            'con.Execute "Update EmpTimes Set OutTime='" & AttTime & "',Hrs=" & dHrs & ",PayableHrs=" & dHrs & " Where EmpID='" & !EmpId & "' AND DT='" & !DT & "'"
            'Call SaveNow(!EmpId, AttDT, Format(AttTime, "HH:NN"))
            
            If Format(!DT, "dddd") = "Friday" Then
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                AttTime = "14:" & Format(iRet, "00")
            Else
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                AttTime = "13:" & Format(25 + iRet, "00")
            End If
            strTime = AttTime
            
            iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
            AttTime = "15:" & Format(30 + iRet, "00")
            dHrs = DateDiff("n", strTime, AttTime) / 60
            con.Execute "Insert Into EmpTimes(EmpID,DT,InTime,OutTime,Hrs,PayableHrs) Values('" & _
             !EmpId & "','" & !DT & "','" & strTime & "','" & AttTime & "'," & dHrs & ",3)"
            
            Debug.Print !EmpId
            .MoveNext
        Loop
        .Close
    End With
    con.CommitTrans
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub DTEdit_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim InTime As String, OutTime As String, Hrs As Double
    Dim dLateHours As Double
    Dim EmpId As String, DT As Date
    
    If KeyCode = 13 Then
        If EditOutTime Then
            OutTime = Format(DTEdit, "HH:NN")
            InTime = con.Execute("Select InTime FROM EmpTimes WHERE EntryID=" & LV.SelectedItem.Tag & "").Fields(0)
            Hrs = DateDiff("n", CDate(InTime), CDate(OutTime)) - 60              ' Lunch Time
            con.Execute "UPDATE EmpTimes SET OutTime='" & OutTime & "',Hrs=" & Round(Hrs / 60, 2) & " WHERE EntryID=" & LV.SelectedItem.Tag & ""
        Else
            InTime = Format(DTEdit, "HH:NN")
            OutTime = con.Execute("Select OutTime FROM EmpTimes WHERE EntryID=" & LV.SelectedItem.Tag & "").Fields(0) & ""
            If OutTime <> "" Then Hrs = DateDiff("n", CDate(InTime), CDate(OutTime)) - 60        ' Lunch Time
            con.Execute "UPDATE EmpTimes Set InTime='" & InTime & "',Hrs=" & Round(Hrs / 60, 2) & " WHERE EntryID=" & LV.SelectedItem.Tag & ""
            
            If LV.SelectedItem.Index = 1 Then 'Update Late Hours If It's Ist Entry Of The Day.\
                dLateHours = DateDiff("n", "08:00", InTime) / 60
                If dLateHours > 10 Then
                    con.Execute "Update AttendanceSheet Set LateHours=" & dLateHours & " Where EmpID='" & cmbEmp.ID & "' And DT='" & DTPicker1 & "'"
                Else
                    con.Execute "Update AttendanceSheet Set LateHours=Null Where EmpID='" & cmbEmp.ID & "' And DT='" & DTPicker1 & "'"
                End If
            End If
        End If
        DTEdit.Visible = False
        Call RefreshList
    End If
    
End Sub

Private Sub LV_Click()
    DTEdit.Visible = False
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo err
    
    If LV.ListItems.Count = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        If MsgBox("Do you really want to delete ?", vbInformation + vbYesNo) = vbYes Then
            
'            con.Execute "Delete From EmpTimes Where EntryID=" & Val(LV.SelectedItem.Tag)
            
'            If LV.ListItems.Count = 1 Then
'                con.Execute "Delete From AttendanceSheet Where EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1 & "'"
'            End If
            con.Execute "Delete From AttendanceSheet Where EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1.value & "'"
            con.Execute "Delete From EmpTimes Where EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1.value & "'"
            
            Call RefreshList
        End If
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_DblClick()
    
    Exit Sub    'Stop The Editing For Now.... Can Still Delete The Att. And Take New.
    If Not EditOutTime Then
        DTEdit.Move LV.Left + LV.ColumnHeaders(2).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(2).Width
        DTEdit = CDate(LV.SelectedItem.SubItems(1))
        DTEdit.Visible = True
        DTEdit.SetFocus
    ElseIf EditOutTime And LV.SelectedItem.SubItems(2) <> "" Then
        DTEdit.Move LV.Left + LV.ColumnHeaders(3).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(3).Width
        DTEdit = CDate(LV.SelectedItem.SubItems(2))
        DTEdit.Visible = True
        DTEdit.SetFocus
    End If
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If X > LV.ColumnHeaders(3).Left Then
        EditOutTime = True
    Else
        EditOutTime = False
    End If
    
End Sub

Private Sub TV_Click()
    If Len(Tv.SelectedItem.key) <> 3 Then
        cmbEmp.ID = Left(Tv.SelectedItem.key, Len(Tv.SelectedItem.key) - 1)
    End If
End Sub

Private Sub TV_Collapse(ByVal Node As MSComctlLib.Node)
    'Node.Image = "ITMIN"
End Sub

Private Sub TV_Expand(ByVal Node As MSComctlLib.Node)
    'Node.Image = "ITM"
End Sub

Private Sub SaveNowOld(EmpId As String, AttDate As Date, AttTime As Date)

    On Error GoTo err

    Dim ExtraDeds As Double, ExtraHrs As Double
    ExtraDeds = 0
    ExtraHrs = 0
    Dim MultipleAllowed As Boolean
    Dim bDeductLHrs As Boolean
    Dim bExempt As Boolean
    Dim DeptID As String
    Dim dLateHrs As Double
    Dim TotalEntries As Integer
    Dim AttRecs As Integer
    Dim bOverTime As Boolean
    Dim strEmpInTime As String, strEmpOutTime As String
    Dim strEmpLunchIn As String, strEmpLunchOut As String
    bOverTime = False
    '''''''''''   This Will Is The Employee Exists Or Not    '''''''''''
    
    With con.Execute("Select DeptID,EmpID,EmpType,ExemptSettings FROM Employees WHERE EmpID='" & EmpId & "'")
        If .EOF Then
            Exit Sub
        Else
            EmpId = ![EmpId] & ""
            Emptype = ![Emptype] & ""
            bExempt = ![ExemptSettings] & ""
            DeptID = ![DeptID] & ""
        End If
        .Close
        
        If bExempt Then
            .Open "Select DeductLHrs From EmpSettings Where EmpID='" & EmpId & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select DeductLHrs From GeneralSettings Where DeptID='" & DeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        
        If .EOF Then
            bDeductLHrs = True
        Else
            bDeductLHrs = ![DeductLHrs]
'            strEmpInTime = Format(!EnterTime, "HH:NN")
'            strEmpOutTime = Format(!ExitTime, "HH:NN")
'            strEmpLunchIn = Format(!LunchInTime, "HH:NN")
'            strEmpLunchOut = Format(!LunchOutTime, "HH:NN")
        End If
        .Close
    End With
    
    Dim Rec As New ADODB.Recordset, CurrTime As String, Hrs As Double, SQL As String, Cond As String
    Dim TimeDiff As Integer

    Cond = "WHERE EmpID='" & EmpId & "' AND DT='" & AttDate & "'"

    CurrTime = Format(AttTime, "HH:NN")

    TimeDiff = Val(con.Execute("Select TimeDiff FROM TimeSettings").Fields(0) & "")
        
    With Rec
        .Open "Select TotalEntries From VEmpDailyEntries Where EmpID='" & EmpId & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            TotalEntries = 0
        Else
            TotalEntries = .Fields(0) & ""
        End If
        .Close
        
        .Open "SELECT * FROM EmpTimes WHERE EntryID=(SELECT Max(EntryID) FROM EmpTimes " & Cond & ")", con, adOpenForwardOnly, adLockReadOnly
            
        If TotalEntries > 0 Then
            
            If TotalEntries = 1 And !OutTime & "" = "" Then 'Means It's Lunch In Entry...
                '12:50 to 1:10
                If CurrTime < #12:25:00 PM# Then
                    ExtraDeds = 0 'DateDiff("n", CurrTime, "12:55") / 60
                ElseIf CurrTime > #12:35:00 PM# Then
                    ExtraDeds = 0 'DateDiff("n", "13:00", CurrTime) / 60
                End If
                
                If DateDiff("n", "8:00", CDate(!InTime)) > 0 And DateDiff("n", "8:00", CDate(!InTime)) < 26 Then
                    ExtraHrs = DateDiff("n", "8:00", CDate(!InTime)) / 60
                ElseIf DateDiff("n", "8:00", CDate(!InTime)) < 0 Then
                    ExtraDeds = ExtraDeds + DateDiff("n", CDate(!InTime), "8:00") / 60
                End If
                
                If Format(AttDate, "dddd") = "Friday" Then
                    ExtraHrs = ExtraHrs + 0.25
                    If CurrTime > #12:15:00 PM# Then
                        Hrs = DateDiff("n", CDate(!InTime), #12:15:00 PM#)
                    Else
                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                    End If
                Else
                    If CurrTime > #12:24:00 PM# Then
                        Hrs = DateDiff("n", CDate(!InTime), #12:30:00 PM#)
                    Else
                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                    End If
                End If
                
            ElseIf TotalEntries = 2 And !OutTime & "" = "" Then 'Means It's Lunch Out Entry
                If Format(AttDate, "dddd") = "Friday" Then
                    ExtraHrs = ExtraHrs + 0.5
                    If CDate(!InTime) > #2:05:00 PM# Then
                        ExtraDeds = 0 'DateDiff("n", "14:30", CDate(!InTime)) / 60
                    ElseIf CDate(!InTime) < #2:00:00 PM# Then
                        ExtraDeds = DateDiff("n", CDate(!InTime), "14:00") / 60
                    End If
                    If DateDiff("n", "14:00", CDate(!InTime)) > 0 And DateDiff("n", "14:00", CDate(!InTime)) < 6 Then
                        ExtraHrs = ExtraHrs + DateDiff("n", "14:00", CDate(!InTime)) / 60
                    End If
                Else
                    If DateDiff("n", "13:30", CDate(!InTime)) > 0 And DateDiff("n", "13:30", CDate(!InTime)) < 6 Then
                        ExtraHrs = DateDiff("n", "13:30", CDate(!InTime)) / 60
                    End If
                    
                    If CDate(!InTime) > #1:35:00 PM# Then
                        ExtraDeds = 0 'DateDiff("n", "14:00", CDate(!InTime)) / 60
                    ElseIf !InTime < #1:30:00 PM# Then
                        ExtraDeds = DateDiff("n", CDate(!InTime), "13:30") / 60
                    End If
                End If
                
                If CurrTime > #5:00:00 PM# Then
                    ExtraDeds = ExtraDeds + 0 'DateDiff("n", "17:00", CurrTime) / 60
                End If
                
                If CurrTime > #5:00:00 PM# Then
                    Hrs = DateDiff("n", CDate(!InTime), #5:00:00 PM#)
                Else
                    Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                End If
            End If
            'Check ... Is This Out Time Entry ?
            
            If TotalEntries < 3 Then
                If !OutTime & "" = "" Then
                    'Hrs = DateDiff("n", CDate(!InTime), CDate(CurrTime))
                
                    'Check for the Minimum Time Difference
                    If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                
                    Hrs = Round(Hrs / 60, 2)
                            
                    'Update Out Time and Do Nothing Else
                
                    SQL = "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & ",PayableHrs=" & (Hrs + ExtraHrs - ExtraDeds) & " WHERE EntryID=" & Val(!EntryID & "") & ""
                    con.Execute SQL
                    'con.Execute "Update AttendanceSheet Set Attendance=1 Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
                    GoTo END_OF_PROCEDURE
                Else
                    'Check for the Minimum Time Difference
                
                    Hrs = DateDiff("n", CDate(!OutTime), CDate(CurrTime))
                    If TotalEntries <> 2 Then
                        If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                    End If
                End If
            End If
        Else
            If bDeductLHrs And DateDiff("n", "08:00", CurrTime) > 10 Then
                dLateHrs = DateDiff("n", "08:00", CurrTime) / 60
            Else
                dLateHrs = 0
            End If
            AttRecs = Val(con.Execute("Select Count(*) From AttendanceSheet Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'").Fields(0).value & "")
            If AttRecs <> 0 Then con.Execute "Delete From AttendanceSheet Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
            
            con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance,LateHours) VALUES('" & EmpId & "','" & AttDate & "',1," & _
             dLateHrs & ")", a
            
        End If

        '.Close
        
        If Emptype = 2 Then
            
            ' If No Previous Entry Found Accept The Entry
            If TotalEntries = 0 Then GoTo ACCEPT_ENTRY
        
            'Remember The Total Entries
        
            ' If this is Second Time Entry Accept It
            If TotalEntries < 2 Then GoTo ACCEPT_ENTRY

        End If
            
        ' If Two Entries Have Been Made ... Check For Over Time Authorization
        If TotalEntries >= 2 Then
            Dim rsOT As New ADODB.Recordset
            Dim bAuthorized As Boolean
            With rsOT
                .Open "Select * From VOverTimeAuth Where EmpID='" & EmpId & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
                ' If Not Authorized for Over Time Neglect This Entry
                If .EOF Then
                    bAuthorized = False
                Else
                    bAuthorized = True
                End If
                rsOT.Close
                Set rsOT = Nothing
                If bAuthorized = False Then
                    MsgBox "Not Authorized.", vbInformation
                    GoTo END_OF_PROCEDURE
                End If
            End With
            'If Val(![Authority] & "") <> 1 Then GoTo END_OF_PROCEDURE
            
            ' Other Wise Accept The Entry
            If !OutTime & "" <> "" Then
                con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry,OverTime,UserName,ComputerName) VALUES('" & EmpId _
                 & "','" & AttDate & "','" & CurrTime & "',1,1,'" & CurrentUserName & "','" & strComputerName & "')"
            Else
                Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                Hrs = Round(Hrs / 60, 2)
                con.Execute "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & ",PayableHrs=" & (Hrs) & " WHERE EntryID=" & Val(!EntryID & "") & ""
            End If
            GoTo END_OF_PROCEDURE
            bOverTime = True
        End If
        
    End With

    ' Enter The New Entry With In Time

ACCEPT_ENTRY:
    
    SQL = "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry,UserName,ComputerName) VALUES('" & EmpId _
     & "','" & AttDate & "','" & CurrTime & "',1,'" & CurrentUserName & "','" & strComputerName & "')"
    
    con.Execute SQL
'    TotalEntries = Val(con.Execute("Select TotalEntries From VEmpDailyEntries Where EmpID='" & EmpID & "' and DT='" & AttDate & "'").Fields(0).Value & "")
'
'    If TotalEntries = 1 Then    'That Means that it's first entry of the day. So Calculate LateHours...
'        If DateDiff("n", "08:00", CurrTime) > 15 Then
'            con.Execute "Update AttendanceSheet Set LateHours=" & DateDiff("n", "08:00", CurrTime) / 60 & " Where EmpID='" & EmpID & "' and DT='" & AttDate & "'"
'        End If
'    End If
    
END_OF_PROCEDURE:
    
    If Rec.State = 1 Then Rec.Close
    Set Rec = Nothing
    
    Exit Sub
err:
    Debug.Print err.Description
    'So That Next Entries Should Not Be Skipped
    Resume Next
    
End Sub

Private Sub SaveNow(EmpId As String, AttDate As Date, AttTime As Date)

    On Error GoTo err

    Dim ExtraDeds As Double, ExtraHrs As Double
    ExtraDeds = 0
    ExtraHrs = 0
    Dim MultipleAllowed As Boolean
    Dim bDeductLHrs As Boolean
    Dim bExempt As Boolean
    Dim DeptID As String
    Dim dLateHrs As Double
    Dim TotalEntries As Integer
    Dim AttRecs As Integer
    Dim bOverTime As Boolean
    Dim strEmpInTime As String, strEmpOutTime As String
    Dim strEmpLunchIn As String, strEmpLunchOut As String
    bOverTime = False
    '''''''''''   This Will Is The Employee Exists Or Not    '''''''''''
    
    With con.Execute("Select DeptID,EmpID,EmpType,ExemptSettings FROM Employees WHERE EmpID='" & EmpId & "'")
        If .EOF Then
            Exit Sub
        Else
            EmpId = ![EmpId] & ""
            Emptype = ![Emptype] & ""
            bExempt = ![ExemptSettings] & ""
            DeptID = ![DeptID] & ""
        End If
        .Close
        
        If bExempt Then
            .Open "Select EnterTime,ExitTime,LunchInTime,LunchOutTime,DeductLHrs From EmpSettings Where EmpID='" & EmpId & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select EnterTime,ExitTime,LunchInTime,LunchOutTime,DeductLHrs From GeneralSettings Where DeptID='" & DeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        
        If .EOF Then
            bDeductLHrs = True
        Else
            bDeductLHrs = ![DeductLHrs]
            strEmpInTime = Format(!EnterTime, "HH:NN") '"08:00" 'Format(!EnterTime, "HH:NN")
            strEmpOutTime = Format(!ExitTime, "HH:NN") '"17:00" 'Format(!ExitTime, "HH:NN")
            strEmpLunchIn = Format(!LunchInTime, "HH:NN") '"12:30" 'Format(!LunchInTime, "HH:NN")
            strEmpLunchOut = Format(!LunchOutTime, "HH:NN") '"13:30" 'Format(!LunchOutTime, "HH:NN")
        End If
        .Close
    End With
    
    Dim Rec As New ADODB.Recordset, CurrTime As String, Hrs As Double, SQL As String, Cond As String
    Dim TimeDiff As Integer

    Cond = "WHERE EmpID='" & EmpId & "' AND DT='" & AttDate & "'"

    CurrTime = Format(AttTime, "HH:NN")

    TimeDiff = Val(con.Execute("Select TimeDiff FROM TimeSettings").Fields(0) & "")
        
    With Rec
        .Open "Select TotalEntries From VEmpDailyEntries Where EmpID='" & EmpId & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            TotalEntries = 0
        Else
            TotalEntries = .Fields(0) & ""
        End If
        .Close
        
        .Open "SELECT * FROM EmpTimes WHERE EntryID=(SELECT Max(EntryID) FROM EmpTimes " & Cond & ")", con, adOpenForwardOnly, adLockReadOnly
            
        If TotalEntries > 0 Then
            
            If TotalEntries = 1 And !OutTime & "" = "" Then 'Means It's Lunch In Entry...
                '12:50 to 1:10
'                If CurrTime < #12:25:00 PM# Then
                If CurrTime < CDate(DateAdd("N", -5, strEmpLunchIn)) Then
                    ExtraDeds = 0 'DateDiff("n", CurrTime, "12:55") / 60
                ElseIf CurrTime > CDate(DateAdd("N", 5, strEmpLunchIn)) Then
                    ExtraDeds = 0 'DateDiff("n", "13:00", CurrTime) / 60
                End If
                
                If DateDiff("n", strEmpInTime, CDate(!InTime)) > 0 And DateDiff("n", strEmpInTime, CDate(!InTime)) < 11 Then
                    ExtraHrs = DateDiff("n", strEmpInTime, CDate(!InTime)) / 60
                ElseIf DateDiff("n", strEmpInTime, CDate(!InTime)) < 0 Then
                    ExtraDeds = ExtraDeds + DateDiff("n", CDate(!InTime), strEmpInTime) / 60
                End If
                
                If Format(AttDate, "dddd") = "Friday" Then
                    ExtraHrs = ExtraHrs + 0.25
                    If CurrTime > #12:15:00 PM# Then
                        Hrs = DateDiff("n", CDate(!InTime), #12:15:00 PM#)
                    Else
                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                    End If
                Else
                    If CurrTime > CDate(DateAdd("N", -6, strEmpLunchIn)) Then '#12:24:00 PM#
                        Hrs = DateDiff("n", CDate(!InTime), strEmpLunchIn) '#12:30:00 PM#
                    Else
                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                    End If
                End If
                
            ElseIf TotalEntries = 2 And !OutTime & "" = "" Then 'Means It's Lunch Out Entry
                If Format(AttDate, "dddd") = "Friday" Then
                    ExtraHrs = ExtraHrs + 0.5
                    If CDate(!InTime) > #2:05:00 PM# Then
                        ExtraDeds = 0 'DateDiff("n", "14:30", CDate(!InTime)) / 60
                    ElseIf CDate(!InTime) < #2:00:00 PM# Then
                        ExtraDeds = DateDiff("n", CDate(!InTime), "14:00") / 60
                    End If
                    If DateDiff("n", "14:00", CDate(!InTime)) > 0 And DateDiff("n", "14:00", CDate(!InTime)) < 6 Then
                        ExtraHrs = ExtraHrs + DateDiff("n", "14:00", CDate(!InTime)) / 60
                    End If
                Else
                    If DateDiff("n", strEmpLunchOut, CDate(!InTime)) > 0 And DateDiff("n", strEmpLunchOut, CDate(!InTime)) < 6 Then '"13:30"
                        ExtraHrs = DateDiff("n", strEmpLunchOut, CDate(!InTime)) / 60 '"13:30"
                    End If
                    
                    If CDate(!InTime) > CDate(DateAdd("n", 5, strEmpLunchOut)) Then  '#1:35:00 PM#
                        ExtraDeds = 0 'DateDiff("n", "14:00", CDate(!InTime)) / 60
                    ElseIf !InTime < strEmpLunchOut Then '#1:30:00 PM#
                        ExtraDeds = DateDiff("n", CDate(!InTime), strEmpLunchOut) / 60 '"13:30"
                    End If
                End If
                
                If CurrTime > strEmpOutTime Then '#5:00:00 PM#
                    ExtraDeds = ExtraDeds + 0 'DateDiff("n", "17:00", CurrTime) / 60
                End If
                
                If CurrTime > strEmpOutTime Then '#5:00:00 PM#
                    Hrs = DateDiff("n", CDate(!InTime), strEmpOutTime) '#5:00:00 PM#
                Else
                    Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                End If
            End If
            'Check ... Is This Out Time Entry ?
            
            If TotalEntries < 3 Then
                If !OutTime & "" = "" Then
                    'Hrs = DateDiff("n", CDate(!InTime), CDate(CurrTime))
                
                    'Check for the Minimum Time Difference
                    If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                
                    Hrs = Round(Hrs / 60, 2)
                            
                    'Update Out Time and Do Nothing Else
                
                    SQL = "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & ",PayableHrs=" & (Hrs + ExtraHrs - ExtraDeds) & " WHERE EntryID=" & Val(!EntryID & "") & ""
                    con.Execute SQL
                    'con.Execute "Update AttendanceSheet Set Attendance=1 Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
                    GoTo END_OF_PROCEDURE
                Else
                    'Check for the Minimum Time Difference
                
                    Hrs = DateDiff("n", CDate(!OutTime), CDate(CurrTime))
                    If TotalEntries <> 2 Then
                        If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                    End If
                End If
            End If
        Else
            If bDeductLHrs And DateDiff("n", strEmpInTime, CurrTime) > 10 Then '"08:00"
                dLateHrs = DateDiff("n", strEmpInTime, CurrTime) / 60 '"08:00"
            Else
                dLateHrs = 0
            End If
            AttRecs = Val(con.Execute("Select Count(*) From AttendanceSheet Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'").Fields(0).value & "")
            If AttRecs <> 0 Then con.Execute "Delete From AttendanceSheet Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
            
            con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance,LateHours) VALUES('" & EmpId & "','" & AttDate & "',1," & _
             dLateHrs & ")", a
            
        End If

        '.Close
        
        If Emptype = 2 Then
            
            ' If No Previous Entry Found Accept The Entry
            If TotalEntries = 0 Then GoTo ACCEPT_ENTRY
        
            'Remember The Total Entries
        
            ' If this is Second Time Entry Accept It
            If TotalEntries < 2 Then GoTo ACCEPT_ENTRY

        End If
            
        ' If Two Entries Have Been Made ... Check For Over Time Authorization
        If TotalEntries >= 2 Then
            Dim rsOT As New ADODB.Recordset
            Dim bAuthorized As Boolean
            With rsOT
                .Open "Select * From VOverTimeAuth Where EmpID='" & EmpId & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
                ' If Not Authorized for Over Time Neglect This Entry
                If .EOF Then
                    bAuthorized = False
                Else
                    bAuthorized = True
                End If
                rsOT.Close
                Set rsOT = Nothing
                If bAuthorized = False Then
                    MsgBox "Not Authorized.", vbInformation
                    GoTo END_OF_PROCEDURE
                End If
            End With
            'If Val(![Authority] & "") <> 1 Then GoTo END_OF_PROCEDURE
            
            ' Other Wise Accept The Entry
            If !OutTime & "" <> "" Then
                con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry,OverTime,UserName,ComputerName) VALUES('" & EmpId _
                 & "','" & AttDate & "','" & CurrTime & "',1,1,'" & CurrentUserName & "','" & strComputerName & "')"
            Else
                Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                Hrs = Round(Hrs / 60, 2)
                con.Execute "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & ",PayableHrs=" & (Hrs) & " WHERE EntryID=" & Val(!EntryID & "") & ""
            End If
            GoTo END_OF_PROCEDURE
            bOverTime = True
        End If
        
    End With

    ' Enter The New Entry With In Time

ACCEPT_ENTRY:
    
    SQL = "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry,UserName,ComputerName) VALUES('" & EmpId _
     & "','" & AttDate & "','" & CurrTime & "',1,'" & CurrentUserName & "','" & strComputerName & "')"
    
    con.Execute SQL
'    TotalEntries = Val(con.Execute("Select TotalEntries From VEmpDailyEntries Where EmpID='" & EmpID & "' and DT='" & AttDate & "'").Fields(0).Value & "")
'
'    If TotalEntries = 1 Then    'That Means that it's first entry of the day. So Calculate LateHours...
'        If DateDiff("n", "08:00", CurrTime) > 15 Then
'            con.Execute "Update AttendanceSheet Set LateHours=" & DateDiff("n", "08:00", CurrTime) / 60 & " Where EmpID='" & EmpID & "' and DT='" & AttDate & "'"
'        End If
'    End If
    
END_OF_PROCEDURE:
    
    If Rec.State = 1 Then Rec.Close
    Set Rec = Nothing
    
    Exit Sub
err:
    Debug.Print err.Description
    'So That Next Entries Should Not Be Skipped
    Resume Next
    
End Sub

Private Function Saved() As Boolean
    
    If DTFrom.value >= DTTo.value Then
        MsgBox "Out Value Could Not Be Greater/Equal In Value  !", vbInformation
        DTFrom.SetFocus
        Exit Function
    End If
    
    If IsPrevTime Then
        MsgBox "You Have Already Taken Attendance For This Time !", vbInformation
        DTFrom.SetFocus
        Exit Function
    End If
    Call StartTrans(con)
    If Not IsNull(DTFrom) Then
        Call SaveNow(cmbEmp.ID, DTPicker1, Format(DTFrom, "HH:NN"))
    End If
    con.CommitTrans
    Call StartTrans(con)
    If Not IsNull(DTTo) Then
        Call SaveNow(cmbEmp.ID, DTPicker1, Format(DTTo, "HH:NN"))
    End If
    con.CommitTrans
    Call RefreshList
    
End Function

Private Function SavedOld() As Boolean

    Dim Affect As Long, First As Integer, Second As Integer, IsFirst As Boolean, IsSecond As Boolean
    Dim Deduction As Integer, Hrs As Double, TFrom As String, TTo As String
    
    Dim rs As New ADODB.Recordset
    Dim bDeductLHrs As Boolean
    Dim bExempt As Boolean
    
    With rs
        .Open "Select ExemptSettings From Employees Where EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        bExempt = !ExemptSettings & ""
        .Close
        If bExempt Then
            .Open "Select DeductLHrs From EmpSettings Where EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select DeductLHrs From GeneralSettings Where DeptID='" & DeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        If .EOF Then
            bDeductLHrs = True
        Else
            bDeductLHrs = ![DeductLHrs]
        End If
        .Close
    End With
    
    dLateHours = 0
    
    SavedOld = False

    'On Error GoTo err
    Dim SQL As String, Cond As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    End If
    If Not IsNull(DTFrom) And Not IsNull(DTTo) Then
        If IsPrevTime Then
            MsgBox "You Have Already Taken Attendance For This Time !", vbInformation
            DTFrom.SetFocus
            Exit Function
        ElseIf DTFrom.value >= DTTo.value Then
            MsgBox "Out Value Could Not Be Greater/Equal In Value  !", vbInformation
            DTFrom.SetFocus
            Exit Function
        
        End If
    
        Cond = "WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1 & "'"
    
        con.BeginTrans
        
        TFrom = Format(DTFrom, "HH:mm")
        TTo = Format(DTTo, "HH:mm")
        
        Hrs = Round(DateDiff("n", TFrom, TTo) / 60, 2)
        
        ''''''''''Will INSERT INTO Leaves'''''''''''
        
        con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,OutTime,Hrs,UserName,ComputerName) VALUES('" & cmbEmp.ID _
            & "','" & DTPicker1 & "','" & Format(DTFrom, "HH:mm") & "','" & Format(DTTo, "HH:mm") & "'," & Hrs & _
            ",'" & CurrentUserName & "','" & strComputerName & "')"
            
        If LV.ListItems.Count = 0 Then
            On Error Resume Next
            If bDeductLHrs And DateDiff("n", "08:00", Format(DTFrom, "HH:mm")) > 15 Then
                dLateHrs = DateDiff("n", "08:00", Format(DTFrom, "HH:mm")) / 60
            Else
                dLateHrs = 0
            End If
            con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance,LateHours) VALUES('" & cmbEmp.ID & "','" & DTPicker1 & "',1," & dLateHrs & ")"
            If err.Number = -2147217900 Then
                If MsgBox("This Employee is Absent on This Date. Update This Entry ?", vbQuestion + vbYesNo) = vbYes Then
                    con.Execute "Delete From AttendanceSheet Where EmpID='" & cmbEmp.ID & "' and DT='" & DTPicker1 & "'"
                    con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & cmbEmp.ID & "','" & DTPicker1 & "',1)"
                End If

            End If
        End If
        On Error GoTo err
        
        con.CommitTrans
    Else
    
        If Not IsNull(DTFrom) Then
            Call TakeAtt(cmbEmp.ID, Format(DTFrom, "HH:NN"))
        ElseIf Not IsNull(DTTo) Then
            Call TakeAtt(cmbEmp.ID, Format(DTTo, "HH:NN"))
        End If
    
    End If
    
    
    Call RefreshList
    SavedOld = True
    SQL = ""
    Exit Function
err:
    con.RollbackTrans
    MsgBox err.Description
End Function

''''''''Will Select The Attendance From AttendanceSheet'''''''''''
Private Sub CheckLeave(First As Integer, Second As Integer, Cond As String)

    First = Val(con.Execute("Select COUNT(*) FROM AttendanceSheet " & Cond & " AND Shift='First'").Fields(0) & "")
    Second = Val(con.Execute("Select COUNT(*) FROM AttendanceSheet " & Cond & " AND Shift='Second'").Fields(0) & "")

End Sub

Private Function IsPrevEntry(Cond As String) As Boolean
On Error Resume Next
Dim Shift As String
Shift = con.Execute("Select Shift FROM Leaves " & Cond & "").Fields(0)
If Shift = "" Then
    IsPrevEntry = False: Exit Function
Else
    If Shift = "Complete Day" Then
        IsPrevEntry = True: Exit Function
    End If
    If Shift = cmbShift Then
        IsPrevEntry = True: Exit Function
    End If
End If
End Function

Private Sub RefreshList()

    Dim Rec As New ADODB.Recordset, ITM As ListItem
    LV.ListItems.Clear
    With Rec
        .Open "SELECT * FROM EmpTimes WHERE DT='" & DTPicker1 & "' AND EmpID='" & cmbEmp.ID & "'", con, adOpenStatic, adLockReadOnly
        For i = 1 To .RecordCount
            Set ITM = LV.ListItems.add(, , "")
            ITM.Tag = !EntryID & ""
            ITM.ListSubItems.add , , Format(!InTime & "", "HH:NN AMPM")
            ITM.ListSubItems.add , , Format(!OutTime & "", "HH:NN AMPM")
            ITM.ListSubItems.add , , getHrsMin(Val(!Hrs & ""))
            .MoveNext
        Next
    End With

End Sub

Private Function IsPrevTime() As Boolean


Dim Temp As New ADODB.Recordset
Dim TimeFrom  As Date, TimeTo As Date, CTimeFrom As Date, CTimeTo As Date

With Temp
    .CursorLocation = adUseClient
    .Open "SELECT * FROM EmpTimes WHERE DT='" & DTPicker1 & "' AND EmpID='" & cmbEmp.ID & "' AND ISNULL(OutTime,'')<>''", con, adOpenStatic, adLockReadOnly
    
    If Not IsNull(DTFrom) Then CTimeFrom = Format(DTFrom, "HH:mm")
    If Not IsNull(DTTo) Then CTimeTo = Format(DTTo, "HH:mm")
            
    If .RecordCount > 0 Then
        For i = 1 To .RecordCount
            TimeFrom = !InTime & ""
            TimeTo = !OutTime & ""
            
            If CTimeFrom >= TimeFrom And CTimeFrom <= TimeTo Then
                IsPrevTime = True
                Exit For
            ElseIf CTimeTo >= TimeFrom And CTimeTo <= TimeTo Then
                IsPrevTime = True
                Exit For
            ElseIf CTimeFrom = TimeFrom Then
                IsPrevTime = True
                Exit For
            End If
            .MoveNext
        Next
    
    End If
    
End With

End Function

Private Sub cmbEmp_matched()
    Call RefreshList
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    
    If bOverTime Then
        If LV.ListItems.Count <> 2 Then
            MsgBox "You can Only Add OverTime.", vbInformation
            Exit Sub
        End If
    End If
    
    If Format(DTPicker1, "ddd") = "Sun" Then
        MsgBox "Can't Add Attendance For Sunday.", vbInformation
        Exit Sub
    End If
    
    If con.Execute("Select DT From Holidays Where DT='" & DTPicker1 & "'").RecordCount > 0 Then
        MsgBox "Can't Add Attendance For Holidays.", vbInformation
        Exit Sub
    End If
    
    If con.Execute("Select DT From Leaves Where EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1 & "'").RecordCount > 0 Then
        MsgBox "Can't Add Attendance, Leave Feeded.", vbInformation
        Exit Sub
    End If
    If Saved Then
         'Unload Me
        'IncreaseSal.Show 1
    End If
    
    
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Public Sub ShowMe(Optional m_bEmployees As Boolean = True, Optional m_bOverTime As Boolean)
    
    bOverTime = m_bOverTime
    bEmployees = m_bEmployees
    If bEmployees Then
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name", "Employees", "EmpID", "WHERE Active=1 AND DeptID In(Select DeptID From Departments Where TempDept=0)"
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name", "Employees", "EmpID", "WHERE Active=1 AND DeptID In(Select DeptID From Departments Where TempDept=1)"
    End If
    
    Call AddHeads
    ShowForm Me, frmFlash, , "Take Attendance"
    
End Sub


Private Sub DTPicker1_Change()
    Call RefreshList
End Sub

Private Sub Form_Load()
    
    bOverTime = False
    bEmployees = False
    'cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 7485
    
    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
     
    DTPicker1 = Date
    DTFrom = Now
    DTTo = Now
    
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub

Private Sub TakeAtt(EmpId As String, AttTime As String)

    '''''''''''   This Will Is The Employee Exists Or Not    '''''''''''

    Dim Rec As New ADODB.Recordset, CurrTime As String, Hrs As Double, SQL As String, Cond As String
    Dim TimeDiff As Integer

    Cond = "WHERE EmpID='" & EmpId & "' AND DT='" & DTPicker1 & "'"
    CurrTime = AttTime 'Format(Now, "HH:NN")

    TimeDiff = Val(con.Execute("Select TimeDiff FROM TimeSettings").Fields(0) & "")

    With Rec
        .Open "SELECT * FROM EmpTimes WHERE EntryID=(SELECT MAX(EntryID) FROM EmpTimes " & Cond & ")", con, adOpenStatic, adLockReadOnly
    
        SQL = "INSERT INTO EmpTimes(EmpID,DT,InTime) VALUES('" & EmpId _
         & "','" & DTPicker1 & "','" & CurrTime & "')"
        If .RecordCount > 0 Then
            If !OutTime & "" = "" Then
                Hrs = DateDiff("n", CDate(!InTime), CDate(CurrTime))
                If Hrs < TimeDiff Then Exit Sub
                Hrs = Round(Hrs / 60, 2)
                SQL = "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & " WHERE EntryID=" & Val(!EntryID & "") & ""
            Else
                Hrs = DateDiff("n", CDate(!OutTime), CDate(CurrTime))
                If Hrs < TimeDiff Then Exit Sub
            End If
        Else
            On Error Resume Next
            con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & EmpId & "','" & DTPicker1 & "',1)"
        
            If err.Number = -2147217900 Then
                If MsgBox("This Employee is Absent on This Date. Update This Entry ?", vbQuestion + vbYesNo) = vbYes Then
                    con.Execute "Delete From AttendanceSheet Where EmpID='" & cmbEmp.ID & "' and DT='" & DTPicker1 & "'"
                    con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & cmbEmp.ID & "','" & DTPicker1 & "',1)"
                End If

            End If
        
        End If

        con.Execute SQL
    End With

End Sub





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' This procedure is made to take attendance of all employees of vision
' technologies for the month of july 2004 automatically at default time
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Private Sub MakeAllPresentForJuly()
    
    Dim DT As Date
    DT = #7/1/2004#
    
    Dim rs As New ADODB.Recordset
    
    rs.Open "Select EmpID From Employees Order By EmpID", con, adOpenForwardOnly, adLockReadOnly
    
    For i = 1 To 31
                
        If Format(DT, "ddd") = "Sun" Then GoTo Next_day
        
        
        
        With rs
            .MoveFirst
            For j = 0 To rs.RecordCount - 1
                con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,OutTime,Hrs) VALUES('" & ![EmpId] _
                & "','" & DT & "','08:00 AM','05:00 PM',8)"
                
                con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & ![EmpId] & "','" & DT & "',1)"
                
                .MoveNext
            Next
    
        End With
        
Next_day:
    DT = DT + 1
    
    Next
    
    
    
    
End Sub
