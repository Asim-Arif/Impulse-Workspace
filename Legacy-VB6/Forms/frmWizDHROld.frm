VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWizDHROld 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CTW Generate DHR Documents Wizard..."
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmWizDHROld.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PIC 
      Height          =   3390
      Index           =   3
      Left            =   60
      ScaleHeight     =   3330
      ScaleWidth      =   7140
      TabIndex        =   19
      Top             =   30
      Visible         =   0   'False
      Width           =   7200
      Begin VB.TextBox txtItemCode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   135
         TabIndex        =   22
         Top             =   1455
         Width           =   1470
      End
      Begin MSComCtl2.DTPicker StartDT 
         Height          =   300
         Left            =   1635
         TabIndex        =   25
         Top             =   1455
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   22609923
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker FinishDT 
         Height          =   300
         Left            =   3255
         TabIndex        =   26
         Top             =   1455
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   22609923
         CurrentDate     =   37250
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "(Item 1 of 1)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   27
         Top             =   510
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Man. End Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3315
         TabIndex        =   24
         Top             =   1245
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Man. Start Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1695
         TabIndex        =   23
         Top             =   1245
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Select Manufacturing Start and End Dates."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   21
         Top             =   225
         Width           =   3570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   20
         Top             =   1245
         Width           =   750
      End
   End
   Begin VB.PictureBox PIC 
      Height          =   3390
      Index           =   4
      Left            =   60
      ScaleHeight     =   3330
      ScaleWidth      =   7140
      TabIndex        =   28
      Top             =   30
      Visible         =   0   'False
      Width           =   7200
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Wizard now have enough information to complete DHR Docuemnts."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   30
         Top             =   165
         Width           =   5670
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "- Click Finish To complete."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   75
         TabIndex        =   29
         Top             =   1905
         Width           =   1845
      End
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6045
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3420
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4650
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2025
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
   End
   Begin VB.PictureBox PIC 
      Height          =   3390
      Index           =   2
      Left            =   60
      ScaleHeight     =   3330
      ScaleWidth      =   7140
      TabIndex        =   15
      Top             =   30
      Visible         =   0   'False
      Width           =   7200
      Begin MSForms.ComboBox cmbInvNo 
         Height          =   285
         Left            =   420
         TabIndex        =   18
         Top             =   1050
         Width           =   1905
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3360;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblInvNo 
         AutoSize        =   -1  'True
         Caption         =   "- Select Proforma No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   825
         Width           =   1545
      End
      Begin VB.Label lblAdd 
         AutoSize        =   -1  'True
         Caption         =   "Generate New Export Package."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   225
         Width           =   2595
      End
   End
   Begin VB.PictureBox PIC 
      Height          =   3390
      Index           =   1
      Left            =   60
      ScaleHeight     =   3330
      ScaleWidth      =   7140
      TabIndex        =   9
      Top             =   30
      Visible         =   0   'False
      Width           =   7200
      Begin VB.OptionButton OptNew 
         Caption         =   "&Edit Export Package."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   14
         Top             =   1545
         Width           =   2640
      End
      Begin VB.OptionButton OptNew 
         Caption         =   "&Create new Export Package."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   13
         Top             =   705
         Value           =   -1  'True
         Width           =   2640
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "You can either create New set of Export Documents or Edit the saved Export Documents."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   225
         Width           =   6435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "This lets you create New Export Package from the Proforma Invocie."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   11
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "This lets you Edit the saved Export Package by using the Invoice No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   10
         Top             =   1815
         Width           =   4950
      End
   End
   Begin VB.PictureBox PIC 
      Height          =   3390
      Index           =   0
      Left            =   67
      ScaleHeight     =   3330
      ScaleWidth      =   7140
      TabIndex        =   4
      Top             =   30
      Width           =   7200
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "- Click Next To continue."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   8
         Top             =   1905
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "- Before proceeding please make sure that you have all the neccessary    material for these documents"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   75
         TabIndex        =   7
         Top             =   945
         Width           =   5175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "- This wizard lets you create all of your DHR Documents in Seconds."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   6
         Top             =   510
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Welcome to the Generate DHR Documents Wizard."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   165
         Width           =   4260
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   60
      X2              =   7265
      Y1              =   3510
      Y2              =   3510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   75
      X2              =   7265
      Y1              =   3525
      Y2              =   3525
   End
End
Attribute VB_Name = "frmWizDHROld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PicIndex As Integer
Dim Editing As Boolean
Dim SQL As String
Dim MyCol As New Collection
Dim CurrentItem As Integer
Dim TotalItems As Integer
Dim TotMinDays As Integer
Private Sub cmbinvno_Change()
    Call cmbinvno_click
End Sub

Private Sub cmbinvno_click()
    
    On Error GoTo err
    
    
    If cmbInvNo.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select MaterialTestDT From FSteelPurchDetail1 Where PInvoice='" & cmbInvNo & "'", con, adOpenForwardOnly, adLockReadOnly
        'StartDT = Date
        If .RecordCount > 0 Then
            If Not IsNull(![MaterialTestDT]) Then
                StartDT = DateAdd("d", 5, ![MaterialTestDT] & "")
            Else
                MsgBox "Test Reports Not Complete.", vbInformation
                Exit Sub
            End If
        Else
            MsgBox "Test Reports Not Complete.", vbInformation
            Exit Sub
        End If
        StartDT.MinDate = ![MaterialTestDT] & ""
        FinishDT = DateAdd("m", 1, StartDT)
        
        
        .Close
        
          
        .Open "Select DeliveryDT,ItemCode From VProformaItems Where PInvoice='" & cmbInvNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        'TotalItems = .RecordCount
        CurrentItem = 0
        
        For i = 1 To MyCol.Count
            MyCol.Remove 1
        Next i
        Dim tempstr As String
        
        Do Until .EOF
            Dim C_Items As New DHRWiz
            
            For i = 1 To MyCol.Count
                Debug.Print MyCol(i).ItemCode
                If MyCol(i).ItemCode = ![ItemCode] Then
                    GoTo MYMOVENEXT
                End If
            Next i
            
            C_Items.ItemCode = ![ItemCode]
            MyCol.Add C_Items
            Set C_Items = Nothing
            'tempstr = tempstr & "," & ![ItemCode]
MYMOVENEXT:
            .MoveNext
        Loop
        .Close
        
        TotalItems = MyCol.Count
        
        .Open "Select DT From CustomInvoice Where CustomInvoice='" & cmbInvNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        If IsNull(.Fields(0)) Then
            MsgBox "Delviery Date Not Defined In Proforma Invoice.", vbInformation
            Exit Sub
        Else
            'FinishDT.MaxDate = CheckHoliday(DateAdd("d", 1, .Fields(0) & ""))
            FinishDT.MaxDate = DateAdd("d", -2, .Fields(0) & "")
        End If
        
    End With
    
    Set rs = Nothing
    
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdBack_Click()

    PicIndex = PicIndex - 1
    
    PIC(PicIndex).Visible = True
    
    If PicIndex <> PIC.UBound Then
        PIC(PicIndex + 1).Visible = False
        cmdNext.Enabled = True
        cmdFinish.Enabled = False
    End If
    
    If PicIndex = 0 Then
        cmdBack.Enabled = False
    End If
    
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Are You Sure To Quit This Wizard?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdFinish_Click()

    On Error GoTo err
    
    Call StartTrans(con)
    
    Dim cmd As New ADODB.Command
    
    con.Execute "Delete From DHR Where PInvoice='" & cmbInvNo.Text & "'"
    con.Execute "Delete From DHRDetail Where PInvoice='" & cmbInvNo.Text & "'"
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_CalcDHR"
        For i = 1 To MyCol.Count
            .Parameters("@PInvoice").value = cmbInvNo.Text
            .Parameters("@ItemID").value = MyCol(i).ItemCode
            .Parameters("@ManStartDT").value = MyCol(i).ManStartDate
            .Parameters("@ManFinishDT").value = MyCol(i).ManEndDate
            .Execute
            
            Dim iQty As Integer
            iQty = Val(con.Execute("Select Sum(WasteQty) From DHR Where PInvoice='" & cmbInvNo & "' and ItemCode='" & MyCol(i).ItemCode & "'").Fields(0).value & "")
            con.Execute "Insert Into DHRDetail(PInvoice,ItemCode,ExtraPer) Values('" & cmbInvNo & "','" & MyCol(i).ItemCode & "'," & iQty & ")"
            
        Next i
    End With
    
    con.Execute "Insert Into DHRDetail1(PInvoice) Values('" & cmbInvNo.Text & "')"
    
    con.CommitTrans
    
    MsgBox "DHR Documents Created Successfuly.", vbInformation
    Unload Me
    
    Exit Sub
err:
    On Error Resume Next
    con.RollbackTrans
    MsgBox err.Description
    
End Sub

Private Sub cmdNext_Click()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    
    If PicIndex = 3 Then
        Dim TotHolidays As Integer
        TotHolidays = 0
        DaysInRange = DateDiff("d", StartDT, FinishDT)
        
        
        rs.Open "Select Count(*) from Holidays Where DT Between '" & StartDT & "' And '" & FinishDT & "'", con, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            TotHolidays = TotHolidays + Val(rs.Fields(0) & "")
        End If
        rs.Close
        
        TotHolidays = TotHolidays + (Int(DaysInRange / 7))
        'Dim DaysToFirstSunday As Integer
        'DaysToFirstSunday = 8 - Weekday(StartDT, vbSunday)
        
        If TotHolidays + TotMinDays > DaysInRange Then
            MsgBox "Days Are Not Enough.", vbInformation '& vbNewLine & "Atleast " & TotHolidays + TotMinDays & " Days Are Required.", vbInformation
            Exit Sub
        End If
        
        MyCol(CurrentItem + 1).ManStartDate = StartDT
        MyCol(CurrentItem + 1).ManEndDate = FinishDT
        CurrentItem = CurrentItem + 1
        
        lblStatus.Caption = "(Item " & CurrentItem + 1 & " of " & TotalItems & ")"
        
        If CurrentItem < TotalItems Then
            txtItemCode = MyCol(CurrentItem).ItemCode
            Exit Sub
        End If
        
    ElseIf PicIndex = 2 Then
    
        If cmbInvNo.MatchFound = False Then
            MsgBox "First Select Proforma No. From The List.", vbInformation
            cmbInvNo.SetFocus
            cmbInvNo.DropDown
            Exit Sub
        Else
            txtItemCode = MyCol(1).ItemCode
            lblStatus.Caption = "(Item " & CurrentItem + 1 & " of " & TotalItems & ")"
        End If
        
    ElseIf PicIndex = 1 Then
        With rs
            .Open SQL, con, adOpenForwardOnly, adLockReadOnly
            cmbInvNo.Clear
            Do Until .EOF
                cmbInvNo.AddItem .Fields(0) & ""
                .MoveNext
            Loop
            .Close
        End With
    End If
    
    PicIndex = PicIndex + 1
    
    PIC(PicIndex).Visible = True
    
    If PicIndex <> 0 Then
        PIC(PicIndex - 1).Visible = False
        cmdBack.Enabled = True
    End If
    
    If PicIndex = PIC.UBound Then
        cmdNext.Enabled = False
        cmdFinish.Enabled = True
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    
    CurrentItem = 0
    PicIndex = 0
    Editing = False
    
    
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Sum(MinDays) From FDAProcesses", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            TotMinDays = Val(.Fields(0) & "")
        End If
        .Close
        
        
    End With
    
    Set rs = Nothing
    
    
    
    SQL = "Select PInvoice From FPInvoice Where PInvoice Not In(Select PInvoice From DHR)"
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub OptNew_Click(Index As Integer)
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    If Index = 0 Then
        lblAdd.Caption = "Edit Saved Export Package."
        lblInvNo.Caption = "Select Proforma No."
        SQL = "Select PInvoice From FPInvoice Where PInvoice Not In (Select PInvoice From DHR)"
        Editing = False

    Else
        lblAdd.Caption = "Add New Export Package."
        lblInvNo.Caption = "Select Invoice No."
        SQL = "Select Distinct PInvoice From DHR"
        Editing = True
        
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

