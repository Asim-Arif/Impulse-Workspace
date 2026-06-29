VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Object = "{D300FD6A-3B4C-404D-869C-673A361E0D43}#2.0#0"; "CTWButton.ocx"
Begin VB.Form frmRebateStatusOLD 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rebate Documents"
   ClientHeight    =   4230
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRebateStatusOLD.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9495
   Begin VB.Frame FAll 
      Height          =   4080
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   9345
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   9345
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Rebate Documents"
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
            Height          =   465
            Index           =   8
            Left            =   75
            TabIndex        =   16
            Top             =   165
            Width           =   9180
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Rebate Documents"
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
            Height          =   465
            Index           =   9
            Left            =   90
            TabIndex        =   17
            Top             =   180
            Width           =   9180
         End
      End
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E9F9FA&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2115
         ScaleHeight     =   255
         ScaleWidth      =   4395
         TabIndex        =   1
         Top             =   3030
         Visible         =   0   'False
         Width           =   4425
         Begin VB.TextBox txtEDT 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -15
            TabIndex        =   4
            Top             =   -15
            Width           =   1425
         End
         Begin VB.TextBox txtEChq 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1395
            TabIndex        =   3
            Top             =   -15
            Width           =   1815
         End
         Begin VB.TextBox txtEAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3195
            TabIndex        =   2
            Top             =   -15
            Width           =   1215
         End
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   330
         Left            =   165
         TabIndex        =   5
         Top             =   1050
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   503
         ListBackColor   =   16777215
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBackColor   =   16777215
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
      End
      Begin MSFlexGridLib.MSFlexGrid FGrid 
         Height          =   1695
         Left            =   2085
         TabIndex        =   6
         Top             =   1860
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   2990
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColor       =   16777215
         ForeColorFixed  =   8388608
         BackColorBkg    =   -2147483633
         GridColor       =   6244673
         GridColorFixed  =   6244673
         GridLinesFixed  =   1
         Appearance      =   0
         FormatString    =   "<         Date           |<          Cheque #          |>     Amount     "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HotButton.CTWButton cmdSave 
         Height          =   375
         Left            =   60
         TabIndex        =   7
         Top             =   3120
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HotPic          =   "frmRebateStatusOLD.frx":014A
         DownPic         =   "frmRebateStatusOLD.frx":0E15
         SimplePic       =   "frmRebateStatusOLD.frx":1B2F
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   300
         Left            =   210
         TabIndex        =   18
         Top             =   2670
         Visible         =   0   'False
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   50331651
         CurrentDate     =   37770
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   360
         Left            =   4890
         TabIndex        =   19
         Top             =   3615
         Width           =   1665
         Caption         =   "Close            "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   14
         Top             =   735
         Width           =   690
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   300
         Left            =   2175
         TabIndex        =   13
         Top             =   1035
         Width           =   1635
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         DisplayStyle    =   3
         Size            =   "2884;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   2190
         TabIndex        =   12
         Top             =   735
         Width           =   585
      End
      Begin MSForms.CheckBox ChkApply 
         Height          =   330
         Left            =   225
         TabIndex        =   11
         Top             =   1860
         Width           =   1710
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3016;582"
         Value           =   "0"
         Caption         =   "Applied For Rebate"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice #"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3885
         TabIndex        =   10
         Top             =   735
         Width           =   690
      End
      Begin MSForms.ComboBox cmbInvoice 
         Height          =   300
         Left            =   3840
         TabIndex        =   9
         Top             =   1035
         Width           =   1515
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         DisplayStyle    =   3
         Size            =   "2672;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Applied Date"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   285
         TabIndex        =   8
         Top             =   2430
         Visible         =   0   'False
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmRebateStatusOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EntryID As Integer
Dim EditBal As Double
Private Sub ChkApply_Click()

    If ChkApply.value Then
        DT.Visible = True
        DT.Enabled = True
        Label3(1).Visible = True
        cmdSave.Visible = True
        DT.SetFocus
    Else
        DT.Visible = False
        Label3(1).Visible = False
        cmdSave.Visible = False
    End If
    
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub cmbCountry_Click()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim iRow As Integer
    cmbCust.Tag = cmbCust.Text & cmbCountry
    With rs
        .Open "Select CustomInvoice,Applied,AppDT from VRebateAppd Where CustCode+Country='" & cmbCust.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbInvoice.Clear
        Do Until .EOF
            iRow = .AbsolutePosition - 1
            cmbInvoice.AddItem ![CustomInvoice] & ""
            'cmbInvoice.List(iRow, 1) = ![Applied]
            'cmbInvoice.List(iRow, 2) = Format(![AppDT], "dd-MM-yyyy")
            .MoveNext
        Loop
        .Close
        
        
    End With
    
    Set rs = Nothing
    Exit Sub
err:

End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Call FillCountries(cmbCust.Text, cmbCountry)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo err
    Exit Sub
    con.Execute "Delete From rptTotalExport"
    With FGrid
        For i = 1 To .Rows - 1
            'con.Execute
        Next i
    End With
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub cmbInvoice_Change()
    Call cmbInvoice_Click
End Sub

Private Sub cmbInvoice_Click()
    On Error GoTo err
    
    With cmbInvoice
        If .MatchFound = False Then Exit Sub
        
        
        'If .List(.ListIndex, 1) Then
        '    ChkApply = True
        '    txtDT.Visible = True
        '    txtDT = .List(.ListIndex, 2)
        'Else
        '    ChkApply = False
        '    'txtDT.Visible = False
        '    txtDT = ""
        'End If
        
    End With
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim IsApplied As Boolean
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdText
        .CommandText = "Select * From VRebateBalance Where CustomInvoice='" & cmbInvoice & "'"
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
    End With
    
    With rs
        
        cmbInvoice.Tag = .Fields(0) & ""
        .Close
        
        .Open "Select Applied,AppDT From Rebate Where CustomInvoice='" & cmbInvoice.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        IsApplied = ![Applied]
        
        If IsApplied Then
            txtDT = Format(![AppDT], "dd-MM-yyyy")
        End If
        
        ChkApply.value = IsApplied
        Label3(1).Visible = IsApplied
        DT.Enabled = IsApplied
        ChkApply.Locked = IsApplied
        cmdSave.Visible = False
        
        
        .Close
        
        
        .Open "Select * From RebateStatus Where CustomInvoice='" & cmbInvoice & "'", con, adOpenForwardOnly, adLockReadOnly
        Call ClearRows(FGrid)
        Dim iRow As Integer
        Do Until .EOF
            iRow = .AbsolutePosition
            If iRow > 1 Then FGrid.AddItem ""
            FGrid.TextMatrix(iRow, 0) = Format(![DT], "dd-MM-yyyy")
            FGrid.TextMatrix(iRow, 1) = ![ChqNo] & ""
            FGrid.TextMatrix(iRow, 2) = Val(![Amt])
            
            FGrid.TextMatrix(iRow, 3) = ![EntryID]
            .MoveNext
        Loop
        .Close
        
        
    End With
    Set rs = Nothing
    
    If Val(cmbInvoice.Tag) > 0 Then
        If FGrid.TextMatrix(FGrid.Rows - 1, 0) <> "" Then
            FGrid.AddItem ""
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub cmdSave_Click()

    On Error GoTo err
    
    con.Execute "Update Rebate Set Applied=1,AppDT='" & _
     DT & "' Where CustomInvoice='" & _
     cmbInvoice.Text & "'"
    cmdSave.Visible = False
    Call cmbInvoice_Click
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub FGrid_DblClick()

    Dim iRow As Integer
    
    On Error GoTo err
    iRow = FGrid.MouseRow
    
    If ChkApply.value = False Then
        MsgBox "Can't Receive Rebate." & vbNewLine & "You Have To Apply First.", vbInformation
        Exit Sub
    End If
    
    With FGrid
        
        .Row = iRow
        .Tag = iRow
        If iRow = 0 Then Exit Sub
            
        If .TextMatrix(iRow, 0) = "" Then
            txtEDT = Format(Date, "dd-MM-yyyy")
            txtEChq = ""
            txtEAmt = cmbInvoice.Tag
            EditBal = 0
        Else
            
            txtEDT = .TextMatrix(iRow, 0) & ""
            txtEChq = .TextMatrix(iRow, 1) & ""
            txtEAmt = .TextMatrix(iRow, 2) & ""
                        
            EditBal = Val(.TextMatrix(iRow, 2))
            EntryID = Val(.TextMatrix(iRow, 3))
        End If
        
        Pic.Move .Left, .Top + .CellTop
        Pic.Visible = True
        
        txtEDT.SetFocus
        
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()

    cmbCust.Enabled = True
    cmbCountry.Enabled = True
    
    cmbCust.ListHeight = 1500
    
    'cmbCust.AddItem "<All>", 0
    cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers"
    FGrid.ColWidth(3) = 0
    DT = Date
    
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub

Private Sub txtEAmt_KeyPress(KeyAscii As Integer)
On Error GoTo err

    Select Case KeyAscii
    
        Case 13
        
            If Val(txtEAmt) > Val(cmbInvoice.Tag) + EditBal Then
                MsgBox "Amount Is More Than The Balance.", vbInformation
                Exit Sub
            ElseIf Val(txtEAmt) < 1 Then
                MsgBox "Invalid Amount", vbInformation
                Exit Sub
            End If
            
            Call StartTrans(con)
            If EditBal > 0 Then
                con.Execute "Delete From  RebateStatus Where EntryID=" & EntryID
            End If
            
            con.Execute "Insert Into RebateStatus(CustomInvoice," & _
             "ChqNo,DT,Amt) Values('" & cmbInvoice & "','" & _
             txtEChq & "','" & CDate(SetSaveFormat(txtEDT)) & "'," & _
             Val(txtEAmt) & ")"
             
            con.CommitTrans
            
            With FGrid
                .TextMatrix(.Tag, 0) = txtEDT
                .TextMatrix(.Tag, 1) = txtEChq
                .TextMatrix(.Tag, 2) = txtEAmt
            End With
            Pic.Visible = False
            
            Call cmbInvoice_Click
            
        Case 27
            Pic.Visible = False
            
    End Select
    
    Call OnlyNums(KeyAscii)
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtEChq_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 13
            If txtEChq = "" Then
                MsgBox "Please Enter Cheque #.", vbInformation
                Exit Sub
            End If
            txtEAmt.SetFocus
        Case 27
            Pic.Visible = False
    End Select
    
End Sub



Private Sub txtEDT_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 13
            txtEChq.SetFocus
        Case 27
            Pic.Visible = False
    End Select
    
End Sub

Private Sub txtEDT_LostFocus()
    On Error Resume Next
    If Not SetDateFormat(txtEDT) Then
        If txtEDT <> "" Then
            txtEDT.SetFocus
        End If
    End If
    
End Sub
