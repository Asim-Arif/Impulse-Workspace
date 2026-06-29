VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmTotalExportOld 
   Appearance      =   0  'Flat
   ClientHeight    =   9780
   ClientLeft      =   180
   ClientTop       =   180
   ClientWidth     =   12405
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
   Icon            =   "frmTotalExportOld.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   12405
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   135
      Width           =   12375
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   12375
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Export"
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
            TabIndex        =   10
            Top             =   135
            Width           =   12225
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Export"
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
            TabIndex        =   11
            Top             =   150
            Width           =   12225
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&R e f r e s h"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8955
         TabIndex        =   1
         Top             =   855
         Width           =   1200
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   330
         Left            =   1710
         TabIndex        =   2
         Top             =   915
         Width           =   2055
         _ExtentX        =   3625
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
         Height          =   7785
         Left            =   165
         TabIndex        =   3
         Top             =   1305
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   13732
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColor       =   16777215
         ForeColorFixed  =   12582912
         BackColorBkg    =   -2147483633
         GridColor       =   6244673
         GridColorFixed  =   6244673
         GridLinesFixed  =   1
         Appearance      =   0
         FormatString    =   "< Customer |^  Country  |^ Invoice # |^ Date |^ Currency |> Amount |> Amt Rcvd  |^ Form E #"
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
      Begin MSComCtl2.DTPicker DTFrom 
         Height          =   300
         Left            =   5820
         TabIndex        =   15
         Top             =   900
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   56950787
         CurrentDate     =   37770
      End
      Begin MSComCtl2.DTPicker DTTo 
         Height          =   300
         Left            =   7380
         TabIndex        =   17
         Top             =   900
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   56950787
         CurrentDate     =   37770
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   7380
         TabIndex        =   16
         Top             =   705
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   5820
         TabIndex        =   14
         Top             =   705
         Width           =   360
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   8910
         TabIndex        =   13
         Top             =   9150
         Width           =   1665
         Caption         =   "Print            "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   360
         Left            =   10605
         TabIndex        =   12
         Top             =   9150
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   1740
         TabIndex        =   8
         Top             =   720
         Width           =   690
      End
      Begin MSForms.ComboBox cmbBy 
         Height          =   300
         Left            =   180
         TabIndex        =   7
         Top             =   915
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "By"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   690
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3810
         TabIndex        =   5
         Top             =   705
         Width           =   330
      End
      Begin MSForms.ComboBox cmbView 
         Height          =   300
         Left            =   3780
         TabIndex        =   4
         Top             =   900
         Width           =   2025
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         DisplayStyle    =   3
         Size            =   "3572;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmTotalExportOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbBy_Change()
    Call cmbBy_Click
End Sub

Private Sub cmbBy_Click()

    If Not cmbBy.MatchFound Then Exit Sub
    
    Dim Sql As String
    If cmbBy.Text = "Customer" Then
        Sql = "Distinct CustCode"
    Else
        Sql = "Distinct Curr"
    End If
    
    Call cmbCust.ClearVals
    Call cmbCust.AddItem("<All>", "0")
    Call cmbCust.AddVals(con, Sql, "ForeignCustomers")
    
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

Private Sub cmdRefresh_Click()
    
    On Error GoTo err
    Call ClearRows(FGrid)
    Dim Cond As String
    If cmbView.MatchFound = False Then cmbView.Text = "<All>"
        
    
    If cmbCust.Text = "<All>" Then
        Cond = ""
    Else
        If cmbBy.Text = "Customer" Then
            Cond = " Where CustCode='" & cmbCust.Text & "'"
        Else
            Cond = " Where Curr='" & cmbCust.Text & "'"
        End If
    End If
    
    If Cond = "" Then
        Cond = " Where "
    Else
        Cond = Cond & " and "
    End If
    Cond = Cond & " DT Between '" & DTFrom & "' and '" & DTTo & "'"
    Dim iRow As Integer
    Dim cmd As New ADODB.Command
    'Dim cmd1 As New ADODB.Command
    
    Dim TotalVal As Double, TotalRCVD As Double
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "SP_InvTotalVal"
        Set .ActiveConnection = con
        
    End With
    
    'With cmd1
    '    .CommandType = adCmdStoredProc
    '    .CommandText = "VTotalRCVD"
    '    Set .ActiveConnection = con
    '    .Parameters.Refresh
    '    .Parameters(0).Type = adVarChar
    '    .Parameters(0).Size = 255
    'End With
    
    Dim rsTotal As New ADODB.Recordset
    
    With rsTotal
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .CursorType = adOpenForwardOnly
    End With
    
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VTotalExport" & Cond & " Order By DT", con, adOpenForwardOnly, adLockReadOnly
        FGrid.Visible = False
        iRow = 0
        Do Until .EOF
        
            cmd.Parameters("@CustomInvoice").value = rs![CustomInvoice] & ""
            Set rsTotal = cmd.Execute
            TotalVal = Val(rsTotal.Fields(0))
            rsTotal.Close
            
            TotalRCVD = Val(![TotalRCVD] & "")
            
            
            If cmbView.ListIndex = 1 Then
                If Not (TotalVal > TotalRCVD) Then GoTo MoveNext
            ElseIf cmbView.ListIndex = 2 Then
                If Not (TotalVal = TotalRCVD) Then GoTo MoveNext
            End If
            
            iRow = iRow + 1
            If iRow > 1 Then FGrid.AddItem ""
            
            FGrid.TextMatrix(iRow, 0) = ![CustCode] & ""
            FGrid.TextMatrix(iRow, 1) = ![Country] & ""
            FGrid.TextMatrix(iRow, 2) = ![CustomInvoice] & ""
            FGrid.TextMatrix(iRow, 3) = Format(![DT], "dd-MM-yyyy") & ""
            FGrid.TextMatrix(iRow, 4) = ![Curr] & ""
            
            
            FGrid.TextMatrix(iRow, 5) = TotalVal
            FGrid.TextMatrix(iRow, 6) = TotalRCVD
            FGrid.TextMatrix(iRow, 7) = ![FormE] & ""
            'FGrid.TextMatrix(iRow, 8) = ![FormE] & ""
            
MoveNext:
            .MoveNext
        Loop
        FGrid.Visible = True
        .Close
    End With
    
    Set rs = Nothing
    Set rsTotal = Nothing
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub Form_Load()
    
    
    DTFrom = DateAdd("m", -1, Date)
    DTTo = Date
    With FGrid
        .ColWidth(0) = 1700
        .ColWidth(1) = 1400
        .ColWidth(2) = 1200
        .ColWidth(3) = 1400
        .ColWidth(4) = 1000
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200
        .ColWidth(7) = 1717
    End With
    
    cmbCust.ListHeight = 1500
    'cmbCat.ListHeight = 1500
    cmbCust.AddItem "<All>", 0
    cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers"
    'Call GetItems
    With cmbBy
        .Clear
        .AddItem "Customer"
        .AddItem "Currency"
    End With
    
    With cmbView
        .Clear
        .AddItem "<All>"
        .AddItem "Outstanding Payments"
        .AddItem "Completed Payments"
    End With
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub
