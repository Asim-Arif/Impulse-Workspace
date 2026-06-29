VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmPayments 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   11880
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
   ScaleHeight     =   8880
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   195
      TabIndex        =   0
      Top             =   1830
      Width           =   2040
      _ExtentX        =   3598
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
      Height          =   5565
      Left            =   195
      TabIndex        =   3
      Top             =   2280
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   9816
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16777215
      ForeColorFixed  =   12582912
      BackColorBkg    =   -2147483633
      GridColor       =   6244673
      GridColorFixed  =   6244673
      GridLines       =   3
      GridLinesFixed  =   3
      Appearance      =   0
      FormatString    =   "<      Invoice #       |<          Date          |>    Amount   "
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
   Begin VB.PictureBox cmdClose 
      Height          =   375
      Left            =   7470
      ScaleHeight     =   315
      ScaleWidth      =   1815
      TabIndex        =   1
      Top             =   8205
      Width           =   1875
   End
   Begin VB.PictureBox cmdPrint 
      Height          =   375
      Left            =   5520
      ScaleHeight     =   315
      ScaleWidth      =   1815
      TabIndex        =   4
      Top             =   8205
      Width           =   1875
   End
   Begin VB.PictureBox cmdAdd 
      Height          =   375
      Left            =   3570
      ScaleHeight     =   315
      ScaleWidth      =   1815
      TabIndex        =   9
      Top             =   8205
      Width           =   1875
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   300
      Left            =   3750
      TabIndex        =   8
      Top             =   1845
      Width           =   1950
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      DisplayStyle    =   3
      Size            =   "3440;529"
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
      Caption         =   "Invoice #"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3795
      TabIndex        =   7
      Top             =   1575
      Width           =   690
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   300
      Left            =   2265
      TabIndex        =   6
      Top             =   1845
      Width           =   1440
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      DisplayStyle    =   3
      Size            =   "2540;529"
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2340
      TabIndex        =   5
      Top             =   1575
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   2
      Top             =   1575
      Width           =   690
   End
End
Attribute VB_Name = "frmPayments"
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
    Call cmbCust.AddVals(con, Sql, "Customers")
    
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    Call ClearRows(FGrid)
    Dim Cond As String
    If cmbCust.Text = "<All>" Then
        Cond = ""
    Else
        If cmbBy.Text = "Customer" Then
            Cond = " Where CustCode='" & cmbCust.Text & "'"
        Else
            Cond = " Where Curr='" & cmbCust.Text & "'"
        End If
    End If
    
    Dim iRow As Integer
    Dim cmd As New ADODB.Command
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "VInvTotalVal"
        Set .ActiveConnection = con
        .Parameters.Refresh
        .Parameters(0).Type = adVarChar
        .Parameters(0).Size = 255
        
    End With
    Dim rsTotal As New ADODB.Recordset
    With rsTotal
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .CursorType = adOpenForwardOnly
    End With
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VTotalExport" & Cond & " Order By DT", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            iRow = .AbsolutePosition
            If iRow > 1 Then FGrid.AddItem ""
            FGrid.TextMatrix(iRow, 0) = ![CustomInvoice] & ""
            FGrid.TextMatrix(iRow, 1) = Format(![DT], "dd-MM-yyyy") & ""
            FGrid.TextMatrix(iRow, 2) = ![Curr] & ""
            cmd.Parameters(0).value = ![CustomInvoice] & ""
            Set rsTotal = cmd.Execute
            FGrid.TextMatrix(iRow, 3) = rsTotal.Fields(0) & ""
            FGrid.TextMatrix(iRow, 4) = ![CustCode] & ""
            FGrid.TextMatrix(iRow, 5) = ![FormE] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
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

Private Sub Form_Load()

    
    
    With FGrid
    
        .ColWidth(0) = 1400
        .ColWidth(1) = 1400
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1400
        .ColWidth(5) = 1600
    
    End With
    
    cmbCust.ListHeight = 1500
    'cmbCat.ListHeight = 1500
    cmbCust.AddItem "<All>", 0
    cmbCust.AddVals con, "Distinct CustCode", "Customers"
    'Call GetItems
    With cmbBy
        .Clear
        .AddItem "Customer"
        .AddItem "Currency"
    End With
End Sub

