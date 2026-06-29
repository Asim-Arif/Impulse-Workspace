VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCorrectBals 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AWM Accounts Balance Correction Utility"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   330
      Left            =   60
      TabIndex        =   2
      Top             =   315
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Auto Correct Balance"
      Default         =   -1  'True
      Height          =   390
      Left            =   3075
      TabIndex        =   0
      Top             =   1545
      Width           =   1995
   End
   Begin MSComctlLib.ProgressBar PBarMain 
      Height          =   330
      Left            =   60
      TabIndex        =   3
      Top             =   855
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Overall Progress..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   4
      Top             =   1215
      Width           =   4950
   End
   Begin VB.Label Label1 
      Caption         =   "Place This File In Database Directory and Run It ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   4950
   End
End
Attribute VB_Name = "frmCorrectBals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Loaded As Boolean


Private Sub Command1_Click()

    Screen.MousePointer = vbHourglass
    Loaded = True
'    Call ReorderSno

    Dim RsAccs As New ADODB.Recordset
    Command1.Enabled = False
    
    With RsAccs
        
        .Open "Select AccNo,OpenBal,AccTitle,Balance From Accounts ", Con, adOpenForwardOnly, adLockOptimistic
        
        PBarMain.Min = 0
        PBarMain.Max = .RecordCount
        For i = 0 To .RecordCount - 1
            !Balance = AdjustBalances(![AccNo] & "", Val(!OpenBal & ""), !AccTitle)
            .MoveNEXT
            PBarMain.Value = i
        Next
         
    End With

    pbar.Value = 0
    MsgBox "Its Done...", vbInformation, "Correction Complete"
    Screen.MousePointer = vbDefault
    Unload Me

End Sub

Private Sub Form_Activate()
    If Not Loaded Then Call Command1_Click
End Sub

Private Sub Form_Load()

    On Error GoTo err
    
    pbar.Value = 0
    
    
    Exit Sub
err:
    MsgBox "Cannot Connect To DataBase." & vbNewLine & " AWM Accounting System May Not Be Installed."
    End
End Sub


Public Function GetMax(Field As String, Table As String, Optional Condition As String) As String

    Dim recmax As New ADODB.Recordset
    recmax.Open "Select max(" & Field & ") from " & Trim(Table) & " " & Condition, Con, adOpenForwardOnly, adLockReadOnly
    If Not IsNull(recmax.Fields(0)) Then
        GetMax = recmax.Fields(0)
    Else
        GetMax = 0
    End If
    recmax.Close
    Set recmax = Nothing
    
End Function

Private Sub ReorderSno()

    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim NewSno As Double
    
    With rs
        .Open "Select Sno,Vdate From VLedger Where  Convert(varchar(15),GetDate(),12)<>Cast(left(Sno,5) As Int)", Con, adOpenForwardOnly, adLockOptimistic
        Label1.Caption = "Verifying Indexing Order"
   
        If .RecordCount > 0 Then pbar.Max = .RecordCount
        For i = 0 To .RecordCount - 1
            Cnd = " Where Sno between " & Val(Format(![Vdate], "yyMMdd") & "0000") & " and " & Val(Format(![Vdate], "yyMMdd") & "9999")
            NewSno = GetMax("Sno", "Vouchers", CStr(Cnd)) + 1
            If NewSno = 1 Then NewSno = Val(Format(![Vdate], "yyMMdd") & "0001")
      
            ![Sno] = NewSno
            .Update
            pbar = i
            .MoveNEXT
        Next
        'MsgBox .RecordCount & " Records  Found To Be Reordered ," & .RecordCount & " Records Reordered ... "
        .Close
        Set rs = Nothing
    End With
    
End Sub

Private Function AdjustBalances(ByVal AccNo As String, ByVal OpBal As Double, ByVal AccTitle As String) As Double

    Dim rsV As New ADODB.Recordset
    Dim Balance  As Double
   
    With rsV
   
        'Get Opening Balance
      
'        .Open "Select OpenBal,Acctitle From Accounts Where Accno='" & AccNo & "'", Con, adOpenForwardOnly, adLockReadOnly
'        If .RecordCount = 0 Then Exit Sub
        Balance = OpBal 'Val(![OpenBal] & "")
        Label1.Caption = "Correcting """ & AccTitle & """"   '"Correcting """ & ![AccTitle] & """"
        Me.Refresh
'        .Close
        pbar.Max = 100
        pbar.Min = 0
        pbar.Value = 0
        'Correct The Balances
        
        .Open "Select AccNo,VchrNo,Sno,Debit,Credit,Balance From Vouchers Where Accno='" & AccNo & "' Order By SNo", Con, adOpenForwardOnly, adLockOptimistic
        If .RecordCount > 0 Then
            pbar.Max = .RecordCount
            pbar.Min = 0
            For i = 0 To .RecordCount - 1
                Balance = Balance + Val(![Debit] & "") - Val(![Credit] & "")
                ![Balance] = Balance
                .MoveNEXT
                pbar = i
            Next
        End If
        .Close
    End With
    Set rsV = Nothing
    'Con.Execute "Update Accounts Set Balance=" & Balance & " Where AccNo='" & AccNo & "'"
    AdjustBalances = Balance
   
End Function
