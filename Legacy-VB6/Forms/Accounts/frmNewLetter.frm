VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmNewLetter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Letter"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9795
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9795
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5148
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   9570
      Begin MSComctlLib.ListView LV 
         Height          =   2895
         Left            =   6105
         TabIndex        =   11
         Top             =   1740
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Customer"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Country"
            Object.Width           =   2117
         EndProperty
      End
      Begin ComboList.Usercontrol1 cmbCompany 
         Height          =   285
         Left            =   1905
         TabIndex        =   1
         Top             =   435
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
      End
      Begin VB.TextBox txts 
         Height          =   348
         Index           =   2
         Left            =   180
         TabIndex        =   2
         Tag             =   "DebitAccName"
         Top             =   1110
         Width           =   9240
      End
      Begin VB.TextBox txts 
         Height          =   2895
         Index           =   0
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   3
         Tag             =   "DebitAccName"
         Top             =   1740
         Width           =   5910
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   330
         Left            =   180
         TabIndex        =   0
         Tag             =   "DT"
         Top             =   420
         Width           =   1710
         _ExtentX        =   3016
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   424017923
         CurrentDate     =   37250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   3
         Left            =   195
         TabIndex        =   10
         Top             =   870
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   204
         Index           =   2
         Left            =   1896
         TabIndex        =   9
         Top             =   228
         Width           =   792
      End
      Begin MSForms.CommandButton cmdSClose 
         Height          =   360
         Left            =   6105
         TabIndex        =   4
         Top             =   4635
         Width           =   1620
         Caption         =   "Save"
         PicturePosition =   327683
         Size            =   "2857;635"
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
         Left            =   7800
         TabIndex        =   5
         Top             =   4635
         Width           =   1620
         Caption         =   "Cancel            "
         PicturePosition =   327683
         Size            =   "2857;635"
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
         Caption         =   "Letter Body Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   0
         Left            =   195
         TabIndex        =   8
         Top             =   1515
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   27
         Left            =   180
         TabIndex        =   7
         Top             =   225
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmNewLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lEntryNo As String
    
Public Sub EditItem(m_lEntryNo As Long)
     

    On Error GoTo err
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    lEntryNo = m_lEntryNo
    
    With rs
        
        .Open "SELECT * FROM Letters WHERE EntryID=" & lEntryNo, con, adOpenForwardOnly, adLockReadOnly
        DT = !DT
        cmbCompany.ID = Val(!Company_RefID & "")
        txts(0) = !LetterText & ""
        
        txts(2) = !Subject & ""
        .Close
        
        
        .Open "SELECT * FROM Letters_Customers WHERE RefID=" & lEntryNo, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            For i = 1 To LV.ListItems.count
                With LV.ListItems(i)
                    If .Text = rs!CustCode & "" And .SubItems(1) = rs!Country & "" Then
                        .Checked = True
                    End If
                End With
            Next
            .MoveNEXT
        Loop
        .Close
    End With
    
    Me.Show 1
    
    
    Exit Sub
err:
    
    MsgBox err.Description
    
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSClose_Click()

    On Error GoTo err
    
    If cmbCompany.MatchFound = False Then
        MsgBox "Please Select Company.", vbInformation
        Exit Sub
    ElseIf Trim(txts(0)) = "" Then
        MsgBox "Please Enter Text.", vbInformation
        Exit Sub
    End If
    
    Dim Sql As String, strVchrNo As String, strVchrDesc As String
    Dim objVchr As New clsVouchers
    
    Call StartTrans(con)
    
    Dim lMyEntryID As Long
    If lEntryNo = 0 Then
        con.Execute "INSERT INTO Letters(Company_RefID,LetterText,DT,UserName,MachineName,Subject) VALUES(" & cmbCompany.ID & ",'" & txts(0) & "','" & DT & _
         "','" & CurrentUserName & "','" & strComputerName & "','" & "','" & txts(2) & "')"
         
        lMyEntryID = GetSingleLongValue("MAX(EntryID)", "Letters", " WHERE MachineName='" & strComputerName & "'")
    Else
        con.Execute "UPDATE Letters SET DT='" & DT & "',Company_RefID=" & cmbCompany.ID & ",LetterText='" & txts(0) & "',Subject='" & txts(2) & "' WHERE EntryID=" & lEntryNo
        con.Execute "DELETE FROM Letters_Customers WHERE RefID=" & lEntryNo
        lMyEntryID = lEntryNo
    End If
    
    Dim i As Integer
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                con.Execute "INSERT INTO Letters_Customers(RefID,CustCode,Country) VALUES(" & lMyEntryID & ",'" & .Text & "','" & .SubItems(1) & "')"
            End If
        End With
    Next
    
    con.CommitTrans
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()

    
    lEntryNo = 0
    DT = Date
    
    cmbCompany.ListHeight = 3000
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID"
                
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT CustCode,Country FROM ForeignCustomers_Marketing ORDER BY CustCode,Country", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !CustCode & "")
            ITM.ListSubItems.add , , !Country & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub
