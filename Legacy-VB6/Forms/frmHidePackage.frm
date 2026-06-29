VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmHidePackage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hide and Lock Package"
   ClientHeight    =   4845
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHidePackage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHidePackage.frx":030A
   ScaleHeight     =   4845
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Commercial Documents"
      Height          =   1950
      Left            =   2520
      TabIndex        =   32
      Top             =   2835
      Width           =   4140
      Begin VB.PictureBox PicOuter 
         Height          =   1500
         Index           =   1
         Left            =   90
         ScaleHeight     =   1440
         ScaleWidth      =   3900
         TabIndex        =   33
         Top             =   390
         Width           =   3960
         Begin VB.VScrollBar VScroll2 
            Height          =   1455
            Left            =   3645
            Max             =   2
            Min             =   1
            TabIndex        =   50
            Top             =   -15
            Value           =   1
            Width           =   270
         End
         Begin VB.PictureBox PicInner 
            BorderStyle     =   0  'None
            Height          =   2655
            Index           =   1
            Left            =   -30
            ScaleHeight     =   2655
            ScaleWidth      =   3690
            TabIndex        =   34
            Top             =   0
            Width           =   3690
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   23
               Left            =   3225
               TabIndex        =   44
               Tag             =   "GSPL"
               Top             =   1485
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   22
               Left            =   2520
               TabIndex        =   43
               Tag             =   "GSPH"
               Top             =   1485
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   21
               Left            =   3225
               TabIndex        =   42
               Tag             =   "CInspectionL"
               Top             =   1125
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   20
               Left            =   2520
               TabIndex        =   41
               Tag             =   "CInspectionH"
               Top             =   1125
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   19
               Left            =   3225
               TabIndex        =   40
               Tag             =   "CPackingListL"
               Top             =   765
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   18
               Left            =   2520
               TabIndex        =   39
               Tag             =   "CPackingListH"
               Top             =   765
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   17
               Left            =   3225
               TabIndex        =   38
               Tag             =   "CInvoiceL"
               Top             =   405
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   16
               Left            =   2520
               TabIndex        =   37
               Tag             =   "CInvoiceH"
               Top             =   405
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   15
               Left            =   3225
               TabIndex        =   36
               Tag             =   "CCoveringLetterL"
               Top             =   45
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   14
               Left            =   2520
               TabIndex        =   35
               Tag             =   "CCoveringLetterH"
               Top             =   45
               Width           =   315
            End
            Begin VB.Label Label2 
               Caption         =   "G.S.P"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   11
               Left            =   90
               TabIndex        =   49
               Top             =   1485
               Width           =   1245
            End
            Begin VB.Label Label2 
               Caption         =   "Inspection Certificate"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   10
               Left            =   90
               TabIndex        =   48
               Top             =   1125
               Width           =   1680
            End
            Begin VB.Label Label2 
               Caption         =   "Packing List"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   9
               Left            =   90
               TabIndex        =   47
               Top             =   765
               Width           =   1245
            End
            Begin VB.Label Label2 
               Caption         =   "Custom Invoice"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   8
               Left            =   90
               TabIndex        =   46
               Top             =   405
               Width           =   1245
            End
            Begin VB.Label Label2 
               Caption         =   "Covering Letter"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   7
               Left            =   75
               TabIndex        =   45
               Top             =   45
               Width           =   1245
            End
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Document Name"
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
         Index           =   1
         Left            =   120
         TabIndex        =   53
         Top             =   195
         Width           =   1395
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hide"
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
         Index           =   1
         Left            =   2520
         TabIndex        =   52
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Lock"
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
         Index           =   1
         Left            =   3195
         TabIndex        =   51
         Top             =   195
         Width           =   390
      End
   End
   Begin VB.Frame FraCustomDocs 
      Caption         =   "Custom Documents"
      Height          =   1965
      Left            =   2520
      TabIndex        =   4
      Top             =   750
      Width           =   4140
      Begin VB.PictureBox PicOuter 
         Height          =   1500
         Index           =   0
         Left            =   90
         ScaleHeight     =   1440
         ScaleWidth      =   3900
         TabIndex        =   5
         Top             =   390
         Width           =   3960
         Begin VB.PictureBox PicInner 
            BorderStyle     =   0  'None
            Height          =   2655
            Index           =   0
            Left            =   -30
            ScaleHeight     =   2655
            ScaleWidth      =   3690
            TabIndex        =   7
            Top             =   0
            Width           =   3690
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   0
               Left            =   2520
               TabIndex        =   21
               Tag             =   "CoveringLetterH"
               Top             =   45
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   1
               Left            =   3225
               TabIndex        =   20
               Tag             =   "CoveringLetterL"
               Top             =   45
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   13
               Left            =   2520
               TabIndex        =   19
               Tag             =   "InvoiceH"
               Top             =   405
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   2
               Left            =   3225
               TabIndex        =   18
               Tag             =   "InvoiceL"
               Top             =   405
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   12
               Left            =   2520
               TabIndex        =   17
               Tag             =   "PackingListH"
               Top             =   765
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   3
               Left            =   3225
               TabIndex        =   16
               Tag             =   "PackingListL"
               Top             =   765
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   11
               Left            =   2520
               TabIndex        =   15
               Tag             =   "ShippingH"
               Top             =   1125
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   4
               Left            =   3225
               TabIndex        =   14
               Tag             =   "ShippingL"
               Top             =   1125
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   10
               Left            =   2520
               TabIndex        =   13
               Tag             =   "PackingLabelsH"
               Top             =   1485
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   5
               Left            =   3225
               TabIndex        =   12
               Tag             =   "PackingLabelsL"
               Top             =   1485
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   9
               Left            =   2520
               TabIndex        =   11
               Tag             =   "FormEH"
               Top             =   1845
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   6
               Left            =   3225
               TabIndex        =   10
               Tag             =   "FormEL"
               Top             =   1845
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   8
               Left            =   2520
               TabIndex        =   9
               Tag             =   "ValuationFormH"
               Top             =   2205
               Width           =   315
            End
            Begin VB.CheckBox Chks 
               Height          =   285
               Index           =   7
               Left            =   3225
               TabIndex        =   8
               Tag             =   "ValuationFormL"
               Top             =   2205
               Width           =   315
            End
            Begin VB.Label Label2 
               Caption         =   "Covering Letter"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   0
               Left            =   75
               TabIndex        =   28
               Top             =   45
               Width           =   1245
            End
            Begin VB.Label Label2 
               Caption         =   "Custom Invoice"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   27
               Top             =   405
               Width           =   1245
            End
            Begin VB.Label Label2 
               Caption         =   "Packing List"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   2
               Left            =   90
               TabIndex        =   26
               Top             =   765
               Width           =   1245
            End
            Begin VB.Label Label2 
               Caption         =   "Shipping Instructions"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   3
               Left            =   90
               TabIndex        =   25
               Top             =   1125
               Width           =   1680
            End
            Begin VB.Label Label2 
               Caption         =   "Packing Labels"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   4
               Left            =   90
               TabIndex        =   24
               Top             =   1485
               Width           =   1245
            End
            Begin VB.Label Label2 
               Caption         =   "Form 'E'"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   5
               Left            =   90
               TabIndex        =   23
               Top             =   1845
               Width           =   1245
            End
            Begin VB.Label Label2 
               Caption         =   "Export Valuation Form"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   6
               Left            =   90
               TabIndex        =   22
               Top             =   2205
               Width           =   1755
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   1455
            Left            =   3645
            Max             =   3
            Min             =   1
            TabIndex        =   6
            Top             =   -15
            Value           =   1
            Width           =   270
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Lock"
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
         Index           =   0
         Left            =   3195
         TabIndex        =   31
         Top             =   195
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hide"
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
         Index           =   0
         Left            =   2520
         TabIndex        =   30
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Document Name"
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
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   195
         Width           =   1395
      End
   End
   Begin MSForms.CommandButton cmdOK 
      Height          =   360
      Left            =   6705
      TabIndex        =   3
      Top             =   3990
      Width           =   1755
      ForeColor       =   0
      Caption         =   "OK     "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmHidePackage.frx":26F4C
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbCustomInvoice 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   300
      Width           =   2370
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4180;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Invoice No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   2
      Top             =   60
      Width           =   2370
   End
   Begin MSForms.CommandButton cmdClose 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   6705
      TabIndex        =   1
      Top             =   4425
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Cancel          "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmHidePackage.frx":2739E
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
Attribute VB_Name = "frmHidePackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCustomInvoice_Change()
    Call cmbCustomInvoice_Click
End Sub

Private Sub cmbCustomInvoice_Click()

    On Error GoTo err
    
    If cmbCustomInvoice.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From ExportPackageLocknHide Where CustomInvoice='" & cmbCustomInvoice & "'", con, adOpenForwardOnly, adLockReadOnly
        
        For i = 0 To Chks.Count - 1
            Chks(i).value = Abs(.Fields(Chks(i).Tag))
        Next i
        
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


Private Sub cmdOK_Click()

    On Error GoTo err
    Dim strSQL As String
    strSQL = ""
    For i = 0 To Chks.Count - 1
        strSQL = strSQL & Chks(i).Tag & "=" & Chks(i).value
        If i <> Chks.Count - 1 Then strSQL = strSQL & ","
        
    Next i
    
    Call StartTrans(con)
    con.Execute "Update ExportPackageLocknHide Set " & strSQL & " Where CustomInvoice='" & cmbCustomInvoice.Text & "'"
    con.CommitTrans
    
    
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select CustomInvoice From CustomInvoice Order By CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
        cmbCustomInvoice.Clear
        Do Until .EOF
            cmbCustomInvoice.AddItem ![CustomInvoice] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub ShowForm(InvoiceNo As String)

    On Error GoTo err
    
    cmbCustomInvoice = InvoiceNo
    
    txtAmt.TabIndex = 0
    Me.Show 1
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub VScroll1_Change()

    '379
    PicInner(0).Top = -((VScroll1.value - 1) * 379)
    
End Sub

Private Sub VScroll1_Scroll()
    PicInner(0).Top = -((VScroll1.value - 1) * 379)
End Sub
Private Sub VScroll2_Change()

    '379
    PicInner(1).Top = -((VScroll2.value - 1) * 379)
    
End Sub

Private Sub VScroll2_Scroll()
    PicInner(1).Top = -((VScroll2.value - 1) * 379)
End Sub
