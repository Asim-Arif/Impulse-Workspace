VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGenExpDocs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AWM Generate Export Documents Wizard"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmGenExpDocs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Index           =   4
      Left            =   60
      ScaleHeight     =   3330
      ScaleWidth      =   7140
      TabIndex        =   30
      Top             =   30
      Visible         =   0   'False
      Width           =   7200
      Begin VB.TextBox txtSpecial 
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
         Left            =   1995
         TabIndex        =   42
         Top             =   2700
         Width           =   4350
      End
      Begin VB.TextBox txtMeasurement 
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
         Left            =   1995
         TabIndex        =   40
         Top             =   2355
         Width           =   4350
      End
      Begin VB.TextBox txtMarks 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   2010
         MultiLine       =   -1  'True
         TabIndex        =   38
         Top             =   1215
         Width           =   4335
      End
      Begin VB.TextBox txtGW 
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
         Left            =   4935
         TabIndex        =   36
         Top             =   870
         Width           =   1410
      End
      Begin VB.TextBox txtNW 
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
         Left            =   2025
         TabIndex        =   34
         Top             =   870
         Width           =   1410
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   2025
         TabIndex        =   31
         Top             =   525
         Width           =   4320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Special Instruction :"
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
         Index           =   9
         Left            =   555
         TabIndex        =   43
         Top             =   2730
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Measurement :"
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
         Index           =   8
         Left            =   900
         TabIndex        =   41
         Top             =   2400
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Marks && Number :"
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
         Index           =   7
         Left            =   705
         TabIndex        =   39
         Top             =   1215
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Gross Weight (Kg) :"
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
         Left            =   3495
         TabIndex        =   37
         Top             =   900
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Net Weight (Kg) :"
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
         Left            =   720
         TabIndex        =   35
         Top             =   900
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Enter Required Values For Shipping Instructions :"
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
         TabIndex        =   33
         Top             =   225
         Width           =   4110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Description Of Goods :"
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
         Left            =   360
         TabIndex        =   32
         Top             =   540
         Width           =   1620
      End
   End
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
      Begin VB.CheckBox chkPack 
         Caption         =   "Create Packing List (Auto)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   29
         Top             =   2850
         Width           =   5595
      End
      Begin VB.TextBox txtFormE 
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
         Left            =   4470
         TabIndex        =   26
         Top             =   765
         Width           =   1830
      End
      Begin VB.TextBox txtBL 
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
         Left            =   1305
         TabIndex        =   24
         Top             =   1185
         Width           =   1830
      End
      Begin MSComCtl2.DTPicker DTInv 
         Height          =   300
         Left            =   1305
         TabIndex        =   22
         Top             =   765
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   55443459
         CurrentDate     =   37809
      End
      Begin MSComCtl2.DTPicker DTBL 
         Height          =   300
         Left            =   4470
         TabIndex        =   28
         Top             =   1185
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   55443459
         CurrentDate     =   37809
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "AWB/BL Date :"
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
         Left            =   3360
         TabIndex        =   27
         Top             =   1230
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Form ""E"" :"
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
         Left            =   3705
         TabIndex        =   25
         Top             =   795
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "AWB/BL No. :"
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
         Left            =   255
         TabIndex        =   23
         Top             =   1215
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Invoice Date :"
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
         Left            =   210
         TabIndex        =   21
         Top             =   795
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Enter Required Values"
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
         TabIndex        =   20
         Top             =   225
         Width           =   1860
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
         Caption         =   "- This wizard lets you create all of your Export Documents in Seconds."
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
         Width           =   5025
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Welcome to the Generate Export Documents Wizard."
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
         Width           =   4455
      End
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
Attribute VB_Name = "frmGenExpDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PicIndex As Integer
Dim Editing As Boolean
Dim Sql As String

Private Sub cmbinvno_Change()
    Call cmbinvno_click
End Sub

Private Sub cmbinvno_click()
    
    On Error GoTo err
    
    If Not Editing Then Exit Sub
    If cmbInvNo.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From FPInvoice Where PInvoice='" & cmbInvNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        DTInv = ![CustomInvDT] & ""
        txtFormE = ![FormENo] & ""
        txtBL = ![BL] & ""
        DTBL = ![BLDT] & ""
        
        txtDescription = ![Goods] & ""
        txtNW = ![NetW] & ""
        txtGW = ![GrossW] & ""
        txtMarks = ![Marks] & ""
        txtMeasurement = ![Measurement] & ""
        txtSpecial = ![SpecialInstr] & ""
        
        .Close
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
    Unload Me
End Sub

Private Sub cmdFinish_Click()

    On Error GoTo err
    
    Call StartTrans(con)
    
  
    
    con.Execute "Update FPInvoice Set CustomInvDT='" & DTInv & "',FormENo='" & _
     txtFormE & "',BL='" & txtBL & "',BLDT='" & DTBL & "',Goods='" & _
     txtDescription & "',NetW='" & txtNW & "',GrossW='" & txtGW & _
     "',Marks='" & txtMarks & "',Measurement='" & txtMeasurement & _
     "',SpecialInstr='" & txtSpecial & "'," & _
     "PackGen=1 Where PInvoice='" & cmbInvNo.Text & "'"
     
    If chkPack.value = vbChecked Then
        Dim cmd As New ADODB.Command
    
        With cmd
            Set .ActiveConnection = con
            .CommandType = adCmdStoredProc
            .CommandText = "SP_CreatePList"
            .Parameters("@PInvoice").value = cmbInvNo.Text
            .Execute RecAffect
        End With
    End If
    
    con.CommitTrans
    
    MsgBox "Package Created Successfuly.", vbInformation
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdNext_Click()
    
    
    Dim rs As New ADODB.Recordset
    If PicIndex = 2 Then
        If cmbInvNo.MatchFound = False Then
            MsgBox "First Select Inv. No. From The List.", vbInformation
            cmbInvNo.SetFocus
            cmbInvNo.DropDown
            Exit Sub
        End If
    ElseIf PicIndex = 1 Then
        With rs
            .Open Sql, con, adOpenForwardOnly, adLockReadOnly
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
    
End Sub

Private Sub Form_Load()

    PicIndex = 0
    Editing = False
    Sql = "Select PInvoice,SMethod From FPInvoice Where IsNull(PackGen,0)=0"
    
End Sub

Private Sub OptNew_Click(Index As Integer)
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    If Index = 0 Then
        lblAdd.Caption = "Edit Saved Export Package."
        lblInvNo.Caption = "Select Proforma No."
        Sql = "Select PInvoice,SMethod From FPInvoice Where IsNull(PackGen,0)=0"
        Editing = False
        txtBL = ""
        txtFormE = ""
        chkPack.value = vbUnchecked
        chkPack.Visible = True
    Else
        lblAdd.Caption = "Add New Export Package."
        lblInvNo.Caption = "Select Invoice No."
        Sql = "Select PInvoice,SMethod From FPInvoice Where IsNull(PackGen,0)=1"
        Editing = True
        chkPack.value = vbUnchecked
        chkPack.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

