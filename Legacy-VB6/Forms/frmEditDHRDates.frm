VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEditDHRDates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Dates Of Steel Purchase Order"
   ClientHeight    =   2805
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5475
   Icon            =   "frmEditDHRDates.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   105
      TabIndex        =   2
      Top             =   15
      Width           =   3435
      Begin MSComCtl2.DTPicker DTPchOrder 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   390
         Width           =   2700
         _ExtentX        =   4763
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
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   50987011
         CurrentDate     =   38043
      End
      Begin MSComCtl2.DTPicker DTSaleInvoice 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1005
         Width           =   2700
         _ExtentX        =   4763
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
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   50987011
         CurrentDate     =   38043
      End
      Begin MSComCtl2.DTPicker DTTestReq 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1635
         Width           =   2700
         _ExtentX        =   4763
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
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   50987011
         CurrentDate     =   38043
      End
      Begin MSComCtl2.DTPicker DTTest 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   2235
         Width           =   2700
         _ExtentX        =   4763
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
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   50987011
         CurrentDate     =   38043
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Material Test Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   2010
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Material Test Request Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   1410
         Width           =   1965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sale Invoice Order Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   780
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Order Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   3
         Top             =   165
         Width           =   1515
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4185
      Picture         =   "frmEditDHRDates.frx":0442
      Top             =   1260
      Width           =   765
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   3630
      TabIndex        =   1
      Top             =   2385
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Cancel       "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmEditDHRDates.frx":0C93
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   360
      Left            =   3630
      TabIndex        =   0
      Top             =   1935
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Save           "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmEditDHRDates.frx":0DA5
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmEditDHRDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PInvoice As String
Public Sub EditDates(strPInvoice As String)
    
    On Error GoTo err
    PInvoice = strPInvoice
    
    Dim rs As New ADODB.Recordset
    With rs
    
        .Open "Select DT,RecDT From FSteelPurchaseOrder Where PInvoice='" & PInvoice & "'", con, adOpenForwardOnly, adLockReadOnly
        DTPchOrder = ![DT] & ""
        DTSaleInvoice = ![RecDT] & ""
        .Close
        
        .Open "Select DT,MaterialTestDT From FSteelPurchDetail1 Where PInvoice='" & PInvoice & "'", con, adOpenForwardOnly, adLockReadOnly
        DTTestReq = ![DT] & ""
        DTTest = ![MaterialTestDT] & ""
        .Close
        
    End With
    Set rs = Nothing
    
    Me.Show 1
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo err
    
    Call StartTrans(con)
    
    con.Execute "Update FSteelPurchaseOrder Set DT='" & DTPchOrder & _
     "',RecDT='" & DTSaleInvoice & "' Where PInvoice='" & PInvoice & "'"
     
    con.Execute "Update FSteelPurchDetail1 Set DT='" & DTTestReq & _
     "',MaterialTestDT='" & DTTest & "' Where PInvoice='" & PInvoice & "'"
     
    con.CommitTrans
     
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

