VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4350
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4305
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6570
      Begin MSComctlLib.ProgressBar Pbar1 
         Height          =   270
         Left            =   1530
         TabIndex        =   5
         Top             =   3945
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Select Database File"
         FileName        =   "Accounts.CTW"
         Filter          =   "CTW Accounts Files|*.CTW"
      End
      Begin VB.Label LBL 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   3555
         TabIndex        =   7
         Top             =   225
         Width           =   2850
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading ..."
         Height          =   195
         Left            =   1545
         TabIndex        =   6
         Top             =   3720
         Width           =   750
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trial Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -1650
         TabIndex        =   4
         Top             =   4200
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label LblW 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   510
         TabIndex        =   1
         Top             =   2835
         Width           =   5610
      End
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   165
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   285
         Width           =   3195
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounting System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   615
         Index           =   0
         Left            =   930
         TabIndex        =   3
         Top             =   1695
         Width           =   4770
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounting System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B1BFC0&
         Height          =   615
         Index           =   1
         Left            =   967
         TabIndex        =   2
         Top             =   1725
         Width           =   4770
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NotFirstTime As Boolean
Private Sub ShowReg()

'''''''''''''''''''''''''''

'DBName = App.Path & "\Press.mdb"
'RptPath = "C:\the backup\new accounts"

   
   DBName = GetSetting(App.Title, "Settings", "DatabasePath", App.Path & "\Accounts.MDB")
   
   'DBName = "c:\Accounts.123"
   
   RptPath = App.Path

'''''''''''''''''''''''''''

TheStart:

If Not ConnectToDatabase() Then
    ServerName = InputBox("Please Enter SQL Server Name...")
      
      GoTo TheStart
 
Else
   'SaveSetting App.Title, "Settings", "DatabasePath", DBName
End If


Load frmMain


Call getFinancialYear

Dim rec As New ADODB.Recordset
rec.Open "Select * From regInfo", con, adOpenStatic
With rec
If rec.RecordCount <> 0 Then
    LBL.Caption = "Licenced To: " & ![ContPerson] & " Of " & ![CompanyName] & vbNewLine & ![Address]
    company = ![CompanyName]
    COMPANYADDRESS = Replace(![Address] & "", vbCrLf, " ")
    AutoDepriciation = ![AutoDepriciation]
End If



                
End With
'LBL.Move (Me.ScaleWidth - LBL.Width) * 4 / 5
frmMain.Show
Label1.Caption = "Loading Data ..."

Me.Show , frmMain
Me.Refresh
End Sub

Private Sub Form_Activate()
If NotFirstTime Then Exit Sub
NotFirstTime = True
Me.Refresh
With Pbar1
If .value > 0 Then Exit Sub
For i = .Min To .Max - 1 Step 3
 .value = i
 If i = 18 Then
    Label1.Caption = "Connecting To Database ..."
    Call ShowReg
 End If
 
 For j = 0 To 30000

Next
  If .value Mod 15 = 0 Then Sleep (90)
Next

Me.Hide

If IsPackageRegistered Then FrmLogin.Show 1, frmMain
    Unload Me
End With

End Sub

Private Sub Form_Load()
LblW.Caption = "Warning...." & vbNewLine & _
               "This Package is not a Free Ware and cannot be distributed illegally." & _
               "For Purchasing or Information Please find free to Contact : " & _
               "Computer Technology World, Shahab Pura Road Sialkot. Phone# : 0432-255960."
End Sub

