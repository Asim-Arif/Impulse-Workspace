VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmPorts1 
   Caption         =   "Countries , Cities & Ports..."
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   Icon            =   "frmPorts1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   420
      Top             =   5670
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":030A
            Key             =   "afghanistan"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":0624
            Key             =   "usa"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":0794
            Key             =   "City"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":0BE6
            Key             =   "Port"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":1038
            Key             =   "Globe"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":148A
            Key             =   "pakistan"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":17A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":1F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":2417
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":2939
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":2D8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":31DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":362F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":3A81
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":3ED3
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":4325
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":4777
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":4BC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":501B
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":546D
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":58BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":5D11
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":6163
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":65B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":6A07
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":6E59
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":72AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":76FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":7B4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":7FA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":83F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":8845
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":8C97
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPorts1.frx":90E9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TV1 
      Height          =   5505
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   9710
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPorts1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call FillTree
End Sub
Private Sub FillTree()

    Dim rs As New ADODB.Recordset
    
    TV1.Nodes.Clear
    TV1.Nodes.Add , , "Globe", "Globe", "Globe"
    With rs
        .Open "Select * From Countries Where CountryName <>'Pak'", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            TV1.Nodes.Add "Globe", tvwChild, ![CountryName], ![CountryName], LCase(.Fields("CountryName").value)
            .MoveNext
        Loop
        .Close
        
        .Open "Select * From Cities Where CityID<>4", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            TV1.Nodes.Add .Fields("CountryName").value, tvwChild, ![CityID] & "C", ![City] & "", "City"
            .MoveNext
        Loop
        .Close
        
        .Open "Select * From Ports Where PortID<>4", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            TV1.Nodes.Add .Fields("CityID").value & "C", tvwChild, ![PortID] & "P", ![Port] & "", "Port"
            .MoveNext
        Loop
        .Close
        
    End With
    
    Set rs = Nothing
    
End Sub

Private Sub TV1_AfterLabelEdit(Cancel As Integer, NewString As String)
    
    On Error GoTo err
    Dim iAffect As Integer
    Dim Sql As String
    
    If TV1.SelectedItem.Image = "City" Then
        Sql = "Update Cities Set City='" & NewString & "' Where CityID=" & Val(TV1.SelectedItem.Key)
    ElseIf TV1.SelectedItem.Image = "Port" Then
        Sql = "Update Ports Set Port='" & NewString & "' Where PortID=" & Val(TV1.SelectedItem.Key)
    Else
        Sql = "Update Countries Set CountryName='" & NewString & "' Where CountryName='" & TV1.SelectedItem.Key & "'"
    End If
    con.Execute Sql, iAffect
    
    If iAffect = 1 Then
        Call FillTree
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub TV1_BeforeLabelEdit(Cancel As Integer)
    If TV1.SelectedItem.Image = "Globe" Then
        Cancel = True
    End If
End Sub
