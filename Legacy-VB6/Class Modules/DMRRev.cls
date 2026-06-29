VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMRRev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'It's For Checking The Revision Entry For Items...

'General Tab
Public GDevDesc As String
Public GUnit As String
Public GUsage As String
Public GType As String
Public GAttributes As String
Public GSize As String
Public GSizeUnit As String
Public GSteelUsed As String
Public GGuageaType As String
Public GPacknLabel As String

'Weight
Public WWeightUnit As String
Public WForgeWeight As String
Public WWasteVis As String
Public WFinWeight As String

'Pricing
Public PCurrency As String
Public PFOB As Double
Public PCIFSea As Double
Public PCIFAir As Double
Public PCnFSea As Double
Public PCnFAir As Double

'Production Processes
Public PrdPrIDs As New Collection

'FDA Processes
Public FDAPrIDs As New Collection

'Components
Public CompIDs As New Collection


