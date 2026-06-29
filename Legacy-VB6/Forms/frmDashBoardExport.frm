VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmDashBoardExport 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Delete Voucher"
   ClientHeight    =   10065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDashBoardExport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   12075
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   1530
      Index           =   10
      Left            =   120
      ScaleHeight     =   1530
      ScaleWidth      =   3915
      TabIndex        =   65
      TabStop         =   0   'False
      Tag             =   "STOCK & PROCURMENT"
      Top             =   8100
      Width           =   3915
      Begin VB.TextBox txtRebateRcvd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   810
         Width           =   1740
      End
      Begin VB.TextBox txtRebate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   810
         Width           =   1740
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Received"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   30
         Left            =   615
         TabIndex        =   70
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "REBATE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00581214&
         Height          =   375
         Index           =   34
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Width           =   3825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outstanding"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   31
         Left            =   2115
         TabIndex        =   67
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   1530
      Index           =   8
      Left            =   4095
      ScaleHeight     =   1530
      ScaleWidth      =   3870
      TabIndex        =   50
      TabStop         =   0   'False
      Tag             =   "STOCK & PROCURMENT"
      Top             =   4875
      Width           =   3870
      Begin VB.TextBox txtOutStandingQtyX2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   810
         Width           =   1710
      End
      Begin VB.TextBox txtOutStandingValueX2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   810
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   21
         Left            =   870
         TabIndex        =   55
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALUE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   20
         Left            =   2400
         TabIndex        =   54
         Top             =   495
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "OUTSTANDING ORDERS X2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00581214&
         Height          =   375
         Index           =   16
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   3870
      End
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   1530
      Index           =   7
      Left            =   135
      ScaleHeight     =   1530
      ScaleWidth      =   6690
      TabIndex        =   44
      TabStop         =   0   'False
      Tag             =   "STOCK & PROCURMENT"
      Top             =   6480
      Width           =   6690
      Begin VB.TextBox txtDueEuro 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   810
         Width           =   1545
      End
      Begin VB.TextBox txtDueGBP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   1785
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   810
         Width           =   1545
      End
      Begin VB.TextBox txtDueUSD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   810
         Width           =   1545
      End
      Begin VB.TextBox txtDuePKR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   4935
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   810
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Euro"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   27
         Left            =   3810
         TabIndex        =   59
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GBP"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   26
         Left            =   2325
         TabIndex        =   57
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "TOTAL OUTSTANDING PAYMENTS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00581214&
         Height          =   375
         Index           =   19
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   6705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USD"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   18
         Left            =   780
         TabIndex        =   48
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PKR"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   17
         Left            =   5445
         TabIndex        =   47
         Top             =   495
         Width           =   435
      End
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   1530
      Index           =   9
      Left            =   6915
      ScaleHeight     =   1530
      ScaleWidth      =   5010
      TabIndex        =   38
      TabStop         =   0   'False
      Tag             =   "STOCK & PROCURMENT"
      Top             =   6480
      Width           =   5010
      Begin VB.TextBox txtDueEuroThisMonth 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   810
         Width           =   1125
      End
      Begin VB.TextBox txtDueGBPThisMonth 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   810
         Width           =   1125
      End
      Begin VB.TextBox txtDuePKRThisMonth 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   3675
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   810
         Width           =   1125
      End
      Begin VB.TextBox txtDueUSDThisMonth 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   810
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Euro"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   29
         Left            =   2850
         TabIndex        =   63
         Top             =   495
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GBP"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   28
         Left            =   1680
         TabIndex        =   61
         Top             =   495
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "PAYMENTS DUE THIS MONTH"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00581214&
         Height          =   375
         Index           =   25
         Left            =   30
         TabIndex        =   43
         Top             =   0
         Width           =   4980
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PKR"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   24
         Left            =   4065
         TabIndex        =   42
         Top             =   495
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USD"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   22
         Left            =   660
         TabIndex        =   41
         Top             =   495
         Width           =   435
      End
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   1530
      Index           =   6
      Left            =   8055
      ScaleHeight     =   1530
      ScaleWidth      =   3870
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "STOCK & PROCURMENT"
      Top             =   4875
      Width           =   3870
      Begin VB.TextBox txtShipThisMonthQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   810
         Width           =   1710
      End
      Begin VB.TextBox txtShipThisMonthValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   810
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   15
         Left            =   825
         TabIndex        =   37
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALUE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   14
         Left            =   2400
         TabIndex        =   36
         Top             =   495
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "ORDERS TO BE SHIPPED WITHIN THIS MONTH"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00581214&
         Height          =   375
         Index           =   13
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   3870
      End
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   1530
      Index           =   5
      Left            =   135
      ScaleHeight     =   1530
      ScaleWidth      =   3870
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "STOCK & PROCURMENT"
      Top             =   4875
      Width           =   3870
      Begin VB.TextBox txtOutStandingValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   810
         Width           =   1710
      End
      Begin VB.TextBox txtOutStandingQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   810
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "OUTSTANDING ORDERS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00581214&
         Height          =   375
         Index           =   12
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   3870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALUE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   11
         Left            =   2400
         TabIndex        =   30
         Top             =   495
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   9
         Left            =   870
         TabIndex        =   29
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   2625
      Index           =   4
      Left            =   6075
      ScaleHeight     =   2625
      ScaleWidth      =   5850
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "STOCK & PROCURMENT"
      Top             =   2175
      Width           =   5850
      Begin MSComctlLib.ListView LVExports 
         Height          =   2235
         Left            =   0
         TabIndex        =   64
         Top             =   375
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S#"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Year"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Pcs"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Avg."
            Object.Width           =   1217
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Amt (USD)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Avg."
            Object.Width           =   1217
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Amt (PKR)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Avg."
            Object.Width           =   1499
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "EXPORTS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00581214&
         Height          =   375
         Index           =   8
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   5865
      End
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   2625
      Index           =   3
      Left            =   135
      ScaleHeight     =   2625
      ScaleWidth      =   5850
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "STOCK & PROCURMENT"
      Top             =   2175
      Width           =   5850
      Begin MSComctlLib.ListView LVDemands 
         Height          =   2235
         Left            =   0
         TabIndex        =   22
         Top             =   375
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S#"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Year"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Pcs"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Avg."
            Object.Width           =   1217
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Amt (USD)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Avg."
            Object.Width           =   1217
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Amt (PKR)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Avg."
            Object.Width           =   1499
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "DEMANDS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00581214&
         Height          =   375
         Index           =   10
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   5865
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   390
      Left            =   8130
      TabIndex        =   20
      Top             =   90
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   227999747
      CurrentDate     =   41125
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   1530
      Index           =   2
      Left            =   8055
      ScaleHeight     =   1530
      ScaleWidth      =   3870
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "STOCK & PROCURMENT"
      Top             =   570
      Width           =   3870
      Begin VB.TextBox txtPaymentsRcvdUSD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   810
         Width           =   1710
      End
      Begin VB.TextBox txtPaymentsRcvdPKR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   810
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USD"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   7
         Left            =   810
         TabIndex        =   19
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PKR"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   6
         Left            =   2565
         TabIndex        =   18
         Top             =   495
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "PAYMENTS RCVD THIS MONTH"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00581214&
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   3870
      End
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   1530
      Index           =   1
      Left            =   4095
      ScaleHeight     =   1530
      ScaleWidth      =   3870
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "STOCK & PROCURMENT"
      Top             =   570
      Width           =   3870
      Begin VB.TextBox txtShippedValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   810
         Width           =   1710
      End
      Begin VB.TextBox txtShippedQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   810
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "SHIPMENTS SENT THIS MONTH"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00581214&
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   3870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALUE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   2
         Left            =   2250
         TabIndex        =   12
         Top             =   495
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   1
         Left            =   690
         TabIndex        =   11
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Close"
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
      Left            =   10710
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9585
      Width           =   1215
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   1530
      Index           =   0
      Left            =   135
      ScaleHeight     =   1530
      ScaleWidth      =   3870
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "STOCK & PROCURMENT"
      Top             =   570
      Width           =   3870
      Begin VB.TextBox txtOrdRcvdQty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   810
         Width           =   1710
      End
      Begin VB.TextBox txtOrdRcvdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   810
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   3
         Left            =   690
         TabIndex        =   7
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALUE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   23
         Left            =   2400
         TabIndex        =   6
         Top             =   495
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "ORDERS RCVD THIS MONTH"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00581214&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3870
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "DASH BOARD"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00581214&
      Height          =   360
      Left            =   165
      TabIndex        =   2
      Top             =   90
      Width           =   11745
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   510
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   11820
   End
End
Attribute VB_Name = "frmDashBoardExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bMouseDown As Boolean
Const lCaptionHeight = 600
Dim lDownX As Long, lDownY As Long
Dim iColNo As Integer
Dim iColNoTemp As Integer
Dim bSaved As Boolean
Dim lEntryID As Long

Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = RGB(0, 200, 0)    'RGB(69, 140, 201) 'RGB(69, 140, 201)
    Pic.Line (0, 0)-(Pic.Width, 0), lBorderColor 'Top
    Pic.Line (0, Pic.Height - 85)-(Pic.Width, Pic.Height - 85), lBorderColor 'Bottom
    Pic.Line (0, 0)-(0, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 85, 0)-(Pic.Width - 85, Pic.Height), lBorderColor  'Right
    
End Sub

Private Sub DrawPicCaptions()

    Dim i As Integer
    Dim lBorderColor As Long, lForeColor As Long
    For i = 0 To PicWC.count - 1
'        If i = 1 Then
'            lBorderColor = RGB(153, 204, 255)
'            lForeColor = 0
'        ElseIf i = 0 Then
'            lBorderColor = RGB(255, 204, 153)
'            lForeColor = 0
'        ElseIf i = 3 Then
'            lBorderColor = RGB(255, 153, 204)
'            lForeColor = 0
'        ElseIf i = 2 Then
'            lBorderColor = RGB(0, 255, 0)
'            lForeColor = 0
'        ElseIf i = 4 Then
'            lBorderColor = RGB(200, 200, 192)
'            lForeColor = 0
'        End If
        lBorderColor = &HFFC0C0
        Call DrawCaptions(PicWC(i), lBorderColor, lForeColor)
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox, lBorderColor As Long, lForeColor As Long)

    Dim lFillColor
    'lBorderColor = 0 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    lFillColor = lBorderColor 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 150)-(Pic.Width, 150), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 300), lFillColor, BF
        Pic.ForeColor = lForeColor 'vbGreen
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 0
        Pic.Print Pic.Tag
    End If
        
    'Pic.Line (0, Pic.Height - 30)-(Pic.Width, Pic.Height - 30), lBorderColor 'Bottom
    Pic.Line (0, Pic.Height - 50)-(Pic.Width, Pic.Height), lFillColor, BF
    
    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor                          'Left
    Pic.Line (Pic.Width - 20, 0)-(Pic.Width - 30, Pic.Height), lBorderColor  'Right
    
    'Pic.Line (0, Pic.Height - 500)-(50, Pic.Height), lBorderColor
'    Me.BackColor = lBorderColor
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub DTPicker1_Change()
    Call UpdateFields
End Sub

Private Sub Form_Load()

'    Call DrawPicCaptions
    bMouseDown = False
    bSaved = False
    lEntryID = 0
    
    Dim myCTL As Control
    For Each myCTL In Controls
        If TypeOf myCTL Is VB.TextBox Then
            myCTL.Text = "-"
        End If
    Next
    DTPicker1.Value = Date
    Call DTPicker1_Change
    
End Sub


Private Sub PicWC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If y < 600 Then
        If Button = 1 Then
            If bMouseDown = False Then
                bMouseDown = True
                lDownX = x
                lDownY = y
            End If
        End If
    End If
End Sub

Private Sub PicWC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMouseDown Then
        Me.Move Me.Left + ((x - lDownX) / 15), Me.Top + ((y - lDownY) / 15)
    End If
End Sub

Private Sub PicWC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = False
End Sub

Public Function ShowMe() As Boolean
    
    'Will hold the handle to the new region.
    Dim hRegion As Long
    '
    'This API can create a Round or Oval Shaped Region.
    hRegion = CreateRoundRectRgn(0, 0, Me.Width / 15, Me.Height / 15, 20, 20)
    '
    'Change the shape of the Form to the newly created Region.
    SetWindowRgn Me.hwnd, hRegion, True
    '
    'Show the new form with the specified shape.
    'frmRegion.Show

    Me.Show
    ShowMe = bSaved
    
End Function


Private Sub UpdateFields()

    On Error GoTo err
    Dim i As Integer, ITM As ListItem, rs As New ADODB.Recordset
    Dim dExchRateExport As Double
    With rs
    
        '.Open "SELECT SUM(Qty) AS OrderQty,SUM(Qty*Price) AS OrderAmt FROM FOrderItems INNER JOIN FCustomerOrders ON FCustomerOrders.orderNo=FOrderItems.OrderNo WHERE MONTH(DT)=" & DTPicker1.Month & " AND YEAR(DT)=" & DTPicker1.year & " AND OrderType='Customer Order' AND CustCode NOT IN('Wrangler','Stock')", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT SUM(Qty) AS OrderQty,SUM(Qty*Price) AS OrderAmt FROM FOrderItems INNER JOIN FCustomerOrders ON FCustomerOrders.OrderNo=FOrderItems.OrderNo WHERE MONTH(DT)=" & DTPicker1.Month & " AND YEAR(DT)=" & DTPicker1.year & " AND OrderType='Customer Order' AND CustCode NOT IN('Wrangler','Stock','HMP') AND FCustomerOrders.OrderNo NOT IN(SELECT OrderNo FROM FCustomerFinalOrders WHERE Cancelled=1)", con, adOpenForwardOnly, adLockReadOnly
        txtOrdRcvdQty = Val(!OrderQty & "")
        txtOrdRcvdValue = Val(!OrderAmt & "")
        .Close
        
        '.Open "SELECT SUM(Qty) AS ShippedQty,SUM(Qty*Price) AS ShippedAmt FROM VCustomInvoice WHERE MONTH(DT)=" & DTPicker1.Month & " AND YEAR(DT)=" & DTPicker1.year & " AND SampleInvoice=0", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT SUM(TotalQty) AS ShippedQty,SUM(TotalAmt) AS ShippedAmt FROM VTotalExport WHERE CustCode NOT IN('HMP') AND ISNULL(GatePassNo,'')<>'' AND MONTH(DT)=" & DTPicker1.Month & " AND YEAR(DT)=" & DTPicker1.year, con, adOpenForwardOnly, adLockReadOnly
        txtShippedQty = Val(!ShippedQty & "")
        txtShippedValue = Val(!ShippedAmt & "")
        .Close
         
        .Open "SELECT SUM(AmtRcvd) AS USDRcvd,SUM(AmtRcvd*ExchRate) AS PKRRcvd FROM FCustPayments WHERE MONTH(DT)=" & DTPicker1.Month & " AND YEAR(DT)=" & DTPicker1.year, con, adOpenForwardOnly, adLockReadOnly
        txtPaymentsRcvdUSD = Val(!USDRcvd & "")
        txtPaymentsRcvdPKR = Val(!PKRRcvd & "")
        .Close
        
        .Open "SELECT SUM((Qty-ShippedQty)) AS OrderQty,SUM((Qty-ShippedQty)*Price) AS OrderAmt FROM VFOrderItemsWithShippedQty WHERE OrderNo IN(SELECT OrderNo FROM VUnShippedOrderList) AND CustCode NOT IN('Wrangler','Stock','HMP')", con, adOpenForwardOnly, adLockReadOnly
        txtOutStandingQty = Val(!OrderQty & "")
        txtOutStandingValue = Round(Val(!OrderAmt & ""))
        .Close
        
        .Open "SELECT SUM((Qty-ShippedQty)) AS OrderQty,SUM((Qty-ShippedQty)*Price) AS OrderAmt FROM VFOrderItemsWithShippedQty WHERE OrderNo IN(SELECT OrderNo FROM VUnShippedOrderList) AND CustCode IN('X2')", con, adOpenForwardOnly, adLockReadOnly
        txtOutStandingQtyX2 = Val(!OrderQty & "")
        txtOutStandingValueX2 = Round(Val(!OrderAmt & ""))
        .Close
        
        .Open "SELECT SUM(Qty) AS OrderQty,SUM(Qty*Price) AS OrderAmt FROM FOrderItems INNER JOIN FCustomerOrders ON FCustomerOrders.orderNo=FOrderItems.OrderNo WHERE MONTH(FCustomerOrders.DeliveryDT)=" & DTPicker1.Month & " AND YEAR(FCustomerOrders.DeliveryDT)=" & DTPicker1.year & " AND OrderType='Customer Order' AND CustCode NOT IN('Wrangler','Stock','HMP')", con, adOpenForwardOnly, adLockReadOnly
        txtShipThisMonthQty = Val(!OrderQty & "")
        txtShipThisMonthValue = Val(!OrderAmt & "")
        .Close
         
        Dim myDT As Date
        myDT = DateSerial(DTPicker1.year, DTPicker1.Month + 1, 0)
        
        .Open "SELECT CustomerCurrency,SUM(CommercialAmtToRcv) AS AmtToRcvUSD,SUM(CommercialAmtToRcv*CustomInvoiceAuthorization.ExchRate) AS AmtToRcvPKR FROM VPaymentStatusForDashBoard INNER JOIN CustomInvoiceAuthorization ON VPaymentStatusForDashBoard.CustomInvoice=CustomInvoiceAuthorization.CustomInvoice WHERE CommercialAmtToRcv>0 AND DateAdd(d,PaymentDays,AWBNoDT)<='" & myDT & "' GROUP BY CustomerCurrency", con, adOpenForwardOnly, adLockReadOnly
        txtDueUSDThisMonth = "0"
        txtDueGBPThisMonth = "0"
        txtDueEuroThisMonth = "0"
        txtDuePKRThisMonth = "0"
        
        Do Until .EOF
            If !CustomerCurrency & "" = "USD" Then
                txtDueUSDThisMonth = Val(!AmtToRcvUSD & "")
            ElseIf !CustomerCurrency & "" = "Euro" Then
                txtDueEuroThisMonth = Val(!AmtToRcvUSD & "")
            ElseIf !CustomerCurrency & "" = "GBP" Then
                txtDueGBPThisMonth = Val(!AmtToRcvUSD & "")
            End If
            txtDuePKRThisMonth = Val(txtDuePKRThisMonth) + Val(!AmtToRcvPKR & "")
            .MoveNEXT
        Loop
        
        .Close
        
        '.Open "SELECT CustomerCurrency,SUM(CommercialAmtToRcv) AS AmtToRcvUSD,SUM(CommercialAmtToRcv*CustomInvoiceAuthorization.ExchRate) AS AmtToRcvPKR FROM VPaymentStatusForDashBoard INNER JOIN CustomInvoiceAuthorization ON VPaymentStatusForDashBoard.CustomInvoice=CustomInvoiceAuthorization.CustomInvoice WHERE CommercialAmtToRcv>0 GROUP BY CustomerCurrency", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT Curr AS CustomerCurrency,SUM(TotalCustomAmtToRcv) AS AmtToRcvUSD,SUM(TotalCustomAmtToRcv*CustomInvoiceAuthorization.ExchRate) AS AmtToRcvPKR FROM VrptCustomInvoice LEFT OUTER JOIN CustomInvoiceAuthorization ON VrptCustomInvoice.CustomInvoice=CustomInvoiceAuthorization.CustomInvoice WHERE FreeSamples=0 GROUP BY Curr", con, adOpenForwardOnly, adLockReadOnly
        txtDueUSD = "0"
        txtDueGBP = "0"
        txtDueEuro = "0"
        txtDuePKR = "0"
        Do Until .EOF
            If Left(!CustomerCurrency & "", 3) = "EUR" Then
                txtDueEuro = Val(!AmtToRcvUSD & "")
                dExchRateExport = GetSingleDoubleValue("ExchRate", "VCurrencyExchangeRates", " WHERE Currency='EURO ¿'")
            ElseIf Left(!CustomerCurrency & "", 3) = "GBP" Then
                txtDueGBP = Val(!AmtToRcvUSD & "")
                dExchRateExport = GetSingleDoubleValue("ExchRate", "VCurrencyExchangeRates", " WHERE Currency='GBP £'")
            Else
                txtDueUSD = txtDueUSD + Val(!AmtToRcvUSD & "")
                dExchRateExport = GetSingleDoubleValue("ExchRate", "VCurrencyExchangeRates", " WHERE Currency='US $'")
            End If
            'txtDuePKR = Val(txtDuePKR) + Val(!AmtToRcvPKR & "")
            txtDuePKR = Val(txtDuePKR) + (Val(!AmtToRcvUSD & "") * dExchRateExport)
            .MoveNEXT
        Loop
        
        .Close
        
        .Open "Select SUM(Balance) AS RebateAmt From VRebateStatus Where CustomInvoice In(Select CustomInvoice From VRebateBalance Where Round(Balance,1)>0)", con_ServerSide, adOpenForwardOnly, adLockReadOnly
        txtRebate = Format(Val(!RebateAmt & ""), "#,##")
        .Close
        .Open "SELECT SUM(Amt) AS RebateRcvdAmt FROM RebateStatus", con, adOpenForwardOnly, adLockReadOnly
        txtRebateRcvd = Format(Val(!RebateRcvdAmt & ""), "#,##")
        .Close
    End With
        
        
    'Yearly Demands
    LVDemands.ListItems.Clear
    LVExports.ListItems.Clear
    Dim lYearFrom As Long
    
    For i = 2012 To DTPicker1.year + 1
        
        Set ITM = LVDemands.ListItems.add(, i & "'")
        ITM.ListSubItems.add , , i
        ITM.ListSubItems.add , , "-"
        ITM.ListSubItems.add , , "-"
        ITM.ListSubItems.add , , "-"
        
        Set ITM = LVExports.ListItems.add(, i & "'")
        ITM.ListSubItems.add , , i
        ITM.ListSubItems.add , , "-"
        ITM.ListSubItems.add , , "-"
        ITM.ListSubItems.add , , "-"
        
    Next
    
    Dim dExchRate As Double, DTForAvg As Date, iMonthsForAvg As Integer
    dExchRate = GetSingleDoubleValue("ExchRate", "CurrencyExchangeRates", " WHERE EntryID IN(SELECT MAX(EntryID) FROM CurrencyExchangeRates WHERE Currency='US $')")
    For i = 1 To LVDemands.ListItems.count
        With rs
            .Open "SELECT SUM(Qty) AS Qty,SUM(Qty*Price) AS Amount FROM FOrderItems INNER JOIN FCustomerOrders ON FCustomerOrders.OrderNo=FOrderItems.OrderNo WHERE YEAR(DT)=" & Val(LVDemands.ListItems(i).SubItems(1)) & " AND OrderType='Customer Order' AND CustCode NOT IN('Wrangler','Stock','HMP') AND FCustomerOrders.OrderNo NOT IN(SELECT OrderNo FROM FCustomerFinalOrders WHERE Cancelled=1)", con, adOpenForwardOnly, adLockReadOnly
            DTForAvg = DateSerial(Val(LVDemands.ListItems(i).SubItems(1)), 12, 1)
            
            If Date < DTForAvg Then
                iMonthsForAvg = Month(Date)
            Else
                iMonthsForAvg = 12
            End If
            
            LVDemands.ListItems(i).SubItems(2) = Val(!Qty & "")
            LVDemands.ListItems(i).SubItems(3) = Round(Val(!Qty & "") / iMonthsForAvg)
            LVDemands.ListItems(i).SubItems(4) = Round(Val(!Amount & ""))
            LVDemands.ListItems(i).SubItems(5) = Round(Val(!Amount & "") / iMonthsForAvg)
            LVDemands.ListItems(i).SubItems(6) = Round(Val(!Amount & "") * dExchRate)
            LVDemands.ListItems(i).SubItems(7) = Round((Val(!Amount & "") * dExchRate) / iMonthsForAvg)
            .Close
        End With
    Next
     
    For i = 1 To LVExports.ListItems.count
        With rs
            '.Open "SELECT SUM(Qty) AS Qty,SUM(Qty*Price) AS Amount,SUM(Qty*Price*ExchRate) AS AmountPKR FROM VCustomInvoice LEFT OUTER JOIN CustomInvoiceAuthorization ON VCustomInvoice.CustomInvoice=CustomInvoiceAuthorization.CustomInvoice WHERE YEAR(DT)=" & LVExports.ListItems(i).SubItems(1), con, adOpenForwardOnly, adLockReadOnly
            .Open "SELECT SUM(TotalQty) AS Qty,SUM(TotalAmt) AS Amount,SUM(TotalAmt*ExchRate_F) AS AmountPKR FROM VTotalExport WHERE CustCode NOT IN('HMP') AND ISNULL(GatePassNo,'')<>'' AND YEAR(DT)=" & LVExports.ListItems(i).SubItems(1), con, adOpenForwardOnly, adLockReadOnly
            DTForAvg = DateSerial(Val(LVExports.ListItems(i).SubItems(1)), 12, 1)
            
            If Date < DTForAvg Then
                iMonthsForAvg = Month(Date)
            Else
                iMonthsForAvg = 12
            End If
            
'            If Left(!Curr & "", 3) = "EUR" Then
'                dExchRateExport = GetSingleDoubleValue("ExchRate", "VCurrencyExchangeRates", " WHERE Currency='EURO ¿'")
'            ElseIf Left(!Curr & "", 3) = "GBP" Then
'                dExchRateExport = GetSingleDoubleValue("ExchRate", "VCurrencyExchangeRates", " WHERE Currency='GBP £'")
'            Else
'                dExchRateExport = GetSingleDoubleValue("ExchRate", "VCurrencyExchangeRates", " WHERE Currency='US $'")
'            End If
            
            LVExports.ListItems(i).SubItems(2) = Val(!Qty & "")
            LVExports.ListItems(i).SubItems(3) = Round(Val(!Qty & "") / iMonthsForAvg)
            LVExports.ListItems(i).SubItems(4) = Round(Val(!Amount & ""))
            LVExports.ListItems(i).SubItems(5) = Round(Val(!Amount & "") / iMonthsForAvg)
            LVExports.ListItems(i).SubItems(6) = Round(Val(!AmountPKR & ""))
            LVExports.ListItems(i).SubItems(7) = Round(Val(!AmountPKR & "") / iMonthsForAvg)
            
            .Close
        End With
    Next
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    Exit Sub

End Sub

