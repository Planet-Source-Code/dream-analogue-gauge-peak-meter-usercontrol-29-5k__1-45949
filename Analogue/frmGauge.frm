VERSION 5.00
Begin VB.Form frmGauge 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analogue Gauge / Peak Meter 1.0.0      Example By: Dream"
   ClientHeight    =   6795
   ClientLeft      =   150
   ClientTop       =   -1965
   ClientWidth     =   7830
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGauge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11.986
   ScaleMode       =   0  'User
   ScaleWidth      =   13.811
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   3720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Simulation"
      Height          =   375
      Left            =   3480
      TabIndex        =   56
      Top             =   3000
      Width           =   1695
   End
   Begin prjGauge.dGauge dGauge2 
      Height          =   630
      Left            =   2160
      TabIndex        =   54
      Top             =   960
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1111
      cBack           =   16777215
      GaugeVisible    =   -1  'True
      pFore           =   8388608
      pbBack          =   16777215
      pMxFore         =   255
      ScrollBarVisible=   -1  'True
      pntrCBCol       =   192
      BorderVis       =   0
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000007&
      Height          =   1095
      Left            =   5760
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   51
      Top             =   840
      Width           =   1335
      Begin prjGauge.dGauge dGauge4 
         Height          =   750
         Left            =   240
         TabIndex        =   52
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1323
         cBack           =   0
         GaugeVisible    =   -1  'True
         PercentBarVisible=   -1  'True
         pFore           =   0
         pbBack          =   16777215
         pMxFore         =   16576
         ggAngle         =   1
         cPntFIL         =   192
         cPntOL          =   192
         pntrCBCol       =   192
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Peak Meter"
      Height          =   375
      Left            =   5040
      TabIndex        =   38
      Top             =   3600
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   34
      Top             =   840
      Width           =   1215
      Begin prjGauge.dGauge dGauge3 
         Height          =   630
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1111
         cBack           =   16777215
         GaugeVisible    =   -1  'True
         pFore           =   8388608
         pbBack          =   16777215
         pMxFore         =   255
         ScrollBarVisible=   -1  'True
         pkMeter         =   -1  'True
         cPntFIL         =   12582912
         cPntOL          =   0
         pntrCBCol       =   12582912
         BorderVis       =   0
      End
   End
   Begin VB.PictureBox picToxic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   -3000
      Picture         =   "frmGauge.frx":1CCA
      ScaleHeight     =   156
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   106
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin prjGauge.dGauge dGauge1 
      Height          =   750
      Left            =   600
      TabIndex        =   53
      Top             =   960
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      cBack           =   16777215
      GaugeVisible    =   -1  'True
      PercentBarVisible=   -1  'True
      pFore           =   12582912
      pbBack          =   16777215
      pMxFore         =   255
      ggAngle         =   1
      cPntFIL         =   0
      cPntOL          =   0
      pntrCBCol       =   0
      BorderVis       =   0
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "240 Degree Arc"
      Height          =   255
      Left            =   5760
      TabIndex        =   55
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label57 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Vote If You Like It !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   49
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Label Label56 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label55 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Select The Width Of The Arc ie: 180 degrees or 240 degrees"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   4680
      Width           =   5415
   End
   Begin VB.Label Label52 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "180 Degree Arc"
      Height          =   255
      Left            =   3720
      TabIndex        =   46
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Hidden PercentBar)"
      Height          =   255
      Left            =   3720
      TabIndex        =   45
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Hidden ScrollBar)"
      Height          =   255
      Left            =   5640
      TabIndex        =   44
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      Caption         =   "5. Pointer Outline Color, Forecolor, Center color"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Line Line9 
      X1              =   3.175
      X2              =   8.467
      Y1              =   0.423
      Y2              =   0.423
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGauge.frx":3BB7
      Height          =   495
      Left            =   360
      TabIndex        =   42
      Top             =   4200
      Width           =   7095
   End
   Begin VB.Line Line8 
      X1              =   8.255
      X2              =   9.525
      Y1              =   10.16
      Y2              =   10.16
   End
   Begin VB.Line Line7 
      X1              =   8.255
      X2              =   9.313
      Y1              =   9.314
      Y2              =   9.314
   End
   Begin VB.Line Line6 
      X1              =   0.212
      X2              =   2.963
      Y1              =   9.102
      Y2              =   9.102
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Controlable Gauge"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Meter Level From 0 to 180 or 0 to 240 for 240 Degree Arc"
      Height          =   495
      Left            =   5520
      TabIndex        =   40
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Displayed Level On Meter"
      Height          =   255
      Left            =   5520
      TabIndex        =   39
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Returns:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4680
      TabIndex        =   37
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Inputs:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4680
      TabIndex        =   36
      Top             =   5040
      Width           =   855
   End
   Begin VB.Line Line5 
      X1              =   10.795
      X2              =   12.488
      Y1              =   8.679
      Y2              =   8.679
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Peak Meter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6120
      TabIndex        =   35
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "180 Degree Arc"
      Height          =   255
      Left            =   1920
      TabIndex        =   33
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "240 Degree Arc"
      Height          =   255
      Left            =   480
      TabIndex        =   32
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblPercent2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   31
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Percent"
      Height          =   255
      Left            =   2040
      TabIndex        =   30
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "2. PercentBar BackColor "
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Line Line4 
      X1              =   0.212
      X2              =   1.482
      Y1              =   11.43
      Y2              =   11.43
   End
   Begin VB.Line Line3 
      X1              =   0.212
      X2              =   1.27
      Y1              =   7.409
      Y2              =   7.409
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "MouseDown And MouseMove Events"
      Height          =   255
      Left            =   960
      TabIndex        =   28
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "(To Parent Container) The Value Of The Gauge (As A Percentage Of The Gauge as Integer ie: Between 0 and 100)"
      Height          =   495
      Left            =   960
      TabIndex        =   27
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Returns:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Inputs:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "3.  The Gauge Itself"
      Height          =   255
      Left            =   2760
      TabIndex        =   24
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "2. The ScrollBar"
      Height          =   255
      Left            =   1560
      TabIndex        =   23
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "1. The PercentBar"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Color Settings:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Visible Settings:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   20
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "3. PercentBar ForeColor "
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "4. PercentBarMax ForeColor (Maximum level, Above 90%)"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Hide/Show PercentBar"
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Hide/Show ScrollBar"
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Hide/Show Gauge"
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "1. UserControl BackColor (Including Gauge BackColor)"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Line Line2 
      X1              =   2.963
      X2              =   7.62
      Y1              =   4.657
      Y2              =   4.657
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Configurable Property Settings:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Line Line1 
      X1              =   0.212
      X2              =   1.27
      Y1              =   10.795
      Y2              =   10.795
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Analogue Gauge / Peak Meter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Percent"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Peak Output - 10%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "You can control the Gauge By MouseDown And Move On Any of the following:"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Scroll Value"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Degree's"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   2160
      Width           =   855
   End
End
Attribute VB_Name = "frmGauge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ####################################################################### _
  Copyright © 2002-2003 Dream-Domain.net _
  ************************************** _
  NOTE: THIS HEADER MUST STAY INTACT.
'    Terms of Agreement: _
  By using this code, you agree to the following terms... _
  1) You may use this code in your own programs (and may compile it into _
  a program and distribute it in compiled format for languages that allow _
  it) freely and with no charge, providing you notify me by e-mail. _
  2) You MAY NOT redistribute this code (for example to a web site) without _
  written permission from the original author. Failure to do so is a _
  violation of copyright laws. _
  3) You may link to this code from another website, but ONLY if it is not _
  wrapped in a frame. _
  4) You will abide by any additional copyright restrictions which the _
  author may have placed in the code or code's description. _
 **********************************
' Analogue Gauge/Peak Meter _
  ----------------- _
  By Dream _
  Date:  17th May 2003 _
  Email:  baddest_attitude@hotmail.com _
 ********************************** _
  Additional Terms of Agreement: _
  You MAY NOT Sell This Code _
  You MAY NOT Sell Any Program Containing This Code _
  You use this code knowing I hold no responsibilities for any results _
  occuring from the use and/or misuse of this code _
  If you make any improvements it would be nice if you would send me a copy. _
 ********************************** _
  Analogue Gauge/Peak Meter _
 ***********************************
' Please Comment And Vote
'Optional
Dim Tops As Long
'Optional
Dim gate As Boolean

Private Sub Command2_Click()
Timer1 = False
End Sub

' This Event Is Raised From The dGauge UserControl - REQUIRED
Private Sub dGauge1_Change()

  '####################################################################################
  'FOR A DEMO YOU CAN ADD MEDIA PLAYER CONTROL AND ADJUST THE VOLUME WITH DGAUGE1
  'Uncomment the two lines referencing the media player control
  'MediaPlayer1.Filename = App.Path & "\1.mp3"
  '####################################################################################

  '#####################################################
  '### ONLY VALUE WE WORK WITH, as Declared Public in the UserControl
  '-----------------------------------------------------
  '### The Percentage: (dGauge1.Percentage)
  '    This value controls everything! its all you need!
  '-----------------------------------------------------
  '## OUR VOLUME range in this case is from 0 to -4999 in this case so we calculate
  ' the volume From The value...    dGauge1.Percentage...
  '## THE VOLUME = -4999 + (4999 / 100 * dGauge1.Percentage) so....
   Dim TheVolume As Integer
   TheVolume = -4999 + (4999 / 100 * dGauge1.Percentage)
  ' MediaPlayer1.volume = TheVolume      'somewhere from 0 to 100
  '#####################################################
  
  ' JUST SOME LABELS - OPTIONAL -volume level/pointer
  ' angle/scrollbar value/peak volume/percent
  'Angle of pointer
  Dim Degr As Integer
   Select Case dGauge1.Percentage
         Case Is < 50
               Degr = 240 + dGauge1.Percentage * 2.42857
               Label3.Caption = Degr & " deg"
         Case Is >= 50
               Degr = (dGauge1.Percentage - 50) * 2.4
               Label3.Caption = Degr & " deg"
              'Label3.Caption = dGauge1.Angle & " deg"
         Case Else: 'blah
   End Select
  
  'Volume level
   Label2.Caption = TheVolume
   
  'ScrollBar Value: Percentage * 1% of ScrollBar Value
   Label1.Caption = dGauge1.Percentage * 2.4 & " /240"
   
  'The Percentage
   Label10.Caption = dGauge1.Percentage & " %"
   
  'Peak Volume Warning Label
   Select Case dGauge1.Percentage
          Case Is > 99
               Label8.Caption = "Peak Volume 100%"
          Case Is > 95
               Label8.Caption = "Peak Volume < 5%"
          Case Is > 90
               With Label8
                   .Visible = True
                   .Caption = "Peak Volume < 10%"
               End With
          Case Is < 91
               Label8.Visible = False
          Case Else: 'blah
   End Select
End Sub

Private Sub dGauge2_change()
  'The Percentage
   lblPercent2.Caption = dGauge2.Percentage & " %"
End Sub

'### BELOW IS FOR PEAK METER DEMO ONLY
Private Sub Command1_Click()
 MsgBox "Play a song in winamp then switch this on"
   Select Case Command1.Caption
          Case "Show Peak Meter"
               gate = False
               Call Peakout
               Command1.Caption = "Stop"
          Case "Stop"
               Call dGauge4.Change_Peek(0)
               gate = False
               Command1.Caption = "Show Peak Meter"
               lblPercent2.Caption = "0 %"
   End Select
End Sub

Private Sub Peakout()
 '##############################
 'OPTIONAL FOR DEMO
 '##############################
 Dim Current
 Dim VU As VULights
 
 Do
 Dim ValU As Single
' get volume
      VU.VolLev = volume / 327.67
      If (volume < 0) Then volume = -volume
      mxcd.dwControlID = outputVolCtrl.dwControlID
      mxcd.item = outputVolCtrl.cMultipleItems
      rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
      CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    
    ' convert volume into perc
      If Tops <> 0 Then ValU = (volume / Tops)
      
      ValU = ValU * 50
      ValU = Int(ValU)
      
    ' Make sure value is not more than 100
      If ValU > 100 Then ValU = 100
      
    ' sleep 1/10th of a second before sending new value
      Current = Timer
      Do While Timer - Current < 0.1: DoEvents: Loop
     
    'This value changes the peak meter level
    '#############################################################################
     dGauge4.Change_Peek CInt(ValU)   'somewhere from 0 to 100
    '#############################################################################

      If gate = True Then Exit Do
      DoEvents
   Loop
End Sub

Private Sub Form_Load()
   On Error Resume Next
  'Open Mixer
   rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
   If ((MMSYSERR_NOERROR <> rc)) Then
       MsgBox "Couldn't open the mixer."
       Exit Sub
   End If
   OK = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
   If (OK = True) Then
   Else
   End If
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1
   ' Set Maximum Volume
   Tops = outputVolCtrl.lMaximum
   
End Sub

Private Sub Form_Unload(cancel As Integer)
  'For Peak Meter Demo Only
   gate = False
   End
End Sub

Private Sub Timer1_Timer()
Dim a As Integer
a = Rnd * 100
 dGauge3.Change_Peek CInt(a)
End Sub
