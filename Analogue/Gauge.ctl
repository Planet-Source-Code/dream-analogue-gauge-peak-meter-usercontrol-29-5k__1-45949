VERSION 5.00
Begin VB.UserControl dGauge 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   55
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   132
   ToolboxBitmap   =   "Gauge.ctx":0000
   Begin VB.PictureBox pic240 
      Height          =   135
      Left            =   840
      Picture         =   "Gauge.ctx":0312
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   240
      Width           =   135
   End
   Begin VB.PictureBox pic180 
      Height          =   135
      Left            =   840
      Picture         =   "Gauge.ctx":1FDC
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   120
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   135
      Left            =   1080
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   1080
      Max             =   240
      SmallChange     =   5
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   0
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   0
      Top             =   0
      Width           =   720
      Begin VB.Shape Shape1 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000001&
         Height          =   75
         Left            =   720
         Shape           =   3  'Circle
         Top             =   720
         Width           =   135
      End
   End
End
Attribute VB_Name = "dGauge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Option Compare Text

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
' Analogue Gauge/Peak Meter And Graph _
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
  Analogue Gauge/Peak Meter And Graph _
 ***********************************
' Please Comment And Vote
' #######################################################################

'For Peak Meter
Private Declare Sub Sleep Lib "kernel32" ( _
              ByVal dwMilliseconds As Long)  'makes it sleep

Private Declare Function Polygon Lib "gdi32" ( _
                   ByVal hdc As Long, _
                 lpPoint As POINTAPI, _
                   ByVal nCount As Long) As Long

Private Declare Function GetSysColor Lib "user32" ( _
                   ByVal nIndex As Long) As Long

'### When this event occurs it is sent back to the parent container ###'
Public Event change()

'Default property values
Const M_PEAK_VIS = False
Const M_PIC_VIS = False
Const M_GAUGE_VIS = False
Const M_SCROLL_VIS = False
Const M_BORDER_VIS = 1
Const ANGLE_180 = 0
Const LAB_MAKER = "Dream."
Const M_CREATOR = "By: Dream Copyright © 2002-2003 Dream-Domain.net"

Const pi                    As Single = 3.141593 '3.14159265358979
Const RADS                  As Single = pi / 180

'Sin/Cos look-up table
Private Rotate(359, 1)      As Single            'how many degrees on a compass? lol
Private OldX                As Integer           'mousemovement
Private OldY                As Integer           'mousemovement
Private Radius              As Integer
Private GaugeRadius         As Integer           'pictureback height/2
Private gValue              As Integer           'gauge value
Private Jump                As Integer           'Amount of change per mousemove
Private Moover              As Boolean           'mousedown occurs

'Property variables
Private bPicVis             As Boolean           'percentbar visible
Private bGaugeVis           As Boolean           'gauge visible
Private bScrollVis          As Boolean           'scrollbar visible
Private bPeakMode           As Boolean           'In Peak Mode
Private dPreGaugeArc        As Integer           'Preset Gauge Arc 180 or 240

'For Peak Meter
Public NewPeak              As Integer           'for peak meter display/fade etc
Public OldPeak              As Integer           'for peak meter display/fade etc

Public Percentage           As Integer           ' < The data available to the
                                                 '   parent container
Private Creator             As Variant           ' Me !

Private cColor              As Long              'control/gauge/percentbar back color
Private picMaxColor         As Long              'percentbarMax forecolor
Private picColor            As Long              'percentbar forecolor
Private pcBColor            As Long              'percentbar backcolor
Private pntFILLC            As Long              'pointer fill color
Private pntOLC              As Long              'pointer outline color
Private pntCENTClr          As Long              'centre color

Private ptSecond()          As POINTAPI          'Polygons point to 90 degree's
Private ptNewSecond()       As POINTAPI          'Rotated polygons

Private Type POINTAPI
   X As Long
   Y As Long
End Type
 
Private Type UserControlProps
    gAngle                  As GaugeAngleConstants 'meter angle
    bBorder                 As BorderStyleConstants
End Type

Private myProps             As UserControlProps  ' cached button properties

Public Enum BorderStyleConstants                 ' border styles
    HideBorder = 0
    ShowBorder = 1
End Enum
Public Enum GaugeAngleConstants                  ' angle styles
    A180 = 0
    A240 = 1
End Enum

Public Property Let BackColor(nColor As OLE_COLOR)
  
  'Sets the backcolor of the control/gauge/progress meter
   cColor = ConvertColor(nColor)
   picBack.BackColor = nColor
   UserControl.BackColor = nColor
   
   PropertyChanged "cBack"

End Property

Public Property Get BackColor() As OLE_COLOR
   
   BackColor = cColor

End Property

Public Property Let BorderVisible(style As BorderStyleConstants)
   
    myProps.bBorder = style
    
    Select Case myProps.bBorder
           Case HideBorder
                UserControl.BorderStyle = 0
           Case ShowBorder
                UserControl.BorderStyle = 1
           Case Else:  'blah@microsoft
    End Select
    
    PropertyChanged "BorderVis"

End Property

Public Property Get BorderVisible() As BorderStyleConstants
    
    BorderVisible = myProps.bBorder

End Property
Private Function ConvertColor(tColor As Long) As Long
  
  'Converts VB color constants to real color values
   If tColor < 0 Then
      ConvertColor = GetSysColor(tColor And &HFF&)
   Else
      ConvertColor = tColor
   End If

End Function

Public Property Let Created(ByVal New_Creator As Variant)
  
  'displays created info in usercontrol properties
    err.Raise 382

End Property

Public Property Get Created() As Variant
    
    Created = Creator

End Property

Public Property Let GaugeAngle(style As GaugeAngleConstants)

  'sets up the angle of the arc the gauge will swing
   myProps.gAngle = style
   
   Select Case myProps.gAngle
       Case A180
            picBack.Picture = pic180
            HScroll1.Top = 30
            Picture1.Top = 40
            dPreGaugeArc = 180
       Case A240
            picBack.Picture = pic240
            HScroll1.Top = 38
            Picture1.Top = 48
            dPreGaugeArc = 240
       Case Else:  'blah@microsoft
   End Select
   
   With picBack
    .Top = 0
    .Left = 0
   End With
   
   HScroll1.Left = 0
   HScroll1.Max = dPreGaugeArc
   Picture1.Left = 0
   
   UserControl_Resize
   
   PropertyChanged "ggANgle"
   
End Property

Public Property Get GaugeAngle() As GaugeAngleConstants
   
   GaugeAngle = myProps.gAngle

End Property

Public Property Let GaugeVisible(ByVal New_GaugeVisible As Boolean)
   
   'Gauge visible true/false
    bGaugeVis = New_GaugeVisible
    picBack.Visible = bGaugeVis
    PropertyChanged "GaugeVisible"

End Property

Public Property Get GaugeVisible() As Boolean
    
    GaugeVisible = bGaugeVis

End Property

Public Property Let PeakMode(ByVal New_PeakMeter As Boolean)
    
   'set as peakmeter true/false
    bPeakMode = New_PeakMeter
    PropertyChanged "pkMeter"

End Property

Public Property Get PeakMode() As Boolean
    
    PeakMode = bPeakMode

End Property

Public Property Let PercentBarBackColor(pbColor As OLE_COLOR)
  
  'Sets the backcolor of the progress meter
   pcBColor = ConvertColor(pbColor)
   Picture1.BackColor = pbColor
   
   PropertyChanged "pbBack"

End Property

Public Property Get PercentBarBackColor() As OLE_COLOR
   
   PercentBarBackColor = pcBColor

End Property

Public Property Let PercentBarForeColor(pColor As OLE_COLOR)
  
  'Sets the forecolor of the progress meter
   picColor = ConvertColor(pColor)
   Picture1.ForeColor = pColor
   
   PropertyChanged "pFore"

End Property

Public Property Get PercentBarForeColor() As OLE_COLOR
   
   PercentBarForeColor = picColor

End Property

Public Property Let PercentBarMaxForeColor(pMColor As OLE_COLOR)
  
  'Sets the MAX forecolor of the progress meter
   picMaxColor = ConvertColor(pMColor)
   Picture1.ForeColor = pMColor
   
   PropertyChanged "pMxFore"

End Property

Public Property Get PercentBarMaxForeColor() As OLE_COLOR
   
   PercentBarMaxForeColor = picMaxColor

End Property

Public Property Let PercentBarVisible(ByVal New_PercentBar As Boolean)
   
   'Progress bar visible true/false
    bPicVis = New_PercentBar
    Picture1.Visible = bPicVis
    UserControl_Resize
    
    PropertyChanged "PercentBarVisible"

End Property

Public Property Get PercentBarVisible() As Boolean
    
    PercentBarVisible = bPicVis

End Property

Public Property Let PointerCenterColor(eColor As OLE_COLOR)
  
  'Sets the centrecolor of the gauge
   pntCENTClr = ConvertColor(eColor)
   Shape1.BackColor = eColor
   Shape1.BorderColor = eColor
   
   PropertyChanged "pntrCBCol"

End Property

Public Property Get PointerCenterColor() As OLE_COLOR
   
   PointerCenterColor = pntCENTClr

End Property

Public Property Let PointerFillColor(wColor As OLE_COLOR)
  
  'Sets the fill color of the pointer
   pntFILLC = ConvertColor(wColor)
   DrawPolygon ptNewSecond()
   
   PropertyChanged "cPntFIL"

End Property

Public Property Get PointerFillColor() As OLE_COLOR
   
   PointerFillColor = pntFILLC

End Property

Public Property Let PointerOutlineColor(dColor As OLE_COLOR)
  
  'Sets the backcolor of the control/gauge/progress meter
   pntOLC = ConvertColor(dColor)
   DrawPolygon ptNewSecond()
   
   PropertyChanged "cPntOL"

End Property

Public Property Get PointerOutlineColor() As OLE_COLOR
   
   PointerOutlineColor = pntOLC

End Property

Public Property Let ScrollBarVisible(ByVal New_ScrollVis As Boolean)
   
   'sets whether or not the srcoll bar is visible  true/false
    bScrollVis = New_ScrollVis
    HScroll1.Visible = bScrollVis
    UserControl_Resize
    
    PropertyChanged "ScrollBarVisible"

End Property

Public Property Get ScrollBarVisible() As Boolean
    
    ScrollBarVisible = bScrollVis

End Property

Private Sub UserControl_InitProperties()
   
   bPicVis = True
   bGaugeVis = True
   bScrollVis = True
   cColor = 16777215
   picMaxColor = 255
   picColor = 8388608
   pcBColor = 16777215
   bPeakMode = False
   
   Me.GaugeAngle = ANGLE_180
   Me.PointerCenterColor = 3
   Me.PointerOutlineColor = 223
   Me.PointerFillColor = 3
   Me.BorderVisible = M_BORDER_VIS

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        Creator = .ReadProperty("Created", M_CREATOR)
        Me.BackColor = .ReadProperty("cBack", Parent.BackColor)
        Me.GaugeVisible = .ReadProperty("GaugeVisible", M_GAUGE_VIS)
        Me.PercentBarForeColor = .ReadProperty("pFore", vbBlue)
        Me.PercentBarBackColor = .ReadProperty("pbBack", Parent.BackColor)
        Me.PercentBarMaxForeColor = .ReadProperty("pMxFore", vbRed)
        Me.PercentBarVisible = .ReadProperty("PercentBarVisible", M_PIC_VIS)
        Me.ScrollBarVisible = .ReadProperty("ScrollBarVisible", M_SCROLL_VIS)
        bPeakMode = .ReadProperty("pkMeter", M_PEAK_VIS)
        Me.GaugeAngle = .ReadProperty("ggANgle", ANGLE_180)
        Me.PointerCenterColor = .ReadProperty("pntrCBCol", 3)
        Me.PointerOutlineColor = .ReadProperty("cPntOL", 223)
        Me.PointerFillColor = .ReadProperty("cPntFIL", 3)
        Me.BorderVisible = .ReadProperty("BorderVis", M_BORDER_VIS)
   End With
   
   DisplayPeak 1
   
   If bPeakMode = True Then HScroll1.Visible = False

End Sub

Private Sub UserControl_Resize()
  
  With UserControl
       Select Case myProps.gAngle
              Case A180                    'If Gauge Angle 180 then..
                   Select Case bScrollVis
                          Case False       'If scrollbar visible = false
                               If bPicVis = False Then  'If percent bar visible = false
                                 .Height = 480
                                Else
                                 .Height = 625         'If percent bar visible = true
                                  Picture1.Top = 30
                               End If
                          Case True        'If scrollbar visible =  true
                               If bPicVis = False Then 'If percent bar visible = false
                                 .Height = 630
                                Else                  'If percent bar visible = true
                                 .Height = 755
                                  Picture1.Top = 40
                               End If
                  End Select
              Case A240                    'If Gauge Angle 240 then..
                   Select Case bScrollVis
                          Case False      'If scrollbar visible = false
                               If bPicVis = False Then 'If percent bar visible = false
                                 .Height = 585
                                Else                   'If percent bar visible = true
                                 .Height = 745
                                  Picture1.Top = 38
                               End If
                          Case True       'If scrollbar visible = true
                               If bPicVis = False Then 'If percent bar visible = false
                                 .Height = 745
                                Else                 'If percent bar visible = true
                                 .Height = 875
                                  Picture1.Top = 48
                                End If
                   End Select
              Case Else:  'blah@microsoft
       End Select
      .Width = 755
  End With
   
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   
   With PropBag
       .WriteProperty "cBack", cColor
       .WriteProperty "Created", Creator, M_CREATOR
       .WriteProperty "GaugeVisible", bGaugeVis, M_GAUGE_VIS
       .WriteProperty "PercentBarVisible", bPicVis, M_PIC_VIS
       .WriteProperty "pFore", picColor
       .WriteProperty "pbBack", pcBColor
       .WriteProperty "pMxFore", picMaxColor
       .WriteProperty "ScrollBarVisible", bScrollVis, M_SCROLL_VIS
       .WriteProperty "pkMeter", bPeakMode, M_PEAK_VIS
       .WriteProperty "ggAngle", myProps.gAngle, ANGLE_180
       .WriteProperty "cPntFIL", pntFILLC, 3
       .WriteProperty "cPntOL", pntOLC, 223
       .WriteProperty "pntrCBCol", pntCENTClr, 3
       .WriteProperty "BorderVis", myProps.bBorder, M_BORDER_VIS
   End With

End Sub
 
Private Sub BuildPolygon()
   Dim n             As Long
   Dim m_HandRadius  As Long
   
   n = IIf(0, 18, 4)
  
  'Points to 3Hr mark (90°)
   ReDim ptSecond(n)
   ReDim ptNewSecond(n)
 
  'Define upper half of pointer
   ptSecond(0).X = -GaugeRadius * 0.01 'pointer tail
   ptSecond(1).Y = GaugeRadius * 0.1   'width of pointer
   ptSecond(2).X = GaugeRadius * 0.7   'Length of pointer
  
  'Replicate upper half to bottom
   MirrorVerticals 3, 4, 4 'From, To, Index

End Sub

Public Sub Change_Peek(Peak As Integer)
  Dim a As Integer
    Peak = (dPreGaugeArc / 100) * Peak
  'For The Peak Meter We Use The Variable 'Percentage' but now its value is
  'from 0 to 180 ( or 0 to 240 ) depending on which control you use,
  '... the 180 Arc or the 240 Arc
  Select Case OldPeak - Peak
  Case Is > 0
   For a = OldPeak To Peak Step -15
       DisplayPeak a
       Sleep 3
   Next a
  Case Is < 0
   For a = OldPeak To Peak Step 15
       DisplayPeak a
       Sleep 3
   Next a
  End Select
    OldPeak = Peak
End Sub

Private Sub DisplayPeak(Peak As Integer)
 Dim angle As Integer
 
  'gValue(Our Meter Level) must now be same as Peak.
   gValue = Peak
   
  'So too the scrollbar value...
  '#####################################################################
  'IMPORTANT NOTE: When there is a change in the scrollbar value, we exit this sub,
  ' as the HScroll1_Change() Sub will call this sub again, in which case there is no
  ' change in the HScroll1.Value as we adjusted as previously mentioned above.
  '(otherwise we would get a stack overflow error from the two subs calling eachother)
  
  'DO NOT ALTER OR REMOVE THESE LINE'S
   If bPeakMode = False Then
      If HScroll1.Value <> Peak Then HScroll1.Value = Peak: Exit Sub
   End If
  '#####################################################################
   
  'Work out the angle of the Gauge Pointer(180 arc or 240)
   Select Case myProps.gAngle
          Case A180
               If Peak + 270 > 359 Then
                  angle = Peak - 90
                Else
                  angle = Peak + 270
               End If
          Case A240
               If Peak + 240 > 359 Then
                  angle = Peak - 120
                Else
                  angle = Peak + 240
               End If
   End Select
   
  'Set up the angle
   RotatePoints ptSecond, ptNewSecond, angle
   
  'Clear prior paints.
   picBack.Cls
   
  'Draw the new pointer direction
   DrawPolygon ptNewSecond()
  
  'Refresh the picture
   picBack.Refresh
   
  'generates error on load and exit so resume next
   On Error Resume Next
   Percentage = Peak / dPreGaugeArc * 100
   
   '### CALL THE PERCENT BAR (If Visible)(PictureBox Replica)
   If bPicVis = True Then PercentBarVal Picture1, Percentage, 100
End Sub

Private Sub DrawPolygon(ptNew() As POINTAPI) ' OutlineColor As Long, FillColor As Long)
   Dim L4      As Integer
   Dim P       As Integer
   Dim hdc     As Long
    
   With picBack
         '##### Draw API polygon
         .ForeColor = pntOLC
         
         'Fill polygon
         .FillColor = pntFILLC
         .FillStyle = 0
         Polygon .hdc, ptNew(0), UBound(ptNew)
   End With
End Sub

Public Sub HScroll1_Change()
  'Im using this as our 'Engine' To Avoid using timers
  'If in PeakMode we skip this sub also the mousemove events below
  
  'DO NOT REMOVE THIS LINE
   If bPeakMode = True Then Exit Sub
   
  'Display the pointer angle
   DisplayPeak HScroll1.Value
   
   RaiseEvent change        '<--- Notify the parent container of a change
                                 'Only if NOT in peak mode as we exit above
                                 'with bPeakMode being True
End Sub

Private Sub picBack_MouseDown( _
            Button As Integer, _
            Shift As Integer, _
            X As Single, _
            Y As Single)
 
 Jump = 6                'amount of change each movement
 Moover = True           'only activate MouseMove Event if mousedown

End Sub

Private Sub picBack_MouseUp( _
            Button As Integer, _
            Shift As Integer, _
            X As Single, _
            Y As Single)
  
  Moover = False         ' MouseDown no longer occurs

End Sub

Private Sub Picture1_MouseDown( _
            Button As Integer, _
            Shift As Integer, _
            X As Single, _
            Y As Single)
  
  Jump = 12              'amount of change each movement
  Moover = True          'only activate MouseMove Event if mousedown

End Sub

Private Sub Picture1_MouseUp( _
            Button As Integer, _
            Shift As Integer, _
            X As Single, _
            Y As Single)
  
  Moover = False         ' MouseDown no longer occurs

End Sub

Private Sub Picture1_MouseMove( _
            Button As Integer, _
            Shift As Integer, _
            X As Single, _
            Y As Single)
  
  Call picBack_MouseMove(Button, Shift, X, Y)   'Call the picBack mousemove and save
                                                'a few bytes in filesize
End Sub

Private Sub picBack_MouseMove( _
            Button As Integer, _
            Shift As Integer, _
            X As Single, _
            Y As Single)
  
  'This little 'Engine' controls the sound volume through the MouseDown event
  'What we want is for the sound to go down with the mouse moving down on the lefthand
  'side of the center of the gauge and the sound to go up with the mouse moving down
  'on the righthand side of the center of the gauge so here we go...
  '(mouse move left = sound down/mouse move right = sound up)
  'PercentBarMouseDown calls this same Sub so we dont want the y co ordinates
  'so we check the size of the variable 'change'
   
   If bPeakMode = True Then Exit Sub   'If in PeakMode And user mousedown and move
                                       'over picBack then exit this sub!!!!
   If Moover = False Then Exit Sub '  <----  If No MouseDown with Mousemove then
 'XXXXXXXXXXXXX                                            'exit this sub
   Select Case X
         Case Is < OldX         'sound goes down as the mouse is moving left
              If gValue - Jump >= 0 Then
                 gValue = gValue - Jump
               Else
                 gValue = 0
              End If
         Case Is > OldX         'sound goes up as the mouse is moving right
              If gValue + Jump > dPreGaugeArc Then
                 gValue = dPreGaugeArc
               Else
                 gValue = gValue + Jump
              End If
         Case Else: 'blah
   End Select
    
   If Jump = 12 Then GoTo NoY   'mousedown is on percentagebar so skip y co-ordinates

'YYYYYYYYYYYYYYY                                  'This is a little trickier
   If X < picBack.Left + (picBack.Width / 2) Then    'if mouse on left of center then
      Select Case Y
            Case Is > OldY                      'if mouse moving down then sound
                 If gValue - Jump >= 0 Then   'goes down
                    gValue = gValue - Jump
                  Else
                    gValue = 0
                 End If
            Case Is < OldY                      'else sound goes up
                 If gValue + Jump > dPreGaugeArc Then
                    gValue = dPreGaugeArc
                  Else
                    gValue = gValue + Jump
                 End If
            Case Else: 'blah
      End Select
   Else ''''''''' mouse is on right hand side of center so .......
      Select Case Y
            Case Is < OldY                      'if mouse moving up sound goes down
                 If gValue - Jump >= 0 Then
                    gValue = gValue - Jump
                  Else
                    gValue = 0
                 End If
            Case Is > OldY                      'elseif mouse moving down then sound
                 If gValue + Jump > dPreGaugeArc Then                 'goes up
                    gValue = dPreGaugeArc
                  Else
                    gValue = gValue + Jump
                 End If
            Case Else: 'blah
       End Select
   End If
     OldY = Y
NoY:
     OldX = X
     DisplayPeak gValue   'Call our sub to display the pointers new direction
End Sub

Private Sub LoadGauge()

   GaugeRadius = picBack.ScaleHeight \ 2    'picBack AutoResize = True(LEAVE IT)
   
   BuildPolygon
   
   'Position center image(anti-aliased in Corel)
   Shape1.Move GaugeRadius - Shape1.Width \ 2, _
                  GaugeRadius - Shape1.Height \ 2
                  
End Sub

Private Sub MirrorVerticals( _
            ByVal Start As Integer, _
            ByVal Finish As Integer, _
            ByVal Idx As Integer)
            
   Dim n As Integer
  
  'Makes a mirror image of the first half of the pointer
   For n = Start To Finish
      ptSecond(n).X = ptSecond(Idx - n).X
      ptSecond(n).Y = -ptSecond(Idx - n).Y
   Next

End Sub
 
Private Sub PercentBarVal( _
            Pic As PictureBox, _
            Done As Variant, _
            Total As Variant)
            
   Dim X
   On Error Resume Next
   
   With Pic
       .Cls    'clear the picture box
       
        Select Case Done
          Case Is > Total / 100 * 90
              .ForeColor = picMaxColor      'we make the forecolor red
          Case Else
              .ForeColor = picColor   'else its blue
        End Select
        
        X = Done / Total * (.Width - 5) 'X is similar to a percentage of total bar width
       'Color in the picturebox to the value of x
        Pic.Line (0, 0)-(.Width, .Height), .BackColor, BF
        Pic.Line (0, 0)-(X, .Height), .ForeColor, BF
       .Refresh    'refresh the picture
   End With

End Sub

Private Sub RotatePoints( _
            Points() As POINTAPI, _
            NewPoints() As POINTAPI, _
            ByVal angle As Single)
            
   Dim i       As Integer
   Dim P       As Integer
   
   P = UBound(Points)
   
   'Use Sin/Cos lookup table Rotate() for speed
   For i = 0 To P
      NewPoints(i).X = Points(i).X * Rotate(angle, 0) + _
                       Points(i).Y * Rotate(angle, 1) + GaugeRadius
      NewPoints(i).Y = -Points(i).X * Rotate(angle, 1) + _
                       Points(i).Y * Rotate(angle, 0) + GaugeRadius
   Next

End Sub

Private Sub UserControl_Initialize()
   Dim L4 As Long
   Dim SinRads As Single
   Dim CosRads As Single
   
   LoadGauge
   
   For L4 = 0 To 359
     '##### Sin/Cos look-up array(singles) for polygon
     '      rotation and Numeral positions
      SinRads = Sin(L4 * RADS)
      CosRads = Cos(L4 * RADS)
      Rotate(L4, 0) = SinRads
      Rotate(L4, 1) = CosRads
   Next

End Sub

