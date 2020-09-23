VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frm_MarqueeClassDemo 
   Caption         =   "Marquee Class Demo"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   7875
   Icon            =   "frm_MarqueeClassDemo.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   365
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDemo 
      Caption         =   "Button Wrap"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   19
      ToolTipText     =   "Note Cls_frmMarquee can only work on controls that are directly on the form."
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame fraMarqueeSettings 
      Caption         =   "Marquee Settings"
      Height          =   3135
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   5175
      Begin VB.PictureBox picCFXPBugFixForm1 
         BorderStyle     =   0  'None
         Height          =   2760
         Left            =   120
         ScaleHeight     =   2760
         ScaleWidth      =   4980
         TabIndex        =   1
         Top             =   240
         Width           =   4980
         Begin VB.CommandButton cmdDemo 
            Caption         =   "Freeze"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   18
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdDemo 
            Caption         =   "Frame Wrap"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   17
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CommandButton cmdDemo 
            Caption         =   "Form Edge"
            Height          =   255
            Index           =   0
            Left            =   3120
            TabIndex        =   16
            Top             =   960
            Width           =   1095
         End
         Begin VB.HScrollBar hscWidth 
            Height          =   255
            Left            =   2640
            Max             =   1000
            TabIndex        =   13
            Top             =   2520
            Value           =   1000
            Width           =   2000
         End
         Begin VB.HScrollBar hscLeft 
            Height          =   255
            Left            =   2640
            Max             =   1000
            TabIndex        =   12
            Top             =   240
            Width           =   2000
         End
         Begin VB.VScrollBar vscHeight 
            Height          =   2000
            Left            =   4680
            Max             =   1000
            TabIndex        =   11
            Top             =   480
            Value           =   1000
            Width           =   255
         End
         Begin VB.VScrollBar vscTop 
            Height          =   2000
            Left            =   2400
            Max             =   1000
            TabIndex        =   10
            Top             =   480
            Width           =   255
         End
         Begin VB.CommandButton cmdDemo 
            Caption         =   "Off"
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   6
            Top             =   2520
            Width           =   975
         End
         Begin VB.CheckBox chkSquare 
            Caption         =   "Square"
            Height          =   195
            Left            =   600
            TabIndex        =   5
            Top             =   720
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin MSComctlLib.Slider sldSpeed 
            Height          =   255
            Left            =   360
            TabIndex        =   2
            Top             =   2160
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            SmallChange     =   10
            Max             =   100
            SelStart        =   50
            TickFrequency   =   10
            Value           =   50
         End
         Begin MSComctlLib.Slider sldWidth 
            Height          =   255
            Left            =   255
            TabIndex        =   4
            Top             =   285
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   450
            _Version        =   393216
            Min             =   5
            Max             =   200
            SelStart        =   50
            TickFrequency   =   20
            Value           =   50
         End
         Begin MSComctlLib.Slider sldHeight 
            Height          =   1575
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   2778
            _Version        =   393216
            Orientation     =   1
            Min             =   5
            Max             =   200
            SelStart        =   50
            TickFrequency   =   20
            Value           =   50
         End
         Begin VB.Label lblWidthHeight 
            Alignment       =   1  'Right Justify
            Caption         =   "Label5"
            Height          =   255
            Left            =   2760
            TabIndex        =   15
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label lblLeftTop 
            Caption         =   "(0, 0)"
            Height          =   255
            Left            =   2760
            TabIndex        =   14
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lblDemo 
            Caption         =   "h e i g h t"
            Height          =   1215
            Index           =   2
            Left            =   0
            TabIndex        =   9
            Top             =   840
            Width           =   135
         End
         Begin VB.Label lblDemo 
            Caption         =   "Width"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label lblDemo 
            Caption         =   "Speed"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   7
            Top             =   1965
            Width           =   975
         End
         Begin VB.Image imgMarq 
            BorderStyle     =   1  'Fixed Single
            Height          =   765
            Left            =   600
            Picture         =   "frm_MarqueeClassDemo.frx":000C
            Stretch         =   -1  'True
            ToolTipText     =   "Clcik to load other images"
            Top             =   1080
            Width           =   765
         End
      End
   End
   Begin MSComDlg.CommonDialog cdl_Demo 
      Left            =   1560
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   6000
      Picture         =   "frm_MarqueeClassDemo.frx":03BF
      Top             =   240
      Width           =   1320
   End
   Begin VB.Image leds 
      Height          =   555
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   555
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufileopt 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin VB.Menu mnufileopt 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnufileopt 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frm_MarqueeClassDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bResizing     As Boolean
Private marquee       As New cls_FormMarqee

Private Sub chkSquare_Click()

  marquee.Square = chkSquare.Value = vbChecked

End Sub

Private Sub cmdDemo_Click(Index As Integer)

  With marquee
    Select Case Index
     Case 0
      .FormEdge
      'Me.Refresh
      'Form_Resize
     Case 1
      .ControlWrap fraMarqueeSettings
      ResetPositions
     Case 2
      Select Case cmdDemo(2).Caption
       Case "Run"
        cmdDemo(2).Caption = "Off"
        .Speed = sldSpeed.Value
        cmdDemo(3).Caption = IIf(Not .Freeze, "Freeze", "Un-Freeze")
        .Run
        'marquee.Run leds, Form1.Slider1.Value
       Case "Off"
        cmdDemo(2).Caption = "Run"
        .Halt
      End Select
     Case 3
      cmdDemo(3).Caption = IIf(.Freeze, "Freeze", "Un-Freeze")
      .Freeze = Not .Freeze
     Case 4
      .ControlWrap cmdDemo(4)
      ResetPositions
    End Select
  End With

End Sub

Private Sub Form_Load()

  Show
  Set leds(0).Picture = imgMarq.Picture
  marquee.Init Me, leds, , , 15, 15
  marquee.ImageFromControl = Image1
  chkSquare.Value = vbChecked
  With marquee
    sldHeight.Value = .UnitHeight
    sldWidth.Value = .UnitHeight
    ResetPositions
    .Run
  End With

End Sub

Private Sub Form_Resize()

  bResizing = True
  fraMarqueeSettings.Left = (Me.ScaleWidth - fraMarqueeSettings.Width) / 2
  fraMarqueeSettings.Top = (Me.ScaleHeight - fraMarqueeSettings.Height) / 2
  cmdDemo(4).Top = fraMarqueeSettings.Top - cmdDemo(4).Height * 2
  cmdDemo(4).Left = (Me.ScaleWidth - cmdDemo(4).Width) / 2
  Refresh
  DoEvents
  marquee.MarqResize
  DoEvents
  ResetPositions
  bResizing = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

  End

End Sub

Private Sub hscLeft_Change()

  If bResizing = False Then
    If hscLeft.Value > hscWidth.Value - marquee.UnitWidth Then
      hscLeft.Value = hscWidth.Value - marquee.UnitWidth
    End If
    lblLeftTop.Caption = " ( " & vscTop.Value & " , " & hscLeft.Value & ")"
    marquee.Move hscLeft.Value
  End If

End Sub

Private Sub hscWidth_Change()

  If bResizing = False Then
    If hscWidth.Value < hscLeft.Value + marquee.UnitWidth Then
      hscWidth.Value = hscLeft.Value ' + marquee.UnitWidth
    End If
    lblWidthHeight.Caption = " ( " & hscWidth.Value & " , " & vscHeight.Value & ")"
    marquee.Move , , hscWidth.Value
  End If

End Sub

Private Sub imgMarq_Click()

  openFile

End Sub

Private Sub mnufileopt_Click(Index As Integer)

  Select Case Index
   Case 0
    openFile
   Case 2
    Unload Me
  End Select

End Sub

Private Sub openFile()

  With cdl_Demo
    .Filter = "Images |*.bmp;*.ico;*.gif;*.jpeg;*.jpg|Icons|*.ico|Gif|*.gif|BMP|*.bmp|JPeg|*.jpg;*.jpeg"
    .ShowOpen
    If Len(.FileName) Then
      imgMarq.Picture = LoadPicture(.FileName)
      marquee.ImageFromFile = .FileName
    End If
  End With

End Sub

Private Sub ResetPositions()

  With marquee
    vscTop.Max = .MaxHeight
    hscLeft.Max = .MaxWidth
    vscHeight.Max = .MaxHeight
    hscWidth.Max = .MaxWidth
    vscTop.Value = .Top
    hscLeft.Value = .Left
    vscHeight.Value = .Height
    hscWidth.Value = .Width
  End With

End Sub

Private Sub sldHeight_Change()

  marquee.UnitHeight = sldHeight.Value
  If chkSquare.Value = vbChecked Then
    sldWidth.Value = sldHeight.Value
  End If

End Sub

Private Sub sldSpeed_Change()

  marquee.Speed = sldSpeed.Value

End Sub

Private Sub sldWidth_Change()

  marquee.UnitWidth = sldWidth.Value
  If chkSquare.Value = vbChecked Then
    sldHeight.Value = sldWidth.Value
  End If

End Sub

Private Sub vscHeight_Change()

  If bResizing = False Then
    If vscHeight.Value < vscTop.Value + marquee.UnitHeight Then
      vscHeight.Value = vscTop.Value + marquee.UnitHeight
    End If
    lblWidthHeight.Caption = " ( " & hscWidth.Value & " , " & vscHeight.Value & ")"
    marquee.Move , , , vscHeight.Value
  End If

End Sub

Private Sub vscTop_Change()

  If bResizing = False Then
    If vscTop.Value > vscHeight.Value - marquee.UnitHeight Then
      vscTop.Value = vscHeight.Value - marquee.UnitHeight
    End If
    lblLeftTop.Caption = " ( " & vscTop.Value & " , " & hscLeft.Value & ")"
    marquee.Move , vscTop.Value
  End If

End Sub

':)Roja's VB Code Fixer V1.1.92 (11/02/2004 1:04:30 PM) 3 + 198 = 201 Lines Thanks Ulli for inspiration and lots of code.
