VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_FormMarqee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private bInitialized           As Boolean ' stop class properties firing prematurely
Private CFm_UnitWidth          As Long    'size of the image control that makes up the marquee
Private CFm_UnitHeight         As Long
Private CFm_Speed              As Long
Private CFm_Running            As Boolean 'stop start the marquee (stop hides marquee
Private CFm_Freeze             As Boolean 'freeze marquee (stays visible
Private m_Led                  As Variant 'holds the array of copies of the basic Image control
Private CFm_Square             As Boolean ' force the images to stay square
Private CFm_Top                As Long    'marquee dimensions
Private CFm_Left               As Long
Private CFm_Width              As Long
Private CFm_Height             As Long
Private WithEvents m_Frm       As Form    ' outside the demo find and uncomment m_Frm_Resize
Attribute m_Frm.VB_VarHelpID = -1
Private CFm_MaxWidth           As Single
Private CFm_MaxHeight          As Single
Private CFm_ControlTarget      As Control
Private m_LastSetMode          As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Class_Initialize()

'set default values

  Speed = 1
  CFm_Square = True
  CFm_Left = 0
  CFm_Top = 0
'Width and Height have to wait until a form is assigned to class
  Halt

End Sub

Private Sub Class_Terminate()

  Dim I As Long

  On Error Resume Next
'clean up
  Halt ' in case still running
  For I = 1 To m_Led.Count ' remove control array (if any)
    Unload m_Led(I)
  Next I
  On Error GoTo 0

End Sub

Public Sub ControlWrap(ctl As Control)

  Set CFm_ControlTarget = ctl
  DoControlWrap

End Sub

Private Sub DoControlWrap()

  With CFm_ControlTarget
    Move .Left - CFm_UnitWidth, .Top - CFm_UnitHeight, .Left + .Width + CFm_UnitHeight, .Top + .Height + CFm_UnitHeight
  End With
  Generate
  m_LastSetMode = 2

End Sub

Public Sub FormEdge()

  Move 0, 0, m_Frm.ScaleWidth, m_Frm.ScaleHeight
  m_LastSetMode = 1
  Generate

End Sub

Public Property Get Freeze() As Boolean

  Freeze = CFm_Freeze

End Property

Public Property Let Freeze(ByVal PropVal As Boolean)

  CFm_Freeze = PropVal
  CFm_Running = Not PropVal
  If CFm_Running Then
    Run
  End If

End Property

Public Sub Generate()

'This routine creates and destroys the array of controls
'should be called when ever you change any setting that affects size

  Dim bTempStop  As Boolean
  Dim bExitFor   As Boolean
  Dim I          As Long
  Dim MWide      As Long
  Dim MHigh      As Long
  If bInitialized Then
    bTempStop = CFm_Running
    Halt
    For I = 1 To m_Led.Count ' remove old set (if any)
      On Error Resume Next
      Unload m_Led(I)
      On Error GoTo 0
      DoEvents
    Next I
    CFm_MaxWidth = m_Frm.ScaleWidth
    CFm_MaxHeight = m_Frm.ScaleHeight
    With m_Led(0)
      .Width = CFm_UnitWidth
      .Height = CFm_UnitHeight
      .Left = CFm_Left
      .Top = CFm_Top
    End With 'm_Led(0)
    MWide = (CFm_Width - CFm_Left) \ CFm_UnitWidth
    MHigh = (CFm_Height - CFm_Top) \ CFm_UnitHeight
    For I = 1 To (MWide * 2 + MHigh * 2)
      DoEvents
      On Error Resume Next
      Load m_Led(I)
      On Error GoTo 0
      With m_Led(I)
        .Width = CFm_UnitWidth
        .Height = CFm_UnitHeight
        If I < MWide Then
'Top
          .Top = CFm_Top
          .Left = m_Led(I - 1).Left + CFm_UnitWidth
         ElseIf I < MWide + MHigh - 1 Then
'Right
          .Top = m_Led(I - 1).Top + CFm_UnitHeight
          .Left = CFm_Width - CFm_UnitWidth
         ElseIf I >= MWide + MHigh - 1 And I < MWide * 2 + MHigh - 2 Then
'Bottom
          .Top = CFm_Height - CFm_UnitHeight
          .Left = m_Led(I - 1).Left - CFm_UnitWidth
         Else
'Left
          .Top = m_Led(I - 1).Top - CFm_UnitHeight
          .Left = CFm_Left
          If .Top <= CFm_Top Then
' allows for slight errors in calculating number of elements
            If I > 0 Then
              On Error Resume Next
              Unload m_Led(I)
              On Error GoTo 0
'This mess avoids the risk of a memory leak if you just used Exit For
'and failed to trigger the End With code
              bExitFor = True
              GoTo Done
            End If
          End If
        End If
Done:
      End With
      If bExitFor Then
        Exit For
      End If
    Next I
    CFm_Running = bTempStop
    If CFm_Freeze Then
      Run
    End If
  End If

End Sub

Public Sub Halt()

  CFm_Freeze = False
  CFm_Running = False

End Sub

Public Property Get Height() As Long

  Height = CFm_Height

End Property

Public Property Let Height(ByVal PropVal As Long)

  CFm_Height = PropVal
  Generate

End Property

Public Property Let ImageFromControl(ByVal PropVal As Control)

'Not used in Demo this allows you to assign an image from a control in your program
'must be called after Init

  m_Led(0).Picture = PropVal.Picture
  LoadPictureToAll

End Property

Public Property Let ImageFromFile(ByVal PropVal As String)

'Allows you to load an image from a file
'must be called after Init

  m_Led(0).Picture = LoadPicture(PropVal)
  LoadPictureToAll

End Property

Public Sub Init(frm As Form, _
                varLed As Variant, _
                Optional Pic As Image, _
                Optional ByVal strFileName As String, _
                Optional ElementW As Long = -1, _
                Optional ElementH As Long = -1)

'NOTE varLed must be an Indexed Image control

  bInitialized = True
  Set m_Frm = frm
  m_Frm.ScaleMode = vbPixels
  CFm_Width = m_Frm.ScaleWidth
  CFm_Height = m_Frm.ScaleHeight
  CFm_MaxWidth = m_Frm.ScaleWidth
  CFm_MaxHeight = m_Frm.ScaleHeight
  Set m_Led = varLed
'set base image size
  If ElementH > 0 Then
    CFm_UnitHeight = ElementH
   Else
    CFm_UnitHeight = m_Led(0).Height
  End If
  If ElementW > 0 Then
    CFm_UnitWidth = ElementW
   Else
    CFm_UnitWidth = m_Led(0).Width
  End If
'There are 3 ways of getting an image into the marquee
  If Not Pic Is Nothing Then
'1 load it from another control
    m_Led(0).Picture = Pic.Picture 'CFm_Image
   ElseIf strFileName <> "" Then
'2 Load it from a file
    m_Led(0).Picture = LoadPicture(strFileName)
   Else
'3 Just load it into the base Image control
    If m_Led(0).Picture.Height = 0 Then
      If m_Led(0).Picture.Width = 0 Then
'and just incase you did none of these
        MsgBox " You need to set a picture cls_FormMarquee class to use"
        Exit Sub
      End If
    End If
  End If
  Generate

End Sub

Public Property Get Left() As Long

  Left = CFm_Left

End Property

Public Property Let Left(ByVal PropVal As Long)

  CFm_Left = PropVal
  Generate

End Property

Private Sub LoadPictureToAll()

  Dim I As Long

  For I = 0 To m_Led.Count - 1
    m_Led(I).Picture = m_Led(0).Picture
  Next I

End Sub

Private Sub m_Frm_Resize()

'<STUB> Reason: 'Uncomment this in your project
''commented out because the demo needs to call
'' MarqResize from the Form's 'real' ReSize event to keep the other controls in sync
'
'  MarqResize
'


End Sub

Public Sub MarqResize()

  If bInitialized Then
    Select Case m_LastSetMode
     Case 0
      Generate
     Case 1
      FormEdge
     Case 2
      DoControlWrap
    End Select
  End If

End Sub

Public Property Get MaxHeight() As Single

  MaxHeight = CFm_MaxHeight

End Property

Public Property Let MaxHeight(ByVal PropVal As Single)

  CFm_MaxHeight = PropVal

End Property

Public Property Get MaxWidth() As Single

  MaxWidth = CFm_MaxWidth

End Property

Public Property Let MaxWidth(ByVal PropVal As Single)

  CFm_MaxWidth = PropVal

End Property

Public Sub Move(Optional mL As Single = -1, _
                Optional mT As Single = -1, _
                Optional mW As Single = -1, _
                Optional mH As Single = -1)

  If mL > -1 Then
    CFm_Left = mL
  End If
  If mT > -1 Then
    CFm_Top = mT
  End If
  If mW > -1 Then
    CFm_Width = mW
  End If
  If mH > -1 Then
    CFm_Height = mH
  End If
  Generate

End Sub

Public Sub Run()

  Dim I As Long

  If CFm_Freeze = False Then
    CFm_Running = True
   Else
    CFm_Freeze = False
  End If
  On Error Resume Next
  Do
    DoEvents
    Sleep CFm_Speed
    For I = 0 To m_Led.Count - 1 Step 2
      If CFm_Running = False Then ' faster stopping
        If CFm_Freeze = False Then 'unless freeze is on as we want a single cycle of the marquee to draw
        End If
        Exit Do
      End If
      DoEvents
      m_Led(I).Visible = True
      m_Led(I).Refresh
      If I + 1 < m_Led.Count - 1 Then
        m_Led(I + 1).Visible = False
      End If
    Next I
    For I = 0 To m_Led.Count - 1 Step 2
      If CFm_Running = False Then ' faster stopping
        If CFm_Freeze = False Then
        End If
        Exit Do
      End If
      DoEvents
      m_Led(I).Visible = False
      If I + 1 < m_Led.Count - 1 Then
        m_Led(I + 1).Visible = True
        m_Led(I + 1).Refresh
      End If
    Next I
    DoEvents
  Loop Until CFm_Running = False
  If CFm_Freeze = False Then
'hide marquee unless freeze is on
    For I = 0 To m_Led.Count - 1
      m_Led(I).Visible = False
    Next I
  End If
  On Error GoTo 0

End Sub

Public Property Get Speed() As Long

  Speed = CFm_Speed

End Property

Public Property Let Speed(ByVal PropVal As Long)

'the higher the value the slower the marquee

  CFm_Speed = PropVal

End Property

Public Property Get Square() As Boolean

  Square = CFm_Square

End Property

Public Property Let Square(ByVal PropVal As Boolean)

  CFm_Square = PropVal

End Property

Public Property Get Top() As Long

  Top = CFm_Top

End Property

Public Property Let Top(ByVal PropVal As Long)

  CFm_Top = PropVal
  Generate

End Property

Public Property Get UnitHeight() As Long

  UnitHeight = CFm_UnitHeight

End Property

Public Property Let UnitHeight(ByVal PropVal As Long)

  CFm_UnitHeight = PropVal
  If CFm_Square Then
    CFm_UnitWidth = CFm_UnitHeight
  End If
  Generate
  If bInitialized Then
    m_Led(0).Visible = True
  End If

End Property

Public Property Get UnitWidth() As Long

  UnitWidth = CFm_UnitWidth

End Property

Public Property Let UnitWidth(ByVal PropVal As Long)

  CFm_UnitWidth = PropVal
  If CFm_Square Then
    CFm_UnitHeight = CFm_UnitWidth
  End If
  Generate
  If bInitialized Then
    m_Led(0).Visible = True
  End If

End Property

Public Property Get Width() As Long

  Width = CFm_Width

End Property

Public Property Let Width(ByVal PropVal As Long)

  CFm_Width = PropVal
  Generate

End Property

':)Roja's VB Code Fixer V1.1.92 (11/02/2004 1:04:27 PM) 19 + 476 = 495 Lines Thanks Ulli for inspiration and lots of code.

