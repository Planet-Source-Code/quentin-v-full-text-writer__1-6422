VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Text-Writer"
   ClientHeight    =   2085
   ClientLeft      =   3585
   ClientTop       =   5280
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fChrSpace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   1620
      Width           =   1275
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bDoEffect As Boolean

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long
Private Type RECT
    left As Long
    tOp As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_BTNFACE = 15
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Private Const DT_DISPFILE = 6            '  Display-file
Private Const DT_EXPANDTABS = &H40
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_INTERNAL = &H1000
Private Const DT_LEFT = &H0
Private Const DT_METAFILE = 5            '  Metafile, VDM
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_PLOTTER = 0             '  Vector plotter
Private Const DT_RASCAMERA = 3           '  Raster camera
Private Const DT_RASDISPLAY = 1          '  Raster display
Private Const DT_RASPRINTER = 2          '  Raster printer
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Private Sub TextEffect( _
        ByVal sText As String, _
        ByVal lX As Long, ByVal lY As Long, _
        Optional ByVal bLoop As Boolean = False, _
        Optional ByVal lStartSpacing As Long = 128, _
        Optional ByVal lEndSpacing As Long = -1, _
        Optional ByVal oColor As OLE_COLOR = vbWindowText _
        )
Dim i As Long
Dim x As Long
Dim lLen As Long
Dim lHDC As Long
Dim hBrush As Long
Static tR As RECT
Dim iDir As Long
Dim bNotFirstTime As Boolean
Dim lTime As Long
Dim lIter As Long
Dim bSlowDown As Boolean
Dim lCOlor As Long
Dim bDoIt As Boolean


    iDir = -1
    i = lStartSpacing
    tR.left = lX: tR.tOp = lY: tR.Right = lX: tR.Bottom = lY
    OleTranslateColor oColor, 0, lCOlor

    hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))
    lLen = Len(sText)
    lHDC = Me.hdc
    SetTextColor lHDC, lCOlor
    bDoIt = True
    
    Do While m_bDoEffect And bDoIt
        lTime = timeGetTime
        If (i < -3) And Not (bLoop) And Not (bSlowDown) Then
            bSlowDown = True
            iDir = 1
            lIter = (i + 4)
        End If
        If (i > 128) Then iDir = -1
        If Not (bLoop) And iDir = 1 Then
            If (i = lEndSpacing) Then
                ' Stop
                bDoIt = False
            Else
                lIter = lIter - 1
                If (lIter <= 0) Then
                    i = i + iDir
                    lIter = (i + 4)
                End If
            End If
        Else
            i = i + iDir
        End If
        FillRect lHDC, tR, hBrush
        x = 32 - (i * lLen)
        SetTextCharacterExtra lHDC, i
        DrawText lHDC, sText, lLen, tR, DT_CALCRECT
        tR.Right = tR.Right + 4
        If (tR.Right > Me.ScaleWidth \ Screen.TwipsPerPixelX) Then tR.Right = Me.ScaleWidth \ Screen.TwipsPerPixelX
        DrawText lHDC, sText, lLen, tR, DT_LEFT
        Me.Refresh
        Do
            DoEvents
        Loop While (timeGetTime - lTime) < 20
    Loop
    DeleteObject hBrush

End Sub


Private Sub cmdOK_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Me.Show
    Me.Refresh
    If Not (m_bDoEffect) Then
        Me.Cls
        Me.Font.Size = 24
        m_bDoEffect = True
        TextEffect "Text-Writer", 12, 12, , 128, -2, RGB(&H80, 0, 0)
        If m_bDoEffect Then
            Me.Font.Size = 14
            TextEffect "The best ever", 36, 52, , 128, , vb3DShadow
        End If
        If m_bDoEffect Then
            Me.Font.Name = "Tahoma"
            Me.Font.Size = 8
            Me.Font.Bold = False
            TextEffect "Text-Writer is an 100% FREEWARE and may not be saled, made by Quentin Vaquette", 49, 86, , 128, 0
        End If
        If m_bDoEffect Then
            TextEffect "Visit my site and learn how to make your own program in VB", 49, 100, , 128, 0
        End If
        m_bDoEffect = False
    Else
        m_bDoEffect = False
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_bDoEffect = False
End Sub
