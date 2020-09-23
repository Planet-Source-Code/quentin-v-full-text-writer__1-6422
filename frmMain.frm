VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Text Writer"
   ClientHeight    =   7785
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10005
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "Icons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   20
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            Object.Tag             =   ""
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            Object.Tag             =   ""
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Left"
            Object.ToolTipText     =   "Align to left"
            Object.Tag             =   ""
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Center"
            Object.ToolTipText     =   "Centre"
            Object.Tag             =   ""
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Right"
            Object.ToolTipText     =   "Align to right"
            Object.Tag             =   ""
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "color"
            Object.ToolTipText     =   "color"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Font"
            Object.ToolTipText     =   "Font"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer asTimer 
      Interval        =   60000
      Left            =   2520
      Tag             =   "5"
      Top             =   3000
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7530
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   13414
            Text            =   "Status"
            TextSave        =   "Status"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   979
            MinWidth        =   970
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   926
            MinWidth        =   919
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   741
            MinWidth        =   751
            TextSave        =   "INS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   4
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   979
            MinWidth        =   970
            TextSave        =   "SCRL"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog comDialog 
      Left            =   4920
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList Icons 
      Left            =   3720
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":09F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":109A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":13EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":173E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1DE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2134
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2486
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":27D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSepar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &as"
      End
      Begin VB.Menu mnuFileSepar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFileSepar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuInsertInsert 
         Caption         =   "&Insert Image"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "P&aste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuEditSepar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAll 
         Caption         =   "&Select All"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbarOpt 
         Caption         =   "&Toolbar Options"
      End
      Begin VB.Menu mnuViewASOpt 
         Caption         =   "&Auto-Save options"
      End
      Begin VB.Menu mnuVeiwSepar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCascade 
         Caption         =   "&Cascade "
      End
      Begin VB.Menu mnuViewHorizon 
         Caption         =   "Tile &horizontal"
      End
      Begin VB.Menu mnuViewVerti 
         Caption         =   "Tile &vertical"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About Text-Writer"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim asDes As String
Dim TimerTF As Boolean
Dim SaveYN As Boolean

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302

Private Sub asTimer_Timer()
    Tag = Tag - 1
    If Tag > 0 Then Exit Sub
    asDes = "c:\windows\temp\" & "Document" & Number & ".txt"
    If SaveYN = False Then doc.TextBox.SaveFile (asDes)
    If SaveYN = True Then doc.TextBox.SaveFile (sFile)
    MsgBox "Auto-Save , ok"
    TimerTF = True
    Tag = 5
End Sub

Private Sub MDIForm_Load()
    Dim DocPath As String
    Dim DocName As String
    Dim asTF As Boolean
    Dim OpenOK As Boolean
    Dim YesNo As String
    
    frmDoc.Hide
    
    'Open the Config.TWC and change the options
    'IF you have any problem with this file then create a file
    'named "TWConfig.TWC" in the windows directory and
    'put True or False on the first and second line
    'You can chnage this in thecontrol panel for autosave
    Dim EnAS As String
    Dim EnMSG As String
    
    Open "c:\windows\TWconfig.TWC" For Input As #5
        Input #5, EnAS
        Input #5, EnMSG
    If EnAS = False Then asTimer.Enabled = False
    If EnAS = True Then asTimer.Enabled = True
    If EnMSG = False Then OpenOK = False
    Close #5
    'Close the config.TWC
    
    TimerTF = False
    asTF = True
    
    On Error GoTo err1
    
    FileCopy "c:\windows\TW_AS.ini", "c:\windows\temp\TW_AS.ini"
    Open "c:\windows\temp\TW_AS.ini" For Input As #1
        Input #1, YesNo
        Input #1, DocPath
        Input #1, DocName
    
    If YesNo = True Then OpenOK = True
    If YesNo = False Then OpenOK = False
    If OpenOK = False Then GoTo err1
    If OpenOK = True Then MsgBox DocName & " " & "at" & " (" & DocPath & ")" & " has been saved by the auto-save", , "Document Retrieving"
    Close #1
    Kill "c:\windows\temp\TW_AS.ini"
    Kill "c:\windows\TW_AS.ini"
    NewDoc
    doc.TextBox.SelFontSize = "11"
err1:
NewDoc
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If TimerTF = False Then Exit Sub
    If sYN = True Then
        Open "c:\windows\TW_AS.ini" For Output As #2
            Print #2, False
        Close #2
    End If
    If sYN = False Then
        Open "c:\windows\TW_AS.ini" For Output As #3
            Print #3, True
            Print #3, asDes
            Print #3, "Document " & Number
        Close #3
    End If
    End
End Sub

Private Sub mnuEditAll_Click()
    doc.TextBox.SelStart = 0
    doc.TextBox.SelLength = Len(doc.TextBox.Text)
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.Clear
    Clipboard.SetText doc.TextBox.SelText, 1
End Sub

Private Sub mnuEditCut_Click()
    Clipboard.Clear
    Clipboard.SetText doc.TextBox.SelText, 1
    doc.TextBox.SelText = ""
End Sub

Private Sub mnuEditDelete_Click()
    doc.TextBox.SelText = ""
End Sub

Private Sub mnuEditPaste_Click()
    doc.TextBox.SelText = ""
    doc.TextBox.SelText = Clipboard.GetText(1)
End Sub

Private Sub mnuEditUndo_Click()

End Sub

Private Sub mnuFileExit_Click()
    If errmsg("Are you shure that you want to quit ?", vbYesNo) = vbYes Then
        End
    End If
    End Sub

Private Sub mnuFileNew_Click()
    NewDoc
End Sub

Private Sub mnuFileOpen_Click()
    Dim oFile As String
    Dim fileTF As Boolean
    
    fileTF = True
    
    With comDialog
        .CancelError = True
        On Error GoTo err1
        .DialogTitle = "Open a file"
        .DefaultExt = "*.txt"
        .Filter = "Text Files (*.txt)|*.txt|Rich Text Formats (*.rtf)|*.rtf"
        .ShowOpen
        oFile = .filename
    End With
    
    If oFile = "" Then fileTF = False
    
    NewDoc
    doc.Caption = oFile
    
    StatusBar1.Panels(1).Text = "Status : Opening a file"
    If fileTF = True Then
        doc.TextBox.LoadFile (oFile)
        StatusBar1.Panels(1).Text = "Status : Opening file succesfull"
    Else: If fileTF = False Then GoTo err1
    End If
    
err1:
    StatusBar1.Panels(1).Text = "Status"
    Exit Sub
End Sub

Private Sub mnuFilePrint_Click()
    With comDialog
        .CancelError = True
        On Error GoTo err2
        .DialogTitle = "Pritn your document"
        .ShowPrinter
    End With
    Printer.Print doc.TextBox.Text
    StatusBar1.Panels(1).Text = "Status : printing your document"
    Printer.EndDoc
    StatusBar1.Panels(1).Text = "Status"
err2: Exit Sub
End Sub

Private Sub mnuFileSave_Click()
    If sYN = True Then
        doc.TextBox.SaveFile (sFile)
        doc.Caption = sFile
        sYN = True
    Else: If sYN = False Then mnuFileSaveAs_Click
    End If
End Sub

Public Sub mnuFileSaveAs_Click()
    Dim fileTF As Boolean
    
    sYN = True
    fileTF = True
    
    With comDialog
        .CancelError = True
        On Error GoTo err1
        .DialogTitle = "save a file"
        .DefaultExt = "*.txt"
        .Filter = "Text Files (*.txt)|*.txt|Rich Text Formats (*.rtf)|*.rtf|My format (*.tst)|*.tst"
        .ShowSave
        sFile = .filename
    End With
    
    If sFile = "" Then fileTF = False
        
    StatusBar1.Panels(1).Text = "Status : saving a file"
    If fileTF = True Then
        doc.TextBox.SaveFile (sFile)
        doc.Caption = sFile
        StatusBar1.Panels(1).Text = "Status : saving file succesfull"
        sYN = True
    Else: If fileTF = False Then GoTo err1
    End If
    
err1:
    StatusBar1.Panels(1).Text = "Status"
    Exit Sub
End Sub

Private Sub SaveTimer_Timer()
    doc.TextBox.SaveFile ("c:\windows\temp\" & doc.Caption & ".txt")
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuInsertInsert_Click()
    Dim iFile As String
    
    Clipboard.Clear
    
    With comDialog
       .DefaultExt = "*.jpg"
        .Filter = "JPG images (*.jpg)|*.jpg|JPEG images (*.jpeg)|*.jpeg|GIF images (*.gif)|*.gif|TIF images|*.tif|Cliparts (*.wmf)|*.wmf"
        .DialogTitle = "Choose an image"
        .ShowOpen
        iFile = .filename
    End With
    frmStuff.Image1.Picture = LoadPicture(iFile)
    Clipboard.SetData frmStuff.Image1.Picture
    SendMessage doc.TextBox.hwnd, WM_PASTE, 0, 0
End Sub

Private Sub mnuToolbarOpt_Click()
    Toolbar.Customize
End Sub

Private Sub mnuViewASOpt_Click()
    frmOptions.Show
End Sub

Private Sub mnuViewCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuViewHorizon_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuViewVerti_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
        
    With doc.TextBox
    Select Case Button.Key
        Case "New"
            NewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "cut"
            Clipboard.Clear
            Clipboard.SetText .SelText
            'put the code so that the SelText is deleted
        Case "Copy"
            Clipboard.Clear
            Clipboard.SetText .SelText
        Case "Paste"
            'Put the code to paste here
            'gebruik het GetText commando van de clipboard
            'maar mischien ook het GetData commando zodat ik plaatjes kan pasten
        Case "Bold"
            If Toolbar.Buttons(11).Value = tbrPressed Then
                .SelBold = True
                Exit Sub
            End If
            If Toolbar.Buttons(11).Value = tbrUnpressed Then
                .SelBold = False
                Exit Sub
            End If
        Case "Italic"
            If Toolbar.Buttons(12).Value = tbrPressed Then
                .SelItalic = True
                Exit Sub
            End If
            If Toolbar.Buttons(12).Value = tbrUnpressed Then
                .SelItalic = False
                Exit Sub
            End If
        Case "Underline"
            If Toolbar.Buttons(13).Value = tbrPressed Then
                .SelUnderline = True
                Exit Sub
            End If
            If Toolbar.Buttons(13).Value = tbrUnpressed Then
                .SelUnderline = False
                Exit Sub
            End If
        Case "left"
            'put the code to align to the left here
            MsgBox "insert the code for left here"
        Case "centre"
            'put the code to centre the text
            MsgBox "insert the code for centre here"
        Case "right"
            'put the code to align to the right
            MsgBox "insert the code for right here"
        Case "color"
            Dim COLOR As String
            With comDialog
                .DialogTitle = "take your color"
                .ShowColor
                COLOR = .COLOR
            End With
            doc.TextBox.SelColor = COLOR
        Case "Font"
            Dim FontN As String
            Dim FontS As String
            
            On Error GoTo err2
            
            With comDialog
                .DialogTitle = "Choose a font"
                .Flags = cdlCFBoth
                .ShowFont
                FontN = .FontName
                FontS = .FontSize
            End With
            .SelFontName = FontN
            .SelFontSize = FontS
            Exit Sub
err2:
MsgBox "an error has occured, operation aborted", , "ERROR"

    End Select
    End With
    End Sub

Private Sub selecColor()
    Dim COLOR As String
    
    With comDialog
        .DialogTitle = "take your color"
        .ShowColor
        COLOR = .COLOR
    End With
    doc.TextBox.SelColor = COLOR
End Sub

