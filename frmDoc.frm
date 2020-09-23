VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDoc 
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   4470
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox TextBox 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5318
      _Version        =   393217
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmDoc.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Form_Resize
End Sub

Private Sub Form_Resize()
    TextBox.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
End Sub

Private Sub TextBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'this is modulo
    If Button Mod 8 = 2 Then
        frmStuff.PopupMenu frmStuff.mnu1
    End If
End Sub
