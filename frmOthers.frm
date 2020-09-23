VERSION 5.00
Begin VB.Form frmStuff 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   4095
      Left            =   360
      Top             =   360
      Width           =   5295
   End
   Begin VB.Menu mnu1 
      Caption         =   "jksdf"
      Begin VB.Menu mnuBold 
         Caption         =   "&Bold"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "&Font"
         Begin VB.Menu mnuFontSmal 
            Caption         =   "Smaller"
         End
         Begin VB.Menu mnuFontLarge 
            Caption         =   "Larger"
         End
         Begin VB.Menu mnuFontChange 
            Caption         =   "Change"
         End
      End
   End
End
Attribute VB_Name = "frmStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu2_Click()

End Sub

Private Sub mnuFontChange_Click()
        Dim FontN As String
        Dim FontS As String
            
            On Error GoTo err2
            
            With frmMain.comDialog
                .DialogTitle = "Choose a font"
                .Flags = cdlCFBoth
                .ShowFont
                FontN = .FontName
                FontS = .FontSize
            doc.TextBox.SelFontName = FontN
            doc.TextBox.SelFontSize = FontS
            Exit Sub
err2:
MsgBox "an error has occured, operation aborted", , "ERROR"
    End With
End Sub

Private Sub mnuFontSmal_Click()
    doc.TextBox.SelFontSize = doc.TextBox.SelFontSize - 3
End Sub
