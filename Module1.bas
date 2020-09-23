Attribute VB_Name = "Module1"
Option Explicit

Public docOnOff As Boolean
Public sYN As Boolean
Public sFile As String
Public doc As frmDoc
Public Number As Long

Public Function errmsg(msg As String, alert As Integer) As Integer
    errmsg = MsgBox(msg, alert, "Tewt-Writer")
End Function

Public Sub NewDoc()
    Number = Number + 1
    Set doc = New frmDoc
    doc.Caption = "Document " & Number
    doc.Show
End Sub
Public Sub OpenDoc()
        Dim oFile As String
    Dim fileTF As Boolean
    
    fileTF = True
    
    With frmMain.comDialog
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
    
    frmMain.StatusBar1.Panels(1).Text = "Status : Opening a file"
    If fileTF = True Then
        doc.TextBox.LoadFile (oFile)
        frmMain.StatusBar1.Panels(1).Text = "Status : Opening file succesfull"
    Else: If fileTF = False Then GoTo err1
    End If
    
err1:
    frmMain.StatusBar1.Panels(1).Text = "Status"
    Exit Sub
End Sub
 Public Sub PrintDoc()
     With frmMain.comDialog
         .CancelError = True
          On Error GoTo err2
         .DialogTitle = "Print your document"
         .ShowPrinter
     End With
     Printer.Print doc.TextBox.Text
     frmMain.StatusBar1.Panels(1).Text = "Status : printing your document"
     Printer.EndDoc
     frmMain.StatusBar1.Panels(1).Text = "Status" '
err2: Exit Sub
End Sub
Public Sub SaveDoc()
       If sYN = True Then
       doc.TextBox.SaveFile (sFile)
       doc.Caption = sFile
       sYN = True
   Else: If sYN = False Then frmMain.mnuFileSaveAs_Click
   End If
End Sub


