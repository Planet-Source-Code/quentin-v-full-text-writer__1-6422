VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Frame framAS 
      Caption         =   "Auto-save"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CheckBox chkEnMSG 
         Caption         =   "Enable message on startup when document has been saved by Auto-save"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.CheckBox chkEnAS 
         Caption         =   "Enable the auto-save each 60sec."
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EnAS As Boolean
Dim EnMSG As Boolean
Dim StateAS As Boolean
Dim StateMSG As Boolean

Private Sub cmdApply_Click()
    If chkEnAS.Value = 1 Then EnAS = True
    If chkEnAS.Value = 0 Then EnAS = False
    'If chkEnAS.Value = 0 Then chkEnMSG.Value = 3
    'If chkEnAS.Value = 1 Then chkEnMSG.Value = 0
    If chkEnMSG.Value = 1 Then EnMSG = True
    If chkEnMSG.Value = 0 Then EnMSG = False
    
    Open "c:\windows\TWconfig.dll" For Output As #4
        Print #4, EnAS
        Print #4, EnMSG
    Close #4
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub Form_Load()
    Open "c:\windows\TWconfig.dll" For Input As #6
        Input #6, StateAS
        Input #6, StateMSG
    Close #6
    If StateAS = True Then chkEnAS.Value = 1
    If StateAS = False Then chkEnAS.Value = 0
    If StateMSG = True Then chkEnMSG.Value = 1
    If StateMSG = False Then chkEnMSG.Value = 0
End Sub

