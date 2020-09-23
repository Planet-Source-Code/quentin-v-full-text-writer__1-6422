VERSION 5.00
Begin VB.Form frmAutoSave 
   Caption         =   "Auto-Save"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtTime 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "save my document in :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "sec."
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label label1 
      Caption         =   "save my document each"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmAutoSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
