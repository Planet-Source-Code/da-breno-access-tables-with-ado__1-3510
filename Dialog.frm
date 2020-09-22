VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select a Database"
   ClientHeight    =   5130
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   5295
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   4815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   4440
      Width           =   4815
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
