VERSION 5.00
Begin VB.Form frmAutoUpdate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Update Reminder"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Should I check for new updates?"
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "days."
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Your signature file is quite old:"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmAutoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub cmdNo_Click()

    Unload Me

End Sub

Private Sub cmdYes_Click()

    frmUpdate.Show , Me
    Unload Me

End Sub

Private Sub Form_Load()

    lblText(1).Caption = DateDiff("d", frmMain.lblText(3).Caption, date)

End Sub

