VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "About"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdExit 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   1875
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   3
      Text            =   "frmAbout.frx":0000
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label lblThanks2 
      BackStyle       =   0  'Transparent
      Caption         =   "Paul"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label lblThanks 
      BackStyle       =   0  'Transparent
      Caption         =   "Patabugen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "cyber_chris235@gmx.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "© Copyright by Cyber Chris"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Open Source Antivirus Project"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Cyber Chris"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Anti Virus 2004"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    lblText(3).Caption = "Version: " & App.major & "." & App.minor & "." & App.Revision

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim myArticleAddr As String

    If MsgBox("Would you please vote on PSC Website in case you like this program?", vbQuestion + vbYesNo, "Your vote will be very well appreciated ...") = vbYes Then
        myArticleAddr = "http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=51899&optCodeRatingValue=5"
        Call ShellExecute(Me.hWnd, "Open", myArticleAddr, vbNullString, vbNullString, 1)
        MsgBox "Thank you very much. I really appreciate that :-) ", , "Thanks a million..."
    End If

End Sub

Private Sub lblCopyright_Click(Index As Integer)

    Call ShellExecute(Me.hWnd, "Open", "mailto:cyber_chris235@gmx.net", vbNullString, "c:\", 1)

End Sub

Private Sub lblThanks2_Click()

    Call ShellExecute(Me.hWnd, "Open", "mailto:wpsjr1@succeed.net", vbNullString, "c:\", 1)

End Sub

Private Sub lblThanks_Click()

    Call ShellExecute(Me.hWnd, "Open", "mailto:dude@patabugen.co.uk", vbNullString, "c:\", 1)
    
End Sub

