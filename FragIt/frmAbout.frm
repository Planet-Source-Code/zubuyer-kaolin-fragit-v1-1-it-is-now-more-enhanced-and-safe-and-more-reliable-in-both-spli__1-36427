VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3255
   ClientLeft      =   2310
   ClientTop       =   1620
   ClientWidth     =   4410
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&OK"
      Height          =   345
      Left            =   3615
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   645
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":000C
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   765
      Index           =   5
      Left            =   165
      TabIndex        =   7
      Top             =   1605
      Width           =   4125
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "For additional information email me at the following address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Index           =   4
      Left            =   165
      TabIndex        =   6
      Top             =   2550
      Width           =   2955
   End
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   1335
      TabIndex        =   5
      Top             =   285
      Width           =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   6
      X2              =   287
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lifeforcez@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   165
      TabIndex        =   4
      Top             =   2940
      Width           =   1665
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":00C4
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   2
      Left            =   165
      TabIndex        =   3
      Top             =   825
      Width           =   4125
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "© 2002 by Muhammad Zubaer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   855
      TabIndex        =   1
      Top             =   495
      Width           =   2190
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FragIt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   585
      TabIndex        =   0
      Top             =   210
      Width           =   660
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   75
      Picture         =   "frmAbout.frx":016D
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''
'FragIt v.1.1                          '
'© Copyright 2002 by Muhammad Zubaer   '
'                                      '
'This is a FREEWARE but this code      '
'is not intend to be used commercially.'
'Although you can use it as you like   '
'in your own project but do not resale '
'it or destroy the original author's   '
'name. If you use this code in your    '
'project then it would be nice to give '
'me some cradits. I've worked hard on  '
'it.
'                                      '
'Warning: There is no warranty provided'
'so use it in your own risk. The author'
'is not responsible for any damage     '
'caused by this code.                  '
'                                      '
'Mail me at the following address if   '
'you have any questions or made any    '
'enhancement.                          '
'lifeforcez@hotmail.com                '
''''''''''''''''''''''''''''''''''''''''

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
  SetWindowPos Me.hWnd, -1, 0&, 0&, 0&, 0&, 1 Or 2 'top most
    Dim i, j As Integer
    
    For i = 1 To 217
        Me.Line (0, Me.ScaleHeight - i)-(Me.ScaleWidth, Me.ScaleHeight - i), _
        RGB((i / 3) + 170, (i / 3) + 170, (i / 3) + 170), BF
    Next i
lblVer.Caption = "v." & App.Major & "." & App.Minor
End Sub
