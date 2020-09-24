VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "FragIt"
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   2  'Custom
   ScaleHeight     =   106
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   45
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   187
      TabIndex        =   6
      Top             =   1290
      Width           =   2835
   End
   Begin VB.PictureBox picCHolder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   0
      Left            =   45
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   232
      TabIndex        =   0
      Top             =   270
      Width           =   3510
      Begin VB.TextBox txtSegSize 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1065
         TabIndex        =   16
         Text            =   "1.38"
         ToolTipText     =   "Text field for segment file size"
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox chkDelete 
         Height          =   195
         Index           =   0
         Left            =   1050
         TabIndex        =   15
         Top             =   705
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.TextBox txtSFName 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1065
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         ToolTipText     =   "Text field for file name"
         Top             =   45
         Width           =   2055
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Original File"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   1305
         TabIndex        =   12
         Top             =   705
         Width           =   1320
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MB"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3195
         TabIndex        =   11
         Top             =   405
         Width           =   240
      End
      Begin VB.Label lblBrowse 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "..."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3165
         MouseIcon       =   "frmMain.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "Browse"
         Top             =   45
         Width           =   270
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fragment Size"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   3
         Top             =   405
         Width           =   1005
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   2
         Top             =   90
         Width           =   705
      End
   End
   Begin VB.PictureBox picCHolder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   1
      Left            =   45
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   232
      TabIndex        =   7
      Top             =   270
      Width           =   3510
      Begin VB.CheckBox chkDelete 
         Height          =   195
         Index           =   1
         Left            =   1050
         TabIndex        =   13
         Top             =   705
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.TextBox txtSegfname 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1065
         OLEDropMode     =   1  'Manual
         TabIndex        =   8
         ToolTipText     =   "Text field for segment file to be merged"
         Top             =   45
         Width           =   2055
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Fragment Files"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   1305
         TabIndex        =   14
         Top             =   705
         Width           =   1530
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fragment file"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   90
         Width           =   900
      End
      Begin VB.Label lblSBrowse 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "..."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3165
         MouseIcon       =   "frmMain.frx":074C
         MousePointer    =   99  'Custom
         TabIndex        =   9
         ToolTipText     =   "Browse"
         Top             =   45
         Width           =   270
      End
   End
   Begin VB.Image imgExit 
      Height          =   165
      Left            =   3375
      MouseIcon       =   "frmMain.frx":0A56
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":0D60
      ToolTipText     =   "Exit fSplit"
      Top             =   60
      Width           =   180
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   165
      Left            =   2670
      Picture         =   "frmMain.frx":0E6E
      Top             =   75
      Width           =   465
   End
   Begin VB.Image imgInfo 
      Height          =   165
      Left            =   3180
      MouseIcon       =   "frmMain.frx":0FD4
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":12DE
      ToolTipText     =   "About fSplit"
      Top             =   60
      Width           =   180
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00000000&
      Height          =   1590
      Left            =   0
      Top             =   0
      Width           =   3600
   End
   Begin VB.Label lblInitialize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Initialize"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2910
      MouseIcon       =   "frmMain.frx":13EC
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Initialize file split or merge"
      Top             =   1290
      Width           =   645
   End
   Begin VB.Image imgTab 
      Height          =   240
      Index           =   0
      Left            =   45
      MouseIcon       =   "frmMain.frx":16F6
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":1A00
      Top             =   45
      Width           =   720
   End
   Begin VB.Image imgTab 
      Height          =   240
      Index           =   1
      Left            =   675
      MouseIcon       =   "frmMain.frx":21CA
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":24D4
      Top             =   45
      Width           =   720
   End
End
Attribute VB_Name = "frmMain"
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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Dim Mode As Integer

Private Sub Form_Load()
If App.PrevInstance Then End 'don't load if loaded

  SetWindowPos Me.hWnd, -1, 0&, 0&, 0&, 0&, 1 Or 2 'top most

    Dim i, j As Integer
    
    For i = 1 To 106
        Me.Line (0, Me.ScaleHeight - i)-(Me.ScaleWidth, Me.ScaleHeight - i), _
        RGB(i + 135, i + 135, i + 135), BF
    Next i
    For i = 1 To 64
    For j = 0 To 1
        picCHolder(j).Line (0, picCHolder(j).ScaleHeight - i)-(picCHolder(j). _
        ScaleWidth, picCHolder(j).ScaleHeight - i), RGB(i + 160, i + 160, i + _
        160), BF
    Next j
    Next i
  
  Mode = 0
  spProg 0
'Get settings from Registry
chkDelete(0).Value = val(GetSetting("FragIt", "Settings", "DeleteOriginalFile", 1))
chkDelete(1).Value = val(GetSetting("FragIt", "Settings", "DeleteFragmentFiles", 1))
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hWnd, &HA1, 2, 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Save settings in the Registry
SaveSetting "FragIt", "Settings", "DeleteOriginalFile", Str$(chkDelete(0).Value)
SaveSetting "FragIt", "Settings", "DeleteFragmentFiles", Str$(chkDelete(1).Value)
End Sub

Private Sub imgInfo_Click()
frmAbout.Show 1, Me
End Sub

Private Sub imgTab_Click(Index As Integer)
Mode = Index
picCHolder(Index).ZOrder
imgTab(Index).ZOrder
End Sub

Private Sub lblBrowse_Click()
  Dim FileDialog As CFileDialog
  Set FileDialog = New CFileDialog
  With FileDialog
    .DialogTitle = "Open"
    .Filter = "All Files (*.*)|*.*"
    .FilterIndex = 0
    .Flags = FleFileMustExist + FleHideReadOnly + FleCreatePrompt
    .hWndParent = Me.hWnd
    If .Show(True) Then
    txtSFName.Text = .FileName
    Else
    Exit Sub
    End If
  End With
End Sub

Private Sub lblBrowse_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
lblBrowse.BackStyle = 1
End Sub

Private Sub lblBrowse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
lblBrowse.BackStyle = 0
End Sub

Private Sub imgExit_Click()
CancelJob = True
CancelAndExit = True
Unload Me
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
lblExit.BackColor = 12632256
End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
lblExit.BackColor = 14737632
End Sub

Private Sub lblInitialize_Click()
Dim i, j As Integer
Dim Fragments As Integer

If lblInitialize.Caption = "Cancel" Then CancelJob = True: _
lblInitialize.Caption = "Initialize": Exit Sub

lblInitialize.Caption = "Cancel"


Select Case Mode
Case 0
    
    'Call the function
    i = SplitFile(txtSFName.Text, val(txtSegSize.Text) * 1048576, chkDelete(0).Value, Fragments)
    'Inform the user about the call success or failure
    Select Case i
    Case 0
        MsgBox "The process completed successfully." & Chr(10) & "The file was split to " & Fragments & " Fragments.", vbExclamation, "Successfully Done"
    Case 1
        MsgBox "An error occured!" & vbCr & "Try entering different Fragment value or check if the file name is correct.", vbExclamation, "Error"
    End Select
Case 1
    
    'Call the function
    j = MergeFiles(txtSegfname.Text, chkDelete(1).Value, Fragments)
 
    'Inform the user about the call success or failure
    Select Case j
    Case 0
        MsgBox "The process completed successfully." & Chr(10) & "The file was merged from " & Fragments & " Fragments.", vbExclamation, "Successfully Done"
    Case 1
        MsgBox "An error occured!" & Chr(10) & "Check if all the fraggment files are in the same directory" & Chr(10) & "or make sure the program is not overwriting any file.", vbExclamation, "Error"
    End Select

End Select

lblInitialize.Caption = "Initialize"
CancelJob = False
End Sub

Private Sub lblInitialize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
lblInitialize.BackStyle = 1
End Sub

Private Sub lblInitialize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
lblInitialize.BackStyle = 0
End Sub

Private Sub lblSBrowse_Click()
  Dim FileDialog As CFileDialog
  Set FileDialog = New CFileDialog
  With FileDialog
    .DialogTitle = "Open"
    .Filter = "All Files (*.*)|*.*"
    .FilterIndex = 0
    .Flags = FleFileMustExist + FleHideReadOnly + FleCreatePrompt
    .hWndParent = Me.hWnd
    If .Show(True) Then
    txtSegfname.Text = .FileName
    Else
    Exit Sub
    End If
  End With
End Sub

Public Sub spProg(val As Integer)
Dim i
picProgress.Cls
For i = 1 To 15
    picProgress.Line (0, picProgress.ScaleHeight - i)-(picProgress.ScaleWidth, picProgress.ScaleHeight - i), RGB((i * 2) + 140, (i * 2) + 140, (i * 2) + 140), BF
Next i

picProgress.Line (30, 2)-((val * 1.55) + 30, 12), &H808080, BF
picProgress.Line (30, 2)-(184, 12), 0, B
picProgress.CurrentX = 3: picProgress.CurrentY = 1
picProgress.Print val & "%"

End Sub

Private Sub lblSBrowse_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
lblSBrowse.BackStyle = 1
End Sub

Private Sub lblSBrowse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
lblSBrowse.BackStyle = 0
End Sub

Private Sub txtSegfname_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
txtSegfname.Text = Data.Files(1)
End Sub

Private Sub txtSFName_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
txtSFName.Text = Data.Files(1)
End Sub
