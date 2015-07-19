VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "SQL Generator"
   ClientHeight    =   6405
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   5070
   Icon            =   "frmMain.frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   6030
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboWrapper 
      Height          =   315
      ItemData        =   "frmMain.frx":4BC64
      Left            =   120
      List            =   "frmMain.frx":4BC66
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtSearchString 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3600
      Width           =   4695
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton cmdParse 
      Caption         =   "Parse"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtValues 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4920
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblCombo 
      Caption         =   "Wrapper"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblGenerator 
      Caption         =   "Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblSearchString 
      Caption         =   "Search String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   +{F4}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuParse 
         Caption         =   "Parse"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    ClearValues
End Sub

Private Sub cmdCopy_Click()
    CopyToClip
End Sub

Private Sub cmdParse_Click()
    Parse
End Sub

Private Sub Form_Load()
    cboWrapper.AddItem ("")
    cboWrapper.AddItem ("'")
End Sub
Private Sub mnuAbout_Click()
    MsgBox "SQL Generator" & vbCrLf _
        & "Copyright " & Chr$(169) & " 2014 Alex Hedley", vbInformation, _
        "About"
End Sub

Private Sub mnuClear_Click()
    ClearValues
End Sub

'Menu Options
Private Sub mnuCopy_Click()
    CopyToClip
End Sub

Private Sub mnuExit_Click()
    'Update the Status Bar
    StatusBar1.Panels(1).Text = "Exiting..."
    'Close Form
    Unload Me
End Sub

Private Sub mnuParse_Click()
    Parse
End Sub

Sub ClearValues()
    txtValues = ""
End Sub

Sub CopyToClip()
    CopyTextToClipboard txtSearchString
    StatusBar1.Panels(1).Text = "Copied"
End Sub

Sub Parse()

    'Update the Status Bar
    StatusBar1.Panels(1).Text = "Parsing..."
    
    'Clear any previous ones
    txtSearchString = ""
    
    Dim arr() As String
    arr = Split(txtValues, vbNewLine)
    
    Dim itm As Variant
    For Each itm In arr
        'MsgBox itm
        'ParseIncidentNumber itm
        If itm <> "" Then
            'txtSearchString = txtSearchString & Chr(39) & itm & Chr(39) & ","
            txtSearchString = txtSearchString & cboWrapper & itm & cboWrapper & ","
        End If
    Next itm
    
    Dim intSSLen As Integer
    intSSLen = Len(txtSearchString)
    
    'Remove the last ","
    If (intSSLen) <> 0 Then
        txtSearchString = Left(txtSearchString, intSSLen - 1)
        'Wrap in ()
        txtSearchString = "(" & txtSearchString & ")"
        'Update the Status Bar
        StatusBar1.Panels(1).Text = "Parsed"
    Else
        'Update the Status Bar
        StatusBar1.Panels(1).Text = "No values to be parsed"
    
    End If
    
End Sub

Private Sub txtSearchString_Change()
    If Len(txtSearchString) = 0 Then
        cmdCopy.Enabled = False
    Else
        cmdCopy.Enabled = True
    End If
End Sub
