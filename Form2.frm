VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Doc Hunter Options"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   LinkTopic       =   "Form2"
   ScaleHeight     =   4665
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Read the saved configuration"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Config"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00000000&
      Caption         =   "Only 10 first matches"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   3360
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Notify line at the document"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2760
      Width           =   3855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Do not make difference between Upp/Low Keys"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set file extension (Example :TXT; DOC)"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set Path (Example: C:\)"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RutaDir = Text1
If Dir(RutaDir, vbDirectory) = "" Then
    Form1.List1.AddItem "Invalid Path " & RutaDir
Else
    Form1.List1.AddItem "Path set to: " & RutaDir
End If

If Check1.Value = vbChecked Then
    Sensitive = True
Else
    Sensitive = False
End If

If Check2.Value = vbChecked Then
    nLinea = True
Else
    nLinea = False
End If
If Check3.Value = vbChecked Then
    First100 = True
Else
    First100 = False
End If
ExtFile = Text2
Form1.Label2 = "Path:" & RutaDir & " -- Extens:" & ExtFile
Unload Me
End Sub
Private Sub Command2_Click()
If MsgBox("Are you sure?", 36) = 6 Then
    If Dir("C:\DOCHUNTER\", vbDirectory) = "" Then
        MkDir ("C:\CABRON")
    End If
    Open "C:\DOCHUNTER\CONFIG.CFG" For Output As #1
    Print #1, "[Difference]"
    If Check1.Value = vbChecked Then
        Print #1, "YES"
    Else
        Print #1, "NO"
    End If
    Print #1, "[NOTIFY]"
    If Check2.Value = vbChecked Then
        Print #1, "YES"
    Else
        Print #1, "NO"
    End If
    Print #1, "[100]"
    If Check3.Value = vbChecked Then
        Print #1, "YES"
    Else
        Print #1, "NO"
    End If
    Print #1, "[PATH]"
    Print #1, Text1
    Print #1, "[EXTENSION]"
    Print #1, Text2
    Close (1)
End If
End Sub
Private Sub Command3_Click()
Dim Linea As String
If Dir("C:\DOCHUNTER\CONFIG.CFG") <> "" Then
    Open "C:\DOCHUNTER\CONFIG.CFG" For Input As #1
    Line Input #1, Linea '**Sensitive
    Line Input #1, Linea
    If Linea = "YES" Then
        Check1.Value = vbChecked
    Else
        Check1.Value = vbUnchecked
    End If
    Line Input #1, Linea '**Notify
    Line Input #1, Linea
    If Linea = "YES" Then
        Check2.Value = vbChecked
    Else
        Check2.Value = vbUnchecked
    End If
    Line Input #1, Linea '**First 100
    Line Input #1, Linea
    If Linea = "YES" Then
        Check3.Value = vbChecked
    Else
        Check3.Value = vbUnchecked
    End If
    Line Input #1, Linea '**PATH
    Line Input #1, Linea
    Text1 = Linea
    Line Input #1, Linea '**Extension
    Line Input #1, Linea
    Text2 = Linea
    Close (1)
Else
    MsgBox "Config.cfg not found", vbExclamation
End If


End Sub

Private Sub Form_Load()
If Sensitive = True Then
    Check1.Value = vbChecked
Else
    Check1.Value = vbUnchecked
End If
If nLinea = True Then
    Check2.Value = vbChecked
Else
    Check2.Value = vbUnchecked
End If
Text1 = RutaDir
Text2 = ""
End Sub

