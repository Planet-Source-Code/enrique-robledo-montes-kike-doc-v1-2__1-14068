VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "DOC HUNTER"
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   1560
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PULSA UNA TECLA PARA COMENZAR"
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   8175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DOC HUNTER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public i As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii > 0 Then
    Form1.Show
    Unload Me
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
End Sub

Private Sub Timer1_Timer()
Dim cadena As String
i = i + 1

cadena = "PRESS A KEY TO START                                                                                                                                       "

Label2.Caption = Mid(cadena, i, Len(cadena)) & Mid(cadena, 1, i)
If i = Len(cadena) Then i = 0

End Sub
