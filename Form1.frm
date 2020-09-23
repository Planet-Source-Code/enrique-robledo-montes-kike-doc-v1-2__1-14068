VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doc Hunter - - Double click over a file name and you open it!!"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Options"
      Height          =   255
      Left            =   8880
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2955
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Label2"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   10080
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Search String"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cadena As String
Public Contador100 As Integer
Private Sub Command1_Click()
Dim Extensiones(9) As String
Dim ContaExtensiones As Integer
Dim i As Integer
List1.Clear
List1.AddItem "Preparing..."
List1.Refresh
'**Fixing the String
ExtFile = Trim(ExtFile)
If Right(ExtFile, 1) = ";" Then
    ExtFile = Left(ExtFile, Len(ExtFile) - 1)
End If
'**Find the File extension
ContaExtensiones = 0
For i = 1 To Len(ExtFile)
    If Mid(ExtFile, i, 1) <> ";" Then
        Extensiones(ContaExtensiones) = Extensiones(ContaExtensiones) & Trim(Mid(ExtFile, i, 1))
    Else
        ContaExtensiones = ContaExtensiones + 1
    End If
Next


Ruta = Dir(RutaDir, vbDirectory)
cadena = Text1
'**Process
Contador100 = 0
Do While Ruta <> "" '***Busca la ultima imagen
    For i = 0 To 9
        If Right(Ruta, 4) = "." & Extensiones(i) Then
            Call BuscaEnArchivo(RutaDir & Ruta)
        End If
        If Extensiones(i) = "" Then Exit For
    Next
    If Contador100 >= 10 Then Exit Do
    Ruta = Dir()
Loop
List1.AddItem "END"
End Sub
Private Sub Command2_Click()
Form2.Show
End Sub
Public Function BuscaEnArchivo(Archivo As String)
Dim Linea As String
Dim Contador As Long
Open Archivo For Input As #1
'**Open the file and search
Do While Not EOF(1)
    Line Input #1, Linea
    Contador = Contador + 1
    If Sensitive = True Then
        Linea = UCase(Linea)
        cadena = UCase(cadena)
    End If
    
    If InStr(1, Linea, cadena) > 0 Then
        If First100 = True Then
            Contador100 = Contador100 + 1
        End If
        List1.AddItem "String " & cadena & " found in"
        If nLinea = False Then
            List1.AddItem Archivo
            List1.Refresh
        Else
            List1.AddItem Archivo & " at line " & Contador
            List1.Refresh
        End If
    End If
Loop
Close (1)
End Function
Private Sub Form_Load()
Text1 = ""
Label2 = ""
End Sub
Private Sub List1_DblClick()
Dim a As Long
a = Shell("C:\WINDOWS\WRITE.EXE " & List1.List(List1.ListIndex), vbNormalFocus)
End Sub
