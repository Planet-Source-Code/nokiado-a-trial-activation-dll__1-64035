VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNag 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Notificación de Versión de Prueba"
   ClientHeight    =   4935
   ClientLeft      =   2910
   ClientTop       =   1860
   ClientWidth     =   6480
   Icon            =   "Nag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6480
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdRegCode 
      Caption         =   "Entre Código &Registro"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdEvaluate 
      Caption         =   "Usar &Version Evaluación"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "&Compre ahora"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   6135
   End
   Begin VB.Label lblTrial 
      Alignment       =   2  'Center
      Caption         =   "Faltan 30 días"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmNag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuy_Click()
'abre la ventana de compra
frmBuy.Show vbModal
End Sub

Private Sub cmdEvaluate_Click()
'-------------------------------------
'Permite evaluar el programa
'-------------------------------------
PreviaInstancia = True
Unload Me
End Sub

Private Sub cmdRegCode_Click()
'abre la ventana de registro
frmRegCode.Show vbModal
End Sub

Private Sub Form_Load()

Dim intDays As Integer
intDays = modRegCode.getTrialDays
PBar.Max = 30

If Not (intDays < 0) Then
    PBar.Value = intDays
    lblTrial = "Te faltan " & CStr(intDays) & " días para usar esta Demo.  Para utilizar este software en su computadora y remover esta pantalla por favor regístrelo."
End If

If intDays <= 0 Then 'Expired
    cmdEvaluate.Enabled = False '  Ya no evalua
    lblTrial = "Tu version de prueba ha expirado.  Por favor registra esta copia.  GRACIAS!"
    lblTrial.ForeColor = vbRed
End If

Label1 = "Bienvenido a la version de prueba del PVOnline, usted puede usar este programa por 30 días."
'Label2 =

'Posiciona los botones aleatoriamente
Call RandomButtons
End Sub

Private Function RandomButtons()
'esta funcion posiciona los botones aleatoriamente
Dim j As Integer, pos(6) As Long, m As Integer
Dim n As Integer

pos(1) = cmdBuy.Left
pos(2) = cmdEvaluate.Left
pos(3) = cmdRegCode.Left

Randomize
j = Int((Rnd * 3) + 1)
cmdBuy.Left = pos(j)

n = 4
For m = 1 To 3
    If m <> j Then
        pos(n) = pos(m)
        n = n + 1
    End If
Next

Do While j < 4
    DoEvents
    Randomize
    j = Int((Rnd * 5) + 1)
Loop
cmdEvaluate.Left = pos(j)
If j = 5 Then
    cmdRegCode.Left = pos(4)
Else
    cmdRegCode.Left = pos(5)
End If
End Function


Private Sub Form_Unload(Cancel As Integer)
PreviaInstancia = True
If modRegCode.getTrialDays < 0 Then Termino = True
End Sub
