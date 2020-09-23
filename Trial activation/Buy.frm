VERSION 5.00
Begin VB.Form frmBuy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro del programa"
   ClientHeight    =   1815
   ClientLeft      =   3000
   ClientTop       =   2580
   ClientWidth     =   4455
   Icon            =   "Buy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdWebsite 
      Caption         =   "Sitio Web"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblID 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Por favor anote su computer ID,se te lo pedirá en el proceso de instalación o llame a su vendedor ."
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmBuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
'con esto se abre el archivo de ayuda
ShellExecute Me.hwnd, "open", "help.chm", 0, 0, 1
Unload Me
End Sub

Private Sub cmdWebsite_Click()
'con esto se abre el navegador
ShellExecute Me.hwnd, "open", "http://www.dalpcorp.com", 0, 0, 1
Unload Me
End Sub

Private Sub Form_Load()
lblID.Caption = "Tu Computer ID : " & modRegCode.getComputerID
End Sub
