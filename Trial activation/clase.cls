VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTrial"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub Show()
    Load frmNag
    frmNag.Show 1
End Sub
Public Property Get PrevInstancia() As Boolean
 PrevInstancia = PreviaInstancia
End Property

Public Property Get EstaRegistrado() As Boolean
    EstaRegistrado = modRegCode.IsRegistered
End Property
Public Property Get Finalizar() As Boolean
 Finalizar = Termino
End Property

