VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContextoConsultaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Contexto
    Cadastro
    Alteracao
    Busca
    Exclusao
End Enum

Private eContextoAtual As Contexto

Public Property Get ContextoAtual() As Contexto
    ContextoAtual = eContextoAtual
End Property

Public Property Let DefineContexto(ByVal vNovoContexto As Contexto)
    eContextoAtual = vNovoContexto
End Property


