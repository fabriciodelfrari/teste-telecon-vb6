VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public class Cliente



Dim lCodigo As Long
Dim sNome As String
Dim sEndereco As String
Dim sCidade As String
Dim sBairro As String
Dim sCpf As String
Dim dLimiteCredito As Double
Dim dValorGasto As Double
Dim bSexo As Boolean


end Class

Dim ctTelefone As ctClienteTelefone







