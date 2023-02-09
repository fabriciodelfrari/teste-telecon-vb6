VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTratamentoMascaras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function sRemoveMascaraCpf(ByVal sCpf As String) As String
    On Error GoTo TratamentoDeErro

    sCpf = Replace(sCpf, ".", "")
    sCpf = Replace(sCpf, "-", "")

    sRemoveMascaraCpf = sCpf

TratamentoDeErro:
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro ao remover a máscara do CPF", vbInformation, "Atenção!"
    End If

End Function
