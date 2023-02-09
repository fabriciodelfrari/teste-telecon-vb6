VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConexaobanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Public Sub ConectarBanco()
    Set cn = New ADODB.Connection
    cn.Open "Provider=SQLOLEDB;Data Source= NBKDEV03;Initial Catalog=TESTEVB6;User ID=trabmaisprati;Password=trabmaisprati;"
End Sub


Public Sub DesconectarBanco()
    cn.Close
    Set cn = Nothing
End Sub


Public Sub FecharRecordset()
    rs.Close
End Sub

Public Function oPesquisaBanco(ByVal sQuery As String)
    Set rs = New ADODB.Recordset
    rs.Open sQuery, cn

    Set oPesquisaBanco = rs

    Set rs = Nothing
End Function

Public Sub InserirNoBanco(ByVal sQuery As String)
    On Error GoTo TratamentoDeErro
    Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sQuery
    End With

    cmd.Execute

TratamentoDeErro:
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro ao cadastrar o cliente. Contate o suporte." & " Erro: " & Err.Number & " - " & Err.Description, vbInformation, "Atenção!"
    End If

    'Verificar como gerar logs de erro.
End Sub

