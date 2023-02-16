Attribute VB_Name = "mdlConexaoBanco"
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ServidorBanco As String
Dim BancoDeDados As String
Dim Usuario As String
Dim Senha As String


Public Sub ConectarBanco()
    Set cn = New ADODB.Connection
    Dim sStringConexao As String

    LerConfiguracoesParaAcessoAoBanco

    sStringConexao = "Provider=SQLOLEDB;Data Source=" & ServidorBanco
    sStringConexao = sStringConexao & ";Initial Catalog=" & BancoDeDados
    sStringConexao = sStringConexao & ";User ID=" & Usuario & ";Password=" & Senha & ";"

    cn.Open sStringConexao

End Sub


Public Sub DesconectarBanco()
    cn.Close
    Set cn = Nothing
End Sub


Public Sub FecharRecordset()
    rs.Close
End Sub

Public Function fPesquisaBanco(ByVal sQuery As String) As ADODB.Recordset

    ConectarBanco

    Set rs = CreateObject("ADODB.Recordset")

    rs.Open sQuery, cn

    Set fPesquisaBanco = rs

    'Set rs = Nothing
    'DesconectarBanco

End Function

Public Sub InserirOuDeletarNoBanco(ByVal sQuery As String)
    On Error GoTo TrataErro
    Dim cmd As ADODB.Command
    Dim iRetorno As Integer

    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sQuery
    End With
    
    cmd.Execute

TrataErro:
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro ao cadastrar o cliente." & " Erro: " & Err.Number & " - " & Err.Description, vbInformation, "Atenção!"
    End If

    'Verificar como gerar logs de erro.
End Sub

Private Sub LerConfiguracoesParaAcessoAoBanco()
    Dim sNomeArquivo As String
    Dim iNumerosDeLinha As Integer
    Dim sLinha As String
    Dim arrLinha() As String

    sNomeArquivo = App.Path & "\config.txt"

    'Abrir o arquivo config.txt
    iNumerosDeLinha = FreeFile
    Open sNomeArquivo For Input As #iNumerosDeLinha

    'Ler as linhas do arquivo
    Do Until EOF(iNumerosDeLinha)
        Line Input #iNumerosDeLinha, sLinha
        arrLinha() = Split(sLinha, " = ")    'Criar chave e valor, dividindo a linha

        'Atribuir os valores as variaveis
        Select Case arrLinha(0)
        Case "Servidor"
            ServidorBanco = arrLinha(1)
        Case "Banco"
            BancoDeDados = arrLinha(1)
        Case "Usuario"
            Usuario = arrLinha(1)
        Case "Senha"
            Senha = arrLinha(1)
        End Select
    Loop

    Close #iNumerosDeLinha
End Sub







