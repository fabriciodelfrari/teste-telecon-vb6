VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBuscaClientes 
   AutoRedraw      =   -1  'True
   Caption         =   "Busca Clientes"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   8760
   Begin MSFlexGridLib.MSFlexGrid grdClientes 
      Height          =   6495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   11456
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmBuscaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clsConexao As New clsConexaobanco
'const's do grid
Private Const lbtCodigo As Byte = 0
Private Const lbtNome As Byte = 1
Private Const lbtCidade As Byte = 2
Private Const lbtBairro As Byte = 3
Private Const lbtContato As Byte = 4
Private Sub Form_Load()
    sFormatatarFlexGrid
    sListarClientes
End Sub
Private Sub sFormatatarFlexGrid()
    With grdClientes
        .Cols = 5
        .FixedCols = 0
        .rows = 1
        .SelectionMode = flexSelectionByRow
        
        .TextMatrix(0, lbtCodigo) = "Codigo"
        .ColWidth(lbtCodigo) = .Width * 0.1
        .TextMatrix(0, lbtNome) = "Nome"
        .ColWidth(lbtNome) = .Width * 0.2
        .TextMatrix(0, lbtCidade) = "Cidade"
        .ColWidth(lbtCidade) = .Width * 0.2
        .TextMatrix(0, lbtBairro) = "Bairro"
        .ColWidth(lbtBairro) = .Width * 0.2
        .TextMatrix(0, lbtContato) = "Contato"
        .ColWidth(lbtContato) = .Width * 0.2
    End With
    
        
End Sub

Private Sub sListarClientes()
    Dim sQuery As String
    Dim rsRetornoBanco As ADODB.Recordset
    On Error GoTo TrataErro
    
    With grdClientes
        sQuery = "SELECT c.*, ct.CodigoArea, ct.Telefone, ct.Observacao FROM Clientes c "
        sQuery = sQuery & "LEFT JOIN ClienteTelefones ct on c.CodCliente = ct.CodCliente"
        
        Set rsRetornoBanco = clsConexao.fPesquisaBanco(sQuery)
        .rows = 1
        
        Do While Not rsRetornoBanco.EOF
            .rows = .rows + 1
            
        .TextMatrix(.rows - 1, lbtCodigo) = rsRetornoBanco("CodCliente")
        .TextMatrix(.rows - 1, lbtNome) = rsRetornoBanco("Nome")
        .TextMatrix(.rows - 1, lbtCidade) = rsRetornoBanco("Cidade")
        .TextMatrix(.rows - 1, lbtBairro) = rsRetornoBanco("Bairro")
        .TextMatrix(.rows - 1, lbtContato) = rsRetornoBanco("CodigoArea") & " - " & rsRetornoBanco("Telefone")
        
        rsRetornoBanco.MoveNext
        Loop
    End With
    
    
TrataErro:
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro ao listar os clientes. " & Err.Number & " - " & Err.Description, vbInformation, "Atenção!"
    End If
End Sub

Private Sub grdClientes_DblClick()
    Dim frmFormAberto As Form
    
    'encontra o form que já está aberto, evitando que abra uma nova janela
   For Each frmFormAberto In Forms
        If frmFormAberto.Name = "frmConsultaClientes" Then
            frmFormAberto.sBuscarCliente (grdClientes.TextMatrix(grdClientes.Row, 0))
            Unload Me
            Exit Sub
        End If
    Next
    
End Sub
