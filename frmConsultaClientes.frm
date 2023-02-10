VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ACTIVETEXT.OCX"
Begin VB.Form frmConsultaClientes 
   Caption         =   "Consulta de Clientes"
   ClientHeight    =   8910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   15960
   Begin rdActiveText.ActiveText txtCodigo 
      Height          =   315
      Left            =   180
      TabIndex        =   35
      Top             =   540
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.Frame frInformacoesFinanceiras 
      Caption         =   "Informações Financeiras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   12180
      TabIndex        =   20
      Top             =   1020
      Width           =   3195
      Begin rdActiveText.ActiveText txtValorGasto 
         Height          =   315
         Left            =   180
         TabIndex        =   32
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtLimiteCredito 
         Height          =   315
         Left            =   180
         TabIndex        =   30
         Top             =   780
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin VB.Label lbValorGasto 
         Caption         =   "Valor Gasto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   31
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label lbLimiteCredito 
         Caption         =   "Limite de Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   29
         Top             =   420
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdUltimoCliente 
      Caption         =   "Último"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   11760
      TabIndex        =   28
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdProximoCliente 
      Caption         =   "Próximo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   10560
      TabIndex        =   27
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   9360
      TabIndex        =   26
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdNovoCliente 
      Caption         =   "Novo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   8160
      TabIndex        =   25
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   6960
      TabIndex        =   24
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   5760
      TabIndex        =   23
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdClienteAnterior 
      Caption         =   "Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   4560
      TabIndex        =   22
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrimeiroCliente 
      Caption         =   "Primeiro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3360
      TabIndex        =   21
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdProcura 
      Height          =   315
      Left            =   1920
      Picture         =   "frmConsultaClientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   540
      Width           =   375
   End
   Begin VB.Frame frDadosGerais 
      Caption         =   "Dados Gerais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   12015
      Begin VB.ComboBox cboSexo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8400
         TabIndex        =   34
         Text            =   "M/F"
         Top             =   720
         Width           =   915
      End
      Begin VB.CommandButton cmdBuscaEndereco 
         Caption         =   "Consultar CEP"
         Height          =   315
         Left            =   3360
         TabIndex        =   18
         Top             =   1560
         Width           =   1575
      End
      Begin rdActiveText.ActiveText txtCpf 
         Height          =   315
         Left            =   5760
         TabIndex        =   2
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   14
         TextMask        =   7
         RawText         =   7
         Mask            =   "###.###.###-##"
         FontName        =   "MS Sans Serif"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText txtNome 
         Height          =   315
         Left            =   300
         TabIndex        =   3
         Top             =   720
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText txtTelefoneContato 
         Height          =   315
         Left            =   5760
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   14
         TextMask        =   9
         RawText         =   9
         Mask            =   "(##)####-#####"
         FontName        =   "MS Sans Serif"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText txtEndereco 
         Height          =   315
         Left            =   300
         TabIndex        =   5
         Top             =   2400
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   55
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText txtCep 
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Top             =   1560
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   9
         TextMask        =   6
         RawText         =   6
         Mask            =   "#####-###"
         FontName        =   "MS Sans Serif"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText txtCidade 
         Height          =   315
         Left            =   300
         TabIndex        =   7
         Top             =   3240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText txtNumero 
         Height          =   315
         Left            =   5760
         TabIndex        =   8
         Top             =   2400
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   6
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   9,75
      End
      Begin rdActiveText.ActiveText txtBairro 
         Height          =   315
         Left            =   3060
         TabIndex        =   9
         Top             =   3240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   9,75
      End
      Begin VB.Label lbSexo 
         Caption         =   "Sexo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8460
         TabIndex        =   33
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lbNome 
         Caption         =   "Nome Completo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   420
         TabIndex        =   17
         Top             =   420
         Width           =   1635
      End
      Begin VB.Label lbCpf 
         Caption         =   "CPF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   5820
         TabIndex        =   16
         Top             =   420
         Width           =   675
      End
      Begin VB.Label lbCidade 
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   2940
         Width           =   675
      End
      Begin VB.Label lbBairro 
         Caption         =   "Bairro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3180
         TabIndex        =   14
         Top             =   2940
         Width           =   675
      End
      Begin VB.Label lbNome 
         Caption         =   "Endereço"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   2100
         Width           =   1635
      End
      Begin VB.Label lbCep 
         Caption         =   "CEP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label lbNumero 
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   5820
         TabIndex        =   11
         Top             =   2100
         Width           =   795
      End
      Begin VB.Label lbTelefoneContato 
         Caption         =   "Telefone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5820
         TabIndex        =   10
         Top             =   1260
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "frmConsultaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conexao As clsConexaobanco

Private Sub cmdClienteAnterior_Click(Index As Integer)
On Error GoTo TrataErro

    'verificar possivel bug em casos de não haver clientes com o codigo inferior
    '(exemplo: cliente atual é 205, e o cliente com o cod 204 foi deletado do banco)
    Dim iCodigoClienteAtual As String
    Dim iCodigoClienteAnterior As Integer
    Dim iCodigoPrimeiro As Integer
    
    iCodigoPrimeiro = fObterPrimeiroCodigoCliente
    iCodigoClienteAtual = txtCodigo.Text
    iCodigoClienteAnterior = Val(iCodigoClienteAtual) - 1
    
    If txtCodigo.Text = Empty Then
        MsgBox "Não é possível ir para o anterior pois não há informações de clientes no formulário", vbInformation, "Atenção!"
        txtCodigo.SetFocus
        Exit Sub
    ElseIf iCodigoClienteAnterior < iCodigoPrimeiro Then
        MsgBox "Não há mais clientes para informar.", vbInformation, "Atenção!"
        Exit Sub
    End If
    
    sLimparCampos
    
    sBuscarClienteAnterior iCodigoClienteAtual
    
TrataErro:
 If Err.Number <> 0 Then
        MsgBox "Erro ao buscar cliente anterior.", vbInformation, "Atenção!"
        End If

End Sub

Private Sub cmdPrimeiroCliente_Click(Index As Integer)
    Dim iCodPrimeiroCliente As String
    iCodPrimeiroCliente = str(fObterPrimeiroCodigoCliente)
    iCodPrimeiroCliente = Trim(iCodPrimeiroCliente)
    sBuscarCliente str(iCodPrimeiroCliente)
End Sub

Private Sub cmdProcura_Click()
    lbCep.Visible = False
    txtCep.Visible = False
    
    sBuscarCliente txtCodigo.Text
End Sub

Private Sub cmdProximoCliente_Click(Index As Integer)
    Dim iUltimoCliente As Integer
    
    If txtCodigo.Text = Empty Then
        MsgBox "Não é possível ir para o anterior pois não há informações de clientes no formulário", vbInformation, "Atenção!"
        txtCodigo.SetFocus
        Exit Sub
    End If

    sBuscarClienteProximo txtCodigo.Text
End Sub

Private Sub cmdUltimoCliente_Click(Index As Integer)
    Dim iCodUltimoCliente As Integer 'funcao ira retornar o ultimo codigo + 1
    iCodUltimoCliente = fObterProximoCodigoCliente - 1
    
    sBuscarCliente str(iCodUltimoCliente)
End Sub

Private Sub Form_Load()
    cmdBuscaEndereco.Visible = False
    cboSexo.AddItem "M"
    cboSexo.AddItem "F"
End Sub

Private Sub cmdGravar_Click(Index As Integer)
    sCadastrarCliente
End Sub

Private Sub cmdNovoCliente_Click(Index As Integer)
On Error GoTo TrataErro
    
    lbCep.Visible = True
    txtCep.Visible = True
    sDestrancarCampos
    sLimparCampos
    cmdBuscaEndereco.Visible = True
    lbValorGasto.Visible = False
    txtValorGasto.Visible = False

TrataErro:
    If Err.Number <> 0 Then
        MsgBox "Erro ao iniciar novo cadastro.", vbInformation, "Atenção!"
        End If
End Sub
Private Sub sBuscarCliente(ByVal lCodCliente As String)
On Error GoTo TrataErro
    
    lCodCliente = Trim(lCodCliente)
    Dim rsRetornoBanco As ADODB.Recordset
    Dim sQuery As String
    Dim sEndereco As String
    Dim sNumeroEndereco As String
    Dim sTelefoneCompleto As String
    Dim arrEndereco() As String
    Set Conexao = New clsConexaobanco
    
    sQuery = "SELECT c.*, ct.CodigoArea, ct.Telefone, ct.Observacao FROM Clientes c "
    sQuery = sQuery & "LEFT JOIN ClienteTelefones ct on c.CodCliente = ct.CodCliente "
    sQuery = sQuery & "WHERE c.CodCliente = " & lCodCliente
    
    
    Set rsRetornoBanco = Conexao.fPesquisaBanco(sQuery)
    
    sTelefoneCompleto = rsRetornoBanco("CodigoArea") & rsRetornoBanco("Telefone")
    sTelefoneCompleto = Format(sTelefoneCompleto, "(##)#####-####")
    arrEndereco() = Split(rsRetornoBanco("Endereco"), ", ")
    
    txtCodigo.Text = rsRetornoBanco("CodCliente")
    txtNome.Text = rsRetornoBanco("Nome")
    txtCpf.Text = rsRetornoBanco("CPF")
    txtEndereco.Text = arrEndereco(0)
    txtNumero.Text = arrEndereco(1)
    txtCidade.Text = rsRetornoBanco("Cidade")
    txtBairro.Text = rsRetornoBanco("Bairro")
    txtTelefoneContato.Text = sTelefoneCompleto
    txtLimiteCredito.Text = rsRetornoBanco("LimiteCredito")
    txtValorGasto.Text = rsRetornoBanco("ValorGasto")
    
    If rsRetornoBanco("Sexo") = "0" Then
        cboSexo.ListIndex = 0
    Else
        cboSexo.ListIndex = 1
    End If
    
    sTrancarCampos
    
TrataErro:
    If Err.Number <> 0 Then
         MsgBox "Ocorreu um erro ao buscar o cliente: " & Err.Description & " - " & Err.Number
    End If
End Sub
Private Sub sBuscarClienteProximo(ByVal lCodCliente As String)
On Error GoTo TrataErro
    
    lCodCliente = Trim(lCodCliente)
    Dim rsRetornoBanco As ADODB.Recordset
    Dim sQuery As String
    Dim sEndereco As String
    Dim sNumeroEndereco As String
    Dim sTelefoneCompleto As String
    Dim arrEndereco() As String
    Set Conexao = New clsConexaobanco
    
    sQuery = "SELECT TOP 1  c.*, ct.CodigoArea, ct.Telefone, ct.Observacao FROM Clientes c "
    sQuery = sQuery & "INNER Join ClienteTelefones ct on c.CodCliente = ct.CodCliente "
    sQuery = sQuery & "WHERE c.CodCliente > " & lCodCliente
    sQuery = sQuery & " ORDER BY c.CodCliente ASC"
    
    
    Set rsRetornoBanco = Conexao.fPesquisaBanco(sQuery)
    
    If rsRetornoBanco.EOF Then
        MsgBox "Não há mais clientes para buscar.", vbInformation, "Atenção!"
        Exit Sub
    End If
    
    sTelefoneCompleto = rsRetornoBanco("CodigoArea") & rsRetornoBanco("Telefone")
    sTelefoneCompleto = Format(sTelefoneCompleto, "(##)#####-####")
    arrEndereco() = Split(rsRetornoBanco("Endereco"), ", ")
    
    txtCodigo.Text = rsRetornoBanco("CodCliente")
    txtNome.Text = rsRetornoBanco("Nome")
    txtCpf.Text = rsRetornoBanco("CPF")
    txtEndereco.Text = arrEndereco(0)
    txtNumero.Text = arrEndereco(1)
    txtCidade.Text = rsRetornoBanco("Cidade")
    txtBairro.Text = rsRetornoBanco("Bairro")
    txtTelefoneContato.Text = sTelefoneCompleto
    txtLimiteCredito.Text = rsRetornoBanco("LimiteCredito")
    txtValorGasto.Text = rsRetornoBanco("ValorGasto")
    
    'If rsRetornoBanco("Sexo") = "0" Then
        'cboSexo.ListIndex = 0
    'Else
       ' cboSexo.ListIndex = 1
   ' End If
    
    sTrancarCampos
    
TrataErro:
    If Err.Number <> 0 Then
         MsgBox "Ocorreu um erro ao buscar o cliente: " & Err.Description & " - " & Err.Number
    End If
End Sub
Private Sub sBuscarClienteAnterior(ByVal lCodCliente As String)
On Error GoTo TrataErro
    
    lCodCliente = Trim(lCodCliente)
    Dim rsRetornoBanco As ADODB.Recordset
    Dim sQuery As String
    Dim sEndereco As String
    Dim sNumeroEndereco As String
    Dim sTelefoneCompleto As String
    Dim arrEndereco() As String
    Set Conexao = New clsConexaobanco
    
    sQuery = "SELECT TOP 1  c.*, ct.CodigoArea, ct.Telefone, ct.Observacao FROM Clientes c "
    sQuery = sQuery & "INNER Join ClienteTelefones ct on c.CodCliente = ct.CodCliente "
    sQuery = sQuery & "WHERE c.CodCliente < " & lCodCliente
    sQuery = sQuery & " ORDER BY c.CodCliente DESC"
    
    
    Set rsRetornoBanco = Conexao.fPesquisaBanco(sQuery)
    
    sTelefoneCompleto = rsRetornoBanco("CodigoArea") & rsRetornoBanco("Telefone")
    sTelefoneCompleto = Format(sTelefoneCompleto, "(##)#####-####")
    arrEndereco() = Split(rsRetornoBanco("Endereco"), ", ")
    
    txtCodigo.Text = rsRetornoBanco("CodCliente")
    txtNome.Text = rsRetornoBanco("Nome")
    txtCpf.Text = rsRetornoBanco("CPF")
    txtEndereco.Text = arrEndereco(0)
    txtNumero.Text = arrEndereco(1)
    txtCidade.Text = rsRetornoBanco("Cidade")
    txtBairro.Text = rsRetornoBanco("Bairro")
    txtTelefoneContato.Text = sTelefoneCompleto
    txtLimiteCredito.Text = rsRetornoBanco("LimiteCredito")
    txtValorGasto.Text = rsRetornoBanco("ValorGasto")
    
    'If rsRetornoBanco("Sexo") = "0" Then
        'cboSexo.ListIndex = 0
    'Else
       ' cboSexo.ListIndex = 1
   ' End If
    
    sTrancarCampos
    
TrataErro:
    If Err.Number <> 0 Then
         MsgBox "Ocorreu um erro ao buscar o cliente: " & Err.Description & " - " & Err.Number
    End If
End Sub

Private Sub sCadastrarCliente()
 On Error GoTo TrataErro
    
    If Not bVerificarCamposVaziosOuExcedentes Then
        Exit Sub
    End If

    Set Conexao = New clsConexaobanco

    Dim sQuery As String
    Dim btSexo As Byte
    Dim sRemoveMascara As clsTratamentoMascaras
    Dim sCpfSemMascara As String
    Dim iProximoCodigo As Integer

    If cboSexo.Text = "M" Then
        btSexo = 0
    ElseIf cboSexo.Text = "F" Then
        btSexo = 1
    End If

    Set sRemoveMascara = New clsTratamentoMascaras
    sCpfSemMascara = sRemoveMascara.sRemoveMascaraCpf(txtCpf.Text)
    
    iProximoCodigo = fObterProximoCodigoCliente

    Conexao.ConectarBanco

    sQuery = "INSERT INTO Clientes(CodCliente, Nome,Endereco,Cidade,Bairro,CPF,LimiteCredito,ValorGasto,Sexo) "
    sQuery = sQuery & "VALUES(" & iProximoCodigo & ",'" & txtNome.Text & "', '" & txtEndereco.Text & ", " & txtNumero.Text & "', '"
    sQuery = sQuery & txtCidade.Text & "', '" & txtBairro.Text & "','"
    sQuery = sQuery & sCpfSemMascara & "'," & txtLimiteCredito.Text & ", 0, " & btSexo & ")"

    Conexao.InserirNoBanco (sQuery)

    Conexao.DesconectarBanco
    
    sCadastrarTelefone txtTelefoneContato.Text, str(iProximoCodigo)
    
    MsgBox "Cliente cadastrado!", vbInformation, "Cadastro"
    
    sLimparCampos

TrataErro:
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro ao cadastrar o cliente: " & Err.Description & " - " & Err.Number
    End If

End Sub
Private Sub sCadastrarTelefone(ByVal sTelefoneCompleto As String, ByVal sCodCliente As String)
 On Error GoTo TrataErro
    Dim iProximoCodClienteTelefone As Integer
    Dim sCodArea As String
    Dim sTelefone As String
    Dim Observacao As String
    Dim sQuery As String
    
    iProximoCodClienteTelefone = fObterProximoCodigoClienteTelefone
    sCodArea = Mid(sTelefoneCompleto, 2, 2)
    sTelefone = Mid(sTelefoneCompleto, 5, 10)
    sTelefone = Replace(sTelefone, "-", "")
    Set Conexao = New clsConexaobanco



    Conexao.ConectarBanco

    sQuery = "INSERT INTO ClienteTelefones (CodClienteTelefone, CodCLiente, CodigoArea, Telefone, Observacao) "
    sQuery = sQuery & "VALUES(" & iProximoCodClienteTelefone & ", " & sCodCliente & ", " & sCodArea & ", "
    sQuery = sQuery & sTelefone & ", '-')"
    

    Conexao.InserirNoBanco (sQuery)

TrataErro:
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro ao cadastrar o número do cliente: " & Err.Description & " - " & Err.Number
    End If

End Sub
Private Function fObterPrimeiroCodigoCliente() As Integer
On Error GoTo TrataErro

    Dim iPrimeiroCodigo As Integer
    Dim rsRetornoBanco As ADODB.Recordset
    Set Conexao = New clsConexaobanco
    
    Set rsRetornoBanco = Conexao.fPesquisaBanco("SELECT MIN(CodCliente) as Primeiro FROM CLIENTES")
    
    iPrimeiroCodigo = Val(rsRetornoBanco("Primeiro"))
    
    fObterPrimeiroCodigoCliente = iPrimeiroCodigo

TrataErro:
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro ao buscar o código do cliente." & Err.Number & " - " & Err.Description, vbInformation, "Atenção!"
    End If
End Function
Private Function fObterProximoCodigoClienteTelefone() As Integer
On Error GoTo TrataErro

    Dim iProxCodigoClienteTelefone As Integer
    Dim rsRetornoBanco As ADODB.Recordset
    Set Conexao = New clsConexaobanco
    
    Set rsRetornoBanco = Conexao.fPesquisaBanco("SELECT MAX(CodClienteTelefone) as Maior FROM ClienteTelefones")
    
    iProxCodigoClienteTelefone = Val(rsRetornoBanco("Maior"))
    
    fObterProximoCodigoClienteTelefone = iProxCodigoClienteTelefone + 1

TrataErro:
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro ao buscar o código do cliente (telefone)." & Err.Number & " - " & Err.Description, vbInformation, "Atenção!"
    End If
End Function
Private Function fObterProximoCodigoCliente() As Integer
On Error GoTo TrataErro

    Dim Conexao As clsConexaobanco
    Dim iProxCodigoCliente As Integer
    Dim rsRetornoBanco As ADODB.Recordset
    Set Conexao = New clsConexaobanco
    
    Set rsRetornoBanco = Conexao.fPesquisaBanco("SELECT MAX(CodCliente) as Maior FROM CLIENTES")
    
    iProxCodigoCliente = Val(rsRetornoBanco("maior"))
    
    fObterProximoCodigoCliente = iProxCodigoCliente + 1

TrataErro:
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro ao buscar o código do cliente." & Err.Number & " - " & Err.Description, vbInformation, "Atenção!"
    End If
End Function

Private Sub cmdBuscaEndereco_Click()
On Error GoTo TrataErro
    If Len(txtCep.Text) < 1 Then
        MsgBox "Campo CEP está vázio ou incompleto. Por favor, verifique.", vbInformation, "Atenção!"
        Exit Sub
    End If

    Dim oCepCliente As Object
    Dim oJsonParse As Object
    Set oCepCliente = CreateObject("WinHttp.WinHttpRequest.5.1")

    oCepCliente.Open "GET", "https://viacep.com.br/ws/" & txtCep.Text & "/json/", False
    oCepCliente.Send
    
    
    If InStr(oCepCliente.ResponseText, "erro") Then
        MsgBox "CEP não localizado. Verifique o CEP ou insira os dados manualmente.", vbInformation, "Atenção!"
        Exit Sub
    End If
    

    If oCepCliente.Status = 200 Then
        ' Parse da resposta para Json
        Set oJsonParse = fParsearJson(oCepCliente.ResponseText)

        ' Inserir as informações nos campos
        txtEndereco.Text = oJsonParse("logradouro")
        txtCidade.Text = oJsonParse("localidade")
        txtBairro.Text = oJsonParse("bairro")
    Else
        MsgBox "Não foi possível localizar o endereço. Por favor, insira os dados manualmente.", vbInformation, "Consulta CEP"
    End If
    
TrataErro:
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro ao localizar o endereço. Por favor, insira os dados manualmente.", vbInformation, "Atenção!"
    End If
End Sub
Private Function fParsearJson(ByVal sObjJson As String) As Object
    Dim obJson As Object
    Set obJson = JSON.parse(sObjJson)
    Set fParsearJson = obJson
End Function

Private Function bVerificarCamposVaziosOuExcedentes() As Boolean

    If Len(txtNome.Text) < 1 Then
        MsgBox "Campo nome está vazio ou incompleto.", vbInformation, "Atenção!"
        bVerificarCamposVaziosOuExcedentes = False
        Exit Function
    ElseIf Len(txtCpf.Text) < 1 Then
        MsgBox "Campo CPF está vazio ou incompleto.", vbInformation, "Atenção!"
        bVerificarCamposVaziosOuExcedentes = False
        Exit Function
    ElseIf Len(txtCep.Text) < 1 Then
        MsgBox "Campo CEP está vázio ou incompleto.", vbInformation, "Atenção!"
        bVerificarCamposVaziosOuExcedentes = False
        Exit Function
    ElseIf Len(txtEndereco.Text) < 5 Then
        MsgBox "Campo Endereço está vázio ou incompleto.", vbInformation, "Atenção!"
        bVerificarCamposVaziosOuExcedentes = False
        Exit Function
    ElseIf Len(txtEndereco.Text) > 60 Then
        MsgBox "Campo Endereço excedeu o limite máximo de caracteres (60). Por favor, abrevie.", vbInformation, "Atenção!"
        bVerificarCamposVaziosOuExcedentes = False
        Exit Function
    ElseIf Len(txtNumero.Text) < 1 Then
        MsgBox "Campo Número está vázio ou incompleto.", vbInformation, "Atenção!"
        txtNumero.SetFocus
        bVerificarCamposVaziosOuExcedentes = False
        Exit Function
    ElseIf Len(txtCidade.Text) < 3 Then
        MsgBox "Campo Cidade está vázio ou incompleto.", vbInformation, "Atenção!"
        bVerificarCamposVaziosOuExcedentes = False
        Exit Function
    ElseIf Len(txtCidade.Text) > 40 Then
        MsgBox "Campo Cidade excedeu o limite máximo de caracteres (40). Por favor, abrevie.", vbInformation, "Atenção!"
        bVerificarCamposVaziosOuExcedentes = False
        Exit Function
    ElseIf Len(txtBairro.Text) < 3 Then
        MsgBox "Campo Bairro está vázio ou incompleto.", vbInformation, "Atenção!"
        bVerificarCamposVaziosOuExcedentes = False
        Exit Function
    ElseIf Len(cboSexo.Text) < 1 Or cboSexo.Text = "M/F" Then
        MsgBox "Campo Sexo está vázio ou definido como padrão. Por favor, verifique.", vbInformation, "Atenção!"
        bVerificarCamposVaziosOuExcedentes = False
        Exit Function
    ElseIf Len(txtTelefoneContato.Text) < 14 Then
        MsgBox "Campo Telefone está vázio ou incompleto.", vbInformation, "Atenção!"
        bVerificarCamposVaziosOuExcedentes = False
        Exit Function
    ElseIf Len(txtLimiteCredito.Text) < 1 Then
        MsgBox "Campo Limite de Crédito está vázio ou incompleto.", vbInformation, "Atenção!"
        txtLimiteCredito.SetFocus
        bVerificarCamposVaziosOuExcedentes = False
        Exit Function
    End If

    bVerificarCamposVaziosOuExcedentes = True

End Function

Private Sub sLimparCampos()
    txtCodigo.Text = ""
    txtNome.Text = ""
    txtCpf.Text = ""
    cboSexo.Clear
    txtCep.Text = ""
    txtEndereco.Text = ""
    txtCidade.Text = ""
    txtBairro.Text = ""
    txtNumero.Text = ""
    txtTelefoneContato.Text = ""
    txtLimiteCredito.Text = ""
    txtValorGasto.Text = ""
End Sub

Private Sub sTrancarCampos()
    txtNome.Locked = True
    txtCpf.Locked = True
    cboSexo.Locked = True
    txtCep.Locked = True
    txtEndereco.Locked = True
    txtCidade.Locked = True
    txtBairro.Locked = True
    txtNumero.Locked = True
    txtTelefoneContato.Locked = True
    txtLimiteCredito.Locked = True
    txtValorGasto.Locked = True
End Sub

Private Sub sDestrancarCampos()
    txtNome.Locked = False
    txtCpf.Locked = False
    cboSexo.Locked = False
    txtCep.Locked = False
    txtEndereco.Locked = False
    txtCidade.Locked = False
    txtBairro.Locked = False
    txtNumero.Locked = False
    txtTelefoneContato.Locked = False
    txtLimiteCredito.Locked = False
    txtValorGasto.Locked = False
End Sub



