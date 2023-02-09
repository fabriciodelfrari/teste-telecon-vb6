VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ACTIVETEXT.OCX"
Begin VB.Form frmCadClientes 
   AutoRedraw      =   -1  'True
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Cadastro de Clientes"
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   11400
   Begin VB.CommandButton cmdTesteConn 
      Caption         =   "teste"
      Height          =   495
      Left            =   4320
      TabIndex        =   22
      Top             =   1200
      Width           =   1635
   End
   Begin VB.ComboBox cboSexo 
      Height          =   360
      Left            =   8160
      TabIndex        =   21
      Text            =   "M/F"
      Top             =   2520
      Width           =   1275
   End
   Begin rdActiveText.ActiveText txtCpf 
      Height          =   315
      Left            =   5820
      TabIndex        =   20
      Top             =   2580
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   60
      TabIndex        =   19
      Top             =   2580
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
      Left            =   7920
      TabIndex        =   16
      Top             =   3480
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
      Left            =   120
      TabIndex        =   11
      Top             =   4440
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
      Left            =   120
      TabIndex        =   9
      Top             =   3600
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
   Begin VB.CommandButton cmdBuscaEndereco 
      Caption         =   "Consultar CEP"
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin rdActiveText.ActiveText txtCodigo 
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
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
   Begin rdActiveText.ActiveText txtCidade 
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   5280
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
      Left            =   5640
      TabIndex        =   13
      Top             =   4440
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
      Left            =   2880
      TabIndex        =   15
      Top             =   5280
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
   Begin MSComctlLib.Toolbar tbrCadClienteFrm 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1588
      ButtonWidth     =   1879
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ilImageListCadastro"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cad. Cliente"
            Key             =   "CadCliente"
            ImageKey        =   "GravCadCliente"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "CancCad"
            ImageKey        =   "CancelarCad"
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ilImageListCadastro 
         Left            =   9600
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadClientes.frx":0000
               Key             =   "GravCadCliente"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadClientes.frx":0614
               Key             =   "Cancelar"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadClientes.frx":0AB9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadClientes.frx":0F5E
               Key             =   "CancelarCad"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lbTelefoneContato 
      Caption         =   "Telefone"
      Height          =   315
      Left            =   7920
      TabIndex        =   17
      Top             =   3180
      Width           =   855
   End
   Begin VB.Label lbNumero 
      Caption         =   "Número"
      Height          =   315
      Index           =   0
      Left            =   5700
      TabIndex        =   14
      Top             =   4080
      Width           =   795
   End
   Begin VB.Label lbCep 
      Caption         =   "CEP"
      Height          =   315
      Left            =   180
      TabIndex        =   10
      Top             =   3300
      Width           =   1155
   End
   Begin VB.Label lbNome 
      Caption         =   "Endereço"
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   4080
      Width           =   1635
   End
   Begin VB.Label lbBairro 
      Caption         =   "Bairro"
      Height          =   315
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   4860
      Width           =   675
   End
   Begin VB.Label lbCidade 
      Caption         =   "Cidade"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   4860
      Width           =   675
   End
   Begin VB.Label lbSexo 
      Caption         =   "Sexo"
      Height          =   315
      Index           =   0
      Left            =   8280
      TabIndex        =   4
      Top             =   2160
      Width           =   675
   End
   Begin VB.Label lbCpf 
      Caption         =   "CPF"
      Height          =   315
      Index           =   1
      Left            =   5880
      TabIndex        =   3
      Top             =   2220
      Width           =   675
   End
   Begin VB.Label lbNome 
      Caption         =   "Nome Completo"
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   2220
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   675
   End
End
Attribute VB_Name = "frmCadClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTesteConn_Click()
    On Error GoTo TratamentoDeErro
    If Not bVerificarCamposVaziosOuExcedentes Then
        Exit Sub
    End If

    Dim Conexao As clsConexaobanco
    Set Conexao = New clsConexaobanco
    Dim rs As ADODB.Recordset

    Dim sQuery As String
    Dim btSexo As Byte
    Dim sRemoveMascara As clsTratamentoMascaras
    Dim sCpfSemMascara As String

    If cboSexo.Text = "M" Then
        btSexo = 0
    ElseIf cboSexo.Text = "F" Then
        btSexo = 1
    End If

    Set sRemoveMascara = New clsTratamentoMascaras
    sCpfSemMascara = sRemoveMascara.sRemoveMascaraCpf(txtCpf.Text)


    Conexao.ConectarBanco

    sQuery = "insert into Clientes(Nome,Endereco,Cidade,Bairro,CPF,LimiteCredito,ValorGasto,Sexo) "
    sQuery = sQuery & "VALUES('" & txtNome.Text & "', '" & txtEndereco.Text & ", " & txtNumero.Text & "', '"
    sQuery = sQuery & txtCidade.Text & "', '" & txtBairro.Text & "','"
    sQuery = sQuery & sCpfSemMascara & "', 1000, 0, " & btSexo & ")"

    Conexao.InserirNoBanco (sQuery)

    Conexao.DesconectarBanco

TratamentoDeErro:
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro ao cadastrar o cliente: " & Err.Description & " - " & Err.Number
    End If

End Sub

Private Sub Form_Load()
    cboSexo.AddItem "M"
    cboSexo.AddItem "F"
End Sub

Private Sub cmdBuscaEndereco_Click()

    If Len(txtCep.Text) < 1 Then
        MsgBox "Campo CEP está vázio ou incompleto. Por favor, verifique.", vbInformation, "Atenção!"
        Exit Sub
    End If

    Dim oCepCliente As Object
    Dim oJsonParse As Object
    Set oCepCliente = CreateObject("WinHttp.WinHttpRequest.5.1")

    oCepCliente.Open "GET", "https://viacep.com.br/ws/" & txtCep.Text & "/json/", False
    oCepCliente.Send

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
    End If

    bVerificarCamposVaziosOuExcedentes = True

End Function
Private Sub sCadastrarCliente()

End Sub


Private Sub tbrCadClienteFrm_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.key = "CadCliente" Then
        sCadastrarCliente
    ElseIf Button.key = "CancCad" Then
        Unload Me

    End If
End Sub


