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
   Begin VB.ComboBox cboSexo 
      Height          =   360
      Left            =   8100
      TabIndex        =   21
      Text            =   "M/F"
      Top             =   2280
      Width           =   915
   End
   Begin rdActiveText.ActiveText txtCpf 
      Height          =   315
      Left            =   5820
      TabIndex        =   20
      Top             =   2280
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
      Top             =   2280
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
      Left            =   60
      TabIndex        =   11
      Top             =   3840
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
      Left            =   60
      TabIndex        =   9
      Top             =   3000
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
      Left            =   3060
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin rdActiveText.ActiveText txtCodigo 
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   1380
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
      Left            =   60
      TabIndex        =   12
      Top             =   4680
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
      Left            =   5580
      TabIndex        =   13
      Top             =   3840
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
      Left            =   2820
      TabIndex        =   15
      Top             =   4680
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
      Left            =   5640
      TabIndex        =   14
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label lbCep 
      Caption         =   "CEP"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   2700
      Width           =   1155
   End
   Begin VB.Label lbNome 
      Caption         =   "Endereço"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   1635
   End
   Begin VB.Label lbBairro 
      Caption         =   "Bairro"
      Height          =   315
      Index           =   2
      Left            =   2940
      TabIndex        =   7
      Top             =   4320
      Width           =   675
   End
   Begin VB.Label lbCidade 
      Caption         =   "Cidade"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   675
   End
   Begin VB.Label lbSexo 
      Caption         =   "Sexo"
      Height          =   315
      Index           =   0
      Left            =   8220
      TabIndex        =   4
      Top             =   1920
      Width           =   675
   End
   Begin VB.Label lbCpf 
      Caption         =   "CPF"
      Height          =   315
      Index           =   1
      Left            =   5880
      TabIndex        =   3
      Top             =   1920
      Width           =   675
   End
   Begin VB.Label lbNome 
      Caption         =   "Nome Completo"
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   1920
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   675
   End
End
Attribute VB_Name = "frmCadClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tbrCadClienteFrm_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.key = "CadCliente" Then
        sCadastrarCliente
    ElseIf Button.key = "CancCad" Then
        Unload Me
    End If
End Sub











