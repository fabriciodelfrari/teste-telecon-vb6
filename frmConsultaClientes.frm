VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ACTIVETEXT.OCX"
Begin VB.Form frmConsultaClientes 
   Caption         =   "Consulta de Clientes"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcura 
      Height          =   315
      Left            =   1980
      Picture         =   "frmConsultaClientes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   600
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
      Height          =   4695
      Left            =   60
      TabIndex        =   2
      Top             =   1200
      Width           =   12015
      Begin VB.CommandButton cmdBuscaEndereco 
         Caption         =   "Consultar CEP"
         Height          =   315
         Left            =   3360
         TabIndex        =   19
         Top             =   1560
         Width           =   1575
      End
      Begin rdActiveText.ActiveText txtCpf 
         Height          =   315
         Left            =   6060
         TabIndex        =   3
         Top             =   720
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
         Left            =   300
         TabIndex        =   4
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
         Left            =   8160
         TabIndex        =   5
         Top             =   1920
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         Left            =   5820
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   18
         Top             =   360
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
         Left            =   6120
         TabIndex        =   17
         Top             =   360
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
         TabIndex        =   16
         Top             =   2880
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
         TabIndex        =   15
         Top             =   2880
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
         TabIndex        =   14
         Top             =   2040
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
         TabIndex        =   13
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
         Left            =   5880
         TabIndex        =   12
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
         Left            =   8220
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
   End
   Begin rdActiveText.ActiveText txtCodigo 
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   600
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
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   1
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

Private Sub TabStrip1_Click()

End Sub

Private Sub Form_Load()
    cmdBuscaEndereco.Visible = False
End Sub

