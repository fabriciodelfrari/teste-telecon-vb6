VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "CRUD"
   ClientHeight    =   7680
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12480
   LinkTopic       =   "mdlCRUD"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlListaImagens 
      Left            =   11700
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":0000
            Key             =   "RelClientes"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":04A6
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":0858
            Key             =   "CadClientes"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrBarraFerramentas 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   1588
      ButtonWidth     =   1561
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "imlListaImagens"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultas"
            Key             =   "Consultas"
            ImageKey        =   "RelClientes"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ConsultaCliente"
                  Text            =   "Clientes"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCadastros 
      Caption         =   "Cadastros"
      Visible         =   0   'False
      Begin VB.Menu mnuCadClientes 
         Caption         =   "Clientes"
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "Relatórios"
      Visible         =   0   'False
      Begin VB.Menu mnuRelProdutos 
         Caption         =   "Produtos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRelClientes 
         Caption         =   "Clientes"
      End
   End
End
Attribute VB_Name = "mdiPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tbrBarraFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.key = "Sair" Then
        End
    End If

End Sub

Private Sub tbrBarraFerramentas_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.key = "ConsultaCliente" Then
        frmConsultaClientes.Show
    End If


End Sub
