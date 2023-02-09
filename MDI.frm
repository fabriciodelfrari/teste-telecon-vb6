VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7380
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12480
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbBarraStatus 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   6945
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Text            =   "Operador"
            TextSave        =   "Operador"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   6068
            MinWidth        =   6068
            Text            =   "Nome Micro"
            TextSave        =   "Nome Micro"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   9322
            MinWidth        =   6068
            Text            =   "Caminho do Banco de Dados"
            TextSave        =   "Caminho do Banco de Dados"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
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
      ButtonWidth     =   2011
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "imlListaImagens"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cad. Clientes"
            Key             =   "CadClientes"
            ImageKey        =   "CadClientes"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Relatórios"
            Key             =   "Relatorios"
            ImageKey        =   "RelClientes"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RelClientes"
                  Text            =   "Relatório de Clientes"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RelProdutos"
                  Text            =   "Relatório de Produtos"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Begin VB.Menu mnuCadClientes 
         Caption         =   "Clientes"
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "Relatórios"
      Begin VB.Menu mnuRelProdutos 
         Caption         =   "Produtos"
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

Private Sub mnuCadClientes_Click()
    frmCadClientes.Show
End Sub

Private Sub mnuRelClientes_Click()
    MsgBox "Você acessou o relatório de clientes pelo menu", vbInformation
End Sub

Private Sub mnuRelProdutos_Click()
    MsgBox "Você acessou o relatório de produtos pelo menu", vbInformation
End Sub

Private Sub tbrBarraFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.key = "CadClientes" Then
        frmCadClientes.Show

    ElseIf Button.key = "Sair" Then
        End
    End If

End Sub

Private Sub tbrBarraFerramentas_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.key = "RelClientes" Then
        MsgBox "Relatorio de clientes", vbInformation
    ElseIf ButtonMenu.key = "RelProdutos" Then
        MsgBox "Relatorio de PRODUTOS", vbInformation
    End If


End Sub
