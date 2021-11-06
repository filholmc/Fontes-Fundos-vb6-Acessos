VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm formMDIAce 
   BackColor       =   &H8000000C&
   Caption         =   "Acessos - Fundos"
   ClientHeight    =   4905
   ClientLeft      =   2775
   ClientTop       =   2040
   ClientWidth     =   8745
   Icon            =   "MDIAce.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fpicAtvHlp 
      Align           =   1  'Align Top
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   8745
      TabIndex        =   2
      Top             =   600
      Width           =   8745
   End
   Begin MSComctlLib.Toolbar ftbrModAce 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "fimlModAce"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Ajuda"
            Object.ToolTipText     =   "Ajuda - F1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Usuari"
            Object.ToolTipText     =   "Usuários - Ctrl+F1"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Fundos"
            Object.ToolTipText     =   "Fundos - Ctrl+F2"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Modulo"
            Object.ToolTipText     =   "Módulos - Ctrl+F3"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Formes"
            Object.ToolTipText     =   "Forms - Ctrl+F4"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Botoes"
            Object.ToolTipText     =   "Botões - Ctrl+F5"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "AceFun"
            Object.ToolTipText     =   "Acesso aos Fundos - Ctrl+F6"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "AceMod"
            Object.ToolTipText     =   "Acesso aos Módulos - Ctrl+F7"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "AceFor"
            Object.ToolTipText     =   "Acesso aos Forms - Ctrl+F8"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "AceBot"
            Object.ToolTipText     =   "Acesso aos Botões - Ctrl+F9"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Copiar"
            Object.ToolTipText     =   "Cópia de Acessos - Ctrl+F11"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "RelUsu"
            Object.ToolTipText     =   "Relação de Usuários - Ctrl+F12"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Fechar"
            Object.ToolTipText     =   "Fechar - Ctrl+F"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList fimlModAce 
      Left            =   90
      Top             =   1230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":0CCA
            Key             =   "Ajuda"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":19A4
            Key             =   "Usuari"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":227E
            Key             =   "Fundos"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":2B58
            Key             =   "Modulo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":3832
            Key             =   "Formes"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":410C
            Key             =   "Botoes"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":49E6
            Key             =   "AceFun"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":52C0
            Key             =   "AceMod"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":5F9A
            Key             =   "AceFor"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":6874
            Key             =   "AceBot"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":714E
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":7A28
            Key             =   "RelUsu"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":8302
            Key             =   "Fechar"
         EndProperty
      EndProperty
   End
   Begin VB.Timer ftmrDatHor 
      Interval        =   5000
      Left            =   90
      Top             =   720
   End
   Begin MSComctlLib.StatusBar fsbrModAce 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14287
            MinWidth        =   14287
            Object.ToolTipText     =   "Área reservada a Mensagens"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1102
            Object.ToolTipText     =   "Data do Servidor"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1429
            MinWidth        =   1411
            Object.ToolTipText     =   "Hora do Servidor"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10142
            MinWidth        =   10142
            Object.ToolTipText     =   "Usuário que está Realizando a Sessão"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList fimlStaBar 
      Left            =   90
      Top             =   1890
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIAce.frx":8BDC
            Key             =   "TemLog"
         EndProperty
      EndProperty
   End
   Begin VB.Menu menu 
      Caption         =   "&Menu"
      Begin VB.Menu menuUsuari 
         Caption         =   "&Usuários"
         Enabled         =   0   'False
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu menuDiv000 
         Caption         =   "-"
      End
      Begin VB.Menu menuFundos 
         Caption         =   "Fu&ndos"
         Enabled         =   0   'False
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu menuDiv001 
         Caption         =   "-"
      End
      Begin VB.Menu menuModulo 
         Caption         =   "&Módulos"
         Enabled         =   0   'False
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu menuFormes 
         Caption         =   "F&orms"
         Enabled         =   0   'False
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu menuBotoes 
         Caption         =   "&Botões"
         Enabled         =   0   'False
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu menuDiv002 
         Caption         =   "-"
      End
      Begin VB.Menu menuAcesso 
         Caption         =   "&Acessos aos"
         Begin VB.Menu menuAceFun 
            Caption         =   "Fu&ndos"
            Enabled         =   0   'False
            Shortcut        =   ^{F6}
         End
         Begin VB.Menu menuDiv003 
            Caption         =   "-"
         End
         Begin VB.Menu menuAceMod 
            Caption         =   "&Módulos"
            Enabled         =   0   'False
            Shortcut        =   ^{F7}
         End
         Begin VB.Menu menuAceFor 
            Caption         =   "F&orms"
            Enabled         =   0   'False
            Shortcut        =   ^{F8}
         End
         Begin VB.Menu menuAceBot 
            Caption         =   "&Botões"
            Enabled         =   0   'False
            Shortcut        =   ^{F9}
         End
      End
      Begin VB.Menu menuDiv004 
         Caption         =   "-"
      End
      Begin VB.Menu menuCopiar 
         Caption         =   "Cópia &de Acessos"
         Enabled         =   0   'False
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu menuDiv005 
         Caption         =   "-"
      End
      Begin VB.Menu menuConsul 
         Caption         =   "&Consultas"
         Begin VB.Menu menuConUsu 
            Caption         =   "&Usuários"
            Enabled         =   0   'False
            Shortcut        =   ^{F12}
         End
         Begin VB.Menu menuDiv006 
            Caption         =   "-"
         End
         Begin VB.Menu menuConFun 
            Caption         =   "Fu&ndos"
            Enabled         =   0   'False
         End
         Begin VB.Menu menuDiv007 
            Caption         =   "-"
         End
         Begin VB.Menu menuConMod 
            Caption         =   "&Módulos"
            Enabled         =   0   'False
         End
         Begin VB.Menu menuConFor 
            Caption         =   "F&orms de um Módulo"
            Enabled         =   0   'False
         End
         Begin VB.Menu menuConBot 
            Caption         =   "&Botões de um Form"
            Enabled         =   0   'False
         End
         Begin VB.Menu menuDiv008 
            Caption         =   "-"
         End
         Begin VB.Menu menuConLog 
            Caption         =   "&Log de um"
            Begin VB.Menu menuConLgF 
               Caption         =   "F&orm"
            End
            Begin VB.Menu menuDiv009 
               Caption         =   "-"
            End
            Begin VB.Menu menuConLgU 
               Caption         =   "&Usuário"
            End
         End
         Begin VB.Menu menuDiv010 
            Caption         =   "-"
         End
         Begin VB.Menu menuConAce 
            Caption         =   "&Acessos de um Usuário a"
            Begin VB.Menu menuConAFu 
               Caption         =   "Fu&ndos"
               Enabled         =   0   'False
            End
            Begin VB.Menu menuDiv011 
               Caption         =   "-"
            End
            Begin VB.Menu menuConAMo 
               Caption         =   "&Módulos"
               Enabled         =   0   'False
            End
            Begin VB.Menu menuConAFo 
               Caption         =   "F&orms de um Módulo"
               Enabled         =   0   'False
            End
            Begin VB.Menu menuConABo 
               Caption         =   "&Botões de um Form"
               Enabled         =   0   'False
            End
         End
         Begin VB.Menu menuDiv012 
            Caption         =   "-"
         End
         Begin VB.Menu menuQueAce 
            Caption         =   "U&suários que acessam um"
            Begin VB.Menu menuConFuA 
               Caption         =   "Fu&ndo"
               Enabled         =   0   'False
            End
            Begin VB.Menu menuDiv013 
               Caption         =   "-"
            End
            Begin VB.Menu menuConMoA 
               Caption         =   "&Módulo"
               Enabled         =   0   'False
            End
            Begin VB.Menu menuConFoA 
               Caption         =   "F&orm"
               Enabled         =   0   'False
            End
            Begin VB.Menu menuConBoA 
               Caption         =   "&Botão"
               Enabled         =   0   'False
            End
         End
      End
      Begin VB.Menu menuDiv014 
         Caption         =   "-"
      End
      Begin VB.Menu menuFechar 
         Caption         =   "&Fechar"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu menuModGes 
      Caption         =   "&Gestor"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu menuModPas 
      Caption         =   "&Passivo"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu menuModAtv 
      Caption         =   "&Ativo"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu menuModCtb 
      Caption         =   "&Contabilidade"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu menuModGer 
      Caption         =   "G&erencial"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu menuModCad 
      Caption         =   "Ca&dastros"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu menuFerram 
      Caption         =   "&Ferramentas"
      Begin VB.Menu menuSenhas 
         Caption         =   "Troca de &Senha"
      End
      Begin VB.Menu menuTroUsu 
         Caption         =   "Troca de &Usuário"
         Shortcut        =   ^U
      End
      Begin VB.Menu menuDiv015 
         Caption         =   "-"
      End
      Begin VB.Menu menuRemAce 
         Caption         =   "Remoção de &Acessos"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuDiv016 
         Caption         =   "-"
      End
      Begin VB.Menu menuSenhaR 
         Caption         =   "&Restauração de Senha"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuDiv017 
         Caption         =   "-"
      End
      Begin VB.Menu menuTempor 
         Caption         =   "Temporização da &Proteção"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu menuAjudas 
      Caption         =   "Aj&uda"
      Begin VB.Menu menuContat 
         Caption         =   "C&ontatos"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuDiv018 
         Caption         =   "-"
      End
      Begin VB.Menu menuContdo 
         Caption         =   "&Conteúdo                   F1"
         Enabled         =   0   'False
      End
      Begin VB.Menu menuDiv019 
         Caption         =   "-"
      End
      Begin VB.Menu menuModAce 
         Caption         =   "&Sobre o Módulo Acessos"
      End
   End
End
Attribute VB_Name = "formMDIAce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbooMaximo As Boolean

Private fclsParGer As clssParGer

Private fclsAceFor As clssAceFor

Private fstrAppCtt As String

Private fstrAppCad As String, fstrAppGer As String, fstrAppGes As String

Private fstrAppPas As String, fstrAppAtv As String, fstrAppCtb As String
Private Sub fsbrModAce_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub ftbrModAce_ButtonClick(ByVal Button As MSComctlLib.Button)
        gintTempoP = 0
        Select Case Button.Key
               Case "Ajuda"
                    SendKeys "{F1}"
               Case "Usuari"
                    formUsuari.SetFocus
               Case "Fundos"
                    formFundos.SetFocus
               Case "Modulo"
                    formModulo.SetFocus
               Case "Formes"
                    formFormes.SetFocus
               Case "Botoes"
                    formBotoes.SetFocus
               Case "AceFun"
                    formAceFun.SetFocus
               Case "AceMod"
                    formAceMod.SetFocus
               Case "AceFor"
                    formAceFor.SetFocus
               Case "AceBot"
                    formAceBot.SetFocus
               Case "Copiar"
                    formCopiar.SetFocus
               Case "RelUsu"
                    formConUsu.SetFocus
               Case "Fechar"
                    rlsEncerrar
        End Select
        Button.Value = tbrPressed
End Sub
Private Sub ftbrModAce_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub ftmrDatHor_Timer()
            gdatServBD = rgfDataDoServidor

            fsbrModAce.Panels(2) = Format(gdatServBD, "dd/mm")
            fsbrModAce.Panels(3) = Format(gdatServBD, "hh:mm:ss")

            gintTempoP = gintTempoP + 5

        If (gbooProAce) Then
        If (fbooMaximo) Then
        If (gintTempoP > gintTempoE) Then
            formProtge.Show 1
        End If
        End If
        End If
End Sub
Private Sub MDIForm_Activate()
        If (gbooCancel) Then
            rlsEncerrar
            Exit Sub
        End If

            rlsHabilitarModulos
            rlsHabilitarForms
            gintTempoP = 0

        If (gstrSenhas = "+") Then formSenhas.Show 1
End Sub
Private Sub MDIForm_Load()
        Select Case gintCodAdm
               Case 4
                    formMDIAce.BackColor = &HFFFFC0
               Case 7
                    formMDIAce.BackColor = &HC0FFFF
               Case Else
                    formMDIAce.BackColor = &HC0FFC0
        End Select

        Set fclsParGer = New clssParGer
        Set gclsModulo = New clssModulo
        Set gclsFundos = New clssFundos
        Set fclsAceFor = New clssAceFor
        Set gclsAceBot = New clssAceBot

            gbytConMod = 0
            gbytConFun = 0
            gbytConFor = 0
            gbytConBot = 0
            fbooMaximo = False
            gbooAjuHab = False
            gintTotRes = Screen.Width / Screen.TwipsPerPixelX + Screen.Height / Screen.TwipsPerPixelY

            fsbrModAce.Panels(1).MinWidth = IIf(gintTotRes = 1400, 8099.7168, 9500.0323)

            ftmrDatHor_Timer
            rlsConsultarTempoDeEspera
            rgsPosicionarAjuda Me, gintForAtu, gbooForLog

            formMDIAce.Caption = gstrNomApl

            fsbrModAce.Panels(1) = "SQL: conectado ao Servidor '" & gstrServBD & "', Banco '" & gstrNomeBD & "'"
            fsbrModAce.Panels(4) = gstrNomUsu
        If (gbooUsuLog) Then _
            fsbrModAce.Panels(4).Picture = formMDIAce.fimlStaBar.ListImages(1).Picture

            fstrAppGes = gstrPthExe & "Gestor.exe"

        If (rgfArquivoExiste(fstrAppGes)) Then menuModGes.Visible = True

            fstrAppPas = gstrPthExe & "Passivo.exe"

        If (rgfArquivoExiste(fstrAppPas)) Then menuModPas.Visible = True

            fstrAppAtv = gstrPthExe & "Ativo.exe"

        If (rgfArquivoExiste(fstrAppAtv)) Then menuModAtv.Visible = True

            fstrAppCtb = gstrPthExe & "Contabil.exe"

        If (rgfArquivoExiste(fstrAppCtb)) Then menuModCtb.Visible = True

            fstrAppGer = gstrPthExe & "Gerencial.exe"

        If (rgfArquivoExiste(fstrAppGer)) Then menuModGer.Visible = True

            fstrAppCad = gstrPthExe & "Cadastros.exe"

        If (rgfArquivoExiste(fstrAppCad)) Then menuModCad.Visible = True

            fstrAppCad = fstrAppCad & " /log"

            fstrAppCtt = gstrPthExe & "Contatos.exe"

        If (rgfArquivoExiste(fstrAppCtt)) Then menuContat.Enabled = True

        If (rgfArquivoExiste(gstrPthAju)) Then
                                               gbooAjuHab = True
                                               menuContdo.Enabled = True
                                               ftbrModAce.Buttons("Ajuda").Enabled = True
        End If
End Sub
Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
        ftbrModAce.Buttons("Ajuda").Value = tbrUnpressed
End Sub
Private Sub MDIForm_Resize()
            fbooMaximo = Not fbooMaximo

        If (gbooProAce) Then
        If (fbooMaximo) Then
        If (gintTempoP > gintTempoE) Then
            formProtge.Show 1
        End If
        End If
        End If
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
        rlsAjustarLogado

        Set fclsParGer = Nothing
        Set gclsModulo = Nothing
        Set gclsFundos = Nothing
        Set gclsUsuari = Nothing
        Set gclsFormes = Nothing
        Set gclsAceMod = Nothing
        Set fclsAceFor = Nothing
        Set gclsAceBot = Nothing
        Set gclsLogado = Nothing
        Set gclsDiario = Nothing
        Set gclsInsAdm = Nothing
        Set gDBCFundos = Nothing
End Sub
Private Sub menuAceBot_Click()
        formAceBot.SetFocus
End Sub
Private Sub menuAceFor_Click()
        formAceFor.SetFocus
End Sub
Private Sub menuAceFun_Click()
        formAceFun.SetFocus
End Sub
Private Sub menuAceMod_Click()
        formAceMod.SetFocus
End Sub
Private Sub menuBotoes_Click()
        formBotoes.SetFocus
End Sub
Private Sub menuConABo_Click()
        formConABo.SetFocus
End Sub
Private Sub menuConAFo_Click()
        formConAFo.SetFocus
End Sub
Private Sub menuConAFu_Click()
        formConAFu.SetFocus
End Sub
Private Sub menuConAMo_Click()
        formConAMo.SetFocus
End Sub
Private Sub menuConBoA_Click()
        formConBoA.SetFocus
End Sub
Private Sub menuConBot_Click()
        formConBot.SetFocus
End Sub
Private Sub menuConFoA_Click()
        formConFoA.SetFocus
End Sub
Private Sub menuConFor_Click()
        formConFor.SetFocus
End Sub
Private Sub menuConFuA_Click()
        formConFuA.SetFocus
End Sub
Private Sub menuConFun_Click()
        formConFun.SetFocus
End Sub
Private Sub menuConLgF_Click()
        formConLgF.SetFocus
End Sub
Private Sub menuConLgU_Click()
        formConLgU.SetFocus
End Sub
Private Sub menuConMoA_Click()
        formConMoA.SetFocus
End Sub
Private Sub menuConMod_Click()
        formConMod.SetFocus
End Sub
Private Sub menuCopiar_Click()
        formCopiar.SetFocus
End Sub
Private Sub menuContat_Click()
        Shell fstrAppCtt, 1
End Sub
Private Sub menuContdo_Click()
        SendKeys "{F1}"
End Sub
Private Sub menuConUsu_Click()
        formConUsu.SetFocus
End Sub
Private Sub menuFechar_Click()
        rlsEncerrar
End Sub
Private Sub menuFormes_Click()
        formFormes.SetFocus
End Sub
Private Sub menuFundos_Click()
        formFundos.SetFocus
End Sub
Private Sub menuModAce_Click()
        formVersao.SetFocus
End Sub
Private Sub menuModAtv_Click()
        rlsChecarLogado
        Shell fstrAppAtv, 1
End Sub
Private Sub menuModCad_Click()
        rlsChecarLogado
        Shell fstrAppCad, 1
End Sub
Private Sub menuModCtb_Click()
        rlsChecarLogado
        Shell fstrAppCtb, 1
End Sub
Private Sub menuModGer_Click()
        rlsChecarLogado
        Shell fstrAppGer, 1
End Sub
Private Sub menuModGes_Click()
        rlsChecarLogado
        Shell fstrAppGes, 1
End Sub
Private Sub menuModPas_Click()
        rlsChecarLogado
        Shell fstrAppPas, 1
End Sub
Private Sub menuModulo_Click()
        formModulo.SetFocus
End Sub
Private Sub menuRemAce_Click()
        formRemAce.SetFocus
End Sub
Private Sub menuSenhaR_Click()
        formSenhaR.SetFocus
End Sub
Private Sub menuSenhas_Click()
        formSenhas.Show 1
End Sub
Private Sub menuTempor_Click()
        formTempor.SetFocus
End Sub
Private Sub menuTroUsu_Click()
        rlsEncerrarFormes
        formTroUsu.Show 1
End Sub
Private Sub menuUsuari_Click()
        formUsuari.SetFocus
End Sub
Private Sub rlsAjustarLogado()
            gbytModAtv = gbytModAtv - 1

        If (gbytModAtv > 0) Then
            gclsLogado.Alterar gstrNomCmp, gintUsuLog, gbytModAtv
        Else
            gclsLogado.Excluir gstrNomCmp
        End If

        If (gDBCFundos.Errors.Count > 0) Then rgsTratarErro Err, gDBCFundos.Errors, Me
End Sub
Private Sub rlsChecarLogado()
        Dim lRStLogado As Recordset

        Set lRStLogado = gclsLogado.Consultar(gstrNomCmp)

        If (lRStLogado.EOF) Then
            gbytModAtv = 1
            gclsLogado.Incluir gstrNomCmp, gintUsuLog, gintAgeLog, " ", "Acessos", gbytModAtv

        If (gDBCFundos.Errors.Count > 0) Then rgsTratarErro Err, gDBCFundos.Errors, Me
        End If
        lRStLogado.Close
End Sub
Private Sub rlsConsultarTempoDeEspera()
        Dim lRStParGer As Recordset

        Set lRStParGer = fclsParGer.Consultar("TempoDeEspera")

            gintTempoE = CInt(lRStParGer!Cteudo) * 60

        Set lRStParGer = fclsParGer.Consultar("ProtegeAcesso")

            gbooProAce = IIf(lRStParGer!Cteudo = "Sim", True, False)

        lRStParGer.Close
End Sub
Private Sub rlsEncerrar()
        Unload Me
End Sub
Private Sub rlsEncerrarFormes()
        Dim lobjFormes As Form

        For Each lobjFormes In Forms
        If (Not TypeOf lobjFormes Is MDIForm) Then
            Unload lobjFormes
               Set lobjFormes = Nothing
        End If
        Next
End Sub
Private Sub rlsHabilitarForms()
        Dim lRStAceFor As Recordset

        Set lRStAceFor = fclsAceFor.ConsultarFormsDeUmUsuarioPorModulo(gintUsuLog, 1)

            ftbrModAce.Buttons("Usuari").Enabled = False
            ftbrModAce.Buttons("Fundos").Enabled = False
            ftbrModAce.Buttons("Modulo").Enabled = False
            ftbrModAce.Buttons("Formes").Enabled = False
            ftbrModAce.Buttons("Botoes").Enabled = False
            ftbrModAce.Buttons("AceFun").Enabled = False
            ftbrModAce.Buttons("AceMod").Enabled = False
            ftbrModAce.Buttons("AceFor").Enabled = False
            ftbrModAce.Buttons("AceBot").Enabled = False
            ftbrModAce.Buttons("Copiar").Enabled = False
            ftbrModAce.Buttons("RelUsu").Enabled = False

            menuUsuari.Enabled = False
            menuFundos.Enabled = False
            menuModulo.Enabled = False
            menuFormes.Enabled = False
            menuBotoes.Enabled = False
            menuAceFun.Enabled = False
            menuAceMod.Enabled = False
            menuAceFor.Enabled = False
            menuAceBot.Enabled = False
            menuCopiar.Enabled = False
            menuConUsu.Enabled = False
            menuConFun.Enabled = False
            menuConMod.Enabled = False
            menuConFor.Enabled = False
            menuConBot.Enabled = False
            menuConLgF.Enabled = False
            menuConLgU.Enabled = False
            menuConAFu.Enabled = False
            menuConAMo.Enabled = False
            menuConAFo.Enabled = False
            menuConABo.Enabled = False
            menuConFuA.Enabled = False
            menuConMoA.Enabled = False
            menuConFoA.Enabled = False
            menuConBoA.Enabled = False
            menuRemAce.Enabled = False
            menuSenhaR.Enabled = False
            menuTempor.Enabled = False
        Do _
            While (Not ((lRStAceFor.EOF)))
            Select Case (lRStAceFor!NomFor)
                   Case "formUsuari"
                         menuUsuari.Enabled = True
                         ftbrModAce.Buttons _
                         ("Usuari").Enabled = True
                   Case "formFundos"
                         menuFundos.Enabled = True
                         ftbrModAce.Buttons _
                         ("Fundos").Enabled = True
                   Case "formModulo"
                         menuModulo.Enabled = True
                         ftbrModAce.Buttons _
                         ("Modulo").Enabled = True
                   Case "formFormes"
                         menuFormes.Enabled = True
                         ftbrModAce.Buttons _
                         ("Formes").Enabled = True
                   Case "formBotoes"
                         menuBotoes.Enabled = True
                         ftbrModAce.Buttons _
                         ("Botoes").Enabled = True
                   Case "formAceFun"
                         menuAceFun.Enabled = True
                         ftbrModAce.Buttons _
                         ("AceFun").Enabled = True
                   Case "formAceMod"
                         menuAceMod.Enabled = True
                         ftbrModAce.Buttons _
                         ("AceMod").Enabled = True
                   Case "formAceFor"
                         menuAceFor.Enabled = True
                         ftbrModAce.Buttons _
                         ("AceFor").Enabled = True
                   Case "formAceBot"
                         menuAceBot.Enabled = True
                         ftbrModAce.Buttons _
                         ("AceBot").Enabled = True
                   Case "formCopiar"
                         menuCopiar.Enabled = True
                         ftbrModAce.Buttons _
                         ("Copiar").Enabled = True

                   Case "formConUsu"
                         menuConUsu.Enabled = True
                         ftbrModAce.Buttons _
                         ("RelUsu").Enabled = True
                   Case "formConFun"
                         menuConFun.Enabled = True
                   Case "formConMod"
                         menuConMod.Enabled = True
                   Case "formConFor"
                         menuConFor.Enabled = True
                   Case "formConBot"
                         menuConBot.Enabled = True
                   Case "formConLgF"
                         menuConLgF.Enabled = True
                   Case "formConLgU"
                         menuConLgU.Enabled = True
                   Case "formConAFu"
                         menuConAFu.Enabled = True
                   Case "formConAMo"
                         menuConAMo.Enabled = True
                   Case "formConAFo"
                         menuConAFo.Enabled = True
                   Case "formConABo"
                         menuConABo.Enabled = True
                   Case "formConFuA"
                         menuConFuA.Enabled = True
                   Case "formConMoA"
                         menuConMoA.Enabled = True
                   Case "formConFoA"
                         menuConFoA.Enabled = True
                   Case "formConBoA"
                         menuConBoA.Enabled = True

                   Case "formRemAce"
                         menuRemAce.Enabled = True
                   Case "formSenhaR"
                         menuSenhaR.Enabled = True
                   Case "formTempor"
                         menuTempor.Enabled = True
            End Select
            lRStAceFor.MoveNext
        Loop
        lRStAceFor.Close
End Sub
Private Sub rlsHabilitarModulos()
        Dim lRStAceMod As Recordset

        Set lRStAceMod = gclsAceMod.ConsultarModulosDeUmUsuario(gintUsuLog)

            menuModCad.Enabled = False
            menuModPas.Enabled = False
            menuModAtv.Enabled = False
            menuModCtb.Enabled = False
            menuModGer.Enabled = False
            menuModGes.Enabled = False
        Do _
            While (Not (lRStAceMod.EOF))
            Select Case lRStAceMod!Numero
                   Case 2
                        menuModCad.Enabled = True
                   Case 4
                        menuModPas.Enabled = True
                   Case 5
                        menuModAtv.Enabled = True
                   Case 6
                        menuModCtb.Enabled = True
                   Case 7
                        menuModGer.Enabled = True
                   Case 9
                        menuModGes.Enabled = True
            End Select
            lRStAceMod.MoveNext
        Loop
        lRStAceMod.Close
End Sub
