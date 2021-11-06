VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formConLgU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log de um Usuário"
   ClientHeight    =   5190
   ClientLeft      =   2790
   ClientTop       =   1755
   ClientWidth     =   8910
   Icon            =   "ConLgU.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   8910
   Begin VB.Frame frmeFormes 
      Caption         =   "  Parâmetros: "
      Height          =   1425
      Left            =   120
      TabIndex        =   7
      Top             =   30
      Width           =   8700
      Begin VB.Frame Frame1 
         Caption         =   "Usuário:"
         Height          =   915
         Left            =   90
         TabIndex        =   0
         Top             =   240
         Width           =   3555
         Begin VB.ComboBox fcboUsuari 
            Height          =   315
            ItemData        =   "ConLgU.frx":08CA
            Left            =   120
            List            =   "ConLgU.frx":08CC
            TabIndex        =   1
            Top             =   450
            Width           =   3300
         End
      End
      Begin VB.CommandButton fcmbF09LCA 
         Caption         =   "F9"
         Height          =   255
         Left            =   7230
         TabIndex        =   12
         Top             =   420
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF11Tab 
         Caption         =   "F11"
         Height          =   255
         Left            =   7680
         TabIndex        =   11
         Top             =   420
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF12Hom 
         Caption         =   "F12"
         Height          =   255
         Left            =   8130
         TabIndex        =   10
         Top             =   420
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton fcmbF05Con 
         Caption         =   "Consultar (F5)"
         Default         =   -1  'True
         Height          =   315
         Left            =   6300
         TabIndex        =   5
         Top             =   690
         Width           =   1120
      End
      Begin VB.CommandButton fcmbEscape 
         Caption         =   "Fechar (Esc)"
         Height          =   315
         Left            =   7440
         TabIndex        =   6
         Top             =   690
         Width           =   1120
      End
      Begin VB.Frame fmePeriod 
         Caption         =   "Período:"
         Height          =   915
         Left            =   3750
         TabIndex        =   2
         Top             =   240
         Width           =   2445
         Begin MSMask.MaskEdBox fmskDatIni 
            Height          =   315
            Left            =   130
            TabIndex        =   3
            Top             =   450
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox fmskDatFim 
            Height          =   315
            Left            =   1260
            TabIndex        =   4
            Top             =   450
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label lblDatFim 
            AutoSize        =   -1  'True
            Caption         =   "Data Final"
            Height          =   195
            Left            =   1260
            TabIndex        =   9
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblDatIni 
            AutoSize        =   -1  'True
            Caption         =   "Data Inicial"
            Height          =   195
            Left            =   135
            TabIndex        =   8
            Top             =   240
            Width           =   795
         End
      End
   End
   Begin MSComctlLib.ListView flsvLogUsu 
      Height          =   3525
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6218
      View            =   3
      Arrange         =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "formConLgU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fbytColInd As Byte

Private fbytColOrd As Byte

Private fbytOrdChv As Byte

Private fdatDatFim As Date

Private fdatDatIni As Date

Private fintNumUsu As Integer

Private fstrSmbOrd As String
Private Function rlfNomeFor(ByVal vintNumFor As Integer)
        Dim lRStFormes As Recordset

        Set lRStFormes = gclsFormes.Consultar(vintNumFor)

            rlfNomeFor = lRStFormes!Descri

        lRStFormes.Close
End Function
Private Function rlfNomeFun(ByVal vbytNumFun As Byte)
        Dim lRStFundos As Recordset

        Set lRStFundos = gclsFundos.Consultar(vbytNumFun)

        If (lRStFundos.EOF) Then
            rlfNomeFun = " "
        Else
            rlfNomeFun = lRStFundos!Codigo
        End If

        lRStFundos.Close
End Function
Private Function rlfNomeMod(ByVal vbytNumMod As Byte)
        Dim lRStModulo As Recordset

        Set lRStModulo = gclsModulo.Consultar(vbytNumMod)

            rlfNomeMod = lRStModulo!Descri

        lRStModulo.Close
End Function
Private Function rlfNomeUsu(ByVal vintNumUsu As Integer)
        Dim lRStUsuari As Recordset

        Set lRStUsuari = gclsUsuari.Consultar(vintNumUsu)

            rlfNomeUsu = lRStUsuari!NomUsu

        lRStUsuari.Close
End Function
Private Sub fcboUsuari_Click()
        If (fcboUsuari.ListIndex = -1) Then Exit Sub

        fintNumUsu = CInt(Mid(fcboUsuari, 1, InStr(fcboUsuari, " ") - 1))

        rlsMontarPeriodo
End Sub
Private Sub fcmbEscape_Click()
        Unload Me
End Sub
Private Sub fcmbF05Con_Click()
        If (fcboUsuari.ListIndex = -1) Then
            rgfMsgBox "Escolha uma opção do campo 'Usuário'", MsgErr
            fcboUsuari.SetFocus
            Exit Sub
        End If

        If (fmskDatIni = "") Then
        If (IsDate(Format(fmskDatFim, "00/00/0000"))) Then
            fmskDatIni = Format(fdatDatFim, "dd/mm/yyyy")
        Else
            rgfMsgBox "Corrija o campo 'Data Inicial'", MsgErr, Me.HelpContextID
            fmskDatIni.SetFocus
            Exit Sub
        End If
        Else
        If (Not IsDate(Format(fmskDatIni, "00/00/0000"))) Then
            rgfMsgBox "Corrija o campo 'Data Inicial'", MsgErr, Me.HelpContextID
            fmskDatIni.SetFocus
            Exit Sub
        End If
        End If

            gdatDatIni = Format(fmskDatIni, "00/00/0000")

        If (Not IsDate(Format(fmskDatFim, "00/00/0000"))) Then
            rgfMsgBox "Corrija o campo 'Data Final'", MsgErr, Me.HelpContextID
            fmskDatFim.SetFocus
            Exit Sub
        End If

            gdatDatFim = Format(fmskDatFim, "00/00/0000")

        If (gdatDatIni < fdatDatIni) Then
            rgfMsgBox "Data Inicial menor que a disponível na Base de Dados", MsgErr
            fmskDatIni.SetFocus
            Exit Sub
        End If

        If (gdatDatIni > gdatDatFim) Then
            rgfMsgBox "A Data Inicial deve ser menor ou igual à Data Final", MsgErr
            fmskDatIni.SetFocus
            Exit Sub
        End If

        If (DateDiff("d", gdatDatIni, gdatDatFim) > 6) Then
            rgfMsgBox "O período entre as duas datas deve ser de até 6 dias corridos", MsgErr
            fmskDatIni.SetFocus
            Exit Sub
        End If

        rlsConsultar
End Sub
Private Sub fcmbF09LCA_Click()
        If (TypeOf ActiveControl Is ComboBox Or TypeOf ActiveControl Is MaskEdBox) Then ActiveControl.Text = ""
End Sub
Private Sub fcmbF11Tab_Click()
        SendKeys "+{TAB}"
End Sub
Private Sub fcmbF12Hom_Click()
        fcboUsuari.SetFocus
End Sub
Private Sub flsvLogUsu_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        gintTempoP = 0
        fbytColInd = ColumnHeader.Index
        fstrSmbOrd = IIf(fbytOrdChv = 1, "/\", "\/")
        fbytOrdChv = IIf(fbytOrdChv = 1, 0, 1)
        rlsMontarColunasNomes
End Sub
Private Sub flsvLogUsu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        gintTempoP = 0
End Sub
Private Sub Form_Activate()
        gintTempoP = 0
        formMDIAce.fsbrModAce.Panels(4).Picture = LoadPicture()
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        gintTempoP = 0
        rgsTratarFuncoes KeyCode, Me
End Sub
Private Sub Form_Load()
        fbytColInd = 1
        fbytOrdChv = 0
        fstrSmbOrd = "/\"

        Me.Width = (Screen.Width) - 120
        Me.Height = Screen.Height - 360 - 1715

        With frmeFormes
            .Top = 90
            .Left = 110
            .Width = Me.Width - 290
            .Height = 1425
        End With

        With flsvLogUsu
            .Top = 1590
            .Left = 90
            .Width = (Me.Width) - 270
            .Height = Me.Height - 540 - 1510
        End With

        rgsCentralizarForm Me
        rgsPosicionarAjuda Me, gintForAtu, gbooForLog

        rlsCarregarUsuarios
End Sub
Private Sub rlsCarregarUsuarios()
        Dim lRStDiario As Recordset

        Set lRStDiario = gclsDiario.ConsultarUsuarios

            fcboUsuari.Clear
        Do _
            While (Not lRStDiario.EOF)
            fcboUsuari.AddItem lRStDiario!NumUsu & " - " & rlfNomeUsu(lRStDiario!NumUsu)
            lRStDiario.MoveNext
        Loop
        lRStDiario.Close
End Sub
Private Sub rlsConsultar()
        Dim lLsILinhas As MSComctlLib.ListItem

        Dim lRStDiario As Recordset

        Set lRStDiario = gclsDiario.ConsultarDeUmUsuario(fintNumUsu, gdatDatIni, gdatDatFim)

            flsvLogUsu.ListItems.Clear

        If (lRStDiario.EOF) Then
            rgfMsgBox "Não há Log deste Usuário no Período", MsgInf
            lRStDiario.Close
            Exit Sub
        End If

        If (gintTotRes < 1792) Then
            gbytScrBar = IIf(lRStDiario.RecordCount < 24, 0, 240)
        Else
            gbytScrBar = IIf(lRStDiario.RecordCount < 34, 0, 240)
        End If

            rlsMontarColunas
        Do _
            While (Not ((lRStDiario.EOF)))
        Set lLsILinhas = flsvLogUsu.ListItems.Add(, , Format(lRStDiario!DatBas, "dd/mm/yyyy hh:mm:ss"))
       With lLsILinhas
           .SubItems(1) = (Format(lRStDiario!DatBas, "yyyymmddhhmmss"))
           .SubItems(2) = lRStDiario!NomCmp
           .SubItems(3) = rlfNomeMod(lRStDiario!NumMod)
           .SubItems(4) = rlfNomeFor(lRStDiario!NumFor)
           .SubItems(5) = lRStDiario!Funcao
           .SubItems(6) = rlfNomeFun(lRStDiario!NumFun)
           .SubItems(7) = lRStDiario!Chaves
           .SubItems(8) = lRStDiario!Cteudo
        End With
            lRStDiario.MoveNext
        Loop
        flsvLogUsu.SetFocus
        lRStDiario.Close
End Sub
Private Sub rlsMontarColunas()
       With flsvLogUsu.ColumnHeaders
           .Clear
           .Add 1, , "", 1800, 0
           .Add 2, , "", 0
           .Add 3, , "", 1700, 0
           .Add 4, , "", 1500, 0
           .Add 5, , "", 2400, 0
           .Add 6, , "", 1150, 0
           .Add 7, , "", 1150, 0
           .Add 8, , "", 1700, 0
           .Add 9, , "", IIf(gintTotRes < 1792, 5120, 3510) - gbytScrBar, 0
        End With
        rlsMontarColunasNomes
End Sub
Private Sub rlsMontarColunasNomes()
            fbytColOrd = fbytColInd - 1

        If (fbytColOrd = 0) Then
            fbytColOrd = 1
        End If

       With flsvLogUsu
           .SortKey = ((fbytColOrd))
           .SortOrder = fbytOrdChv

            With .ColumnHeaders
                 .Item(1).Text = "Data"
                 .Item(3).Text = "Computador"
                 .Item(4).Text = "Módulo"
                 .Item(5).Text = "Form"
                 .Item(6).Text = "Ação"
                 .Item(7).Text = "Fundo"
                 .Item(8).Text = "Id"
                 .Item(9).Text = "Conteúdo"

                  Select Case fbytColInd
                         Case 1
                             .Item(1).Text = "Data  " & fstrSmbOrd
                         Case 3
                             .Item(3).Text = "Computador  " & fstrSmbOrd
                         Case 4
                             .Item(4).Text = "Módulo  " & fstrSmbOrd
                         Case 5
                             .Item(5).Text = "Form  " & fstrSmbOrd
                         Case 6
                             .Item(6).Text = "Ação  " & fstrSmbOrd
                         Case 7
                             .Item(7).Text = "Fundo  " & fstrSmbOrd
                         Case 8
                             .Item(8).Text = "Id  " & fstrSmbOrd
                         Case 9
                             .Item(9).Text = "Conteúdo  " & fstrSmbOrd
                  End Select
            End With
        End With
End Sub
Private Sub rlsMontarPeriodo()
        Dim lRStDiario As Recordset

        Set lRStDiario = gclsDiario.ConsultarDataUsuario(fintNumUsu, "ASC")

            fdatDatIni = Format(lRStDiario!DatBas, "dd/mm/yyyy")
            fmskDatIni = Format(lRStDiario!DatBas, "dd/mm/yyyy")

        Set lRStDiario = gclsDiario.ConsultarDataUsuario(fintNumUsu, "DESC")

            fdatDatFim = Format(DateAdd("d", -6, DateAdd("d", 1, lRStDiario!DatBas)), "dd/mm/yyyy")
            fdatDatFim = IIf(fdatDatIni > fdatDatFim, fdatDatIni, fdatDatFim)
            fmskDatFim = Format(DateAdd("d", 1, lRStDiario!DatBas), "dd/mm/yyyy")

        lRStDiario.Close
End Sub
